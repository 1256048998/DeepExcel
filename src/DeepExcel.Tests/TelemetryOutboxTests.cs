using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Account;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 遥测发件箱：事件先落盘，服务端收下才删；进程崩溃、加载失败时记下的事件下次启动补发；
    /// 多个 Excel 进程共用一个发件箱时不重复发送。
    /// </summary>
    public class TelemetryOutboxTests : IDisposable
    {
        private readonly string _root;
        private readonly string _outbox;

        public TelemetryOutboxTests()
        {
            _root = Path.Combine(Path.GetTempPath(), "DeepExcelOutboxTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            _outbox = Path.Combine(_root, "telemetry", "pending.jsonl");
        }

        public void Dispose()
        {
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        private sealed class Server : HttpMessageHandler
        {
            public bool Fail;
            public readonly List<string> TelemetryBodies = new List<string>();

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var path = request.RequestUri.AbsolutePath;
                if (path.EndsWith("/telemetry"))
                {
                    if (Fail) return new HttpResponseMessage(HttpStatusCode.InternalServerError);
                    TelemetryBodies.Add(await request.Content.ReadAsStringAsync());
                    return Json(new Dictionary<string, object> { ["accepted"] = 1 });
                }
                if (path.EndsWith("/endpoint"))
                {
                    return Json(new Dictionary<string, object>
                    {
                        ["mode"] = "byok",
                        ["base_url"] = null,
                        ["auth_header"] = null,
                        ["expires_at"] = DateTimeOffset.UtcNow.ToUnixTimeSeconds() + 900,
                        ["refresh_after_seconds"] = 900,
                        ["entitlement"] = new Dictionary<string, object> { ["plan"] = "beta", ["status"] = "active" },
                    });
                }
                return Json(new Dictionary<string, object>
                {
                    ["access_token"] = "access-1",
                    ["refresh_token"] = "refresh-1",
                    ["expires_at"] = DateTimeOffset.UtcNow.ToUnixTimeSeconds() + 1800,
                });
            }

            private static HttpResponseMessage Json(object body) => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"),
            };
        }

        private TelemetryReporter SignedIn(Server server)
        {
            var session = new SessionManager(new TokenVault(Path.Combine(_root, "vault", Guid.NewGuid().ToString("N"))),
                url => new AuthClient(url, server));
            session.SignInAsync("https://api.example.com", "a@b.com", "password-1234").GetAwaiter().GetResult();
            Assert.Equal(SessionState.SignedIn, session.State);
            return new TelemetryReporter(session, () => new AuthClient("https://api.example.com", server), _outbox);
        }

        private TelemetryReporter SignedOut() =>
            new TelemetryReporter(new SessionManager(new TokenVault(Path.Combine(_root, "vault-out"))), outboxPath: _outbox);

        private static List<string> SentEventTypes(Server server) =>
            server.TelemetryBodies
                .SelectMany(body => JsonDocument.Parse(body).RootElement.GetProperty("events").EnumerateArray())
                .Select(e => e.GetProperty("event_type").GetString())
                .ToList();

        [Fact]
        public void Recorded_events_survive_the_process()
        {
            using (var first = SignedOut())
            {
                first.Record("task_complete", new Dictionary<string, object> { ["outcome"] = "success" });
                first.Record("tool_error", new Dictionary<string, object> { ["tool_name"] = "read_range" });
            }
            // 「下次启动」：新的 reporter 读同一个发件箱
            using (var next = SignedOut())
            {
                Assert.Equal(2, next.BufferedCount);
            }
        }

        [Fact]
        public void Events_are_removed_only_after_the_server_accepts_them()
        {
            var server = new Server { Fail = true };
            using (var reporter = SignedIn(server))
            {
                reporter.Record("session_start", new Dictionary<string, object> { ["host"] = "excel" });
                reporter.FlushAsync().GetAwaiter().GetResult();
                Assert.Equal(1, reporter.BufferedCount);  // 服务端报错：留着
                Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(_outbox), "*.sending"));

                server.Fail = false;
                reporter.FlushAsync().GetAwaiter().GetResult();
                Assert.Equal(0, reporter.BufferedCount);
                Assert.Equal(new[] { "session_start" }, SentEventTypes(server));
                Assert.False(File.Exists(_outbox));
            }
        }

        [Fact]
        public void A_load_failure_written_without_a_reporter_goes_out_on_the_next_start()
        {
            TelemetryReporter.AppendToOutbox("startup_error", new Dictionary<string, object>
            {
                ["diagnostic_code"] = StartupErrorCodes.LoadFailed,
            }, _outbox);
            var server = new Server();
            using (var reporter = SignedIn(server))
            {
                reporter.FlushAsync().GetAwaiter().GetResult();
            }
            Assert.Equal(new[] { "startup_error" }, SentEventTypes(server));
            var payload = JsonDocument.Parse(server.TelemetryBodies[0]).RootElement.GetProperty("events")[0].GetProperty("payload");
            Assert.Equal("E-LOAD-001", payload.GetProperty("diagnostic_code").GetString());
        }

        [Fact]
        public void A_batch_claimed_by_a_process_that_died_is_sent_by_the_next_one()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_outbox));
            // pid 这么大的进程不存在：认领了批次然后崩溃的进程
            File.WriteAllText(_outbox + ".2147483000.deadbeef.sending",
                "{\"event_type\":\"tool_error\",\"occurred_at\":\"2026-09-25T00:00:00Z\",\"install_id\":\"x\",\"payload\":{}}\n");
            var server = new Server();
            using (var reporter = SignedIn(server))
            {
                reporter.Record("session_start", null);
                reporter.FlushAsync().GetAwaiter().GetResult();
            }
            Assert.Equal(new[] { "tool_error", "session_start" }, SentEventTypes(server));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(_outbox), "*.sending"));
        }

        [Fact]
        public void A_batch_claimed_by_a_live_process_is_left_alone()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_outbox));
            var live = _outbox + "." + System.Diagnostics.Process.GetCurrentProcess().Id + ".cafe.sending";
            File.WriteAllText(live, "{\"event_type\":\"tool_error\",\"occurred_at\":\"t\",\"install_id\":\"x\",\"payload\":{}}\n");
            var server = new Server();
            using (var reporter = SignedIn(server))
            {
                reporter.Record("session_start", null);
                reporter.FlushAsync().GetAwaiter().GetResult();
            }
            Assert.Equal(new[] { "session_start" }, SentEventTypes(server));  // 别人正在发的不重复发
            Assert.True(File.Exists(live));
        }

        [Fact]
        public void A_line_cut_short_by_a_crash_is_dropped_not_sent()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_outbox));
            File.WriteAllText(_outbox,
                "{\"event_type\":\"session_start\",\"occurred_at\":\"t\",\"install_id\":\"x\",\"payload\":{}}\n" +
                "{\"event_type\":\"task_comp");
            using (var reporter = SignedOut())
            {
                Assert.Equal(1, reporter.BufferedCount);
            }
        }

        [Fact]
        public void The_outbox_stays_bounded_across_processes()
        {
            for (var i = 0; i < TelemetryReporter.MaxBuffered + 50; i++)
            {
                TelemetryReporter.AppendToOutbox("tool_error", new Dictionary<string, object> { ["error_code"] = "e" + i }, _outbox);
            }
            using (var reporter = SignedOut())
            {
                Assert.Equal(TelemetryReporter.MaxBuffered, reporter.BufferedCount);
            }
            // 丢的是最旧的
            Assert.Contains("\"e249\"", File.ReadAllText(_outbox));
            Assert.DoesNotContain("\"e0\"", File.ReadAllText(_outbox));
        }

        [Fact]
        public void Startup_codes_fit_the_server_allowlist_pattern()
        {
            var pattern = new System.Text.RegularExpressions.Regex(@"^E-[A-Z]+-\d{3}$");
            var codes = new[] { StartupErrorCodes.LoadFailed, StartupErrorCodes.BridgeInitFailed,
                StartupErrorCodes.WebViewInitFailed, StartupErrorCodes.SidecarCrashed }
                .Concat(StartupErrorCodes.Engine.Values).ToList();
            Assert.All(codes, c => Assert.Matches(pattern, c));
            Assert.Equal(codes.Count, codes.Distinct().Count());
        }
    }
}
