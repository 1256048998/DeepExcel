using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The update cycle end to end, with a fake transport.
    ///
    /// Two things are asserted here that no other test covers: that a version
    /// which repeatedly fails to install stops being offered, and that nothing
    /// reportable ever carries a URL, a path, or exception text. The second is a
    /// privacy commitment, and the natural way to write this code violates it —
    /// the obvious thing to report on failure is the message, and every message
    /// in this subsystem quotes either the feed URL or the staging path.
    /// </summary>
    public class UpdateServiceTests : IDisposable
    {
        private const string FeedUrl = "https://feed.example.com/api/v1/updates/latest";
        private const string PackageUrl = "https://updates.example.com/DeepExcel.Setup.exe";
        private static readonly byte[] Package = Encoding.UTF8.GetBytes("pretend installer bytes");

        private readonly string _root;
        private readonly UpdateTestKey _key = new UpdateTestKey();
        private readonly List<UpdateEvent> _events = new List<UpdateEvent>();

        public UpdateServiceTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(), "DeepExcelUpdateServiceTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            _key.Dispose();
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        // ------------------------------------------------------------------

        private sealed class FakeHandler : HttpMessageHandler
        {
            private readonly string _manifest;
            private readonly byte[] _package;
            private readonly HttpStatusCode _feedStatus;

            public FakeHandler(string manifest, byte[] package,
                               HttpStatusCode feedStatus = HttpStatusCode.OK)
            {
                _manifest = manifest;
                _package = package;
                _feedStatus = feedStatus;
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request, CancellationToken cancellationToken)
            {
                string url = request.RequestUri.ToString();
                if (url.StartsWith(FeedUrl, StringComparison.Ordinal))
                {
                    if (_feedStatus != HttpStatusCode.OK)
                    {
                        return Task.FromResult(new HttpResponseMessage(_feedStatus));
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent(_manifest, Encoding.UTF8, "application/json"),
                    });
                }
                if (url == PackageUrl && _package != null)
                {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ByteArrayContent(_package),
                    });
                }
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }
        }

        private UpdateService Build(
            string manifest, byte[] package = null, HttpStatusCode feedStatus = HttpStatusCode.OK)
        {
            byte[] body = package ?? Package;
            return new UpdateService(
                new UpdateOptions { Enabled = true, FeedUrl = FeedUrl, Channel = "stable" },
                _key.Verifier,
                // A new handler per call: UpdateDownloader owns and disposes it.
                () => new UpdateDownloader(new FakeHandler(manifest, body, feedStatus)),
                _root,
                _events.Add);
        }

        private static UpdateStatus Run(UpdateService service)
        {
            return service.CheckAndStageAsync(null, CancellationToken.None).GetAwaiter().GetResult();
        }

        private UpdateEvent Last => _events[_events.Count - 1];

        /// <summary>A version strictly newer than whatever this assembly reports.</summary>
        private static string NewerVersion()
        {
            Version current = typeof(UpdateService).Assembly.GetName().Version;
            return string.Format("{0}.{1}.{2}", current.Major, current.Minor, current.Build + 1);
        }

        // ------------------------------------------------------------------

        [Fact]
        public void Stages_a_signed_newer_release()
        {
            string version = NewerVersion();
            UpdateService service = Build(_key.SignFor(Package, version));

            Assert.Equal(UpdateStatus.Ready, Run(service));
            Assert.Equal(version, service.Staged.Release.Version);
            Assert.True(File.Exists(Path.Combine(_root, version, UpdateStage.PackageFileName)));
            // Stored verbatim so the updater re-verifies the signature itself.
            Assert.True(File.Exists(Path.Combine(_root, version, UpdateStage.ManifestFileName)));

            Assert.Equal("check", Last.Phase);
            Assert.Equal("ready", Last.Outcome);
            Assert.Equal(version, Last.ToVersion);
        }

        [Fact]
        public void Reports_up_to_date_when_the_feed_offers_nothing_newer()
        {
            UpdateService service = Build(_key.SignFor(Package, "0.0.1"));

            Assert.Equal(UpdateStatus.UpToDate, Run(service));
            Assert.Equal("check", Last.Phase);
            Assert.Equal("up_to_date", Last.Outcome);
        }

        [Fact]
        public void Being_up_to_date_clears_a_previous_version_s_failures()
        {
            // The upgrade landed. Whatever it took to get here is history, and
            // leaving the count behind would handicap the next release.
            new UpdateJournal(_root).RecordAttempt("0.0.1");
            UpdateService service = Build(_key.SignFor(Package, "0.0.1"));

            Run(service);

            Assert.Equal(0, new UpdateJournal(_root).AttemptsFor("0.0.1"));
        }

        [Fact]
        public void Stops_offering_a_version_that_keeps_failing_to_install()
        {
            string version = NewerVersion();
            var journal = new UpdateJournal(_root);
            for (int i = 0; i < UpdateJournal.MaxAttempts; i++)
            {
                journal.RecordAttempt(version);
            }
            UpdateService service = Build(_key.SignFor(Package, version));

            Assert.Equal(UpdateStatus.Blocked, Run(service));
            Assert.Contains("安全软件", service.StatusDetail);
            Assert.Equal("blocked", Last.Outcome);
            Assert.Equal("max_attempts", Last.ReasonCode);
        }

        [Fact]
        public void A_blocked_version_is_still_named_so_the_user_can_install_it_by_hand()
        {
            string version = NewerVersion();
            var journal = new UpdateJournal(_root);
            for (int i = 0; i < UpdateJournal.MaxAttempts; i++)
            {
                journal.RecordAttempt(version);
            }
            UpdateService service = Build(_key.SignFor(Package, version));

            Run(service);

            // Hiding it entirely would leave the user silently stuck on an old
            // build with no idea why.
            Assert.Equal(version, service.Staged?.Release?.Version);
        }

        [Fact]
        public void Refuses_to_launch_a_version_that_has_exhausted_its_attempts()
        {
            string version = NewerVersion();
            UpdateService service = Build(_key.SignFor(Package, version));
            Run(service);

            var journal = new UpdateJournal(_root);
            for (int i = 0; i < UpdateJournal.MaxAttempts; i++)
            {
                journal.RecordAttempt(version);
            }

            // The panel already hides the button; this is the check that makes
            // hiding it irrelevant.
            Assert.False(service.LaunchInstaller(out string error));
            Assert.Contains("手动下载", error);
        }

        [Fact]
        public void Reports_an_upgrade_the_updater_completed()
        {
            new UpdateJournal(_root).WriteReceipt("0.5.0", "0.6.0");
            UpdateService service = Build(_key.SignFor(Package, "0.0.1"));

            AppliedReceipt receipt = service.ReportPendingUpgrade();

            Assert.NotNull(receipt);
            Assert.Equal("apply", Last.Phase);
            Assert.Equal("installed", Last.Outcome);
            Assert.Equal("0.5.0", Last.FromVersion);
            Assert.Equal("0.6.0", Last.ToVersion);
            // Reported once: a replayed receipt would inflate the adoption number.
            Assert.Null(service.ReportPendingUpgrade());
        }

        [Fact]
        public void A_transport_failure_reports_a_code_rather_than_a_message()
        {
            UpdateService service = Build("", feedStatus: HttpStatusCode.ServiceUnavailable);

            Assert.Equal(UpdateStatus.Failed, Run(service));
            Assert.Equal("feed_http_503", Last.ReasonCode);
            // The message is what the user sees; it names the server.
            Assert.Contains("503", service.StatusDetail);
        }

        [Fact]
        public void A_rejected_manifest_reports_the_rejection_reason()
        {
            var payload = UpdateTestKey.Payload(version: NewerVersion(), channel: "beta");
            UpdateService service = Build(_key.Sign(payload));

            Assert.Equal(UpdateStatus.Failed, Run(service));
            Assert.Equal(nameof(UpdateRejection.WrongChannel), Last.ReasonCode);
        }

        [Fact]
        public void A_package_of_the_wrong_length_is_refused_before_it_is_downloaded()
        {
            string version = NewerVersion();
            UpdateService service = Build(
                _key.SignFor(Package, version), Encoding.UTF8.GetBytes("different length entirely"));

            Assert.Equal(UpdateStatus.Failed, Run(service));
            Assert.Equal("download", Last.Phase);
            Assert.Equal("package_size_mismatch", Last.ReasonCode);
        }

        [Fact]
        public void A_package_of_the_right_length_but_wrong_content_is_caught_by_the_digest()
        {
            string version = NewerVersion();
            // Same length, different bytes: the size check passes and only the
            // digest stands between this and an executed installer.
            byte[] impostor = Encoding.UTF8.GetBytes("PRETEND INSTALLER BYTES");
            Assert.Equal(Package.Length, impostor.Length);

            UpdateService service = Build(_key.SignFor(Package, version), impostor);

            Assert.Equal(UpdateStatus.Failed, Run(service));
            Assert.Equal("download", Last.Phase);
            Assert.Equal("package_digest_mismatch", Last.ReasonCode);
            // Left on disk it would be re-verified next start and fail forever.
            Assert.False(File.Exists(Path.Combine(_root, version, UpdateStage.PackageFileName)));
        }

        [Fact]
        public void Telemetry_never_carries_a_url_a_path_or_exception_text()
        {
            // Drive every reporting path there is, then look at everything that
            // came out of them.
            string version = NewerVersion();
            Run(Build(_key.SignFor(Package, version)));                                 // ready
            Run(Build(_key.SignFor(Package, "0.0.1")));                                 // up_to_date
            Run(Build("", feedStatus: HttpStatusCode.ServiceUnavailable));              // transport
            Run(Build("this is not a manifest"));                                       // malformed
            Run(Build(_key.Sign(UpdateTestKey.Payload(version: version, channel: "beta"))));
            Run(Build(_key.SignFor(Package, version), Encoding.UTF8.GetBytes("wrong")));
            new UpdateJournal(_root).WriteReceipt("0.5.0", version);
            Build(_key.SignFor(Package, "0.0.1")).ReportPendingUpgrade();

            Assert.NotEmpty(_events);

            // Asserted as a whitelist rather than a blacklist of forbidden
            // substrings: a token must be an identifier and a version must
            // parse as one. A URL, a Windows path and an exception message all
            // fail that on punctuation or whitespace, including messages in
            // Chinese, which is what this subsystem actually produces.
            foreach (UpdateEvent recorded in _events)
            {
                foreach (string token in new[]
                         {
                             recorded.Phase, recorded.Outcome, recorded.ReasonCode,
                         })
                {
                    if (string.IsNullOrEmpty(token)) continue;
                    Assert.Matches("^[A-Za-z_][A-Za-z0-9_]{0,39}$", token);
                }
                foreach (string reported in new[] { recorded.FromVersion, recorded.ToVersion })
                {
                    if (string.IsNullOrEmpty(reported)) continue;
                    Assert.True(
                        ReleaseVersion.TryParse(reported, out _),
                        "Telemetry version field is not a version: " + reported);
                }
            }
        }

        [Fact]
        public void Every_reported_outcome_is_one_the_server_allows()
        {
            // The server drops unknown values silently, so a typo here would
            // produce an event stream that is quietly missing a whole category.
            var serverAllows = new HashSet<string>
            {
                "up_to_date", "ready", "failed", "blocked", "installed", "started",
            };
            var phasesAllowed = new HashSet<string> { "check", "download", "launch", "apply" };

            string version = NewerVersion();
            Run(Build(_key.SignFor(Package, version)));
            Run(Build(_key.SignFor(Package, "0.0.1")));
            Run(Build("", feedStatus: HttpStatusCode.ServiceUnavailable));
            new UpdateJournal(_root).WriteReceipt("0.5.0", version);
            Build(_key.SignFor(Package, "0.0.1")).ReportPendingUpgrade();

            Assert.NotEmpty(_events);
            Assert.All(_events, e => Assert.Contains(e.Outcome, serverAllows));
            Assert.All(_events, e => Assert.Contains(e.Phase, phasesAllowed));
        }

        [Fact]
        public void A_second_check_reuses_an_already_verified_download()
        {
            string version = NewerVersion();
            Run(Build(_key.SignFor(Package, version)));

            // No package this time: if it tried to download again it would 404.
            UpdateService service = Build(_key.SignFor(Package, version), package: null);

            Assert.Equal(UpdateStatus.Ready, Run(service));
        }

        [Fact]
        public void An_unusable_configuration_reports_disabled_without_reaching_the_network()
        {
            var service = new UpdateService(
                new UpdateOptions { Enabled = false, FeedUrl = FeedUrl },
                _key.Verifier,
                () => throw new InvalidOperationException("must not be constructed"),
                _root,
                _events.Add);

            Assert.Equal(UpdateStatus.Disabled, Run(service));
            Assert.Empty(_events);
        }

        [Fact]
        public void Staging_a_new_version_removes_the_previous_one()
        {
            string older = NewerVersion();
            Run(Build(_key.SignFor(Package, older)));
            Assert.True(Directory.Exists(Path.Combine(_root, older)));

            Version current = typeof(UpdateService).Assembly.GetName().Version;
            string newer = string.Format(
                "{0}.{1}.{2}", current.Major, current.Minor, current.Build + 2);
            Run(Build(_key.SignFor(Package, newer)));

            // Otherwise every release leaves another installer on the disk.
            Assert.False(Directory.Exists(Path.Combine(_root, older)));
            Assert.True(Directory.Exists(Path.Combine(_root, newer)));
        }
    }
}
