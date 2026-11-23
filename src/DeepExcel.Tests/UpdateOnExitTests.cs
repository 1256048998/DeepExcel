using System;
using System.Collections.Generic;
using System.IO;
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
    /// Installing a ready update when Excel closes, without being asked.
    ///
    /// Closing Excel is the one moment the installer's precondition is met for
    /// free, so a user who never clicks the banner still gets the update. What
    /// that path must not do is anything the user would notice: no dialog, no
    /// relaunched Excel, and no attempt burned while another Excel window still
    /// has DeepExcel loaded -- that one will apply the update when it closes.
    /// </summary>
    public class UpdateOnExitTests : IDisposable
    {
        private const string FeedUrl = "https://feed.example.com/api/v1/updates/latest";
        private static readonly byte[] Package = Encoding.UTF8.GetBytes("pretend installer bytes");

        private readonly string _root;
        private readonly UpdateTestKey _key = new UpdateTestKey();
        private readonly List<UpdateEvent> _events = new List<UpdateEvent>();

        public UpdateOnExitTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(), "DeepExcelOnExitTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            _key.Dispose();
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        private sealed class Feed : HttpMessageHandler
        {
            private readonly string _manifest;

            public Feed(string manifest)
            {
                _manifest = manifest;
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request, CancellationToken cancellationToken)
            {
                HttpContent content = request.RequestUri.ToString().StartsWith(FeedUrl, StringComparison.Ordinal)
                    ? (HttpContent)new StringContent(_manifest, Encoding.UTF8, "application/json")
                    : new ByteArrayContent(Package);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = content });
            }
        }

        private static string NewerVersion()
        {
            Version current = typeof(UpdateService).Assembly.GetName().Version;
            return string.Format("{0}.{1}.{2}", current.Major, current.Minor, current.Build + 1);
        }

        private UpdateService ReadyService()
        {
            string manifest = _key.SignFor(Package, NewerVersion());
            var service = new UpdateService(
                new UpdateOptions { Enabled = true, FeedUrl = FeedUrl, Channel = "stable" },
                _key.Verifier,
                () => new UpdateDownloader(new Feed(manifest)),
                _root,
                _events.Add);
            Assert.Equal(UpdateStatus.Ready,
                         service.CheckAndStageAsync(null, CancellationToken.None).GetAwaiter().GetResult());
            return service;
        }

        [Fact]
        public void Nothing_staged_means_nothing_to_apply()
        {
            var service = new UpdateService(
                new UpdateOptions { Enabled = true, FeedUrl = FeedUrl, Channel = "stable" },
                _key.Verifier,
                () => new UpdateDownloader(new Feed("")),
                _root,
                _events.Add);

            Assert.Equal(ExitApplyDecision.NotReady, service.ApplyOnExit());
            Assert.Empty(_events);
        }

        [Fact]
        public void Another_process_with_DeepExcel_loaded_defers_to_that_process()
        {
            UpdateService service = ReadyService();
            string leases = UpdateLease.DirectoryFor(_root);
            Directory.CreateDirectory(leases);
            // Held open without delete sharing, exactly as a live add-in holds it.
            using (new FileStream(Path.Combine(leases, "99999" + UpdateLease.FileExtension),
                                  FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                Assert.Equal(ExitApplyDecision.InUseElsewhere, service.ApplyOnExit());
            }

            // Deferring costs nothing: the version is still offered in full.
            Assert.Equal(0, new UpdateJournal(_root).AttemptsFor(service.Staged.Release.Version));
        }

        [Fact]
        public void This_process_s_own_lease_does_not_count_as_somebody_else()
        {
            UpdateService service = ReadyService();
            using (UpdateLease.Acquire(_root, UpdateService.InstalledVersion))
            {
                // Gets past the lease gate; in the test run there is no updater
                // binary beside the assembly, so the launch itself fails.
                Assert.NotEqual(ExitApplyDecision.InUseElsewhere, service.ApplyOnExit());
            }
        }

        [Fact]
        public void An_exhausted_version_is_not_retried_on_exit()
        {
            UpdateService service = ReadyService();
            for (int i = 0; i < UpdateJournal.MaxAttempts; i++)
            {
                new UpdateJournal(_root).RecordAttempt(service.Staged.Release.Version);
            }

            Assert.Equal(ExitApplyDecision.Exhausted, service.ApplyOnExit());
        }

        [Fact]
        public void An_unusable_configuration_never_applies_on_exit()
        {
            UpdateService service = ReadyService();
            var disabled = new UpdateService(
                new UpdateOptions { Enabled = false, FeedUrl = FeedUrl },
                _key.Verifier,
                () => new UpdateDownloader(new Feed("")),
                _root);

            Assert.Equal(ExitApplyDecision.NotReady, disabled.ApplyOnExit());
            Assert.NotNull(service.Staged);
        }

        [Fact]
        public void The_exit_path_is_silent_and_does_not_reopen_Excel()
        {
            UpdateService service = ReadyService();

            string onExit = service.BuildUpdaterArguments(service.Staged, userInitiated: false, hostExecutable: @"C:\Office\EXCEL.EXE");
            string clicked = service.BuildUpdaterArguments(service.Staged, userInitiated: true, hostExecutable: @"C:\Office\EXCEL.EXE");

            // Nobody is waiting for an answer on exit, so a failure must not pop
            // a dialog, and the user just closed Excel on purpose.
            Assert.DoesNotContain("--notify", onExit);
            Assert.DoesNotContain("--relaunch", onExit);
            Assert.Contains("--wait-pid", onExit);
            // The clicked path is unchanged.
            Assert.Contains("--notify", clicked);
            Assert.Contains("--relaunch \"C:\\Office\\EXCEL.EXE\"", clicked);
        }

        [Fact]
        public void A_launch_already_made_this_session_is_not_repeated_on_exit()
        {
            // The banner's "restart and install" launches the updater and then
            // asks Excel to quit -- which runs the exit hook. A second updater
            // racing the first would install twice.
            UpdateService service = ReadyService();
            service.MarkLaunchedForTest();

            Assert.Equal(ExitApplyDecision.AlreadyLaunched, service.ApplyOnExit());
        }
    }
}
