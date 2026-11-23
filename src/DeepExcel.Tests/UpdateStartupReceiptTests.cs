using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// "The new version actually started" -- written by the add-in, read by the
    /// updater.
    ///
    /// An installer returning 0 says the files were copied, not that the
    /// result loads. After a user-initiated update the updater relaunches
    /// Excel and waits for this receipt; when it never comes, the new version
    /// is put back to the old one.
    /// </summary>
    public class UpdateStartupReceiptTests : IDisposable
    {
        private readonly string _root;
        private readonly UpdateTestKey _key = new UpdateTestKey();

        public UpdateStartupReceiptTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(), "DeepExcelStartupReceiptTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            _key.Dispose();
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        [Fact]
        public void No_receipt_until_a_build_has_started()
        {
            Assert.Null(new UpdateJournal(_root).ReadStartupReceipt());
        }

        [Fact]
        public void A_receipt_names_the_version_the_process_and_the_time()
        {
            DateTime before = DateTime.UtcNow.AddSeconds(-1);
            new UpdateJournal(_root).WriteStartupReceipt("0.6.2.0");

            StartupReceipt receipt = new UpdateJournal(_root).ReadStartupReceipt();

            Assert.NotNull(receipt);
            Assert.Equal("0.6.2.0", receipt.Version);
            Assert.Equal(Process.GetCurrentProcess().Id, receipt.ProcessId);
            Assert.True(receipt.WrittenAtUtc >= before, "receipt time " + receipt.WrittenAtUtc.ToString("o"));
            Assert.Equal(DateTimeKind.Utc, receipt.WrittenAtUtc.Kind);
        }

        [Fact]
        public void The_latest_start_overwrites_the_previous_receipt()
        {
            // Not consumed on read: the updater only ever asks "has this
            // version started since I installed it", and every start answers.
            var journal = new UpdateJournal(_root);
            journal.WriteStartupReceipt("0.6.1.0");
            journal.WriteStartupReceipt("0.6.2.0");

            Assert.Equal("0.6.2.0", journal.ReadStartupReceipt().Version);
            Assert.Equal("0.6.2.0", journal.ReadStartupReceipt().Version);
        }

        [Fact]
        public void A_receipt_proves_a_version_started_only_after_a_given_moment()
        {
            var journal = new UpdateJournal(_root);
            journal.WriteStartupReceipt("0.6.2.0");

            Assert.True(journal.HasStartedSince("0.6.2", DateTime.UtcNow.AddMinutes(-1)));
            // An old receipt for the same version is from before the install.
            Assert.False(journal.HasStartedSince("0.6.2", DateTime.UtcNow.AddMinutes(1)));
            // The previous version starting proves nothing about the new one.
            Assert.False(journal.HasStartedSince("0.6.3", DateTime.UtcNow.AddMinutes(-1)));
        }

        [Fact]
        public void A_corrupt_receipt_reads_as_no_receipt()
        {
            File.WriteAllText(Path.Combine(_root, UpdateJournal.StartupReceiptFileName), "{ not json");

            Assert.Null(new UpdateJournal(_root).ReadStartupReceipt());
            Assert.False(new UpdateJournal(_root).HasStartedSince("0.6.2", DateTime.MinValue));
        }

        [Fact]
        public void The_service_records_the_running_build_as_started()
        {
            var service = new UpdateService(
                new UpdateOptions { Enabled = false },
                _key.Verifier,
                () => new UpdateDownloader(new HttpClientHandler()),
                _root);

            // Even with updating switched off: an update the user installed by
            // hand is still waiting to hear whether it started.
            service.RecordSuccessfulStart();

            Assert.Equal(UpdateService.InstalledVersion, new UpdateJournal(_root).ReadStartupReceipt().Version);
        }
    }
}
