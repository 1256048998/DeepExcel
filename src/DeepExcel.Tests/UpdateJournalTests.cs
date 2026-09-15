using System;
using System.IO;
using System.Text;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Update bookkeeping that has to survive a process restart.
    ///
    /// Two failure modes motivate all of this. An install that cannot succeed on
    /// a given machine — antivirus is the usual cause — stays staged and stays
    /// verifiable, so without a count it is offered again on every single start.
    /// And an install that *does* succeed is known only to the updater, which
    /// exits immediately afterwards, so without a receipt nobody can ever say
    /// how many clients took an update.
    /// </summary>
    public class UpdateJournalTests : IDisposable
    {
        private readonly string _root;
        private readonly UpdateJournal _journal;

        public UpdateJournalTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(), "DeepExcelJournalTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            _journal = new UpdateJournal(_root);
        }

        public void Dispose()
        {
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        [Fact]
        public void A_version_starts_with_no_attempts()
        {
            Assert.Equal(0, _journal.AttemptsFor("0.6.0"));
            Assert.False(_journal.IsExhausted("0.6.0"));
        }

        [Fact]
        public void Attempts_accumulate_and_eventually_exhaust()
        {
            for (int i = 1; i <= UpdateJournal.MaxAttempts; i++)
            {
                _journal.RecordAttempt("0.6.0");
                Assert.Equal(i, _journal.AttemptsFor("0.6.0"));
            }

            Assert.True(_journal.IsExhausted("0.6.0"));
        }

        [Fact]
        public void Attempts_survive_a_new_journal_instance()
        {
            // The count only matters across restarts; an in-memory counter would
            // reset with Excel and never stop anything.
            _journal.RecordAttempt("0.6.0");
            _journal.RecordAttempt("0.6.0");

            Assert.Equal(2, new UpdateJournal(_root).AttemptsFor("0.6.0"));
        }

        [Fact]
        public void A_different_version_starts_over()
        {
            _journal.RecordAttempt("0.6.0");
            _journal.RecordAttempt("0.6.0");
            _journal.RecordAttempt("0.6.0");
            Assert.True(_journal.IsExhausted("0.6.0"));

            // A later release is a different package and may well install fine;
            // punishing it for its predecessor's failures would strand the user.
            Assert.Equal(0, _journal.AttemptsFor("0.6.1"));
            Assert.False(_journal.IsExhausted("0.6.1"));
        }

        [Fact]
        public void Clearing_resets_the_count()
        {
            _journal.RecordAttempt("0.6.0");
            _journal.ClearAttempts();

            Assert.Equal(0, _journal.AttemptsFor("0.6.0"));
        }

        [Fact]
        public void A_receipt_survives_to_the_next_process()
        {
            _journal.WriteReceipt("0.5.0", "0.6.0");

            AppliedReceipt receipt = new UpdateJournal(_root).TakeReceipt();

            Assert.NotNull(receipt);
            Assert.Equal("0.5.0", receipt.FromVersion);
            Assert.Equal("0.6.0", receipt.ToVersion);
            Assert.False(string.IsNullOrEmpty(receipt.AppliedAtUtc));
        }

        [Fact]
        public void A_receipt_is_reported_once_and_only_once()
        {
            _journal.WriteReceipt("0.5.0", "0.6.0");

            Assert.NotNull(_journal.TakeReceipt());
            // Replaying it on every start would inflate the one number anyone
            // looks at: how many clients actually upgraded.
            Assert.Null(_journal.TakeReceipt());
        }

        [Fact]
        public void No_receipt_is_not_an_error()
        {
            Assert.Null(_journal.TakeReceipt());
        }

        [Fact]
        public void The_receipt_lives_beside_the_version_directories_not_inside_one()
        {
            // Prune deletes version directories after an upgrade. A receipt
            // stored inside one would be deleted before it could be read.
            Directory.CreateDirectory(Path.Combine(_root, "0.6.0"));
            _journal.WriteReceipt("0.5.0", "0.6.0");
            _journal.RecordAttempt("0.6.0");

            UpdateStage.Prune(_root, null);

            Assert.False(Directory.Exists(Path.Combine(_root, "0.6.0")));
            Assert.NotNull(new UpdateJournal(_root).TakeReceipt());
        }

        [Theory]
        [InlineData("")]
        [InlineData("not json")]
        [InlineData("[1,2,3]")]
        [InlineData("{\"version\":123,\"attempts\":\"lots\"}")]
        public void A_corrupt_journal_degrades_to_no_history(string content)
        {
            // This runs during Excel startup. Throwing here would turn a
            // truncated write into a broken add-in load.
            File.WriteAllText(
                Path.Combine(_root, UpdateJournal.JournalFileName), content, new UTF8Encoding(false));

            Assert.Equal(0, _journal.AttemptsFor("0.6.0"));
            Assert.False(_journal.IsExhausted("0.6.0"));
        }

        [Fact]
        public void A_corrupt_receipt_degrades_to_no_receipt()
        {
            File.WriteAllText(
                Path.Combine(_root, UpdateJournal.ReceiptFileName), "{{{", new UTF8Encoding(false));

            Assert.Null(_journal.TakeReceipt());
        }

        [Fact]
        public void A_receipt_with_no_target_version_is_ignored()
        {
            File.WriteAllText(
                Path.Combine(_root, UpdateJournal.ReceiptFileName),
                "{\"from\":\"0.5.0\"}", new UTF8Encoding(false));

            Assert.Null(_journal.TakeReceipt());
        }

        [Fact]
        public void Writing_works_even_if_the_root_does_not_exist_yet()
        {
            string fresh = Path.Combine(_root, "never-created");
            var journal = new UpdateJournal(fresh);

            journal.RecordAttempt("0.6.0");

            Assert.Equal(1, journal.AttemptsFor("0.6.0"));
        }
    }
}
