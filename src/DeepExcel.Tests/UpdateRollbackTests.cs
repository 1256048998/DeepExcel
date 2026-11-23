using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Going back to the version that was working.
    ///
    /// An install can fail half way -- antivirus killing the installer is the
    /// usual cause -- and leave a mix of old and new files that loads as
    /// neither. The only thing that reliably undoes that is the previous
    /// version's own installer, so the last version that was confirmed to
    /// start keeps its signed manifest and package in updates\rollback\.
    /// Nothing in that folder is ever run without being re-verified against
    /// the signature first: it is as user-writable as the stage itself.
    /// </summary>
    public class UpdateRollbackTests : IDisposable
    {
        private static readonly byte[] Package = Encoding.UTF8.GetBytes("previous installer bytes");

        private readonly string _root;
        private readonly UpdateTestKey _key = new UpdateTestKey();
        private readonly List<UpdateEvent> _events = new List<UpdateEvent>();

        public UpdateRollbackTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(), "DeepExcelRollbackTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            _key.Dispose();
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        /// <summary>Writes a complete stage for <paramref name="version"/>, as UpdateService leaves it.</summary>
        private string StageVersion(string version, byte[] package = null)
        {
            string directory = UpdateStage.DirectoryFor(_root, version);
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, UpdateStage.PackageFileName), package ?? Package);
            File.WriteAllText(
                Path.Combine(directory, UpdateStage.ManifestFileName),
                _key.SignFor(package ?? Package, version));
            return directory;
        }

        private UpdateService Service()
        {
            return new UpdateService(
                new UpdateOptions { Enabled = true, FeedUrl = "https://feed.example.com/latest", Channel = "stable" },
                _key.Verifier,
                () => new UpdateDownloader(new HttpClientHandler()),
                _root,
                _events.Add);
        }

        // ---- the manifest ----------------------------------------------------

        [Fact]
        public void An_installed_version_s_manifest_verifies_even_though_it_is_not_newer()
        {
            // Evaluate() refuses anything not strictly newer -- correct for an
            // update, and exactly wrong for a rollback, whose whole point is to
            // go back to the version that was installed.
            UpdateCheckResult result = UpdateManifest.EvaluateInstalled(
                _key.SignFor(Package, "0.6.1"), "0.6.1.0", "stable", _key.Verifier);

            Assert.True(result.Accepted, result.Detail);
            Assert.Equal("0.6.1", result.Release.Version);
        }

        [Fact]
        public void A_rollback_manifest_for_any_other_version_is_refused()
        {
            // The only version a rollback may restore is the one that was
            // installed when the update began. Anything else is a downgrade
            // wearing a rollback's clothes.
            UpdateCheckResult older = UpdateManifest.EvaluateInstalled(
                _key.SignFor(Package, "0.5.0"), "0.6.1", "stable", _key.Verifier);
            UpdateCheckResult newer = UpdateManifest.EvaluateInstalled(
                _key.SignFor(Package, "0.6.2"), "0.6.1", "stable", _key.Verifier);

            Assert.Equal(UpdateRejection.VersionMismatch, older.Rejection);
            Assert.Equal(UpdateRejection.VersionMismatch, newer.Rejection);
        }

        [Fact]
        public void A_rollback_manifest_still_has_to_carry_a_valid_signature()
        {
            string manifest = _key.SignFor(Package, "0.6.1");
            using (var other = new UpdateTestKey())
            {
                UpdateCheckResult result = UpdateManifest.EvaluateInstalled(
                    manifest, "0.6.1", "stable", other.Verifier);

                Assert.False(result.Accepted);
                Assert.Equal(UpdateRejection.UnknownKey, result.Rejection);
            }
        }

        [Fact]
        public void A_rollback_manifest_from_another_channel_is_refused()
        {
            string manifest = _key.Sign(UpdateTestKey.Payload(version: "0.6.1", channel: "beta"));

            UpdateCheckResult result = UpdateManifest.EvaluateInstalled(
                manifest, "0.6.1", "stable", _key.Verifier);

            Assert.Equal(UpdateRejection.WrongChannel, result.Rejection);
        }

        // ---- the rollback slot -----------------------------------------------

        [Fact]
        public void A_confirmed_version_is_kept_for_rollback_and_verifies()
        {
            StageVersion("0.6.1");

            Assert.True(UpdateStage.KeepForRollback(_root, "0.6.1.0", "stable", _key.Verifier));

            UpdateCheckResult kept = UpdateStage.LoadRollback(_root, "0.6.1.0", "stable", _key.Verifier);
            Assert.True(kept.Accepted, kept.Detail);
            Assert.True(File.Exists(Path.Combine(
                UpdateStage.RollbackDirectoryFor(_root, "0.6.1"), UpdateStage.PackageFileName)));
        }

        [Fact]
        public void Keeping_a_new_version_drops_the_previous_one()
        {
            // One previous version is all a rollback ever needs, and each one is
            // a full installer.
            StageVersion("0.6.1");
            UpdateStage.KeepForRollback(_root, "0.6.1", "stable", _key.Verifier);
            StageVersion("0.6.2");
            UpdateStage.KeepForRollback(_root, "0.6.2", "stable", _key.Verifier);

            Assert.False(Directory.Exists(UpdateStage.RollbackDirectoryFor(_root, "0.6.1")));
            Assert.True(UpdateStage.LoadRollback(_root, "0.6.2", "stable", _key.Verifier).Accepted);
        }

        [Fact]
        public void A_stage_that_does_not_verify_is_not_kept()
        {
            string directory = StageVersion("0.6.1");
            File.WriteAllBytes(Path.Combine(directory, UpdateStage.PackageFileName),
                               Encoding.UTF8.GetBytes("PREVIOUS INSTALLER BYTES"));

            Assert.False(UpdateStage.KeepForRollback(_root, "0.6.1", "stable", _key.Verifier));
            Assert.False(Directory.Exists(UpdateStage.RollbackDirectoryFor(_root, "0.6.1")));
        }

        [Fact]
        public void A_failed_keep_leaves_the_existing_rollback_in_place()
        {
            // Dropping the old slot before the new one is known good would turn
            // a bad download into "no rollback at all".
            StageVersion("0.6.1");
            UpdateStage.KeepForRollback(_root, "0.6.1", "stable", _key.Verifier);

            Assert.False(UpdateStage.KeepForRollback(_root, "0.6.2", "stable", _key.Verifier));

            Assert.True(UpdateStage.LoadRollback(_root, "0.6.1", "stable", _key.Verifier).Accepted);
        }

        [Fact]
        public void A_rollback_package_swapped_after_it_was_kept_is_refused()
        {
            StageVersion("0.6.1");
            UpdateStage.KeepForRollback(_root, "0.6.1", "stable", _key.Verifier);
            File.WriteAllBytes(
                Path.Combine(UpdateStage.RollbackDirectoryFor(_root, "0.6.1"), UpdateStage.PackageFileName),
                Encoding.UTF8.GetBytes("PREVIOUS INSTALLER BYTES"));

            UpdateCheckResult result = UpdateStage.LoadRollback(_root, "0.6.1", "stable", _key.Verifier);

            Assert.False(result.Accepted);
            Assert.Equal(UpdateRejection.BadDigest, result.Rejection);
        }

        [Fact]
        public void No_rollback_slot_is_reported_rather_than_thrown()
        {
            // The first update after a manual install has nothing to go back to.
            UpdateCheckResult result = UpdateStage.LoadRollback(_root, "0.6.1", "stable", _key.Verifier);

            Assert.False(result.Accepted);
        }

        [Fact]
        public void Pruning_stages_never_touches_the_rollback_slot()
        {
            StageVersion("0.6.1");
            UpdateStage.KeepForRollback(_root, "0.6.1", "stable", _key.Verifier);

            UpdateStage.Prune(_root, null);

            Assert.True(UpdateStage.LoadRollback(_root, "0.6.1", "stable", _key.Verifier).Accepted);
        }

        // ---- the journal -----------------------------------------------------

        [Fact]
        public void A_rollback_note_is_read_once()
        {
            var journal = new UpdateJournal(_root);
            journal.WriteRollback("0.6.1", "0.6.2", "installer_failed");

            RollbackNote note = journal.TakeRollback();

            Assert.NotNull(note);
            Assert.Equal("0.6.1", note.RestoredVersion);
            Assert.Equal("0.6.2", note.FailedVersion);
            Assert.Equal("installer_failed", note.ReasonCode);
            Assert.Null(journal.TakeRollback());
        }

        // ---- the service -----------------------------------------------------

        [Fact]
        public void A_rollback_the_updater_performed_is_reported_once()
        {
            new UpdateJournal(_root).WriteRollback("0.6.1", "0.6.2", "installer_failed");

            Service().ReportPendingUpgrade();
            Service().ReportPendingUpgrade();

            UpdateEvent reported = Assert.Single(_events);
            Assert.Equal("apply", reported.Phase);
            Assert.Equal("rolled_back", reported.Outcome);
            Assert.Equal("installer_failed", reported.ReasonCode);
            Assert.Equal("0.6.1", reported.FromVersion);
            Assert.Equal("0.6.2", reported.ToVersion);
        }

        [Fact]
        public void The_running_build_keeps_itself_for_rollback_once_its_upgrade_is_confirmed()
        {
            // This code running at all is the proof the new version starts, so
            // this is the moment it becomes the version to go back to.
            string running = UpdateService.InstalledVersion;
            StageVersion(running);
            new UpdateJournal(_root).WriteReceipt("0.0.1", running);

            Service().ReportPendingUpgrade();

            Assert.True(UpdateStage.LoadRollback(_root, running, "stable", _key.Verifier).Accepted);
        }

        [Fact]
        public void A_receipt_for_some_other_version_does_not_replace_the_rollback()
        {
            // The receipt says 0.6.9 installed, but 0.6.9 is not what is running
            // -- whatever happened, it was not confirmed to start.
            StageVersion("0.6.9");
            new UpdateJournal(_root).WriteReceipt("0.0.1", "0.6.9");

            Service().ReportPendingUpgrade();

            Assert.False(Directory.Exists(UpdateStage.RollbackDirectoryFor(_root, "0.6.9")));
        }
    }
}
