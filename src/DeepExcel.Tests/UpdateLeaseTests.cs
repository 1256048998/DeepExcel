using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// "Is DeepExcel in use?" answered by the add-in itself instead of by
    /// process names.
    ///
    /// The updater used to refuse to install whenever any process called
    /// EXCEL, et or wps existed. That blocks on WPS Writer, which never loads
    /// DeepExcel, and on any Excel that has the add-in disabled — an update
    /// that waits forever for something unrelated. Every process that loads
    /// the add-in now holds a lease file open; a lease nobody holds is stale.
    /// </summary>
    public class UpdateLeaseTests : IDisposable
    {
        private readonly string _root;

        public UpdateLeaseTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(), "DeepExcelLeaseTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        private static int OwnPid => Process.GetCurrentProcess().Id;

        [Fact]
        public void A_held_lease_names_the_process_holding_it()
        {
            using (UpdateLease lease = UpdateLease.Acquire(_root, "0.6.2.0"))
            {
                Assert.NotNull(lease);
                LeaseHolder holder = Assert.Single(UpdateLease.FindHolders(_root));
                Assert.Equal(OwnPid, holder.ProcessId);
            }
        }

        [Fact]
        public void A_released_lease_leaves_no_holder_and_no_file()
        {
            using (UpdateLease.Acquire(_root, "0.6.2.0"))
            {
            }

            Assert.Empty(UpdateLease.FindHolders(_root));
            Assert.Empty(Directory.GetFiles(UpdateLease.DirectoryFor(_root)));
        }

        [Fact]
        public void A_lease_file_nobody_holds_open_is_stale_and_is_removed()
        {
            // What a power cut leaves behind. Delete-on-close covers crashes —
            // the kernel closes the handle — but not a machine that lost power
            // before the file system recorded the delete.
            string directory = UpdateLease.DirectoryFor(_root);
            Directory.CreateDirectory(directory);
            string stale = Path.Combine(directory, "4242" + UpdateLease.FileExtension);
            File.WriteAllText(stale, "EXCEL 0.6.1.0");

            Assert.Empty(UpdateLease.FindHolders(_root));
            Assert.False(File.Exists(stale));
        }

        [Fact]
        public void A_holder_cannot_have_its_lease_deleted_out_from_under_it()
        {
            // The whole mechanism rests on this: the check is "try to delete
            // it", and only a file somebody still has open survives that.
            using (UpdateLease.Acquire(_root, "0.6.2.0"))
            {
                UpdateLease.FindHolders(_root);
                UpdateLease.FindHolders(_root);

                Assert.Single(UpdateLease.FindHolders(_root));
            }
        }

        [Fact]
        public void No_lease_directory_means_leases_were_never_written_here()
        {
            // Distinguished from "nobody holds one" so the updater can fall back
            // to the old process-name check instead of assuming nothing runs.
            Assert.False(UpdateLease.IsTracked(_root));

            using (UpdateLease.Acquire(_root, "0.6.2.0"))
            {
                Assert.True(UpdateLease.IsTracked(_root));
            }
            Assert.True(UpdateLease.IsTracked(_root));
        }

        [Fact]
        public void Acquiring_twice_in_one_process_does_not_throw()
        {
            // OnConnection runs again when the user re-ticks the add-in in the
            // COM Add-ins dialog. The first lease is still held; the second
            // attempt must degrade to null, never take Excel's startup down.
            using (UpdateLease first = UpdateLease.Acquire(_root, "0.6.2.0"))
            {
                UpdateLease second = UpdateLease.Acquire(_root, "0.6.2.0");

                Assert.NotNull(first);
                Assert.Null(second);
                Assert.Single(UpdateLease.FindHolders(_root));
            }
        }

        [Fact]
        public void An_unusable_root_degrades_to_no_lease()
        {
            string file = Path.Combine(_root, "not-a-directory");
            File.WriteAllText(file, "x");

            Assert.Null(UpdateLease.Acquire(file, "0.6.2.0"));
            Assert.Empty(UpdateLease.FindHolders(file));
        }

        [Fact]
        public void Files_that_are_not_leases_are_left_alone()
        {
            string directory = UpdateLease.DirectoryFor(_root);
            Directory.CreateDirectory(directory);
            string other = Path.Combine(directory, "readme.txt");
            File.WriteAllText(other, "x");
            string unparsable = Path.Combine(directory, "abc" + UpdateLease.FileExtension);
            File.WriteAllText(unparsable, "x");

            Assert.Empty(UpdateLease.FindHolders(_root));
            Assert.True(File.Exists(other));
            Assert.True(File.Exists(unparsable));
        }

        [Fact]
        public void Holders_are_reported_once_each()
        {
            using (UpdateLease.Acquire(_root, "0.6.2.0"))
            {
                Assert.Equal(1, UpdateLease.FindHolders(_root).Select(h => h.ProcessId).Distinct().Count());
            }
        }
    }
}
