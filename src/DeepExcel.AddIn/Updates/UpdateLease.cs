using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>A live process that has DeepExcel loaded.</summary>
    public sealed class LeaseHolder
    {
        public int ProcessId { get; set; }

        /// <summary>Best effort; null when the process cannot be inspected.</summary>
        public string ProcessName { get; set; }
    }

    /// <summary>
    /// A file every process that loads the add-in holds open for its lifetime.
    ///
    /// The question the updater has to answer before running the installer is
    /// "does any process still have DeepExcel's files loaded?". It used to
    /// answer a different one -- "does any process called EXCEL, et or wps
    /// exist?" -- which is wrong in both directions: WPS Writer (wps.exe) never
    /// loads DeepExcel yet blocked every update, as did any Excel with the
    /// add-in switched off. Only the add-in knows it is loaded, so it says so:
    ///
    ///     &lt;stage root&gt;\leases\&lt;pid&gt;.lease
    ///
    /// opened with delete-on-close and without FILE_SHARE_DELETE. While the
    /// process lives nobody else can delete the file; when it exits -- cleanly,
    /// by crash or by being killed -- the kernel closes the handle and the file
    /// disappears. So the check is simply "try to delete each lease": what
    /// survives is held. A file that can be deleted was left by a machine that
    /// lost power, and deleting it is the cleanup.
    ///
    /// Shared with DeepExcel.Updater.exe: GAC references only.
    /// </summary>
    public sealed class UpdateLease : IDisposable
    {
        public const string DirectoryName = "leases";
        public const string FileExtension = ".lease";

        private FileStream _stream;

        private UpdateLease(FileStream stream)
        {
            _stream = stream;
        }

        public static string DirectoryFor(string stageRoot)
        {
            return Path.Combine(stageRoot, DirectoryName);
        }

        /// <summary>
        /// Whether leases are kept under this root at all.
        ///
        /// A build that predates leases never creates the directory, and "no
        /// lease directory" must not be read as "nothing is running" -- the
        /// updater falls back to the old name check instead.
        /// </summary>
        public static bool IsTracked(string stageRoot)
        {
            try
            {
                return !string.IsNullOrEmpty(stageRoot) && Directory.Exists(DirectoryFor(stageRoot));
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Takes this process's lease. Returns null instead of throwing: this
        /// runs during Excel's startup, and losing the lease costs a
        /// conservative updater, never a working add-in.
        /// </summary>
        public static UpdateLease Acquire(string stageRoot, string version)
        {
            try
            {
                string directory = DirectoryFor(stageRoot);
                Directory.CreateDirectory(directory);
                Process self = Process.GetCurrentProcess();
                string path = Path.Combine(
                    directory,
                    self.Id.ToString(CultureInfo.InvariantCulture) + FileExtension);

                // FileShare.Read and nothing else: in particular no Delete, which
                // is what makes the file undeletable while held.
                var stream = new FileStream(
                    path, FileMode.Create, FileAccess.ReadWrite, FileShare.Read,
                    256, FileOptions.DeleteOnClose);
                byte[] note = new UTF8Encoding(false).GetBytes(
                    SafeName(self) + " " + (version ?? "") + " " +
                    DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
                stream.Write(note, 0, note.Length);
                stream.Flush();
                return new UpdateLease(stream);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Every process that currently holds a lease. Stale leases are removed
        /// on the way. Never throws; an unreadable directory reports nobody.
        /// </summary>
        public static IList<LeaseHolder> FindHolders(string stageRoot)
        {
            var holders = new List<LeaseHolder>();
            string[] files;
            try
            {
                string directory = DirectoryFor(stageRoot);
                if (!Directory.Exists(directory))
                {
                    return holders;
                }
                files = Directory.GetFiles(directory, "*" + FileExtension);
            }
            catch (Exception)
            {
                return holders;
            }

            foreach (string file in files)
            {
                string stem = Path.GetFileNameWithoutExtension(file);
                if (!int.TryParse(stem, NumberStyles.None, CultureInfo.InvariantCulture, out int pid))
                {
                    // Not ours to judge, and not ours to delete.
                    continue;
                }
                try
                {
                    File.Delete(file);
                    // Deleted: nobody had it open.
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    holders.Add(new LeaseHolder { ProcessId = pid, ProcessName = NameOf(pid) });
                }
            }
            return holders;
        }

        public void Dispose()
        {
            FileStream stream = _stream;
            _stream = null;
            try { stream?.Dispose(); } catch (Exception) { }
        }

        private static string NameOf(int pid)
        {
            try
            {
                using (Process process = Process.GetProcessById(pid))
                {
                    return process.ProcessName;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string SafeName(Process process)
        {
            try { return process.ProcessName; }
            catch (Exception) { return "?"; }
        }
    }
}
