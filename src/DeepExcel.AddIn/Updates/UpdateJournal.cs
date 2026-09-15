using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>A note left by the updater for the version that replaces it.</summary>
    public sealed class AppliedReceipt
    {
        public string FromVersion { get; set; }
        public string ToVersion { get; set; }
        public string AppliedAtUtc { get; set; }
    }

    /// <summary>
    /// The two pieces of update state that have to outlive a process.
    ///
    /// **Attempts.** An update that fails to install is not a one-off: the
    /// package is still staged and still verifies, so the next start finds it
    /// ready and offers it again. Antivirus blocking the installer produces
    /// exactly this, and the roadmap says so explicitly -- 360 / 火绒 will kill
    /// an install with no explanation the user can see. Without a count, that
    /// becomes an unbounded loop of a prompt that can never succeed.
    ///
    /// **Receipt.** The process that knows an update succeeded is the updater,
    /// which then exits and relaunches Excel. The new build has no other way to
    /// learn it was just upgraded, so "how many clients actually took the
    /// update" -- the one question the feature has to answer -- would be
    /// unanswerable. The updater leaves a note; the next start picks it up,
    /// reports it, and deletes it.
    ///
    /// Both live in the staging *root* rather than a version directory, because
    /// <see cref="UpdateStage.Prune"/> deletes version directories and these
    /// have to survive that.
    ///
    /// Shared with DeepExcel.Updater.exe: GAC references only.
    /// </summary>
    public sealed class UpdateJournal
    {
        public const string JournalFileName = "journal.json";
        public const string ReceiptFileName = "applied.json";

        /// <summary>
        /// Three, then stop asking.
        ///
        /// One failure is plausibly transient -- a file lock, a dropped
        /// connection. Three in a row is something structural on that machine,
        /// and continuing to offer an install that cannot succeed just trains
        /// the user to dismiss the banner.
        /// </summary>
        public const int MaxAttempts = 3;

        private readonly string _root;

        public UpdateJournal(string stageRoot)
        {
            _root = stageRoot ?? throw new ArgumentNullException(nameof(stageRoot));
        }

        private string JournalPath => Path.Combine(_root, JournalFileName);

        /// <summary>How many times installing this exact version has been started.</summary>
        public int AttemptsFor(string version)
        {
            Dictionary<string, object> state = ReadObject(JournalPath);
            if (state == null || GetString(state, "version") != version)
            {
                return 0;
            }
            return (int)Math.Max(0, Math.Min(int.MaxValue, GetInt64(state, "attempts")));
        }

        /// <summary>True once this version has burned through its attempts.</summary>
        public bool IsExhausted(string version)
        {
            return AttemptsFor(version) >= MaxAttempts;
        }

        /// <summary>
        /// Counted when the updater is launched, not when it reports back.
        ///
        /// The updater runs after this process is gone, so there is nobody left
        /// to hear that it failed. Counting the attempt up front means a crash,
        /// a kill by antivirus, or a silent installer failure all still burn an
        /// attempt instead of looping forever.
        /// </summary>
        public void RecordAttempt(string version)
        {
            int previous = AttemptsFor(version);
            WriteObject(JournalPath, new Dictionary<string, object>
            {
                ["version"] = version,
                ["attempts"] = previous + 1,
                ["last_utc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
            });
        }

        /// <summary>Called once an upgrade is confirmed, so a later release starts clean.</summary>
        public void ClearAttempts()
        {
            TryDelete(JournalPath);
        }

        /// <summary>Written by the updater immediately after a successful install.</summary>
        public void WriteReceipt(string fromVersion, string toVersion)
        {
            Directory.CreateDirectory(_root);
            WriteObject(Path.Combine(_root, ReceiptFileName), new Dictionary<string, object>
            {
                ["from"] = fromVersion ?? "",
                ["to"] = toVersion ?? "",
                ["utc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
            });
        }

        /// <summary>
        /// Reads and consumes the receipt. Returns null when there is none.
        ///
        /// Deleted on read so one upgrade is reported once, even if telemetry
        /// later fails to reach the server -- a receipt replayed on every start
        /// would inflate the only number anyone looks at.
        /// </summary>
        public AppliedReceipt TakeReceipt()
        {
            string path = Path.Combine(_root, ReceiptFileName);
            Dictionary<string, object> state = ReadObject(path);
            TryDelete(path);
            if (state == null)
            {
                return null;
            }
            string to = GetString(state, "to");
            if (string.IsNullOrEmpty(to))
            {
                return null;
            }
            return new AppliedReceipt
            {
                FromVersion = GetString(state, "from"),
                ToVersion = to,
                AppliedAtUtc = GetString(state, "utc"),
            };
        }

        // ------------------------------------------------------------------
        // Every operation here is best-effort. This is bookkeeping for a
        // background feature; a corrupt or unreadable file must degrade to
        // "no history", never to an exception on somebody's Excel startup.

        private static Dictionary<string, object> ReadObject(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    return null;
                }
                var info = new FileInfo(path);
                if (info.Length > 8 * 1024)
                {
                    return null;
                }
                string json = File.ReadAllText(path, new UTF8Encoding(false));
                return new JavaScriptSerializer { MaxJsonLength = 8 * 1024 }
                    .DeserializeObject(json) as Dictionary<string, object>;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void WriteObject(string path, Dictionary<string, object> state)
        {
            try
            {
                Directory.CreateDirectory(_root);
                File.WriteAllText(
                    path,
                    new JavaScriptSerializer().Serialize(state),
                    new UTF8Encoding(false));
            }
            catch (Exception)
            {
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        private static string GetString(Dictionary<string, object> map, string key)
        {
            return map != null && map.TryGetValue(key, out object value) ? value as string : null;
        }

        private static long GetInt64(Dictionary<string, object> map, string key)
        {
            if (map == null || !map.TryGetValue(key, out object raw) || raw == null)
            {
                return 0;
            }
            if (raw is int i) return i;
            if (raw is long l) return l;
            if (raw is decimal d && d == decimal.Truncate(d) && d >= 0 && d < int.MaxValue) return (long)d;
            return 0;
        }
    }
}
