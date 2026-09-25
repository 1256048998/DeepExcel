using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace DeepExcel.AddIn.Account
{
    /// <summary>
    /// Buffers telemetry and ships it in batches.
    ///
    /// Three rules, all of which exist because telemetry must never be able to
    /// hurt the product it measures:
    ///
    ///   1. Nothing here blocks a user action. Recording an event is one short
    ///      file append.
    ///   2. Every failure is swallowed. A telemetry outage is not a user problem.
    ///   3. The buffer is bounded. An offline week must not grow the outbox
    ///      without limit; the oldest events are dropped instead.
    ///
    /// The buffer is an outbox file (pending.jsonl, one event per line) rather
    /// than memory. An event is written there when it is recorded and removed
    /// only after the server accepted it, so the events that matter most -- the
    /// ones recorded just before Excel crashed or the add-in failed to load --
    /// are not lost with the process; they go out on the next start. Several
    /// Excel processes share the file: a flush claims it by renaming it to a
    /// per-process *.sending file, so no event is sent twice, and a claim left
    /// behind by a process that died is put back on the next flush.
    ///
    /// The privacy allowlist is enforced on the server (see
    /// server/app/routers/telemetry.py). What is emitted here is chosen to match
    /// it, but the server is what guarantees it -- this client is unsigned and
    /// modifiable, so its filtering is a convenience, not a control.
    /// </summary>
    public sealed class TelemetryReporter : IDisposable
    {
        internal const int MaxBuffered = 200;
        private const int FlushAtCount = 20;
        private static readonly TimeSpan FlushInterval = TimeSpan.FromMinutes(5);
        /// <summary>First flush soon after start: that is when last session's leftovers go out.</summary>
        private static readonly TimeSpan FirstFlushDelay = TimeSpan.FromSeconds(30);
        private const string OutboxMutexName = @"Local\DeepExcel.TelemetryOutbox";
        private const string ClaimSuffix = ".sending";

        private readonly SessionManager _session;
        private readonly Func<AuthClient> _clientFactory;
        private readonly string _outboxPath;
        private readonly Timer _timer;
        private int _flushing;
        private bool _disposed;

        public TelemetryReporter(SessionManager session, Func<AuthClient> clientFactory = null, string outboxPath = null)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _clientFactory = clientFactory;
            _outboxPath = outboxPath ?? DefaultOutboxPath;
            InstallId = LoadOrCreateInstallId();
            _timer = new Timer(_ => FireAndForgetFlush(), null, FirstFlushDelay, FlushInterval);
        }

        public static string DefaultOutboxPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DeepExcel", "telemetry", "pending.jsonl");

        /// <summary>
        /// A random value generated once per installation. It deliberately
        /// carries no machine fingerprint: it answers "how many installs" and
        /// nothing about whose machine this is.
        /// </summary>
        public string InstallId { get; }

        public bool Enabled { get; set; } = true;

        /// <summary>Events waiting in the outbox (not counting a batch in flight)</summary>
        public int BufferedCount => ReadOutbox(_outboxPath).Count;

        public void Record(string eventType, IDictionary<string, object> payload = null)
        {
            if (!Enabled || _disposed || string.IsNullOrEmpty(eventType))
            {
                return;
            }
            var pending = Append(_outboxPath, eventType, payload, InstallId);
            if (pending >= FlushAtCount) FireAndForgetFlush();
        }

        /// <summary>
        /// Writes one event straight to the outbox, for the moments when there
        /// is no reporter to hand it to: the add-in failing to load, before any
        /// account session exists. The next reporter to start sends it.
        /// Synchronous and swallowing, like everything here.
        /// </summary>
        public static void AppendToOutbox(string eventType, IDictionary<string, object> payload, string outboxPath = null)
        {
            if (string.IsNullOrEmpty(eventType)) return;
            Append(outboxPath ?? DefaultOutboxPath, eventType, payload, LoadOrCreateInstallId());
        }

        /// <summary>Returns how many events are now waiting, or 0 when the write failed.</summary>
        private static int Append(string path, string eventType, IDictionary<string, object> payload, string installId)
        {
            try
            {
                var line = JsonSerializer.Serialize(new
                {
                    event_type = eventType,
                    occurred_at = DateTime.UtcNow.ToString("o"),
                    install_id = installId,
                    payload = payload ?? new Dictionary<string, object>()
                });
                var count = 0;
                WithOutbox(() =>
                {
                    var existing = ReadOutbox(path);
                    if (existing.Count >= MaxBuffered)
                    {
                        // Drop oldest. Losing old telemetry is strictly better
                        // than growing without bound while the server is unreachable.
                        existing.Add(line);
                        existing = existing.Skip(existing.Count - MaxBuffered).ToList();
                        RewriteOutbox(path, existing);
                    }
                    else
                    {
                        AppendLines(path, new[] { line });
                        existing.Add(line);
                    }
                    count = existing.Count;
                });
                return count;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private void FireAndForgetFlush()
        {
            // Never awaited on a user path.
            Task.Run(async () =>
            {
                try { await FlushAsync().ConfigureAwait(false); }
                catch (Exception) { }
            });
        }

        public async Task FlushAsync(CancellationToken cancellationToken = default)
        {
            if (!Enabled || _disposed)
            {
                return;
            }
            if (_session.State != SessionState.SignedIn)
            {
                // No account, or offline. Events stay in the outbox until there
                // is somewhere to send them.
                return;
            }
            // One flush at a time per process: the timer and a burst of
            // records can both ask for one.
            if (Interlocked.Exchange(ref _flushing, 1) == 1)
            {
                return;
            }

            string claim = null;
            try
            {
                var endpoint = await _session.GetEndpointAsync(cancellationToken).ConfigureAwait(false);
                if (endpoint == null)
                {
                    return;
                }

                claim = Claim(_outboxPath);
                if (claim == null)
                {
                    return;
                }
                var batch = ReadOutbox(claim);
                if (batch.Count > 0)
                {
                    var events = batch.Select(line => JsonDocument.Parse(line).RootElement.Clone()).ToList();
                    var client = _clientFactory != null
                        ? _clientFactory()
                        : new AuthClient(_session.ServerUrl);
                    using (client)
                    {
                        await client.SendTelemetryAsync(
                            await AccessTokenAsync(cancellationToken).ConfigureAwait(false),
                            new { events },
                            cancellationToken).ConfigureAwait(false);
                    }
                }
                TryDelete(claim);
                claim = null;
            }
            catch (Exception)
            {
                // Put the batch back; the next flush (or the next start) retries.
            }
            finally
            {
                if (claim != null) Release(_outboxPath, claim);
                Interlocked.Exchange(ref _flushing, 0);
            }
        }

        /// <summary>
        /// Takes the whole outbox for sending by renaming it. Claims left by
        /// processes that no longer run are folded back first, so their events
        /// ride along in this batch.
        /// </summary>
        private static string Claim(string path)
        {
            string claim = null;
            WithOutbox(() =>
            {
                RecoverOrphans(path);
                if (!File.Exists(path)) return;
                claim = $"{path}.{Process.GetCurrentProcess().Id}.{Guid.NewGuid():N}{ClaimSuffix}";
                File.Move(path, claim);
            });
            return claim;
        }

        /// <summary>Sending failed: the claimed events go back in front of anything recorded meanwhile.</summary>
        private static void Release(string path, string claim)
        {
            WithOutbox(() =>
            {
                var merged = ReadOutbox(claim);
                merged.AddRange(ReadOutbox(path));
                if (merged.Count > MaxBuffered) merged = merged.Skip(merged.Count - MaxBuffered).ToList();
                RewriteOutbox(path, merged);
                TryDelete(claim);
            });
        }

        /// <summary>Caller holds the outbox mutex.</summary>
        private static void RecoverOrphans(string path)
        {
            var directory = Path.GetDirectoryName(path);
            if (!Directory.Exists(directory)) return;
            var prefix = Path.GetFileName(path) + ".";
            foreach (var file in Directory.GetFiles(directory, prefix + "*" + ClaimSuffix))
            {
                var middle = Path.GetFileName(file).Substring(prefix.Length);
                var pidText = middle.Split('.')[0];
                if (int.TryParse(pidText, out var pid) && IsRunning(pid)) continue;  // still sending
                var merged = ReadOutbox(file);
                merged.AddRange(ReadOutbox(path));
                if (merged.Count > MaxBuffered) merged = merged.Skip(merged.Count - MaxBuffered).ToList();
                RewriteOutbox(path, merged);
                TryDelete(file);
            }
        }

        private static bool IsRunning(int pid)
        {
            try
            {
                using (var process = Process.GetProcessById(pid))
                {
                    return !process.HasExited;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void WithOutbox(Action action)
        {
            try
            {
                using (var mutex = new Mutex(false, OutboxMutexName))
                {
                    var owned = false;
                    try
                    {
                        try { owned = mutex.WaitOne(TimeSpan.FromSeconds(2)); }
                        catch (AbandonedMutexException) { owned = true; }
                        // Without the mutex, still write: a rare lost line beats a lost event.
                        action();
                    }
                    finally
                    {
                        if (owned) mutex.ReleaseMutex();
                    }
                }
            }
            catch (Exception) { }
        }

        private static List<string> ReadOutbox(string path)
        {
            var lines = new List<string>();
            try
            {
                if (!File.Exists(path)) return lines;
                foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
                {
                    var line = raw.Trim();
                    if (line.Length == 0) continue;
                    try
                    {
                        // A line cut short by a crash mid-write is dropped, not sent.
                        using (var doc = JsonDocument.Parse(line))
                        {
                            if (doc.RootElement.ValueKind != JsonValueKind.Object ||
                                !doc.RootElement.TryGetProperty("event_type", out _)) continue;
                        }
                        lines.Add(line);
                    }
                    catch (JsonException) { }
                }
            }
            catch (Exception) { }
            return lines;
        }

        private static void AppendLines(string path, IEnumerable<string> lines)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.AppendAllText(path, string.Concat(lines.Select(l => l + "\n")), new UTF8Encoding(false));
        }

        private static void RewriteOutbox(string path, List<string> lines)
        {
            if (lines.Count == 0)
            {
                TryDelete(path);
                return;
            }
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var temp = path + ".tmp";
            File.WriteAllText(temp, string.Concat(lines.Select(l => l + "\n")), new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, null);
            else File.Move(temp, path);
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (Exception) { }
        }

        /// <summary>
        /// SessionManager owns token renewal; asking it for the endpoint first
        /// guarantees a fresh access token without duplicating that logic here.
        /// </summary>
        private async Task<string> AccessTokenAsync(CancellationToken cancellationToken)
        {
            await _session.GetEndpointAsync(cancellationToken).ConfigureAwait(false);
            return _session.CurrentAccessToken;
        }

        private static string LoadOrCreateInstallId()
        {
            try
            {
                var path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "DeepExcel", "install-id.txt");
                if (File.Exists(path))
                {
                    var existing = File.ReadAllText(path).Trim();
                    if (existing.Length == 32)
                    {
                        return existing;
                    }
                }
                var created = Guid.NewGuid().ToString("N");
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, created);
                return created;
            }
            catch (Exception)
            {
                // Non-persistent id is better than failing: the install simply
                // counts as new next time.
                return Guid.NewGuid().ToString("N");
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            _timer.Dispose();
        }
    }
}
