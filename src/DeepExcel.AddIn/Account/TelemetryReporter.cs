using System;
using System.Collections.Generic;
using System.IO;
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
    ///   1. Nothing here blocks a user action. Recording an event is an enqueue.
    ///   2. Every failure is swallowed. A telemetry outage is not a user problem.
    ///   3. The buffer is bounded. An offline week must not grow memory without
    ///      limit; the oldest events are dropped instead.
    ///
    /// The privacy allowlist is enforced on the server (see
    /// server/app/routers/telemetry.py). What is emitted here is chosen to match
    /// it, but the server is what guarantees it -- this client is unsigned and
    /// modifiable, so its filtering is a convenience, not a control.
    /// </summary>
    public sealed class TelemetryReporter : IDisposable
    {
        private const int MaxBuffered = 200;
        private const int FlushAtCount = 20;
        private static readonly TimeSpan FlushInterval = TimeSpan.FromMinutes(5);

        private readonly SessionManager _session;
        private readonly Func<AuthClient> _clientFactory;
        private readonly List<object> _buffer = new List<object>();
        private readonly object _lock = new object();
        private readonly Timer _timer;
        private bool _disposed;

        public TelemetryReporter(SessionManager session, Func<AuthClient> clientFactory = null)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _clientFactory = clientFactory;
            _timer = new Timer(_ => FireAndForgetFlush(), null, FlushInterval, FlushInterval);
            InstallId = LoadOrCreateInstallId();
        }

        /// <summary>
        /// A random value generated once per installation. It deliberately
        /// carries no machine fingerprint: it answers "how many installs" and
        /// nothing about whose machine this is.
        /// </summary>
        public string InstallId { get; }

        public bool Enabled { get; set; } = true;

        public int BufferedCount
        {
            get { lock (_lock) { return _buffer.Count; } }
        }

        public void Record(string eventType, IDictionary<string, object> payload = null)
        {
            if (!Enabled || _disposed || string.IsNullOrEmpty(eventType))
            {
                return;
            }

            lock (_lock)
            {
                if (_buffer.Count >= MaxBuffered)
                {
                    // Drop oldest. Losing old telemetry is strictly better than
                    // growing without bound while the server is unreachable.
                    _buffer.RemoveAt(0);
                }
                _buffer.Add(new
                {
                    event_type = eventType,
                    occurred_at = DateTime.UtcNow.ToString("o"),
                    install_id = InstallId,
                    payload = payload ?? new Dictionary<string, object>()
                });

                if (_buffer.Count < FlushAtCount)
                {
                    return;
                }
            }
            FireAndForgetFlush();
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
                // No account, or offline. Events stay buffered until there is
                // somewhere to send them.
                return;
            }

            List<object> batch;
            lock (_lock)
            {
                if (_buffer.Count == 0)
                {
                    return;
                }
                batch = new List<object>(_buffer);
                _buffer.Clear();
            }

            try
            {
                var endpoint = await _session.GetEndpointAsync(cancellationToken).ConfigureAwait(false);
                if (endpoint == null)
                {
                    Requeue(batch);
                    return;
                }

                var client = _clientFactory != null
                    ? _clientFactory()
                    : new AuthClient(_session.ServerUrl);
                using (client)
                {
                    await client.SendTelemetryAsync(
                        await AccessTokenAsync(cancellationToken).ConfigureAwait(false),
                        new { events = batch },
                        cancellationToken).ConfigureAwait(false);
                }
            }
            catch (Exception)
            {
                // Put them back so a transient outage does not lose data, but
                // only up to the cap -- see Record.
                Requeue(batch);
            }
        }

        private void Requeue(List<object> batch)
        {
            lock (_lock)
            {
                var room = MaxBuffered - _buffer.Count;
                if (room <= 0)
                {
                    return;
                }
                var keep = batch.Count <= room ? batch : batch.GetRange(batch.Count - room, room);
                _buffer.InsertRange(0, keep);
            }
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
