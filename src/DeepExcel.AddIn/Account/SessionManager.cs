using System;
using System.Threading;
using System.Threading.Tasks;

namespace DeepExcel.AddIn.Account
{
    public enum SessionState
    {
        SignedOut,
        SignedIn,
        /// <summary>Signed in previously, server unreachable, inside the grace window.</summary>
        Offline,
        /// <summary>Session was rejected; the user must sign in again.</summary>
        Expired
    }

    /// <summary>
    /// Owns the signed-in session and the endpoint configuration.
    ///
    /// The single rule this class enforces: the routing decision comes from the
    /// server and is cached, never computed locally. Everything else here exists
    /// to keep that true when the network is unreliable.
    /// </summary>
    public sealed class SessionManager
    {
        /// <summary>
        /// How long the add-in keeps working after the server becomes
        /// unreachable.
        ///
        /// This is not a security hole. In BYOK mode the user's own key does the
        /// work, so there is nothing to withhold. In hosted mode the proxy
        /// authenticates every request itself, so a stale local session buys
        /// nothing. What it does buy is that a server outage, or a user on a
        /// plane, does not brick a spreadsheet tool.
        /// </summary>
        public static readonly TimeSpan OfflineGrace = TimeSpan.FromDays(7);

        /// <summary>Re-fetch slightly before expiry so a request never races it.</summary>
        private static readonly TimeSpan RenewalMargin = TimeSpan.FromMinutes(1);

        private readonly TokenVault _vault;
        private readonly Func<string, AuthClient> _clientFactory;
        private readonly Func<DateTimeOffset> _clock;
        private readonly SemaphoreSlim _gate = new SemaphoreSlim(1, 1);

        private AuthClient _client;
        private string _accessToken;
        private DateTimeOffset _accessTokenExpiresAt;
        private string _refreshToken;
        private EndpointConfig _endpoint;
        private DateTime _lastVerifiedUtc;

        public SessionManager(
            TokenVault vault = null,
            Func<string, AuthClient> clientFactory = null,
            Func<DateTimeOffset> clock = null)
        {
            _vault = vault ?? new TokenVault();
            _clientFactory = clientFactory ?? (url => new AuthClient(url));
            _clock = clock ?? (() => DateTimeOffset.UtcNow);
        }

        public SessionState State { get; private set; } = SessionState.SignedOut;
        public string Email { get; private set; }
        public string ServerUrl { get; private set; }

        public event Action<SessionState> StateChanged;

        /// <summary>
        /// The current routing decision, or null when there is none.
        /// Callers must treat null as "not ready", never as "use BYOK".
        /// </summary>
        public EndpointConfig CurrentEndpoint
        {
            get { return _endpoint; }
        }

        /// <summary>
        /// Access token for calling the account API. Callers must ask for the
        /// endpoint first so renewal happens in one place rather than being
        /// re-implemented by each caller.
        /// </summary>
        public string CurrentAccessToken
        {
            get { return _accessToken; }
        }

        // ------------------------------------------------------------------
        // Lifecycle
        // ------------------------------------------------------------------

        /// <summary>
        /// Restores a previous session at startup. Never throws: a failure here
        /// must leave the user at the sign-in screen, not break the add-in.
        /// </summary>
        public async Task<SessionState> RestoreAsync(CancellationToken cancellationToken = default)
        {
            TokenVault.PersistedSession stored;
            if (!_vault.TryLoad(out stored) || string.IsNullOrEmpty(stored.ServerUrl))
            {
                SetState(SessionState.SignedOut);
                return State;
            }

            ServerUrl = stored.ServerUrl;
            Email = stored.Email;
            _refreshToken = stored.RefreshToken;
            _lastVerifiedUtc = stored.LastVerifiedUtc;
            _client = _clientFactory(ServerUrl);

            try
            {
                await RefreshSessionAsync(cancellationToken).ConfigureAwait(false);
                SetState(SessionState.SignedIn);
            }
            catch (AuthException ex) when (ex.Failure == AuthFailure.Network)
            {
                // Offline. Keep going if the last confirmed contact is recent.
                SetState(WithinGrace() ? SessionState.Offline : SessionState.Expired);
            }
            catch (AuthException)
            {
                // Rejected rather than unreachable: the stored token is useless.
                _vault.Clear();
                _refreshToken = null;
                SetState(SessionState.Expired);
            }

            return State;
        }

        public async Task SignInAsync(
            string serverUrl, string email, string password,
            CancellationToken cancellationToken = default)
        {
            var client = _clientFactory(serverUrl);
            TokenPair tokens = await client.LoginAsync(email, password, cancellationToken)
                .ConfigureAwait(false);
            await AdoptAsync(client, serverUrl, email, tokens, cancellationToken).ConfigureAwait(false);
        }

        public async Task RegisterAsync(
            string serverUrl, string email, string password, string inviteCode,
            CancellationToken cancellationToken = default)
        {
            var client = _clientFactory(serverUrl);
            TokenPair tokens = await client
                .RegisterAsync(email, password, inviteCode, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            await AdoptAsync(client, serverUrl, email, tokens, cancellationToken).ConfigureAwait(false);
        }

        public async Task SignOutAsync(CancellationToken cancellationToken = default)
        {
            if (_client != null && !string.IsNullOrEmpty(_refreshToken))
            {
                await _client.LogoutAsync(_refreshToken, cancellationToken).ConfigureAwait(false);
            }
            _vault.Clear();
            _accessToken = null;
            _refreshToken = null;
            _endpoint = null;
            Email = null;
            SetState(SessionState.SignedOut);
        }

        // ------------------------------------------------------------------
        // Routing
        // ------------------------------------------------------------------

        /// <summary>
        /// The endpoint configuration to hand to the sidecar, refreshing it when
        /// the cached one has expired.
        ///
        /// When the server is unreachable this keeps returning the cached value
        /// for the grace window. The alternative -- failing closed -- would mean
        /// a brief server hiccup stops the user editing their own spreadsheet
        /// with their own API key.
        /// </summary>
        public async Task<EndpointConfig> GetEndpointAsync(CancellationToken cancellationToken = default)
        {
            if (State == SessionState.SignedOut || State == SessionState.Expired)
            {
                return null;
            }

            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                if (_endpoint != null && _clock() < _endpoint.ExpiresAt - RenewalMargin)
                {
                    return _endpoint;
                }

                try
                {
                    await RefreshSessionAsync(cancellationToken).ConfigureAwait(false);
                    SetState(SessionState.SignedIn);
                    return _endpoint;
                }
                catch (AuthException ex) when (ex.Failure == AuthFailure.Network)
                {
                    if (_endpoint != null && WithinGrace())
                    {
                        SetState(SessionState.Offline);
                        return _endpoint;
                    }
                    throw;
                }
                catch (AuthException ex) when (ex.Failure == AuthFailure.SessionExpired
                                               || ex.Failure == AuthFailure.AccountDisabled)
                {
                    _vault.Clear();
                    _endpoint = null;
                    SetState(SessionState.Expired);
                    throw;
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        // ------------------------------------------------------------------

        private async Task AdoptAsync(
            AuthClient client, string serverUrl, string email, TokenPair tokens,
            CancellationToken cancellationToken)
        {
            _client = client;
            ServerUrl = client.ServerUrl;
            Email = email;
            _accessToken = tokens.AccessToken;
            _accessTokenExpiresAt = tokens.ExpiresAt;
            _refreshToken = tokens.RefreshToken;
            _lastVerifiedUtc = DateTime.UtcNow;

            await FetchEndpointAsync(cancellationToken).ConfigureAwait(false);
            Persist();
            SetState(SessionState.SignedIn);
        }

        private async Task RefreshSessionAsync(CancellationToken cancellationToken)
        {
            if (_client == null || string.IsNullOrEmpty(_refreshToken))
            {
                throw new AuthException(AuthFailure.SessionExpired, "没有可用的登录状态。");
            }

            if (_accessToken == null || _clock() >= _accessTokenExpiresAt - RenewalMargin)
            {
                TokenPair tokens = await _client.RefreshAsync(_refreshToken, cancellationToken)
                    .ConfigureAwait(false);
                _accessToken = tokens.AccessToken;
                _accessTokenExpiresAt = tokens.ExpiresAt;
                // The server rotates refresh tokens, so the new one must be
                // persisted or the next restore will present a revoked token.
                _refreshToken = tokens.RefreshToken;
            }

            await FetchEndpointAsync(cancellationToken).ConfigureAwait(false);
            _lastVerifiedUtc = DateTime.UtcNow;
            Persist();
        }

        private async Task FetchEndpointAsync(CancellationToken cancellationToken)
        {
            EndpointConfig config = await _client.GetEndpointAsync(_accessToken, cancellationToken)
                .ConfigureAwait(false);

            if (config == null || !config.IsUsable())
            {
                // Applying an incomplete hosted config would point the sidecar
                // at nothing. Keep the previous one and surface the fault.
                throw new AuthException(
                    AuthFailure.ServerError, "服务器返回的出口配置不完整。");
            }
            _endpoint = config;
        }

        private void Persist()
        {
            _vault.Save(new TokenVault.PersistedSession
            {
                ServerUrl = ServerUrl,
                RefreshToken = _refreshToken,
                Email = Email,
                LastVerifiedUtc = _lastVerifiedUtc
            });
        }

        private bool WithinGrace()
        {
            if (_lastVerifiedUtc == default)
            {
                return false;
            }
            return DateTime.UtcNow - _lastVerifiedUtc < OfflineGrace;
        }

        private void SetState(SessionState state)
        {
            if (State == state)
            {
                return;
            }
            State = state;
            Action<SessionState> handler = StateChanged;
            if (handler != null)
            {
                try { handler(state); }
                catch (Exception) { /* a UI subscriber must not break the session */ }
            }
        }
    }
}
