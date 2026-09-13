using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Account;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The client half of the endpoint-routing contract.
    ///
    /// The behaviour under test is the one that makes M6 a configuration change
    /// rather than a client release: the client applies whatever the server
    /// returns and never decides the routing mode itself.
    /// </summary>
    public class SessionManagerTests : IDisposable
    {
        private readonly string _vaultDirectory;

        public SessionManagerTests()
        {
            _vaultDirectory = Path.Combine(Path.GetTempPath(), "DeepExcelSessionTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_vaultDirectory);
        }

        public void Dispose()
        {
            try { Directory.Delete(_vaultDirectory, true); } catch (Exception) { }
        }

        // ------------------------------------------------------------------
        // Fake transport
        // ------------------------------------------------------------------

        private sealed class FakeHandler : HttpMessageHandler
        {
            public Func<HttpRequestMessage, HttpResponseMessage> Responder;
            public bool Offline;
            public readonly List<string> Requests = new List<string>();

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request, CancellationToken cancellationToken)
            {
                Requests.Add(request.Method + " " + request.RequestUri.AbsolutePath);
                if (Offline)
                {
                    throw new HttpRequestException("simulated network failure");
                }
                return Task.FromResult(Responder(request));
            }
        }

        private static HttpResponseMessage Json(HttpStatusCode status, object body)
        {
            return new HttpResponseMessage(status)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
            };
        }

        private static object TokenPairBody(long expiresInSeconds = 1800, string refresh = "refresh-1")
        {
            return new Dictionary<string, object>
            {
                ["access_token"] = "access-1",
                ["refresh_token"] = refresh,
                ["expires_at"] = DateTimeOffset.UtcNow.ToUnixTimeSeconds() + expiresInSeconds
            };
        }

        private static object EndpointBody(string mode, string baseUrl, string authHeader, long ttl = 900)
        {
            return new Dictionary<string, object>
            {
                ["mode"] = mode,
                ["base_url"] = baseUrl,
                ["auth_header"] = authHeader,
                ["expires_at"] = DateTimeOffset.UtcNow.ToUnixTimeSeconds() + ttl,
                ["refresh_after_seconds"] = ttl,
                ["entitlement"] = new Dictionary<string, object>
                {
                    ["plan"] = "beta",
                    ["status"] = "active",
                    ["task_limit"] = null,
                    ["tasks_used"] = 0,
                    ["tasks_remaining"] = null
                }
            };
        }

        private SessionManager Build(FakeHandler handler, Func<DateTimeOffset> clock = null)
        {
            var vault = new TokenVault(_vaultDirectory);
            return new SessionManager(
                vault,
                url => new AuthClient(url, handler),
                clock);
        }

        // ------------------------------------------------------------------

        [Fact]
        public async Task SignIn_AppliesByokConfigFromServer()
        {
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("byok", null, null))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var manager = Build(handler);

            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            Assert.Equal(SessionState.SignedIn, manager.State);
            Assert.Equal(RoutingMode.Byok, manager.CurrentEndpoint.Mode);
            Assert.Null(manager.CurrentEndpoint.BaseUrl);
        }

        [Fact]
        public async Task SignIn_AppliesHostedConfigWithoutAnyClientSideRule()
        {
            // The same client build, the same call. Only the server's answer
            // differs -- which is exactly what makes the M6 switch a server
            // change rather than a release.
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK,
                        EndpointBody("hosted", "https://proxy.example.com/v1", "Bearer proxy-token"))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var manager = Build(handler);

            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            Assert.Equal(RoutingMode.Hosted, manager.CurrentEndpoint.Mode);
            Assert.Equal("https://proxy.example.com/v1", manager.CurrentEndpoint.BaseUrl);
            Assert.Equal("Bearer proxy-token", manager.CurrentEndpoint.AuthHeader);
        }

        [Fact]
        public async Task IncompleteHostedConfigIsRejectedRatherThanApplied()
        {
            // A hosted config with no base_url would point the sidecar at null.
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("hosted", null, null))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var manager = Build(handler);

            var error = await Assert.ThrowsAsync<AuthException>(
                () => manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234"));
            Assert.Equal(AuthFailure.ServerError, error.Failure);
            Assert.Null(manager.CurrentEndpoint);
        }

        [Fact]
        public void UnknownFutureModeFallsBackToByokInsteadOfBreaking()
        {
            // An older client must keep working against a newer server.
            var config = new EndpointConfig { ModeRaw = "some-future-mode" };
            Assert.Equal(RoutingMode.Byok, config.Mode);
            Assert.True(config.IsUsable());
        }

        [Fact]
        public async Task CachedEndpointIsReusedUntilItNearlyExpires()
        {
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("byok", null, null, ttl: 900))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var manager = Build(handler);
            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            int before = handler.Requests.Count;
            for (int i = 0; i < 5; i++)
            {
                await manager.GetEndpointAsync();
            }

            // Re-fetching on every model call would put the account server in
            // the hot path of every user action.
            Assert.Equal(before, handler.Requests.Count);
        }

        [Fact]
        public async Task ExpiredEndpointIsRefetched()
        {
            var now = DateTimeOffset.UtcNow;
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("byok", null, null, ttl: 900))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var manager = Build(handler, () => now);
            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            int before = handler.Requests.Count;
            now = now.AddSeconds(1000); // past the cached expiry
            await manager.GetEndpointAsync();

            Assert.True(handler.Requests.Count > before,
                "an expired endpoint config must be re-fetched: that expiry is the server's kill switch");
        }

        [Fact]
        public async Task NetworkFailureKeepsWorkingInsideTheGraceWindow()
        {
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("byok", null, null, ttl: 60))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var now = DateTimeOffset.UtcNow;
            var manager = Build(handler, () => now);
            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            handler.Offline = true;
            now = now.AddSeconds(600); // cached config has expired

            EndpointConfig config = await manager.GetEndpointAsync();

            // A server hiccup must not stop someone editing their own workbook
            // with their own API key.
            Assert.NotNull(config);
            Assert.Equal(SessionState.Offline, manager.State);
        }

        [Fact]
        public async Task RejectedSessionClearsStoredCredentials()
        {
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("byok", null, null, ttl: 60))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var now = DateTimeOffset.UtcNow;
            var manager = Build(handler, () => now);
            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            // The server now rejects the session, e.g. the account was disabled.
            handler.Responder = _ => Json(HttpStatusCode.Unauthorized,
                new Dictionary<string, object> { ["detail"] = "Refresh token is not valid" });
            now = now.AddSeconds(600);

            await Assert.ThrowsAsync<AuthException>(() => manager.GetEndpointAsync());
            Assert.Equal(SessionState.Expired, manager.State);

            // A rejection is not a network blip: the stored token must go, or
            // every restart retries a token the server has already revoked.
            TokenVault.PersistedSession stored;
            Assert.False(new TokenVault(_vaultDirectory).TryLoad(out stored));
        }

        [Fact]
        public async Task RotatedRefreshTokenIsPersisted()
        {
            // The server rotates refresh tokens. Persisting the old one would
            // make the next restore present an already-revoked token.
            string issued = "refresh-1";
            var handler = new FakeHandler();
            handler.Responder = request =>
            {
                if (request.RequestUri.AbsolutePath.EndsWith("/endpoint"))
                {
                    return Json(HttpStatusCode.OK, EndpointBody("byok", null, null, ttl: 60));
                }
                if (request.RequestUri.AbsolutePath.EndsWith("/refresh"))
                {
                    issued = "refresh-2";
                }
                return Json(HttpStatusCode.OK, TokenPairBody(refresh: issued));
            };

            var now = DateTimeOffset.UtcNow;
            var manager = Build(handler, () => now);
            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            now = now.AddSeconds(3600); // force both token and endpoint renewal
            await manager.GetEndpointAsync();

            TokenVault.PersistedSession stored;
            Assert.True(new TokenVault(_vaultDirectory).TryLoad(out stored));
            Assert.Equal("refresh-2", stored.RefreshToken);
        }

        [Fact]
        public async Task RestoreWithoutStoredSessionSignsOut()
        {
            var handler = new FakeHandler { Responder = _ => Json(HttpStatusCode.OK, TokenPairBody()) };
            var manager = Build(handler);

            Assert.Equal(SessionState.SignedOut, await manager.RestoreAsync());
        }

        [Fact]
        public async Task SignOutClearsEverything()
        {
            var handler = new FakeHandler
            {
                Responder = request => request.RequestUri.AbsolutePath.EndsWith("/endpoint")
                    ? Json(HttpStatusCode.OK, EndpointBody("byok", null, null))
                    : Json(HttpStatusCode.OK, TokenPairBody())
            };
            var manager = Build(handler);
            await manager.SignInAsync("https://api.example.com", "a@b.com", "password-1234");

            await manager.SignOutAsync();

            Assert.Equal(SessionState.SignedOut, manager.State);
            Assert.Null(manager.CurrentEndpoint);
            TokenVault.PersistedSession stored;
            Assert.False(new TokenVault(_vaultDirectory).TryLoad(out stored));
        }

        [Fact]
        public async Task NoEndpointIsReturnedWhenSignedOut()
        {
            var handler = new FakeHandler { Responder = _ => Json(HttpStatusCode.OK, TokenPairBody()) };
            var manager = Build(handler);

            // Callers must treat null as "not ready", never as "use BYOK".
            Assert.Null(await manager.GetEndpointAsync());
        }

        // ------------------------------------------------------------------
        // Token storage
        // ------------------------------------------------------------------

        [Fact]
        public void VaultRoundTripsAndDoesNotStorePlaintext()
        {
            var vault = new TokenVault(_vaultDirectory);
            Assert.True(vault.Save(new TokenVault.PersistedSession
            {
                ServerUrl = "https://api.example.com",
                RefreshToken = "super-secret-refresh-token",
                Email = "a@b.com",
                LastVerifiedUtc = DateTime.UtcNow
            }));

            byte[] raw = File.ReadAllBytes(vault.Path_ForTests);
            string asText = Encoding.UTF8.GetString(raw);
            Assert.DoesNotContain("super-secret-refresh-token", asText);

            TokenVault.PersistedSession loaded;
            Assert.True(vault.TryLoad(out loaded));
            Assert.Equal("super-secret-refresh-token", loaded.RefreshToken);
        }

        [Fact]
        public void TamperedVaultFailsClosed()
        {
            var vault = new TokenVault(_vaultDirectory);
            vault.Save(new TokenVault.PersistedSession
            {
                ServerUrl = "https://api.example.com",
                RefreshToken = "original-token",
                Email = "a@b.com",
                LastVerifiedUtc = DateTime.UtcNow
            });

            // Replacing the file with plaintext JSON must not be a way to inject
            // a session -- that is the failure mode that makes DPAPI pointless.
            File.WriteAllText(vault.Path_ForTests,
                "{\"refresh_token\":\"attacker-token\",\"server_url\":\"https://evil.example\"}");

            TokenVault.PersistedSession loaded;
            Assert.False(vault.TryLoad(out loaded));
            Assert.Null(loaded);
        }

        [Fact]
        public void EmptyTokenIsNotPersisted()
        {
            var vault = new TokenVault(_vaultDirectory);
            Assert.False(vault.Save(new TokenVault.PersistedSession { RefreshToken = "" }));
            Assert.False(File.Exists(vault.Path_ForTests));
        }
    }
}
