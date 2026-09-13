using System;
using DeepExcel.AddIn.Account;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The single place in the client where BYOK and hosted differ.
    ///
    /// The properties worth guarding are about credentials going to the wrong
    /// party: the user's provider key must never reach the DeepExcel proxy, and
    /// the proxy token must never reach a third-party provider.
    /// </summary>
    public class SidecarRoutingTests
    {
        private static EndpointConfig Byok()
        {
            return new EndpointConfig { ModeRaw = "byok" };
        }

        private static EndpointConfig Hosted(string baseUrl = "https://api.deepexcel.com/v1",
                                             string authHeader = "Bearer proxy-token-abc")
        {
            return new EndpointConfig { ModeRaw = "hosted", BaseUrl = baseUrl, AuthHeader = authHeader };
        }

        [Fact]
        public void NoAccountServerBehavesExactlyAsBeforeAccountsExisted()
        {
            // A local-only install is a supported configuration, not an error.
            var routing = SidecarRoutingResolver.Resolve(
                null, "https://api.deepseek.com/anthropic", "deepseek-v4", "sk-local-key");

            Assert.Equal("byok", routing.Mode);
            Assert.Equal("https://api.deepseek.com/anthropic", routing.BaseUrl);
            Assert.Equal("deepseek-v4", routing.Model);
            Assert.Equal("sk-local-key", routing.ApiKey);
            Assert.Null(routing.AuthToken);
        }

        [Fact]
        public void ByokUsesTheLocalProviderAndKey()
        {
            var routing = SidecarRoutingResolver.Resolve(
                Byok(), "https://api.moonshot.cn/anthropic", "kimi-k2", "sk-user-key");

            Assert.Equal("byok", routing.Mode);
            Assert.Equal("https://api.moonshot.cn/anthropic", routing.BaseUrl);
            Assert.Equal("sk-user-key", routing.ApiKey);
            Assert.Null(routing.AuthToken);
        }

        [Fact]
        public void HostedNeverForwardsTheUsersProviderKey()
        {
            // Sending the user's own key to our proxy would expose it to one
            // more party and bypass metering at the same time.
            var routing = SidecarRoutingResolver.Resolve(
                Hosted(), "https://api.anthropic.com", "claude-opus-5", "sk-user-secret-key");

            Assert.Equal("hosted", routing.Mode);
            Assert.Null(routing.ApiKey);
            Assert.Equal("proxy-token-abc", routing.AuthToken);
            Assert.Equal("https://api.deepexcel.com/v1", routing.BaseUrl);
        }

        [Fact]
        public void HostedIgnoresTheLocalProviderBaseUrl()
        {
            // Otherwise a stale local provider setting would send the proxy
            // token to a third-party provider.
            var routing = SidecarRoutingResolver.Resolve(
                Hosted(), "https://api.deepseek.com/anthropic", "claude-opus-5", "sk-user-key");

            Assert.Equal("https://api.deepexcel.com/v1", routing.BaseUrl);
        }

        [Fact]
        public void ModelIsAlwaysTheUsersChoiceInBothModes()
        {
            // Picking the model is a product decision that belongs to the user;
            // the proxy validates it rather than choosing it.
            const string chosen = "claude-opus-5";
            Assert.Equal(chosen, SidecarRoutingResolver.Resolve(Byok(), "https://x", chosen, "k").Model);
            Assert.Equal(chosen, SidecarRoutingResolver.Resolve(Hosted(), "https://x", chosen, "k").Model);
        }

        [Fact]
        public void SchemeIsStrippedSoTheHeaderIsNotDoubled()
        {
            // The SDK adds "Bearer " itself; passing the full header value would
            // produce "Authorization: Bearer Bearer abc".
            Assert.Equal("abc", SidecarRoutingResolver.StripScheme("Bearer abc"));
            Assert.Equal("abc", SidecarRoutingResolver.StripScheme("bearer abc"));
            Assert.Equal("abc", SidecarRoutingResolver.StripScheme("  Bearer   abc  "));
            // A server that sends a bare token still works.
            Assert.Equal("abc", SidecarRoutingResolver.StripScheme("abc"));
            Assert.Null(SidecarRoutingResolver.StripScheme(null));
            Assert.Null(SidecarRoutingResolver.StripScheme(""));
        }

        [Fact]
        public void IncompleteHostedConfigThrowsInsteadOfFallingBackToTheLocalKey()
        {
            // Falling back would silently bill the user's own provider account
            // for traffic they expected us to carry -- and hosted users usually
            // have no local key configured at all.
            foreach (var broken in new[]
            {
                Hosted(baseUrl: null),
                Hosted(authHeader: null),
                Hosted(baseUrl: "", authHeader: "")
            })
            {
                var error = Assert.Throws<AuthException>(
                    () => SidecarRoutingResolver.Resolve(broken, "https://api.anthropic.com", "m", "sk-key"));
                Assert.Equal(AuthFailure.ServerError, error.Failure);
            }
        }

        [Fact]
        public void UnknownFutureModeResolvesToByokRatherThanFailing()
        {
            // An older client must keep working against a newer server.
            var routing = SidecarRoutingResolver.Resolve(
                new EndpointConfig { ModeRaw = "some-future-mode" },
                "https://api.anthropic.com", "claude-opus-5", "sk-key");

            Assert.Equal("byok", routing.Mode);
            Assert.Equal("sk-key", routing.ApiKey);
        }

        [Fact]
        public void MissingLocalBaseUrlFallsBackToAnthropic()
        {
            var routing = SidecarRoutingResolver.Resolve(Byok(), null, "claude-opus-5", "sk-key");
            Assert.Equal("https://api.anthropic.com", routing.BaseUrl);
        }

        [Fact]
        public void IsHostedReflectsTheResolvedMode()
        {
            Assert.True(SidecarRoutingResolver.Resolve(Hosted(), "https://x", "m", "k").IsHosted);
            Assert.False(SidecarRoutingResolver.Resolve(Byok(), "https://x", "m", "k").IsHosted);
            Assert.False(SidecarRoutingResolver.Resolve(null, "https://x", "m", "k").IsHosted);
        }
    }
}
