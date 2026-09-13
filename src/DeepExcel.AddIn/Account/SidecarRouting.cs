using System;

namespace DeepExcel.AddIn.Account
{
    /// <summary>
    /// The concrete values handed to the sidecar for one session.
    /// </summary>
    public sealed class SidecarRouting
    {
        public string BaseUrl { get; set; }
        public string Model { get; set; }

        /// <summary>BYOK only. Sent as ANTHROPIC_API_KEY.</summary>
        public string ApiKey { get; set; }

        /// <summary>
        /// Hosted only. The bare token with no scheme prefix, sent as
        /// ANTHROPIC_AUTH_TOKEN so the SDK emits `Authorization: Bearer ...`.
        /// </summary>
        public string AuthToken { get; set; }

        /// <summary>"byok" or "hosted"; carried through for diagnostics.</summary>
        public string Mode { get; set; }

        public bool IsHosted
        {
            get { return string.Equals(Mode, "hosted", StringComparison.OrdinalIgnoreCase); }
        }
    }

    /// <summary>
    /// Turns the server's routing decision plus the local provider settings into
    /// the values the sidecar needs.
    ///
    /// This is the ONLY place in the client where BYOK and hosted differ. Keeping
    /// it in one pure function is what makes the M6 switch reviewable: there is a
    /// single function to read, and a test for every branch of it.
    /// </summary>
    public static class SidecarRoutingResolver
    {
        /// <param name="endpoint">
        /// The server's answer, or null when no account server is configured.
        /// Null means "local-only install", which must keep working exactly as
        /// it did before accounts existed.
        /// </param>
        /// <param name="localBaseUrl">Provider base URL from config.json.</param>
        /// <param name="localModel">Model the user selected.</param>
        /// <param name="localApiKey">DPAPI-decrypted provider key.</param>
        public static SidecarRouting Resolve(
            EndpointConfig endpoint, string localBaseUrl, string localModel, string localApiKey)
        {
            // The model is always the user's choice, in both modes. The proxy
            // validates it; it is not the server's job to pick it.
            var model = localModel;

            if (endpoint == null || endpoint.Mode == RoutingMode.Byok)
            {
                return new SidecarRouting
                {
                    Mode = "byok",
                    BaseUrl = string.IsNullOrEmpty(localBaseUrl) ? "https://api.anthropic.com" : localBaseUrl,
                    Model = model,
                    ApiKey = localApiKey ?? "",
                    AuthToken = null
                };
            }

            // Hosted. Refuse rather than fall back: silently using the local key
            // against a hosted account would bill the user's own provider
            // account without telling them, and hosted users usually have no
            // local key at all.
            if (!endpoint.IsUsable())
            {
                throw new AuthException(
                    AuthFailure.ServerError,
                    "服务器要求使用托管转发，但未提供转发地址或凭据。");
            }

            return new SidecarRouting
            {
                Mode = "hosted",
                BaseUrl = endpoint.BaseUrl,
                Model = model,
                // The proxy authenticates the request; the user's own key must
                // never be sent to it.
                ApiKey = null,
                AuthToken = StripScheme(endpoint.AuthHeader)
            };
        }

        /// <summary>
        /// The server sends a complete header value ("Bearer abc"), but the SDK
        /// wants the bare token and adds the scheme itself. Sending the full
        /// value would produce "Authorization: Bearer Bearer abc".
        /// </summary>
        internal static string StripScheme(string authHeader)
        {
            if (string.IsNullOrEmpty(authHeader))
            {
                return null;
            }
            var trimmed = authHeader.Trim();
            const string bearer = "Bearer ";
            if (trimmed.StartsWith(bearer, StringComparison.OrdinalIgnoreCase))
            {
                return trimmed.Substring(bearer.Length).Trim();
            }
            return trimmed;
        }
    }
}
