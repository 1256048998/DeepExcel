using System;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Account
{
    /// <summary>
    /// Where the sidecar should send model traffic.
    ///
    /// This value is ALWAYS whatever the server returned. The client must never
    /// derive it from the plan, the account state, or anything else it knows
    /// locally: the add-in is unsigned and modifiable, so a client-side rule is
    /// not an enforcement mechanism, and hard-coding the decision here is what
    /// would force a full client release when hosted routing is switched on.
    /// </summary>
    public enum RoutingMode
    {
        /// <summary>Use the locally configured provider and API key.</summary>
        Byok,

        /// <summary>Route through the DeepExcel proxy using the supplied header.</summary>
        Hosted
    }

    public sealed class TokenPair
    {
        [JsonPropertyName("access_token")]
        public string AccessToken { get; set; }

        [JsonPropertyName("refresh_token")]
        public string RefreshToken { get; set; }

        [JsonPropertyName("expires_at")]
        public long ExpiresAtUnix { get; set; }

        [JsonIgnore]
        public DateTimeOffset ExpiresAt
        {
            get { return DateTimeOffset.FromUnixTimeSeconds(ExpiresAtUnix); }
        }
    }

    public sealed class EntitlementInfo
    {
        [JsonPropertyName("plan")]
        public string Plan { get; set; }

        [JsonPropertyName("status")]
        public string Status { get; set; }

        [JsonPropertyName("task_limit")]
        public int? TaskLimit { get; set; }

        [JsonPropertyName("tasks_used")]
        public int TasksUsed { get; set; }

        [JsonPropertyName("tasks_remaining")]
        public int? TasksRemaining { get; set; }
    }

    /// <summary>
    /// The server's answer to "where does model traffic go".
    ///
    /// Cached until <see cref="ExpiresAt"/> and then re-fetched. That expiry is
    /// the server's kill switch: access can be withdrawn within one interval
    /// without shipping anything.
    /// </summary>
    public sealed class EndpointConfig
    {
        [JsonPropertyName("mode")]
        public string ModeRaw { get; set; }

        [JsonPropertyName("base_url")]
        public string BaseUrl { get; set; }

        /// <summary>A complete header value, scheme included.</summary>
        [JsonPropertyName("auth_header")]
        public string AuthHeader { get; set; }

        [JsonPropertyName("expires_at")]
        public long ExpiresAtUnix { get; set; }

        [JsonPropertyName("entitlement")]
        public EntitlementInfo Entitlement { get; set; }

        [JsonPropertyName("refresh_after_seconds")]
        public int RefreshAfterSeconds { get; set; }

        [JsonIgnore]
        public DateTimeOffset ExpiresAt
        {
            get { return DateTimeOffset.FromUnixTimeSeconds(ExpiresAtUnix); }
        }

        [JsonIgnore]
        public RoutingMode Mode
        {
            get
            {
                // Anything unrecognised falls back to BYOK rather than throwing.
                // A future server that introduces a third mode must not brick
                // older clients; BYOK keeps them working with their own key.
                return string.Equals(ModeRaw, "hosted", StringComparison.OrdinalIgnoreCase)
                    ? RoutingMode.Hosted
                    : RoutingMode.Byok;
            }
        }

        /// <summary>
        /// A hosted config without routing details is unusable and must not be
        /// applied, or the sidecar would send requests to a null endpoint.
        /// </summary>
        public bool IsUsable()
        {
            if (Mode == RoutingMode.Byok)
            {
                return true;
            }
            return !string.IsNullOrEmpty(BaseUrl) && !string.IsNullOrEmpty(AuthHeader);
        }
    }

    public sealed class AccountInfo
    {
        [JsonPropertyName("id")]
        public long Id { get; set; }

        [JsonPropertyName("email")]
        public string Email { get; set; }

        [JsonPropertyName("display_name")]
        public string DisplayName { get; set; }

        [JsonPropertyName("status")]
        public string Status { get; set; }

        [JsonPropertyName("entitlement")]
        public EntitlementInfo Entitlement { get; set; }
    }

    public sealed class ServerMeta
    {
        [JsonPropertyName("invite_required")]
        public bool InviteRequired { get; set; }

        [JsonPropertyName("hosted_routing_available")]
        public bool HostedRoutingAvailable { get; set; }

        [JsonPropertyName("telemetry_enabled")]
        public bool TelemetryEnabled { get; set; }

        [JsonPropertyName("min_client_version")]
        public string MinClientVersion { get; set; }
    }

    /// <summary>Structured failure so the UI can say something specific.</summary>
    public sealed class AuthException : Exception
    {
        public AuthException(AuthFailure failure, string message, string reason = null)
            : base(message)
        {
            Failure = failure;
            Reason = reason;
        }

        public AuthFailure Failure { get; }

        /// <summary>Machine-readable reason from the server, when it supplied one.</summary>
        public string Reason { get; }
    }

    public enum AuthFailure
    {
        /// <summary>Server unreachable. Distinct from a rejection: offline is recoverable.</summary>
        Network,
        InvalidCredentials,
        /// <summary>Session is gone; the user has to sign in again.</summary>
        SessionExpired,
        AccountDisabled,
        /// <summary>Subscription lapsed or quota exhausted; carries a Reason.</summary>
        EntitlementProblem,
        InviteRequired,
        EmailAlreadyRegistered,
        ServerError
    }
}
