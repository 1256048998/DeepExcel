using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace DeepExcel.AddIn.Account
{
    /// <summary>
    /// HTTP client for the DeepExcel account server.
    ///
    /// Translates transport and status codes into <see cref="AuthException"/> so
    /// callers never have to reason about HTTP. The distinction that matters
    /// most is network failure versus rejection: the first is recoverable and
    /// must not sign the user out, the second must.
    /// </summary>
    public sealed class AuthClient : IDisposable
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        private readonly HttpClient _http;

        public AuthClient(string serverUrl, HttpMessageHandler handler = null)
        {
            if (string.IsNullOrWhiteSpace(serverUrl))
            {
                throw new ArgumentException("serverUrl is required", nameof(serverUrl));
            }

            ServerUrl = serverUrl.TrimEnd('/');
            _http = handler == null ? new HttpClient() : new HttpClient(handler, disposeHandler: true);
            _http.BaseAddress = new Uri(ServerUrl + "/");
            // Long enough for a cold container, short enough that a hung server
            // does not freeze the task pane.
            _http.Timeout = TimeSpan.FromSeconds(20);
        }

        public string ServerUrl { get; }

        public Task<ServerMeta> GetMetaAsync(CancellationToken cancellationToken = default)
        {
            return SendAsync<ServerMeta>(HttpMethod.Get, "api/v1/meta", null, null, cancellationToken);
        }

        public Task<TokenPair> RegisterAsync(
            string email, string password, string inviteCode = null, string displayName = null,
            CancellationToken cancellationToken = default)
        {
            var body = new Dictionary<string, object>
            {
                ["email"] = email,
                ["password"] = password
            };
            if (!string.IsNullOrWhiteSpace(inviteCode)) body["invite_code"] = inviteCode.Trim();
            if (!string.IsNullOrWhiteSpace(displayName)) body["display_name"] = displayName;
            return SendAsync<TokenPair>(HttpMethod.Post, "api/v1/auth/register", body, null, cancellationToken);
        }

        public Task<TokenPair> LoginAsync(
            string email, string password, CancellationToken cancellationToken = default)
        {
            var body = new Dictionary<string, object> { ["email"] = email, ["password"] = password };
            return SendAsync<TokenPair>(HttpMethod.Post, "api/v1/auth/login", body, null, cancellationToken);
        }

        public Task<TokenPair> RefreshAsync(string refreshToken, CancellationToken cancellationToken = default)
        {
            var body = new Dictionary<string, object> { ["refresh_token"] = refreshToken };
            return SendAsync<TokenPair>(HttpMethod.Post, "api/v1/auth/refresh", body, null, cancellationToken);
        }

        public async Task LogoutAsync(string refreshToken, CancellationToken cancellationToken = default)
        {
            var body = new Dictionary<string, object> { ["refresh_token"] = refreshToken };
            try
            {
                await SendAsync<object>(HttpMethod.Post, "api/v1/auth/logout", body, null, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (AuthException)
            {
                // Signing out locally must succeed even when the server cannot
                // be reached; the token expires on its own.
            }
        }

        public Task<AccountInfo> GetAccountAsync(string accessToken, CancellationToken cancellationToken = default)
        {
            return SendAsync<AccountInfo>(HttpMethod.Get, "api/v1/auth/me", null, accessToken, cancellationToken);
        }

        public Task<EndpointConfig> GetEndpointAsync(string accessToken, CancellationToken cancellationToken = default)
        {
            return SendAsync<EndpointConfig>(
                HttpMethod.Get, "api/v1/session/endpoint", null, accessToken, cancellationToken);
        }

        public Task SendHeartbeatAsync(
            string accessToken, string installId, string clientVersion, string host,
            CancellationToken cancellationToken = default)
        {
            string query = string.Format(
                "api/v1/session/heartbeat?install_id={0}&client_version={1}&host={2}",
                Uri.EscapeDataString(installId ?? ""),
                Uri.EscapeDataString(clientVersion ?? ""),
                Uri.EscapeDataString(host ?? ""));
            return SendAsync<object>(HttpMethod.Post, query, null, accessToken, cancellationToken);
        }

        public Task SendTelemetryAsync(
            string accessToken, object batch, CancellationToken cancellationToken = default)
        {
            return SendAsync<object>(HttpMethod.Post, "api/v1/telemetry", batch, accessToken, cancellationToken);
        }

        // ------------------------------------------------------------------

        private async Task<T> SendAsync<T>(
            HttpMethod method, string path, object body, string accessToken,
            CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(method, path))
            {
                if (body != null)
                {
                    request.Content = new StringContent(
                        JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
                }
                if (!string.IsNullOrEmpty(accessToken))
                {
                    request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
                }

                HttpResponseMessage response;
                try
                {
                    response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
                }
                catch (TaskCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    throw new AuthException(AuthFailure.Network, "连接服务器超时。");
                }
                catch (HttpRequestException ex)
                {
                    throw new AuthException(AuthFailure.Network, "无法连接服务器：" + ex.Message);
                }

                using (response)
                {
                    string payload = response.Content == null
                        ? ""
                        : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (response.IsSuccessStatusCode)
                    {
                        if (typeof(T) == typeof(object) || string.IsNullOrWhiteSpace(payload))
                        {
                            return default;
                        }
                        try
                        {
                            return JsonSerializer.Deserialize<T>(payload, JsonOptions);
                        }
                        catch (JsonException ex)
                        {
                            throw new AuthException(
                                AuthFailure.ServerError, "服务器返回了无法解析的响应：" + ex.Message);
                        }
                    }

                    throw Translate(response.StatusCode, payload, path);
                }
            }
        }

        /// <summary>Maps a status code onto something the UI can act on.</summary>
        private static AuthException Translate(HttpStatusCode status, string payload, string path)
        {
            string detail = ExtractDetail(payload, out string reason);

            switch (status)
            {
                case HttpStatusCode.Unauthorized:
                    // On the login path this is a wrong password; anywhere else
                    // the stored session is no longer usable.
                    return path.EndsWith("login", StringComparison.OrdinalIgnoreCase)
                        ? new AuthException(AuthFailure.InvalidCredentials, "邮箱或密码不正确。")
                        : new AuthException(AuthFailure.SessionExpired, "登录状态已失效，请重新登录。");

                case HttpStatusCode.Forbidden:
                    if (!string.IsNullOrEmpty(detail) &&
                        detail.IndexOf("Invite", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return new AuthException(AuthFailure.InviteRequired, "需要有效的邀请码。");
                    }
                    return new AuthException(AuthFailure.AccountDisabled, "账号已被停用。");

                case HttpStatusCode.Conflict:
                    return new AuthException(AuthFailure.EmailAlreadyRegistered, "该邮箱已注册。");

                case HttpStatusCode.PaymentRequired:
                    return new AuthException(
                        AuthFailure.EntitlementProblem,
                        string.IsNullOrEmpty(detail) ? "订阅状态异常。" : detail,
                        reason);

                case HttpStatusCode.ServiceUnavailable:
                    // The server said hosted routing is configured for this
                    // account but unavailable. Silently using BYOK here would
                    // hide a deployment fault behind provider errors the user
                    // cannot diagnose.
                    return new AuthException(
                        AuthFailure.ServerError,
                        string.IsNullOrEmpty(detail) ? "服务暂时不可用。" : detail,
                        reason);

                default:
                    return new AuthException(
                        AuthFailure.ServerError,
                        string.Format("服务器错误（{0}）{1}", (int)status,
                            string.IsNullOrEmpty(detail) ? "" : "：" + detail));
            }
        }

        /// <summary>
        /// FastAPI's `detail` is a string for simple errors and an object for the
        /// structured ones this server raises.
        /// </summary>
        private static string ExtractDetail(string payload, out string reason)
        {
            reason = null;
            if (string.IsNullOrWhiteSpace(payload))
            {
                return null;
            }
            try
            {
                using (var document = JsonDocument.Parse(payload))
                {
                    if (!document.RootElement.TryGetProperty("detail", out var detail))
                    {
                        return null;
                    }
                    if (detail.ValueKind == JsonValueKind.String)
                    {
                        return detail.GetString();
                    }
                    if (detail.ValueKind == JsonValueKind.Object)
                    {
                        if (detail.TryGetProperty("reason", out var reasonElement))
                        {
                            reason = reasonElement.GetString();
                        }
                        if (detail.TryGetProperty("message", out var messageElement))
                        {
                            return messageElement.GetString();
                        }
                        return reason;
                    }
                }
            }
            catch (JsonException)
            {
            }
            return null;
        }

        public void Dispose()
        {
            _http.Dispose();
        }
    }
}
