using System;
using System.Threading.Tasks;
using DeepExcel.AddIn.Account;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Account-related bridge messages.
    ///
    /// Kept in its own file because the account layer is optional: an install
    /// with no server configured never reaches any of this, and the rest of
    /// MessageBridge should not have to know that accounts exist.
    /// </summary>
    public partial class MessageBridge
    {
        /// <summary>
        /// Blocks the UI thread on an async call.
        ///
        /// The bridge protocol is synchronous request/response, and the panel is
        /// already showing a spinner while it waits. AuthClient has a 20-second
        /// timeout, so this cannot hang indefinitely.
        /// </summary>
        private static T RunSync<T>(Func<Task<T>> work)
        {
            return Task.Run(work).GetAwaiter().GetResult();
        }

        private static void RunSync(Func<Task> work)
        {
            Task.Run(work).GetAwaiter().GetResult();
        }

        private SessionManager RequireSession()
        {
            if (AccountSession == null)
            {
                AccountSession = new SessionManager();
            }
            return AccountSession;
        }

        /// <summary>Reads an optional string field, tolerating a missing payload.</summary>
        private static string ReadString(Message msg, string name)
        {
            if (msg?.Payload == null)
            {
                return null;
            }
            if (!msg.Payload.Value.TryGetProperty(name, out var element))
            {
                return null;
            }
            return element.ValueKind == System.Text.Json.JsonValueKind.String
                ? element.GetString()
                : null;
        }

        /// <summary>
        /// Pushes the current routing to every open workbook's sidecar.
        ///
        /// Sign-in, sign-out and session restore all change where model traffic
        /// should go. Without this the already-running sidecars keep using the
        /// previous routing -- after sign-out that means a revoked proxy token.
        /// </summary>
        private void ResendConfigToAllSessions()
        {
            try
            {
                foreach (var kvp in _sessions)
                {
                    if (kvp.Value?.Sidecar != null)
                    {
                        SendConfigToSession(kvp.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "ResendConfigToAllSessions failed: " + ex.Message);
            }
        }

        /// <summary>Current sign-in state; the panel renders from this.</summary>
        private string HandleAccountStatus()
        {
            try
            {
                var session = AccountSession;
                if (session == null)
                {
                    return MakeResponse("account_status", new
                    {
                        state = "signed_out",
                        server_url = (string)null,
                        email = (string)null,
                        mode = (string)null,
                        entitlement = (object)null
                    });
                }

                var endpoint = session.CurrentEndpoint;
                return MakeResponse("account_status", new
                {
                    state = session.State.ToString().ToLowerInvariant(),
                    server_url = session.ServerUrl,
                    email = session.Email,
                    mode = endpoint == null ? null : endpoint.ModeRaw,
                    entitlement = endpoint?.Entitlement
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleAccountStatus failed", ex);
                return MakeError("读取登录状态失败");
            }
        }

        /// <summary>
        /// Whether this deployment requires an invite code, so the panel can
        /// render the right form without shipping a build per policy.
        /// </summary>
        private string HandleAccountServerMeta(Message msg)
        {
            var serverUrl = ReadString(msg, "server_url");
            if (string.IsNullOrWhiteSpace(serverUrl))
            {
                return MakeError("请填写服务器地址");
            }

            try
            {
                using (var client = new AuthClient(serverUrl))
                {
                    var meta = RunSync(() => client.GetMetaAsync());
                    return MakeResponse("account_server_meta", new
                    {
                        invite_required = meta?.InviteRequired ?? true,
                        hosted_routing_available = meta?.HostedRoutingAvailable ?? false,
                        telemetry_enabled = meta?.TelemetryEnabled ?? false,
                        min_client_version = meta?.MinClientVersion
                    });
                }
            }
            catch (AuthException ex)
            {
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleAccountServerMeta failed", ex);
                return MakeError("无法连接该服务器地址");
            }
        }

        private string HandleAccountSignIn(Message msg)
        {
            var serverUrl = ReadString(msg, "server_url");
            var email = ReadString(msg, "email");
            var password = ReadString(msg, "password");
            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(email) ||
                string.IsNullOrEmpty(password))
            {
                return MakeError("请填写服务器地址、邮箱和密码");
            }

            try
            {
                var session = RequireSession();
                RunSync(() => session.SignInAsync(serverUrl, email, password));
                // The endpoint may have changed, so every open workbook needs
                // the new routing before its next request.
                ResendConfigToAllSessions();
                return HandleAccountStatus();
            }
            catch (AuthException ex)
            {
                Logger.Instance.Warning("MessageBridge", "Sign-in failed: " + ex.Failure);
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleAccountSignIn failed", ex);
                return MakeError("登录失败，请稍后重试");
            }
        }

        private string HandleAccountRegister(Message msg)
        {
            var serverUrl = ReadString(msg, "server_url");
            var email = ReadString(msg, "email");
            var password = ReadString(msg, "password");
            var inviteCode = ReadString(msg, "invite_code");
            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(email) ||
                string.IsNullOrEmpty(password))
            {
                return MakeError("请填写服务器地址、邮箱和密码");
            }

            try
            {
                var session = RequireSession();
                RunSync(() => session.RegisterAsync(serverUrl, email, password, inviteCode));
                ResendConfigToAllSessions();
                return HandleAccountStatus();
            }
            catch (AuthException ex)
            {
                Logger.Instance.Warning("MessageBridge", "Registration failed: " + ex.Failure);
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleAccountRegister failed", ex);
                return MakeError("注册失败，请稍后重试");
            }
        }

        /// <summary>
        /// Consumption for a hosted account.
        ///
        /// Read from the proxy, which is the authoritative counter -- the
        /// client's own tally is advisory and could be stale or spoofed.
        /// </summary>
        private string HandleAccountUsage()
        {
            try
            {
                var session = AccountSession;
                var endpoint = session?.CurrentEndpoint;
                if (session == null || endpoint == null)
                {
                    return MakeError("请先登录账号");
                }
                if (endpoint.Mode != Account.RoutingMode.Hosted || string.IsNullOrEmpty(endpoint.BaseUrl))
                {
                    // A BYOK account's traffic never reaches us, so there is
                    // nothing for us to report.
                    return MakeResponse("account_usage", new { applicable = false });
                }

                var usage = RunSync(async () =>
                {
                    using (var http = new System.Net.Http.HttpClient
                    {
                        Timeout = TimeSpan.FromSeconds(15)
                    })
                    {
                        var request = new System.Net.Http.HttpRequestMessage(
                            System.Net.Http.HttpMethod.Get,
                            endpoint.BaseUrl.TrimEnd('/') + "/usage");
                        request.Headers.TryAddWithoutValidation("Authorization", endpoint.AuthHeader);
                        var response = await http.SendAsync(request).ConfigureAwait(false);
                        response.EnsureSuccessStatusCode();
                        return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    }
                });

                using (var document = System.Text.Json.JsonDocument.Parse(usage))
                {
                    var root = document.RootElement;
                    return MakeResponse("account_usage", new
                    {
                        applicable = true,
                        period_days = root.GetProperty("period_days").GetInt32(),
                        calls = root.GetProperty("calls").GetInt32(),
                        input_tokens = root.GetProperty("input_tokens").GetInt32(),
                        output_tokens = root.GetProperty("output_tokens").GetInt32(),
                        tasks_used = root.GetProperty("tasks_used").GetInt32()
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "HandleAccountUsage failed: " + ex.Message);
                return MakeError("无法获取用量");
            }
        }

        private string HandleAccountSignOut()
        {
            try
            {
                if (AccountSession != null)
                {
                    RunSync(() => AccountSession.SignOutAsync());
                    // Back to the local provider configuration. Without this the
                    // sidecar would keep using a proxy token that is now revoked.
                    ResendConfigToAllSessions();
                }
                return HandleAccountStatus();
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleAccountSignOut failed", ex);
                return MakeError("退出登录失败");
            }
        }

        /// <summary>
        /// Restores a stored session at add-in startup. Never throws: a failure
        /// here must leave the user signed out, not break the panel.
        /// </summary>
        public void RestoreAccountSessionAsync()
        {
            Task.Run(async () =>
            {
                try
                {
                    var session = RequireSession();
                    var state = await session.RestoreAsync().ConfigureAwait(false);
                    Logger.Instance.Info("MessageBridge", "Account session restored: " + state);
                    if (state == SessionState.SignedIn || state == SessionState.Offline)
                    {
                        ResendConfigToAllSessions();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Instance.Warning("MessageBridge", "Account restore failed: " + ex.Message);
                }
            });
        }
    }
}
