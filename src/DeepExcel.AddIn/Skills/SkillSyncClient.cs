using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Account;

namespace DeepExcel.AddIn.Skills
{
    public sealed class RemoteSkillSummary
    {
        [JsonPropertyName("skill_id")]
        public string SkillId { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("updated_at")]
        public string UpdatedAt { get; set; }

        /// <summary>Non-null once the owner has explicitly shared it.</summary>
        [JsonPropertyName("share_code")]
        public string ShareCode { get; set; }
    }

    /// <summary>
    /// Syncs the local skill library with the account server.
    ///
    /// Everything is scrubbed by <see cref="SkillScrubber"/> before it leaves.
    /// The server scrubs again -- doing it here means the data never leaves the
    /// machine, doing it there means a modified client cannot skip it.
    /// </summary>
    public sealed class SkillSyncClient
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        private readonly SessionManager _session;

        public SkillSyncClient(SessionManager session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        private async Task<HttpClient> CreateClientAsync(CancellationToken cancellationToken)
        {
            // Asking for the endpoint first is what renews the access token; the
            // renewal logic lives in SessionManager and is not duplicated here.
            await _session.GetEndpointAsync(cancellationToken).ConfigureAwait(false);
            var token = _session.CurrentAccessToken;
            if (string.IsNullOrEmpty(token) || string.IsNullOrEmpty(_session.ServerUrl))
            {
                throw new AuthException(AuthFailure.SessionExpired, "请先登录账号。");
            }

            var client = new HttpClient { BaseAddress = new Uri(_session.ServerUrl.TrimEnd('/') + "/") };
            client.Timeout = TimeSpan.FromSeconds(20);
            client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", "Bearer " + token);
            return client;
        }

        public async Task<RemoteSkillSummary> UploadAsync(
            Skill skill, CancellationToken cancellationToken = default)
        {
            if (skill == null)
            {
                throw new ArgumentNullException(nameof(skill));
            }

            // The local copy keeps its working defaults; only the copy that
            // travels is stripped.
            var safe = SkillScrubber.ForUpload(skill);
            var body = new Dictionary<string, object>
            {
                ["skill_id"] = safe.Id,
                ["name"] = safe.Name,
                ["document"] = safe
            };

            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            using (var content = new StringContent(
                JsonSerializer.Serialize(body, JsonOptions), Encoding.UTF8, "application/json"))
            {
                var response = await client.PutAsync("api/v1/skills", content, cancellationToken)
                    .ConfigureAwait(false);
                return await ReadAsync<RemoteSkillSummary>(response).ConfigureAwait(false);
            }
        }

        public async Task<List<RemoteSkillSummary>> ListAsync(
            CancellationToken cancellationToken = default)
        {
            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            {
                var response = await client.GetAsync("api/v1/skills", cancellationToken)
                    .ConfigureAwait(false);
                return await ReadAsync<List<RemoteSkillSummary>>(response).ConfigureAwait(false)
                       ?? new List<RemoteSkillSummary>();
            }
        }

        public async Task<Skill> DownloadAsync(
            string skillId, CancellationToken cancellationToken = default)
        {
            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            {
                var response = await client
                    .GetAsync("api/v1/skills/" + Uri.EscapeDataString(skillId), cancellationToken)
                    .ConfigureAwait(false);
                return await ReadAsync<Skill>(response).ConfigureAwait(false);
            }
        }

        public async Task DeleteAsync(string skillId, CancellationToken cancellationToken = default)
        {
            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            {
                var response = await client
                    .DeleteAsync("api/v1/skills/" + Uri.EscapeDataString(skillId), cancellationToken)
                    .ConfigureAwait(false);
                await ReadAsync<object>(response).ConfigureAwait(false);
            }
        }

        /// <summary>Publishes a skill and returns its share code.</summary>
        public async Task<string> ShareAsync(string skillId, CancellationToken cancellationToken = default)
        {
            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            {
                var response = await client.PostAsync(
                    "api/v1/skills/" + Uri.EscapeDataString(skillId) + "/share", null, cancellationToken)
                    .ConfigureAwait(false);
                var payload = await ReadAsync<Dictionary<string, string>>(response).ConfigureAwait(false);
                return payload != null && payload.TryGetValue("share_code", out var code) ? code : null;
            }
        }

        public async Task UnshareAsync(string skillId, CancellationToken cancellationToken = default)
        {
            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            {
                var response = await client.PostAsync(
                    "api/v1/skills/" + Uri.EscapeDataString(skillId) + "/unshare", null, cancellationToken)
                    .ConfigureAwait(false);
                await ReadAsync<object>(response).ConfigureAwait(false);
            }
        }

        /// <summary>Imports someone else's shared skill.</summary>
        public async Task<Skill> ImportAsync(string shareCode, CancellationToken cancellationToken = default)
        {
            using (var client = await CreateClientAsync(cancellationToken).ConfigureAwait(false))
            {
                var response = await client.GetAsync(
                    "api/v1/skills/shared/" + Uri.EscapeDataString(shareCode), cancellationToken)
                    .ConfigureAwait(false);
                return await ReadAsync<Skill>(response).ConfigureAwait(false);
            }
        }

        private static async Task<T> ReadAsync<T>(HttpResponseMessage response)
        {
            using (response)
            {
                var payload = response.Content == null
                    ? ""
                    : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    throw new AuthException(
                        response.StatusCode == System.Net.HttpStatusCode.Unauthorized
                            ? AuthFailure.SessionExpired
                            : AuthFailure.ServerError,
                        ExtractDetail(payload) ?? ("同步失败（" + (int)response.StatusCode + "）"));
                }

                if (typeof(T) == typeof(object) || string.IsNullOrWhiteSpace(payload))
                {
                    return default;
                }
                return JsonSerializer.Deserialize<T>(payload, JsonOptions);
            }
        }

        private static string ExtractDetail(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
            {
                return null;
            }
            try
            {
                using (var document = JsonDocument.Parse(payload))
                {
                    if (document.RootElement.TryGetProperty("detail", out var detail))
                    {
                        return detail.ValueKind == JsonValueKind.String
                            ? detail.GetString()
                            : detail.ToString();
                    }
                }
            }
            catch (JsonException)
            {
            }
            return null;
        }
    }
}
