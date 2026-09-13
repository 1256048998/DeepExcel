using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Web.Script.Serialization;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>Why a manifest was not accepted. Every refusal has a distinct reason.</summary>
    public enum UpdateRejection
    {
        None = 0,
        /// <summary>No public key compiled in, so nothing can be trusted.</summary>
        NoKeyConfigured,
        TooLarge,
        MalformedJson,
        UnsupportedSchema,
        /// <summary>Signed by a key this build does not know.</summary>
        UnknownKey,
        BadSignature,
        WrongProduct,
        WrongChannel,
        InsecureUrl,
        BadDigest,
        BadSize,
        BadVersion,
        /// <summary>Authentic, but not newer than what is installed.</summary>
        NotNewer,
        /// <summary>Too old to be upgraded in place; the user must reinstall.</summary>
        ManualUpgradeRequired,
    }

    /// <summary>An authenticated, newer release the client may install.</summary>
    public sealed class UpdateRelease
    {
        public string Version { get; set; }
        public string Channel { get; set; }
        public string Url { get; set; }
        public string Sha256 { get; set; }
        public long Size { get; set; }
        public string ReleasedAt { get; set; }
        public string Notes { get; set; }

        /// <summary>
        /// The manifest exactly as received. Staged verbatim so the updater can
        /// re-verify the signature itself rather than trusting a re-encoding.
        /// </summary>
        public string RawManifest { get; set; }
    }

    public sealed class UpdateCheckResult
    {
        public bool Accepted => Rejection == UpdateRejection.None && Release != null;
        public UpdateRejection Rejection { get; set; }
        public string Detail { get; set; }
        public UpdateRelease Release { get; set; }

        public static UpdateCheckResult Reject(UpdateRejection rejection, string detail)
        {
            return new UpdateCheckResult { Rejection = rejection, Detail = detail };
        }
    }

    /// <summary>
    /// Parses and authenticates an update manifest.
    ///
    /// Wire format — the signed bytes are carried as opaque base64 rather than
    /// as a nested JSON object:
    ///
    ///     { "schema": 1, "key_id": "...", "signature": "&lt;b64&gt;", "payload": "&lt;b64 of UTF-8 JSON&gt;" }
    ///
    /// so the client verifies the signature over the exact bytes that were
    /// signed and only then parses them. Signing a re-serialized object instead
    /// would make correctness depend on two implementations agreeing on key
    /// order, escaping and number formatting — a classic source of
    /// signature-bypass bugs. Here there is nothing to canonicalize.
    ///
    /// **Everything above the signature check is untrusted input; everything
    /// below it is authenticated.** Keep that boundary where it is.
    ///
    /// Parsing goes through JavaScriptSerializer (System.Web.Extensions, in the
    /// GAC) rather than System.Text.Json so that this file also compiles into
    /// DeepExcel.Updater.exe, which must run as a lone binary copied out of the
    /// install directory — the directory it is about to let the installer
    /// overwrite. No type resolver is supplied, so deserialization produces
    /// inert dictionaries and strings.
    /// </summary>
    public static class UpdateManifest
    {
        public const int Schema = 1;
        public const string Product = "DeepExcel";
        public const string DefaultChannel = "stable";

        /// <summary>A manifest is a few hundred bytes; this is three orders of magnitude of headroom.</summary>
        public const int MaxManifestBytes = 64 * 1024;

        /// <summary>Refuse to plan a download larger than any plausible installer.</summary>
        public const long MaxPackageBytes = 512L * 1024 * 1024;

        public static UpdateCheckResult Evaluate(
            string manifestJson, string installedVersion, string channel, IManifestVerifier verifier)
        {
            if (verifier == null || !verifier.IsConfigured)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.NoKeyConfigured,
                    "此版本未内置更新签名公钥，自动更新已停用。");
            }
            if (string.IsNullOrWhiteSpace(manifestJson))
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "更新清单为空。");
            }
            if (manifestJson.Length > MaxManifestBytes)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.TooLarge,
                    string.Format(CultureInfo.InvariantCulture,
                        "更新清单过大（{0} 字节）。", manifestJson.Length));
            }

            Dictionary<string, object> envelope = TryParseObject(manifestJson);
            if (envelope == null)
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "更新清单不是合法 JSON 对象。");
            }

            if (!TryGetInt64(envelope, "schema", out long schema) || schema != Schema)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.UnsupportedSchema, "更新清单版本不受支持，请手动升级。");
            }

            string keyId = GetString(envelope, "key_id");
            if (!string.Equals(keyId, verifier.KeyId, StringComparison.Ordinal))
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.UnknownKey,
                    string.Format(CultureInfo.InvariantCulture,
                        "更新清单由未知密钥签名（{0}）。", Describe(keyId)));
            }

            if (!TryDecodeBase64(GetString(envelope, "payload"), out byte[] payloadBytes) ||
                payloadBytes.Length == 0 || payloadBytes.Length > MaxManifestBytes)
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "更新清单负载无法解码。");
            }
            if (!TryDecodeBase64(GetString(envelope, "signature"), out byte[] signature))
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "更新清单签名无法解码。");
            }

            // ---- trust boundary -------------------------------------------
            if (!verifier.Verify(payloadBytes, signature))
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.BadSignature, "更新清单签名校验失败，已拒绝。");
            }
            // ---- everything below here is authenticated --------------------

            Dictionary<string, object> payload = TryParseObject(DecodeUtf8(payloadBytes));
            if (payload == null)
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "已签名负载不是合法 JSON 对象。");
            }

            if (!TryGetInt64(payload, "schema", out long payloadSchema) || payloadSchema != Schema)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.UnsupportedSchema, "已签名负载版本不受支持，请手动升级。");
            }

            // Binds the signature to one product and one channel, so a manifest
            // signed for something else cannot be replayed at this client.
            if (!string.Equals(GetString(payload, "product"), Product, StringComparison.Ordinal))
            {
                return UpdateCheckResult.Reject(UpdateRejection.WrongProduct, "更新清单不属于本产品。");
            }

            string wantedChannel = string.IsNullOrWhiteSpace(channel) ? DefaultChannel : channel.Trim();
            string manifestChannel = GetString(payload, "channel");
            if (!string.Equals(manifestChannel, wantedChannel, StringComparison.Ordinal))
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.WrongChannel,
                    string.Format(CultureInfo.InvariantCulture,
                        "更新通道不匹配（清单 {0}，本机 {1}）。",
                        Describe(manifestChannel), wantedChannel));
            }

            string versionText = GetString(payload, "version");
            if (!ReleaseVersion.TryParse(versionText, out ReleaseVersion candidate))
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.BadVersion,
                    string.Format(CultureInfo.InvariantCulture, "版本号无法解析：{0}", Describe(versionText)));
            }
            if (!ReleaseVersion.TryParse(installedVersion, out ReleaseVersion installed))
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.BadVersion,
                    string.Format(CultureInfo.InvariantCulture,
                        "本机版本号无法解析：{0}", Describe(installedVersion)));
            }

            // Strictly newer. This is also the downgrade defence: an attacker who
            // can replay an older but validly signed manifest still cannot walk
            // the client back to a version with a known defect.
            if (candidate <= installed)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.NotNewer,
                    string.Format(CultureInfo.InvariantCulture,
                        "已是最新版本（本机 {0}，最新 {1}）。", installedVersion, versionText));
            }

            string floorText = GetString(payload, "minimum_upgradable_version");
            if (!string.IsNullOrWhiteSpace(floorText))
            {
                if (!ReleaseVersion.TryParse(floorText, out ReleaseVersion floor))
                {
                    return UpdateCheckResult.Reject(
                        UpdateRejection.BadVersion,
                        string.Format(CultureInfo.InvariantCulture,
                            "最低可升级版本无法解析：{0}", Describe(floorText)));
                }
                if (installed < floor)
                {
                    return UpdateCheckResult.Reject(
                        UpdateRejection.ManualUpgradeRequired,
                        string.Format(CultureInfo.InvariantCulture,
                            "本机 {0} 过旧，无法自动升级到 {1}，请手动重新安装。",
                            installedVersion, versionText));
                }
            }

            string url = GetString(payload, "url");
            string urlProblem = ValidateUrl(url);
            if (urlProblem != null)
            {
                return UpdateCheckResult.Reject(UpdateRejection.InsecureUrl, urlProblem);
            }

            string digest = NormalizeDigest(GetString(payload, "sha256"));
            if (digest == null)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.BadDigest, "更新包 SHA-256 不是 64 位十六进制。");
            }

            if (!TryGetInt64(payload, "size", out long size) || size <= 0 || size > MaxPackageBytes)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.BadSize, "更新包大小超出允许范围。");
            }

            return new UpdateCheckResult
            {
                Rejection = UpdateRejection.None,
                Release = new UpdateRelease
                {
                    Version = versionText,
                    Channel = manifestChannel,
                    Url = url,
                    Sha256 = digest,
                    Size = size,
                    ReleasedAt = GetString(payload, "released_at"),
                    Notes = Truncate(GetString(payload, "notes"), 2000),
                    RawManifest = manifestJson,
                }
            };
        }

        /// <summary>
        /// HTTPS only, no embedded credentials, no fragment.
        ///
        /// The signature already authorises wherever the URL points, so the host
        /// is deliberately not pinned — releases can move to a CDN or to GitHub
        /// Releases without a client rebuild. What is not negotiable is that the
        /// bytes arrive over TLS.
        /// </summary>
        private static string ValidateUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return "更新包地址为空。";
            }
            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri uri))
            {
                return "更新包地址不是合法 URL。";
            }
            if (!string.Equals(uri.Scheme, "https", StringComparison.OrdinalIgnoreCase))
            {
                return "更新包地址必须是 https。";
            }
            if (!string.IsNullOrEmpty(uri.UserInfo))
            {
                return "更新包地址不得内嵌凭据。";
            }
            return null;
        }

        /// <summary>Lowercase 64-hex, or null when it is not a SHA-256 digest.</summary>
        public static string NormalizeDigest(string digest)
        {
            if (digest == null || digest.Length != 64)
            {
                return null;
            }
            var builder = new StringBuilder(64);
            foreach (char c in digest)
            {
                if (c >= '0' && c <= '9') builder.Append(c);
                else if (c >= 'a' && c <= 'f') builder.Append(c);
                else if (c >= 'A' && c <= 'F') builder.Append(char.ToLowerInvariant(c));
                else return null;
            }
            return builder.ToString();
        }

        // ------------------------------------------------------------------

        private static Dictionary<string, object> TryParseObject(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }
            try
            {
                var serializer = new JavaScriptSerializer { MaxJsonLength = MaxManifestBytes };
                return serializer.DeserializeObject(json) as Dictionary<string, object>;
            }
            catch (ArgumentException)
            {
                return null;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static string DecodeUtf8(byte[] bytes)
        {
            try
            {
                // Throwing encoder: invalid UTF-8 must not be silently turned
                // into replacement characters inside a signed document.
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static bool TryDecodeBase64(string text, out byte[] bytes)
        {
            bytes = null;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            try
            {
                bytes = Convert.FromBase64String(text.Trim());
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static string GetString(Dictionary<string, object> map, string key)
        {
            return map != null && map.TryGetValue(key, out object value) ? value as string : null;
        }

        private static bool TryGetInt64(Dictionary<string, object> map, string key, out long value)
        {
            value = 0;
            if (map == null || !map.TryGetValue(key, out object raw) || raw == null)
            {
                return false;
            }
            // JavaScriptSerializer widens JSON numbers to int, long or decimal
            // depending on magnitude. A floating-point size or schema is a
            // malformed manifest, not something to round.
            if (raw is int i) { value = i; return true; }
            if (raw is long l) { value = l; return true; }
            if (raw is decimal d)
            {
                if (d != decimal.Truncate(d) || d > long.MaxValue || d < long.MinValue) return false;
                value = (long)d;
                return true;
            }
            return false;
        }

        private static string Describe(string value)
        {
            return string.IsNullOrEmpty(value) ? "(空)" : Truncate(value, 64);
        }

        private static string Truncate(string value, int max)
        {
            if (value == null || value.Length <= max)
            {
                return value;
            }
            return value.Substring(0, max) + "…";
        }
    }
}
