using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Manifest authentication.
    ///
    /// The client is unsigned, so a poisoned update channel is the one way to
    /// get arbitrary code onto a user's machine with the product's blessing.
    /// Every check the verifier makes is asserted here, including the ones for
    /// manifests that are correctly signed but still must not be installed.
    /// </summary>
    public class UpdateManifestTests
    {
        private const string InstalledVersion = "0.5.0";
        private const string Digest = "9f2c4f1c0b6d5e3a7c8d9e0f1a2b3c4d5e6f708192a3b4c5d6e7f8091a2b3c4d";

        private readonly RSACng _key = new RSACng(2048);
        private readonly IManifestVerifier _verifier;

        public UpdateManifestTests()
        {
            RSAParameters parameters = _key.ExportParameters(false);
            _verifier = RsaManifestVerifier.FromBase64(
                Convert.ToBase64String(parameters.Modulus),
                Convert.ToBase64String(parameters.Exponent));
        }

        private static Dictionary<string, object> BasePayload()
        {
            return new Dictionary<string, object>
            {
                ["schema"] = 1,
                ["product"] = "DeepExcel",
                ["channel"] = "stable",
                ["version"] = "0.6.0",
                ["url"] = "https://updates.example.com/DeepExcel.Setup.exe",
                ["sha256"] = Digest,
                ["size"] = 50_000_000,
                ["released_at"] = "2026-09-13T00:00:00Z",
                ["notes"] = "",
            };
        }

        private string Sign(Dictionary<string, object> payload, string keyIdOverride = null)
        {
            byte[] payloadBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload));
            byte[] signature = _key.SignData(
                payloadBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pss);
            return JsonSerializer.Serialize(new Dictionary<string, object>
            {
                ["schema"] = 1,
                ["key_id"] = keyIdOverride ?? _verifier.KeyId,
                ["signature"] = Convert.ToBase64String(signature),
                ["payload"] = Convert.ToBase64String(payloadBytes),
            });
        }

        private UpdateCheckResult Evaluate(Dictionary<string, object> payload, string keyId = null)
        {
            return UpdateManifest.Evaluate(Sign(payload, keyId), InstalledVersion, "stable", _verifier);
        }

        [Fact]
        public void Accepts_a_newer_signed_release()
        {
            UpdateCheckResult result = Evaluate(BasePayload());

            Assert.True(result.Accepted, result.Detail);
            Assert.Equal("0.6.0", result.Release.Version);
            Assert.Equal(Digest, result.Release.Sha256);
            Assert.Equal(50_000_000, result.Release.Size);
        }

        [Fact]
        public void Keeps_the_manifest_verbatim_for_restaging()
        {
            string manifest = Sign(BasePayload());

            UpdateCheckResult result = UpdateManifest.Evaluate(
                manifest, InstalledVersion, "stable", _verifier);

            // The updater re-verifies from the staged copy, so it has to be the
            // bytes that were authenticated, not a re-encoding of them.
            Assert.Equal(manifest, result.Release.RawManifest);
        }

        [Fact]
        public void Rejects_a_payload_edited_after_signing()
        {
            string manifest = Sign(BasePayload());
            using (JsonDocument document = JsonDocument.Parse(manifest))
            {
                var tampered = BasePayload();
                tampered["url"] = "https://evil.example.com/DeepExcel.Setup.exe";
                string forged = JsonSerializer.Serialize(new Dictionary<string, object>
                {
                    ["schema"] = 1,
                    ["key_id"] = document.RootElement.GetProperty("key_id").GetString(),
                    ["signature"] = document.RootElement.GetProperty("signature").GetString(),
                    ["payload"] = Convert.ToBase64String(
                        Encoding.UTF8.GetBytes(JsonSerializer.Serialize(tampered))),
                });

                UpdateCheckResult result = UpdateManifest.Evaluate(
                    forged, InstalledVersion, "stable", _verifier);

                Assert.Equal(UpdateRejection.BadSignature, result.Rejection);
            }
        }

        [Fact]
        public void Rejects_a_signature_from_another_key()
        {
            using (var other = new RSACng(2048))
            {
                var payload = BasePayload();
                byte[] payloadBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload));
                string forged = JsonSerializer.Serialize(new Dictionary<string, object>
                {
                    ["schema"] = 1,
                    // Claims our key id, so this gets past the cheap check and
                    // has to be caught by the signature itself.
                    ["key_id"] = _verifier.KeyId,
                    ["signature"] = Convert.ToBase64String(other.SignData(
                        payloadBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pss)),
                    ["payload"] = Convert.ToBase64String(payloadBytes),
                });

                UpdateCheckResult result = UpdateManifest.Evaluate(
                    forged, InstalledVersion, "stable", _verifier);

                Assert.Equal(UpdateRejection.BadSignature, result.Rejection);
            }
        }

        [Fact]
        public void Rejects_an_unknown_key_id()
        {
            Assert.Equal(UpdateRejection.UnknownKey, Evaluate(BasePayload(), "0123456789abcdef").Rejection);
        }

        [Fact]
        public void Rejects_pkcs1_signatures()
        {
            // PSS is not decoration: accepting whatever padding the manifest
            // happens to carry would let a signer downgrade the scheme.
            var payload = BasePayload();
            byte[] payloadBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload));
            string manifest = JsonSerializer.Serialize(new Dictionary<string, object>
            {
                ["schema"] = 1,
                ["key_id"] = _verifier.KeyId,
                ["signature"] = Convert.ToBase64String(_key.SignData(
                    payloadBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1)),
                ["payload"] = Convert.ToBase64String(payloadBytes),
            });

            Assert.Equal(
                UpdateRejection.BadSignature,
                UpdateManifest.Evaluate(manifest, InstalledVersion, "stable", _verifier).Rejection);
        }

        [Theory]
        [InlineData("0.5.0", UpdateRejection.NotNewer)]   // same as installed
        [InlineData("0.4.9", UpdateRejection.NotNewer)]   // older
        [InlineData("0.5.0.1", UpdateRejection.None)]     // fourth component still counts
        [InlineData("1.0.0", UpdateRejection.None)]
        public void Only_a_strictly_newer_version_is_installed(string version, UpdateRejection expected)
        {
            var payload = BasePayload();
            payload["version"] = version;

            Assert.Equal(expected, Evaluate(payload).Rejection);
        }

        [Fact]
        public void Refuses_to_upgrade_a_build_below_the_manifest_floor()
        {
            var payload = BasePayload();
            payload["minimum_upgradable_version"] = "0.5.1";

            UpdateCheckResult result = Evaluate(payload);

            Assert.Equal(UpdateRejection.ManualUpgradeRequired, result.Rejection);
            Assert.Contains("重新安装", result.Detail);
        }

        [Fact]
        public void Allows_an_upgrade_exactly_at_the_floor()
        {
            var payload = BasePayload();
            payload["minimum_upgradable_version"] = InstalledVersion;

            Assert.True(Evaluate(payload).Accepted);
        }

        [Theory]
        [InlineData("http://updates.example.com/s.exe")]
        [InlineData("ftp://updates.example.com/s.exe")]
        [InlineData("file://C:/temp/s.exe")]
        [InlineData("https://user:secret@updates.example.com/s.exe")]
        [InlineData("/relative/path.exe")]
        [InlineData("")]
        public void Rejects_a_url_that_is_not_plain_https(string url)
        {
            var payload = BasePayload();
            payload["url"] = url;

            Assert.Equal(UpdateRejection.InsecureUrl, Evaluate(payload).Rejection);
        }

        [Theory]
        [InlineData("")]
        [InlineData("abc")]
        [InlineData("zz9c4f1c0b6d5e3a7c8d9e0f1a2b3c4d5e6f708192a3b4c5d6e7f8091a2b3c4d")]
        public void Rejects_a_digest_that_is_not_sha256(string digest)
        {
            var payload = BasePayload();
            payload["sha256"] = digest;

            Assert.Equal(UpdateRejection.BadDigest, Evaluate(payload).Rejection);
        }

        [Fact]
        public void Accepts_an_uppercase_digest_and_normalises_it()
        {
            var payload = BasePayload();
            payload["sha256"] = Digest.ToUpperInvariant();

            Assert.Equal(Digest, Evaluate(payload).Release.Sha256);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(2L * 1024 * 1024 * 1024)]
        public void Rejects_an_implausible_package_size(long size)
        {
            var payload = BasePayload();
            payload["size"] = size;

            Assert.Equal(UpdateRejection.BadSize, Evaluate(payload).Rejection);
        }

        [Fact]
        public void Rejects_a_fractional_size_rather_than_rounding_it()
        {
            var payload = BasePayload();
            payload["size"] = 1024.5;

            Assert.Equal(UpdateRejection.BadSize, Evaluate(payload).Rejection);
        }

        [Fact]
        public void Rejects_a_manifest_for_another_product()
        {
            var payload = BasePayload();
            payload["product"] = "NotDeepExcel";

            Assert.Equal(UpdateRejection.WrongProduct, Evaluate(payload).Rejection);
        }

        [Fact]
        public void Rejects_a_manifest_from_another_channel()
        {
            var payload = BasePayload();
            payload["channel"] = "beta";

            Assert.Equal(UpdateRejection.WrongChannel, Evaluate(payload).Rejection);
        }

        [Fact]
        public void Rejects_a_future_schema_instead_of_guessing()
        {
            var payload = BasePayload();
            payload["schema"] = 2;

            Assert.Equal(UpdateRejection.UnsupportedSchema, Evaluate(payload).Rejection);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("not json")]
        [InlineData("[1,2,3]")]
        [InlineData("{\"schema\":1,\"key_id\":\"x\"}")]
        public void Rejects_junk_without_throwing(string manifest)
        {
            UpdateCheckResult result = UpdateManifest.Evaluate(
                manifest, InstalledVersion, "stable", _verifier);

            Assert.False(result.Accepted);
            Assert.False(string.IsNullOrEmpty(result.Detail));
        }

        [Fact]
        public void Rejects_a_manifest_larger_than_the_cap()
        {
            string oversized = new string('x', UpdateManifest.MaxManifestBytes + 1);

            Assert.Equal(
                UpdateRejection.TooLarge,
                UpdateManifest.Evaluate(oversized, InstalledVersion, "stable", _verifier).Rejection);
        }

        [Fact]
        public void A_build_with_no_compiled_in_key_accepts_nothing()
        {
            // This is the shipping default, so it had better be the safe one.
            IManifestVerifier none = RsaManifestVerifier.FromBase64("", "");

            UpdateCheckResult result = UpdateManifest.Evaluate(
                Sign(BasePayload()), InstalledVersion, "stable", none);

            Assert.Equal(UpdateRejection.NoKeyConfigured, result.Rejection);
            Assert.False(none.IsConfigured);
        }

        [Fact]
        public void The_shipped_default_has_no_key()
        {
            // Fails the moment a real key is committed, which is the point:
            // the public key belongs in a release build, not in the repository.
            Assert.False(EmbeddedUpdateKey.Verifier.IsConfigured);
        }

        [Theory]
        [InlineData("not base64", "AQAB")]
        [InlineData("AQAB", "AQAB")]            // 3-byte modulus: far below the floor
        [InlineData("", "AQAB")]
        public void A_malformed_compiled_in_key_disables_updates_instead_of_throwing(
            string modulus, string exponent)
        {
            Assert.False(RsaManifestVerifier.FromBase64(modulus, exponent).IsConfigured);
        }

        [Fact]
        public void Key_id_is_derived_from_the_key_material()
        {
            RSAParameters parameters = _key.ExportParameters(false);
            string expected = RsaManifestVerifier.ComputeKeyId(parameters.Modulus, parameters.Exponent);

            Assert.Equal(expected, _verifier.KeyId);
            Assert.Equal(16, _verifier.KeyId.Length);
        }
    }
}
