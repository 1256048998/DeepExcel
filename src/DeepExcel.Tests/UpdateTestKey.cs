using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DeepExcel.AddIn.Updates;

namespace DeepExcel.Tests
{
    /// <summary>
    /// An ephemeral signing key plus manifest construction, for tests that need
    /// an authentic manifest rather than a hand-written one.
    ///
    /// Generated per instance and never persisted. Cross-language agreement with
    /// the real Python signer is a separate concern, covered by
    /// <see cref="UpdateVectorTests"/>.
    /// </summary>
    public sealed class UpdateTestKey : IDisposable
    {
        public const string SampleDigest =
            "9f2c4f1c0b6d5e3a7c8d9e0f1a2b3c4d5e6f708192a3b4c5d6e7f8091a2b3c4d";

        private readonly RSACng _key = new RSACng(2048);

        public UpdateTestKey()
        {
            RSAParameters parameters = _key.ExportParameters(false);
            Verifier = RsaManifestVerifier.FromBase64(
                Convert.ToBase64String(parameters.Modulus),
                Convert.ToBase64String(parameters.Exponent));
        }

        public IManifestVerifier Verifier { get; }

        public static Dictionary<string, object> Payload(
            string version = "0.6.0", string sha256 = SampleDigest, long size = 50_000_000,
            string channel = "stable", string url = "https://updates.example.com/DeepExcel.Setup.exe")
        {
            return new Dictionary<string, object>
            {
                ["schema"] = 1,
                ["product"] = "DeepExcel",
                ["channel"] = channel,
                ["version"] = version,
                ["url"] = url,
                ["sha256"] = sha256,
                ["size"] = size,
                ["released_at"] = "2026-09-14T00:00:00Z",
                ["notes"] = "",
            };
        }

        public string Sign(Dictionary<string, object> payload)
        {
            byte[] payloadBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload));
            byte[] signature = _key.SignData(
                payloadBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pss);
            return JsonSerializer.Serialize(new Dictionary<string, object>
            {
                ["schema"] = 1,
                ["key_id"] = Verifier.KeyId,
                ["signature"] = Convert.ToBase64String(signature),
                ["payload"] = Convert.ToBase64String(payloadBytes),
            });
        }

        /// <summary>A manifest describing a package whose bytes are given.</summary>
        public string SignFor(byte[] package, string version = "0.6.0", string url = null)
        {
            using (var sha = SHA256.Create())
            {
                var digest = new StringBuilder(64);
                foreach (byte b in sha.ComputeHash(package))
                {
                    digest.Append(b.ToString("x2"));
                }
                return Sign(Payload(
                    version: version,
                    sha256: digest.ToString(),
                    size: package.Length,
                    url: url ?? "https://updates.example.com/DeepExcel.Setup.exe"));
            }
        }

        public void Dispose()
        {
            _key.Dispose();
        }
    }
}
