using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Cross-language vectors: manifests produced by the real Python signer.
    ///
    /// <see cref="UpdateManifestTests"/> signs with RSACng, which proves the
    /// verifier is self-consistent and nothing more. Release manifests are
    /// signed by scripts/update_signing.py, and the two implementations have to
    /// agree on details no unit test on either side would catch alone -- RSA-PSS
    /// salt length above all, where .NET fixes it at the digest size and
    /// cryptography's default MAX_LENGTH produces signatures .NET silently
    /// refuses.
    ///
    /// Regenerate after any manifest format change:
    ///     python scripts/update_signing.py testvectors
    /// </summary>
    public class UpdateVectorTests
    {
        private const string VectorFile = "update-test-vectors.json";

        private static string LocateVectors()
        {
            // run-tests-csharp.ps1 stages src/DeepExcel.Tests/fixtures next to the
            // compiled test assembly.
            string assemblyDirectory = Path.GetDirectoryName(
                new Uri(typeof(UpdateVectorTests).Assembly.CodeBase).LocalPath);
            foreach (string candidate in new[]
                     {
                         Path.Combine(assemblyDirectory, "fixtures", VectorFile),
                         Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fixtures", VectorFile),
                     })
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            throw new FileNotFoundException(
                "Update test vectors are missing. Regenerate with: " +
                "python scripts/update_signing.py testvectors");
        }

        public static IEnumerable<object[]> Cases()
        {
            using (JsonDocument document = JsonDocument.Parse(File.ReadAllText(LocateVectors())))
            {
                foreach (JsonElement element in document.RootElement.GetProperty("cases").EnumerateArray())
                {
                    yield return new object[] { element.GetProperty("name").GetString() };
                }
            }
        }

        private static (JsonDocument Document, IManifestVerifier Verifier) Load()
        {
            JsonDocument document = JsonDocument.Parse(File.ReadAllText(LocateVectors()));
            JsonElement key = document.RootElement.GetProperty("key");
            IManifestVerifier verifier = RsaManifestVerifier.FromBase64(
                key.GetProperty("modulus_b64").GetString(),
                key.GetProperty("exponent_b64").GetString());
            return (document, verifier);
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void Python_signed_manifest_gets_the_expected_verdict(string caseName)
        {
            var (document, verifier) = Load();
            using (document)
            {
                JsonElement root = document.RootElement;
                string installedVersion = root.GetProperty("installed_version").GetString();
                string channel = root.GetProperty("channel").GetString();

                JsonElement testCase = default;
                foreach (JsonElement element in root.GetProperty("cases").EnumerateArray())
                {
                    if (element.GetProperty("name").GetString() == caseName)
                    {
                        testCase = element;
                        break;
                    }
                }
                Assert.Equal(JsonValueKind.Object, testCase.ValueKind);

                var expected = (UpdateRejection)Enum.Parse(
                    typeof(UpdateRejection), testCase.GetProperty("expect").GetString());
                string manifest = testCase.GetProperty("manifest").GetRawText();

                UpdateCheckResult result = UpdateManifest.Evaluate(
                    manifest, installedVersion, channel, verifier);

                Assert.Equal(expected, result.Rejection);
                if (expected == UpdateRejection.None)
                {
                    Assert.True(result.Accepted, result.Detail);
                    Assert.Equal(
                        testCase.GetProperty("expect_version").GetString(), result.Release.Version);
                }
            }
        }

        [Fact]
        public void Key_id_agrees_across_languages()
        {
            var (document, verifier) = Load();
            using (document)
            {
                // If this drifts, every released manifest is rejected as
                // UnknownKey and updates die silently in the field.
                Assert.Equal(
                    document.RootElement.GetProperty("key").GetProperty("key_id").GetString(),
                    verifier.KeyId);
            }
        }

        [Fact]
        public void The_vector_file_covers_every_rejection_reason_the_verifier_can_return()
        {
            var (document, _) = Load();
            using (document)
            {
                var covered = new HashSet<string>();
                foreach (JsonElement element in document.RootElement.GetProperty("cases").EnumerateArray())
                {
                    covered.Add(element.GetProperty("expect").GetString());
                }

                // Reasons that cannot be expressed by a correctly signed manifest
                // served over the wire, so they are covered by UpdateManifestTests
                // rather than by a vector.
                var localOnly = new HashSet<string>
                {
                    nameof(UpdateRejection.NoKeyConfigured),
                    nameof(UpdateRejection.TooLarge),
                    nameof(UpdateRejection.MalformedJson),
                    nameof(UpdateRejection.BadVersion),
                };

                var missing = new List<string>();
                foreach (string name in Enum.GetNames(typeof(UpdateRejection)))
                {
                    if (!localOnly.Contains(name) && !covered.Contains(name))
                    {
                        missing.Add(name);
                    }
                }

                Assert.True(
                    missing.Count == 0,
                    "Rejection reasons with no cross-language vector: " + string.Join(", ", missing) +
                    ". Add a case to build_test_vectors() in scripts/update_signing.py.");
            }
        }
    }
}
