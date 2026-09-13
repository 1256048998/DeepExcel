using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The staging directory.
    ///
    /// A verified download still sits in a user-writable folder until the
    /// installer runs, so what is asserted here is that the bytes are checked
    /// again at that point, and that a version string coming out of a manifest
    /// can never be used as a path.
    /// </summary>
    public class UpdateStageTests : IDisposable
    {
        /// <summary>SHA-256 of the empty input; a constant worth not deriving from the code under test.</summary>
        private const string EmptyDigest =
            "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855";

        private readonly string _root;

        public UpdateStageTests()
        {
            _root = Path.Combine(Path.GetTempPath(), "DeepExcelUpdateTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        private string WriteFile(string name, byte[] content)
        {
            string path = Path.Combine(_root, name);
            File.WriteAllBytes(path, content);
            return path;
        }

        private static string Sha256(byte[] content)
        {
            using (var sha = SHA256.Create())
            {
                var builder = new StringBuilder(64);
                foreach (byte b in sha.ComputeHash(content))
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

        [Fact]
        public void Computes_the_same_digest_the_manifest_publishes()
        {
            Assert.Equal(EmptyDigest, UpdateStage.ComputeSha256(WriteFile("empty.bin", new byte[0])));
        }

        [Fact]
        public void Accepts_a_package_that_matches_the_manifest()
        {
            byte[] content = Encoding.UTF8.GetBytes("installer bytes");
            string path = WriteFile("ok.exe", content);
            var release = new UpdateRelease { Sha256 = Sha256(content), Size = content.Length };

            Assert.True(UpdateStage.VerifyPackage(path, release, out string detail), detail);
        }

        [Fact]
        public void Rejects_a_package_swapped_after_download()
        {
            byte[] original = Encoding.UTF8.GetBytes("installer bytes");
            var release = new UpdateRelease { Sha256 = Sha256(original), Size = original.Length };
            // Same length, different content: only the digest catches this, which
            // is why the size check is not the whole check.
            string path = WriteFile("swapped.exe", Encoding.UTF8.GetBytes("INSTALLER BYTES"));

            Assert.False(UpdateStage.VerifyPackage(path, release, out string detail));
            Assert.Contains("SHA-256", detail);
        }

        [Fact]
        public void Rejects_a_truncated_package_before_hashing_it()
        {
            byte[] original = Encoding.UTF8.GetBytes("installer bytes");
            var release = new UpdateRelease { Sha256 = Sha256(original), Size = original.Length };
            string path = WriteFile("short.exe", Encoding.UTF8.GetBytes("inst"));

            Assert.False(UpdateStage.VerifyPackage(path, release, out string detail));
            Assert.Contains("大小", detail);
        }

        [Fact]
        public void Reports_a_missing_package_rather_than_throwing()
        {
            var release = new UpdateRelease { Sha256 = EmptyDigest, Size = 0 };

            Assert.False(UpdateStage.VerifyPackage(
                Path.Combine(_root, "absent.exe"), release, out string detail));
            Assert.False(string.IsNullOrEmpty(detail));
        }

        [Theory]
        [InlineData("..")]
        [InlineData("../../Windows")]
        [InlineData("0.5.0/../../etc")]
        [InlineData("CON")]
        [InlineData("")]
        [InlineData("latest")]
        public void A_version_from_a_manifest_can_never_become_a_path(string version)
        {
            // The version string is attacker-influenced before the signature is
            // checked by anything downstream, so it is validated rather than
            // sanitised.
            Assert.Throws<ArgumentException>(() => UpdateStage.DirectoryFor(_root, version));
        }

        [Fact]
        public void Builds_a_directory_for_a_real_version()
        {
            Assert.Equal(
                Path.Combine(_root, "0.6.0"),
                UpdateStage.DirectoryFor(_root, "0.6.0"));
        }

        [Fact]
        public void Prune_keeps_the_named_version_and_drops_the_rest()
        {
            Directory.CreateDirectory(Path.Combine(_root, "0.5.0"));
            Directory.CreateDirectory(Path.Combine(_root, "0.6.0"));
            Directory.CreateDirectory(Path.Combine(_root, "0.4.9"));
            File.WriteAllText(Path.Combine(_root, "0.5.0", "DeepExcel.Setup.exe"), "x");

            UpdateStage.Prune(_root, "0.6.0");

            Assert.True(Directory.Exists(Path.Combine(_root, "0.6.0")));
            Assert.False(Directory.Exists(Path.Combine(_root, "0.5.0")));
            Assert.False(Directory.Exists(Path.Combine(_root, "0.4.9")));
        }

        [Fact]
        public void Prune_leaves_anything_that_is_not_a_version_directory_alone()
        {
            string stranger = Path.Combine(_root, "not-a-version");
            Directory.CreateDirectory(stranger);

            UpdateStage.Prune(_root, null);

            Assert.True(Directory.Exists(stranger));
        }

        [Fact]
        public void Prune_on_a_missing_root_is_a_no_op()
        {
            UpdateStage.Prune(Path.Combine(_root, "never-created"), "0.6.0");
        }

        [Fact]
        public void Load_reports_a_missing_manifest_instead_of_throwing()
        {
            IManifestVerifier verifier = RsaManifestVerifier.FromBase64(
                Convert.ToBase64String(new byte[256]), "AQAB");

            UpdateCheckResult result = UpdateStage.Load(_root, "0.5.0", "stable", verifier);

            Assert.False(result.Accepted);
            Assert.Contains("更新清单", result.Detail);
        }

        [Fact]
        public void Default_root_is_under_local_appdata()
        {
            string root = UpdateStage.DefaultRoot();

            Assert.StartsWith(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                root,
                StringComparison.OrdinalIgnoreCase);
            Assert.EndsWith(Path.Combine("DeepExcel", "updates"), root, StringComparison.OrdinalIgnoreCase);
        }
    }
}
