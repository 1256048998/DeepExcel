using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>A downloaded, verified release sitting on disk, ready to install.</summary>
    public sealed class StagedUpdate
    {
        public string Directory { get; set; }
        public string PackagePath { get; set; }
        public UpdateRelease Release { get; set; }
    }

    /// <summary>
    /// Layout of, and checks over, the local staging directory.
    ///
    ///     %LOCALAPPDATA%\DeepExcel\updates\&lt;version&gt;\
    ///         update.json             the manifest verbatim, signature intact
    ///         DeepExcel.Setup.exe     the package
    ///         DeepExcel.Updater.exe   a copy of the updater, run from here
    ///         updater.log             what the last attempt did
    ///
    /// The updater runs from this directory rather than from the install
    /// directory on purpose: Inno Setup cannot overwrite a running executable,
    /// so an updater launched in place would block the very install it started.
    ///
    /// Shared with DeepExcel.Updater.exe, so nothing here may reference anything
    /// outside mscorlib / System / System.Core / System.Web.Extensions.
    /// </summary>
    public static class UpdateStage
    {
        public const string ManifestFileName = "update.json";
        public const string PackageFileName = "DeepExcel.Setup.exe";
        public const string UpdaterFileName = "DeepExcel.Updater.exe";
        public const string LogFileName = "updater.log";
        public const string PartialSuffix = ".part";

        /// <summary>
        /// %LOCALAPPDATA%, not %APPDATA%: a half-downloaded installer is machine
        /// state, and pushing it through a roaming profile would be rude.
        /// </summary>
        public static string DefaultRoot()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DeepExcel", "updates");
        }

        public static string DirectoryFor(string root, string version)
        {
            if (string.IsNullOrEmpty(root)) throw new ArgumentException("root is required", nameof(root));
            if (!ReleaseVersion.TryParse(version, out _))
            {
                // The version string becomes a directory name, so it is checked
                // rather than sanitised: no traversal, no device names, nothing
                // to reason about.
                throw new ArgumentException("version must be dotted numeric: " + version, nameof(version));
            }
            return Path.Combine(root, version);
        }

        /// <summary>
        /// Re-reads and re-authenticates a staged update.
        ///
        /// The updater calls this immediately before running the installer even
        /// though the add-in already verified everything at download time. The
        /// package sits on disk in a user-writable directory in between, and a
        /// check that happens once at download is a check that does not cover
        /// the moment of execution.
        /// </summary>
        public static UpdateCheckResult Load(
            string stageDirectory, string installedVersion, string channel, IManifestVerifier verifier)
        {
            string manifestPath = Path.Combine(stageDirectory, ManifestFileName);
            string manifestJson;
            try
            {
                var info = new FileInfo(manifestPath);
                if (!info.Exists)
                {
                    return UpdateCheckResult.Reject(
                        UpdateRejection.MalformedJson, "缺少更新清单：" + manifestPath);
                }
                if (info.Length > UpdateManifest.MaxManifestBytes)
                {
                    return UpdateCheckResult.Reject(UpdateRejection.TooLarge, "更新清单过大。");
                }
                manifestJson = File.ReadAllText(manifestPath, new UTF8Encoding(false));
            }
            catch (IOException ex)
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "读取更新清单失败：" + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "读取更新清单失败：" + ex.Message);
            }

            return UpdateManifest.Evaluate(manifestJson, installedVersion, channel, verifier);
        }

        /// <summary>Confirms the file on disk is byte-for-byte what the manifest promised.</summary>
        public static bool VerifyPackage(string packagePath, UpdateRelease release, out string detail)
        {
            detail = null;
            if (release == null)
            {
                detail = "缺少更新清单信息。";
                return false;
            }

            FileInfo info;
            try
            {
                info = new FileInfo(packagePath);
                if (!info.Exists)
                {
                    detail = "更新包不存在：" + packagePath;
                    return false;
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                detail = "无法访问更新包：" + ex.Message;
                return false;
            }

            if (info.Length != release.Size)
            {
                detail = string.Format(CultureInfo.InvariantCulture,
                    "更新包大小不符（应为 {0} 字节，实为 {1} 字节）。", release.Size, info.Length);
                return false;
            }

            string actual;
            try
            {
                actual = ComputeSha256(packagePath);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                detail = "无法读取更新包：" + ex.Message;
                return false;
            }

            if (!string.Equals(actual, release.Sha256, StringComparison.Ordinal))
            {
                detail = string.Format(CultureInfo.InvariantCulture,
                    "更新包 SHA-256 不符（应为 {0}，实为 {1}）。", release.Sha256, actual);
                return false;
            }
            return true;
        }

        public static string ComputeSha256(string path)
        {
            using (var sha = SHA256.Create())
            using (var stream = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 64, FileOptions.SequentialScan))
            {
                byte[] hash = sha.ComputeHash(stream);
                var builder = new StringBuilder(64);
                foreach (byte b in hash)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

        /// <summary>
        /// Deletes every staged version except the one named.
        ///
        /// Only direct children whose names parse as versions are touched, so a
        /// mangled root cannot turn this into an arbitrary delete. Failures are
        /// ignored: leftover bytes are a disk-space annoyance, while throwing
        /// here would abort an update that is otherwise fine.
        /// </summary>
        public static void Prune(string root, string keepVersion)
        {
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root))
            {
                return;
            }
            string[] children;
            try
            {
                children = Directory.GetDirectories(root);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                return;
            }

            foreach (string child in children)
            {
                string name = Path.GetFileName(child);
                if (!ReleaseVersion.TryParse(name, out _))
                {
                    continue;
                }
                if (keepVersion != null && string.Equals(name, keepVersion, StringComparison.Ordinal))
                {
                    continue;
                }
                try
                {
                    Directory.Delete(child, recursive: true);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                }
            }
        }
    }
}
