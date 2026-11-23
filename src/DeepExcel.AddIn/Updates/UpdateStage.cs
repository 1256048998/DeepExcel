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

        // ------------------------------------------------------------------
        // Rollback slot
        //
        //     <root>\rollback\<version>\update.json + DeepExcel.Setup.exe
        //
        // The last version confirmed to start, kept so a failed update can be
        // undone by that version's own installer. Its own installer, rather
        // than a copy of the install directory, because the install directory
        // also holds the user's snapshots, history and workbooks: restoring a
        // copy would roll those back too. One version only: a rollback only
        // ever goes back one step, and each slot is a full installer.
        //
        // Not a version-named child of <root>, so Prune() never touches it.

        public const string RollbackDirectoryName = "rollback";

        public static string RollbackDirectoryFor(string root, string version)
        {
            return DirectoryFor(Path.Combine(root, RollbackDirectoryName), version);
        }

        /// <summary>
        /// Copies a staged version into the rollback slot, verified, and only
        /// then drops whatever the slot held before. Returns false -- leaving
        /// the previous slot untouched -- when the stage is missing or does not
        /// verify. Never throws.
        /// </summary>
        public static bool KeepForRollback(
            string root, string version, string channel, IManifestVerifier verifier)
        {
            string target = null;
            try
            {
                string source = FindVersionDirectory(root, version);
                if (source == null)
                {
                    return false;
                }
                UpdateCheckResult staged = LoadVerified(source, version, channel, verifier);
                if (!staged.Accepted)
                {
                    return false;
                }

                target = RollbackDirectoryFor(root, staged.Release.Version);
                string incoming = target + PartialSuffix;
                if (Directory.Exists(incoming)) Directory.Delete(incoming, recursive: true);
                Directory.CreateDirectory(incoming);
                File.Copy(Path.Combine(source, PackageFileName), Path.Combine(incoming, PackageFileName));
                File.Copy(Path.Combine(source, ManifestFileName), Path.Combine(incoming, ManifestFileName));
                // The copy is what will be run, so the copy is what gets checked.
                if (!LoadVerified(incoming, version, channel, verifier).Accepted)
                {
                    Directory.Delete(incoming, recursive: true);
                    return false;
                }

                if (Directory.Exists(target)) Directory.Delete(target, recursive: true);
                Directory.Move(incoming, target);

                foreach (string other in Directory.GetDirectories(Path.Combine(root, RollbackDirectoryName)))
                {
                    if (ReleaseVersion.TryParse(Path.GetFileName(other), out _) &&
                        !string.Equals(other, target, StringComparison.OrdinalIgnoreCase))
                    {
                        try { Directory.Delete(other, recursive: true); }
                        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { }
                    }
                }
                return true;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException ||
                                       ex is ArgumentException)
            {
                return false;
            }
        }

        /// <summary>
        /// The rollback package for <paramref name="installedVersion"/>, fully
        /// re-verified: signature, product, channel, exact version, and the
        /// package bytes. Anything less is a rejection, never an exception.
        /// </summary>
        public static UpdateCheckResult LoadRollback(
            string root, string installedVersion, string channel, IManifestVerifier verifier)
        {
            string directory;
            try
            {
                directory = FindVersionDirectory(Path.Combine(root, RollbackDirectoryName), installedVersion);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException ||
                                       ex is ArgumentException)
            {
                directory = null;
            }
            if (directory == null)
            {
                return UpdateCheckResult.Reject(
                    UpdateRejection.MalformedJson, "没有 v" + installedVersion + " 的回滚包。");
            }
            return LoadVerified(directory, installedVersion, channel, verifier);
        }

        private static UpdateCheckResult LoadVerified(
            string directory, string installedVersion, string channel, IManifestVerifier verifier)
        {
            string manifestPath = Path.Combine(directory, ManifestFileName);
            string manifestJson;
            try
            {
                var info = new FileInfo(manifestPath);
                if (!info.Exists)
                {
                    return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "缺少回滚清单。");
                }
                if (info.Length > UpdateManifest.MaxManifestBytes)
                {
                    return UpdateCheckResult.Reject(UpdateRejection.TooLarge, "回滚清单过大。");
                }
                manifestJson = File.ReadAllText(manifestPath, new UTF8Encoding(false));
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                return UpdateCheckResult.Reject(UpdateRejection.MalformedJson, "读取回滚清单失败：" + ex.Message);
            }

            UpdateCheckResult result = UpdateManifest.EvaluateInstalled(
                manifestJson, installedVersion, channel, verifier);
            if (!result.Accepted)
            {
                return result;
            }
            if (!VerifyPackage(Path.Combine(directory, PackageFileName), result.Release, out string detail))
            {
                return UpdateCheckResult.Reject(UpdateRejection.BadDigest, detail);
            }
            return result;
        }

        /// <summary>
        /// The child of <paramref name="parent"/> named for <paramref name="version"/>.
        ///
        /// Matched by version rather than by string: the running build reports
        /// "0.6.1.0" and the manifest that installed it said "0.6.1".
        /// </summary>
        private static string FindVersionDirectory(string parent, string version)
        {
            if (!ReleaseVersion.TryParse(version, out ReleaseVersion wanted) || !Directory.Exists(parent))
            {
                return null;
            }
            foreach (string child in Directory.GetDirectories(parent))
            {
                if (ReleaseVersion.TryParse(Path.GetFileName(child), out ReleaseVersion found) &&
                    found == wanted)
                {
                    return child;
                }
            }
            return null;
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
