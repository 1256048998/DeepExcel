using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Updates
{
    public enum UpdateStatus
    {
        /// <summary>No compiled-in key, no feed configured, or switched off.</summary>
        Disabled,
        Idle,
        Checking,
        UpToDate,
        Downloading,
        /// <summary>Downloaded and verified; waiting for the user to restart.</summary>
        Ready,
        Failed,
    }

    /// <summary>
    /// Drives the update cycle for the add-in: check, download, verify, stage.
    ///
    /// The point of the whole feature, from the roadmap: SmartScreen only warns
    /// about files carrying a mark-of-the-web, which a browser download has and
    /// a download made by an already-installed program does not. So the first
    /// install hurts once and every later version arrives without friction.
    /// That is also why nothing here is allowed to be noisy — a failed check
    /// must cost the user nothing and must never block startup.
    ///
    /// Installing is deliberately not automatic. The installer needs Excel
    /// closed, and closing an application with the user's unsaved work in it is
    /// not a decision this code gets to make.
    /// </summary>
    public sealed class UpdateService
    {
        private const string LogCategory = "Update";

        private readonly Func<UpdateDownloader> _downloaderFactory;
        private readonly IManifestVerifier _verifier;
        private readonly string _stageRoot;
        private readonly object _gate = new object();

        public UpdateService(
            UpdateOptions options = null,
            IManifestVerifier verifier = null,
            Func<UpdateDownloader> downloaderFactory = null,
            string stageRoot = null)
        {
            Options = options ?? UpdateOptions.FromConfig();
            _verifier = verifier ?? EmbeddedUpdateKey.Verifier;
            _downloaderFactory = downloaderFactory ?? (() => new UpdateDownloader());
            _stageRoot = stageRoot ?? UpdateStage.DefaultRoot();
            Status = IsUsable ? UpdateStatus.Idle : UpdateStatus.Disabled;
            StatusDetail = IsUsable ? null : DisabledReason();
        }

        public UpdateOptions Options { get; }

        public UpdateStatus Status { get; private set; }

        public string StatusDetail { get; private set; }

        /// <summary>0..1 while downloading.</summary>
        public double Progress { get; private set; }

        public StagedUpdate Staged { get; private set; }

        /// <summary>
        /// The running build's version.
        ///
        /// Read from the assembly rather than from a constant so it cannot
        /// disagree with what was actually shipped — the add-in's version is
        /// already the single source of truth for registration and packaging.
        /// </summary>
        public static string InstalledVersion
        {
            get
            {
                Version version = typeof(UpdateService).Assembly.GetName().Version;
                return version == null ? "0.0.0.0" : version.ToString();
            }
        }

        public bool IsUsable =>
            Options.Enabled && _verifier.IsConfigured && !string.IsNullOrWhiteSpace(Options.FeedUrl);

        private string DisabledReason()
        {
            if (!Options.Enabled) return "自动更新已在设置中关闭。";
            if (!_verifier.IsConfigured) return "此版本未内置更新签名公钥，自动更新不可用。";
            return "未配置更新源地址，自动更新不可用。";
        }

        /// <summary>
        /// Checks the feed and, when something newer is signed for this channel,
        /// downloads and verifies it into the staging directory.
        ///
        /// Never throws. A machine that cannot reach the update server is a
        /// machine that keeps working on the version it has.
        /// </summary>
        public async Task<UpdateStatus> CheckAndStageAsync(
            IProgress<double> progress = null, CancellationToken cancellationToken = default)
        {
            if (!IsUsable)
            {
                return SetStatus(UpdateStatus.Disabled, DisabledReason());
            }

            SetStatus(UpdateStatus.Checking, null);
            try
            {
                string manifestJson;
                using (UpdateDownloader downloader = _downloaderFactory())
                {
                    manifestJson = await downloader
                        .FetchManifestAsync(Options.FeedUrl, cancellationToken)
                        .ConfigureAwait(false);

                    UpdateCheckResult check = UpdateManifest.Evaluate(
                        manifestJson, InstalledVersion, Options.Channel, _verifier);

                    if (!check.Accepted)
                    {
                        // NotNewer is the ordinary outcome, not a problem.
                        if (check.Rejection == UpdateRejection.NotNewer)
                        {
                            UpdateStage.Prune(_stageRoot, null);
                            return SetStatus(UpdateStatus.UpToDate, check.Detail);
                        }
                        Logger.Instance.Warning(
                            LogCategory,
                            string.Format(CultureInfo.InvariantCulture,
                                "更新清单被拒绝（{0}）：{1}", check.Rejection, check.Detail));
                        return SetStatus(UpdateStatus.Failed, check.Detail);
                    }

                    UpdateRelease release = check.Release;
                    string stageDirectory = UpdateStage.DirectoryFor(_stageRoot, release.Version);
                    Directory.CreateDirectory(stageDirectory);
                    string packagePath = Path.Combine(stageDirectory, UpdateStage.PackageFileName);

                    // A previous session may already have finished this exact
                    // version; re-verify rather than trusting that it is intact.
                    if (!UpdateStage.VerifyPackage(packagePath, release, out _))
                    {
                        SetStatus(UpdateStatus.Downloading, release.Version);
                        var relay = new Progress<double>(value =>
                        {
                            Progress = value;
                            progress?.Report(value);
                        });
                        await downloader
                            .DownloadPackageAsync(release, packagePath, relay, cancellationToken)
                            .ConfigureAwait(false);

                        if (!UpdateStage.VerifyPackage(packagePath, release, out string detail))
                        {
                            TryDelete(packagePath);
                            Logger.Instance.Error(LogCategory, "更新包校验失败：" + detail);
                            return SetStatus(UpdateStatus.Failed, detail);
                        }
                    }

                    // Written last: the manifest's presence is what marks the
                    // stage complete, and it is stored verbatim so the updater
                    // re-checks the signature rather than trusting this process.
                    File.WriteAllText(
                        Path.Combine(stageDirectory, UpdateStage.ManifestFileName),
                        release.RawManifest, new System.Text.UTF8Encoding(false));

                    UpdateStage.Prune(_stageRoot, release.Version);

                    lock (_gate)
                    {
                        Staged = new StagedUpdate
                        {
                            Directory = stageDirectory,
                            PackagePath = packagePath,
                            Release = release,
                        };
                    }
                    Progress = 1.0;
                    Logger.Instance.Info(
                        LogCategory, "已就绪：v" + release.Version + "（重启 Excel 后安装）");
                    return SetStatus(UpdateStatus.Ready, release.Version);
                }
            }
            catch (OperationCanceledException)
            {
                return SetStatus(UpdateStatus.Idle, null);
            }
            catch (UpdateTransportException ex)
            {
                Logger.Instance.Warning(LogCategory, ex.Message);
                return SetStatus(UpdateStatus.Failed, ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error(LogCategory, "检查更新失败", ex);
                return SetStatus(UpdateStatus.Failed, ex.Message);
            }
        }

        /// <summary>
        /// Hands a staged update to DeepExcel.Updater.exe and returns.
        ///
        /// The updater is copied into the staging directory and launched from
        /// there: Inno Setup cannot overwrite a running executable, so an
        /// updater started from the install directory would block the install it
        /// just triggered. The updater waits for this process to exit on its
        /// own; it never terminates Excel.
        /// </summary>
        public bool LaunchInstaller(out string error)
        {
            error = null;
            StagedUpdate staged;
            lock (_gate)
            {
                staged = Staged;
            }
            if (staged == null)
            {
                error = "没有已就绪的更新。";
                return false;
            }

            try
            {
                string installDirectory = Path.GetDirectoryName(
                    new Uri(typeof(UpdateService).Assembly.CodeBase).LocalPath);
                string source = Path.Combine(installDirectory, UpdateStage.UpdaterFileName);
                if (!File.Exists(source))
                {
                    error = "缺少更新程序：" + source;
                    Logger.Instance.Error(LogCategory, error);
                    return false;
                }

                string runner = Path.Combine(staged.Directory, UpdateStage.UpdaterFileName);
                File.Copy(source, runner, overwrite: true);

                var arguments = new System.Text.StringBuilder();
                arguments.Append("--stage \"").Append(staged.Directory).Append("\"");
                arguments.Append(" --installed-version ").Append(InstalledVersion);
                arguments.Append(" --channel ").Append(Options.Channel);
                arguments.Append(" --wait-pid ").Append(
                    Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture));
                // The user just asked for this and is waiting, so a failure is
                // worth a dialog. Nothing else passes this flag: an updater that
                // can pop a modal box unprompted would interrupt people at
                // random for reasons they cannot place.
                arguments.Append(" --notify");

                string host = TryGetHostExecutable();
                if (host != null)
                {
                    arguments.Append(" --relaunch \"").Append(host).Append("\"");
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = runner,
                    Arguments = arguments.ToString(),
                    WorkingDirectory = staged.Directory,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                });
                Logger.Instance.Info(LogCategory, "已启动更新程序，等待 Excel 退出");
                return true;
            }
            catch (Exception ex)
            {
                error = "启动更新程序失败：" + ex.Message;
                Logger.Instance.Error(LogCategory, error, ex);
                return false;
            }
        }

        private static string TryGetHostExecutable()
        {
            try
            {
                return Process.GetCurrentProcess().MainModule?.FileName;
            }
            catch (Exception)
            {
                // Not worth failing an update over; the user can reopen Excel.
                return null;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
            }
        }

        private UpdateStatus SetStatus(UpdateStatus status, string detail)
        {
            Status = status;
            StatusDetail = detail;
            return status;
        }
    }

    /// <summary>Where updates come from, and whether to look.</summary>
    public sealed class UpdateOptions
    {
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// Empty by default. Combined with the empty compiled-in public key this
        /// means a build that nobody configured never contacts anything — no
        /// default domain to squat, no update path to take over.
        /// </summary>
        public string FeedUrl { get; set; } = "";

        public string Channel { get; set; } = UpdateManifest.DefaultChannel;

        public static UpdateOptions FromConfig()
        {
            try
            {
                Config.UpdateSettings settings = Config.ConfigManager.Instance.Current?.Update;
                if (settings == null)
                {
                    return new UpdateOptions();
                }
                return new UpdateOptions
                {
                    Enabled = settings.Enabled,
                    FeedUrl = settings.FeedUrl ?? "",
                    Channel = string.IsNullOrWhiteSpace(settings.Channel)
                        ? UpdateManifest.DefaultChannel
                        : settings.Channel.Trim(),
                };
            }
            catch (Exception)
            {
                return new UpdateOptions();
            }
        }
    }
}
