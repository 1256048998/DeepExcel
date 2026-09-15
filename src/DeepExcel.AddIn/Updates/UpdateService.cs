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
        /// <summary>Installed and failed too many times; stop offering it.</summary>
        Blocked,
    }

    /// <summary>
    /// One reportable moment in the update cycle.
    ///
    /// Every field is a version number or a fixed token. No URLs, no paths, no
    /// exception text — those are the three things that reliably smuggle user
    /// data into telemetry, and the privacy line on this product forbids all of
    /// them. <see cref="UpdateService.StatusDetail"/> contains all three and is
    /// deliberately not part of this type.
    /// </summary>
    public sealed class UpdateEvent
    {
        /// <summary>check | download | launch | apply</summary>
        public string Phase { get; set; }

        /// <summary>up_to_date | ready | failed | blocked | installed | started</summary>
        public string Outcome { get; set; }

        /// <summary>An <see cref="UpdateRejection"/> name or an UpdateTransportException code.</summary>
        public string ReasonCode { get; set; }

        public string FromVersion { get; set; }
        public string ToVersion { get; set; }
        public long DurationMs { get; set; }
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

        /// <summary>How long the updater waits for Excel when a user asked for it.</summary>
        public const int UserInitiatedWaitSeconds = 900;

        private readonly Func<UpdateDownloader> _downloaderFactory;
        private readonly IManifestVerifier _verifier;
        private readonly string _stageRoot;
        private readonly UpdateJournal _journal;
        private readonly Action<UpdateEvent> _telemetry;
        private readonly object _gate = new object();

        public UpdateService(
            UpdateOptions options = null,
            IManifestVerifier verifier = null,
            Func<UpdateDownloader> downloaderFactory = null,
            string stageRoot = null,
            Action<UpdateEvent> telemetry = null)
        {
            Options = options ?? UpdateOptions.FromConfig();
            _verifier = verifier ?? EmbeddedUpdateKey.Verifier;
            _downloaderFactory = downloaderFactory ?? (() => new UpdateDownloader());
            _stageRoot = stageRoot ?? UpdateStage.DefaultRoot();
            _journal = new UpdateJournal(_stageRoot);
            // Injected rather than reached for, so this class stays testable
            // without an account, a network or a server.
            _telemetry = telemetry;
            Status = IsUsable ? UpdateStatus.Idle : UpdateStatus.Disabled;
            StatusDetail = IsUsable ? null : DisabledReason();
        }

        /// <summary>
        /// Reports an upgrade that already happened, if the updater left a note.
        ///
        /// This is the only way the product can answer "did anyone actually take
        /// the update" — the process that knows is the updater, and it exited
        /// before this build started. Safe to call regardless of whether
        /// updating is currently enabled: the upgrade is a past fact.
        /// </summary>
        public AppliedReceipt ReportPendingUpgrade()
        {
            try
            {
                AppliedReceipt receipt = _journal.TakeReceipt();
                if (receipt == null)
                {
                    return null;
                }
                // A version that installed is a version that will not be
                // offered again, so its failure count is spent history.
                _journal.ClearAttempts();
                Logger.Instance.Info(
                    LogCategory, "已完成升级：" + receipt.FromVersion + " → " + receipt.ToVersion);
                Emit("apply", "installed", null, receipt.FromVersion, receipt.ToVersion, 0);
                return receipt;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning(LogCategory, "读取升级回执失败：" + ex.Message);
                return null;
            }
        }

        private void Emit(
            string phase, string outcome, string reasonCode,
            string fromVersion, string toVersion, long durationMs)
        {
            if (_telemetry == null)
            {
                return;
            }
            try
            {
                _telemetry(new UpdateEvent
                {
                    Phase = phase,
                    Outcome = outcome,
                    ReasonCode = reasonCode,
                    FromVersion = fromVersion,
                    ToVersion = toVersion,
                    DurationMs = durationMs,
                });
            }
            catch (Exception ex)
            {
                // Telemetry is never allowed to be the reason an update fails.
                Logger.Instance.Warning(LogCategory, "更新遥测上报失败：" + ex.Message);
            }
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
            var clock = Stopwatch.StartNew();
            // Tracked so a transport failure is attributed to the step it
            // actually happened in. Reporting a failed download as a failed
            // check would point whoever reads the numbers at the wrong system:
            // the feed would look broken when the package host is.
            string phase = "check";
            string targetVersion = null;
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
                            _journal.ClearAttempts();
                            Emit("check", "up_to_date", null,
                                 InstalledVersion, null, clock.ElapsedMilliseconds);
                            return SetStatus(UpdateStatus.UpToDate, check.Detail);
                        }
                        Logger.Instance.Warning(
                            LogCategory,
                            string.Format(CultureInfo.InvariantCulture,
                                "更新清单被拒绝（{0}）：{1}", check.Rejection, check.Detail));
                        Emit("check", "failed", check.Rejection.ToString(),
                             InstalledVersion, null, clock.ElapsedMilliseconds);
                        return SetStatus(UpdateStatus.Failed, check.Detail);
                    }

                    UpdateRelease release = check.Release;

                    // An update that has already failed to install several times
                    // is not going to start working on the next restart. It is
                    // still staged and still verifies, so without this it would
                    // be offered again every single session — antivirus blocking
                    // the installer produces exactly that loop.
                    int attempts = _journal.AttemptsFor(release.Version);
                    if (attempts >= UpdateJournal.MaxAttempts)
                    {
                        Logger.Instance.Warning(
                            LogCategory,
                            string.Format(CultureInfo.InvariantCulture,
                                "v{0} 已连续 {1} 次安装未成功，停止自动提示", release.Version, attempts));
                        Emit("launch", "blocked", "max_attempts",
                             InstalledVersion, release.Version, clock.ElapsedMilliseconds);
                        // Staged is kept so the panel can name the version the
                        // user now has to install by hand. Status is what stops
                        // it being offered; LaunchInstaller re-checks anyway.
                        lock (_gate)
                        {
                            Staged = new StagedUpdate
                            {
                                Directory = UpdateStage.DirectoryFor(_stageRoot, release.Version),
                                PackagePath = Path.Combine(
                                    UpdateStage.DirectoryFor(_stageRoot, release.Version),
                                    UpdateStage.PackageFileName),
                                Release = release,
                            };
                        }
                        return SetStatus(
                            UpdateStatus.Blocked,
                            string.Format(CultureInfo.InvariantCulture,
                                "v{0} 已连续 {1} 次安装未成功，可能被安全软件拦截。请手动下载安装。",
                                release.Version, attempts));
                    }

                    string stageDirectory = UpdateStage.DirectoryFor(_stageRoot, release.Version);
                    Directory.CreateDirectory(stageDirectory);
                    string packagePath = Path.Combine(stageDirectory, UpdateStage.PackageFileName);

                    // A previous session may already have finished this exact
                    // version; re-verify rather than trusting that it is intact.
                    if (!UpdateStage.VerifyPackage(packagePath, release, out _))
                    {
                        phase = "download";
                        targetVersion = release.Version;
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
                            Emit("download", "failed", "package_digest_mismatch",
                                 InstalledVersion, release.Version, clock.ElapsedMilliseconds);
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
                    Emit("check", "ready", null,
                         InstalledVersion, release.Version, clock.ElapsedMilliseconds);
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
                // ex.Code, never ex.Message: the message quotes the feed URL.
                Emit(phase, "failed", ex.Code,
                     InstalledVersion, targetVersion, clock.ElapsedMilliseconds);
                return SetStatus(UpdateStatus.Failed, ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error(LogCategory, "检查更新失败", ex);
                // The exception type, not its message. An unexpected exception
                // is the most likely place for a path to leak into telemetry.
                Emit(phase, "failed", ex.GetType().Name,
                     InstalledVersion, targetVersion, clock.ElapsedMilliseconds);
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
            // Defence in depth: the panel hides the button once a version is
            // blocked, but the panel is not a security boundary and the check
            // that stops the loop belongs next to the thing it guards.
            if (_journal.IsExhausted(staged.Release.Version))
            {
                error = string.Format(
                    CultureInfo.InvariantCulture,
                    "v{0} 已连续 {1} 次安装未成功，请手动下载安装。",
                    staged.Release.Version, UpdateJournal.MaxAttempts);
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
                // Fifteen minutes, not the three-minute default. Closing Excel
                // means answering a save prompt per dirty workbook, and a user
                // who steps away mid-prompt should not come back to an update
                // that quietly gave up and left no sign it was ever running.
                arguments.Append(" --wait-seconds ").Append(
                    UserInitiatedWaitSeconds.ToString(CultureInfo.InvariantCulture));
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

                // Counted here, not when the updater reports back — by then this
                // process is gone. An installer killed by antivirus, a crash, a
                // silent failure: all of them still burn an attempt, which is
                // what stops the offer from repeating forever.
                _journal.RecordAttempt(staged.Release.Version);
                Logger.Instance.Info(
                    LogCategory,
                    string.Format(CultureInfo.InvariantCulture,
                        "已启动更新程序（第 {0} 次尝试），等待 Excel 退出",
                        _journal.AttemptsFor(staged.Release.Version)));
                Emit("launch", "started", null, InstalledVersion, staged.Release.Version, 0);
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
