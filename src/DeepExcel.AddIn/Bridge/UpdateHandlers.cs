using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Account;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Updates;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Update-related bridge messages.
    ///
    /// Kept apart from the rest of the bridge for the same reason the account
    /// handlers are: an install with no update feed configured never reaches any
    /// of this.
    /// </summary>
    public partial class MessageBridge
    {
        /// <summary>
        /// Wait before the first check.
        ///
        /// Excel's startup is already contended — COM registration, the sidecar
        /// pre-warm, WebView2 — and the network stack is not reliably up the
        /// instant an add-in loads. Nothing about an update is urgent enough to
        /// join that queue.
        /// </summary>
        private static readonly TimeSpan FirstCheckDelay = TimeSpan.FromMinutes(1);

        /// <summary>
        /// Then every six hours, for as long as Excel stays open.
        ///
        /// Checking once at startup was wrong for the way this product is
        /// actually used: Excel commonly stays open for days, so a single failed
        /// check at 9am meant no update for the rest of the week, and a release
        /// published at noon reached nobody until they happened to restart.
        /// </summary>
        private static readonly TimeSpan RecheckInterval = TimeSpan.FromHours(6);

        private UpdateService _updateService;
        private Timer _updateTimer;
        private int _updateCheckRunning;
        private readonly object _updateGate = new object();

        private UpdateService Updates
        {
            get
            {
                lock (_updateGate)
                {
                    return _updateService ?? (_updateService = new UpdateService(telemetry: ReportUpdateEvent));
                }
            }
        }

        /// <summary>
        /// Sends one update event, if this install has an account to send it under.
        ///
        /// Every field is a version number or a fixed token; see
        /// <see cref="UpdateEvent"/> for why <c>StatusDetail</c> is not among
        /// them.
        /// </summary>
        private void ReportUpdateEvent(UpdateEvent update)
        {
            if (update == null)
            {
                return;
            }
            TelemetryReporter reporter = Telemetry;
            if (reporter == null)
            {
                // Signed out. The update itself still works — the feed is
                // deliberately unauthenticated — we just cannot count it.
                return;
            }

            var payload = new Dictionary<string, object>
            {
                ["phase"] = update.Phase,
                ["outcome"] = update.Outcome,
                ["from_version"] = update.FromVersion,
            };
            if (!string.IsNullOrEmpty(update.ReasonCode)) payload["reason_code"] = update.ReasonCode;
            if (!string.IsNullOrEmpty(update.ToVersion)) payload["to_version"] = update.ToVersion;
            if (update.DurationMs > 0) payload["duration_ms"] = update.DurationMs;

            reporter.Record("update_event", payload);
        }

        /// <summary>
        /// Starts the update cycle: report any completed upgrade, then check on
        /// a schedule.
        ///
        /// Called from add-in startup. It must never throw, never block, and
        /// never be the reason Excel is slow to open a workbook: the entire
        /// value of the updater is that the user does not participate in it.
        /// </summary>
        public void BeginBackgroundUpdateCheck()
        {
            try
            {
                UpdateService service = Updates;

                // Runs even when updating is switched off or unconfigured: the
                // receipt records an upgrade that already happened, and it is
                // the only evidence that the update path works end to end.
                Task.Run(() =>
                {
                    try { service.ReportPendingUpgrade(); }
                    catch (Exception ex)
                    {
                        Logger.Instance.Warning("Update", "升级回执上报失败：" + ex.Message);
                    }
                });

                if (!service.IsUsable)
                {
                    Logger.Instance.Debug("Update", "跳过更新检查：" + service.StatusDetail);
                    return;
                }

                lock (_updateGate)
                {
                    if (_updateTimer != null)
                    {
                        return;
                    }
                    _updateTimer = new Timer(
                        _ => RunUpdateCheck(), null, FirstCheckDelay, RecheckInterval);
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("Update", "无法启动更新检查：" + ex.Message);
            }
        }

        /// <summary>One scheduled check. Overlapping ticks are dropped, not queued.</summary>
        private void RunUpdateCheck()
        {
            // A check on a slow link can outlast the interval. Re-entering would
            // have two downloads writing the same staging file.
            if (Interlocked.CompareExchange(ref _updateCheckRunning, 1, 0) != 0)
            {
                return;
            }
            try
            {
                UpdateService service = Updates;

                // Ready: already downloaded and verified, waiting on the user.
                // Blocked: failed to install repeatedly, so re-checking would
                // only re-stage something that cannot install on this machine.
                if (service.Status == UpdateStatus.Ready || service.Status == UpdateStatus.Blocked)
                {
                    return;
                }
                service.CheckAndStageAsync(null, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("Update", "后台更新检查异常：" + ex.Message);
            }
            finally
            {
                Interlocked.Exchange(ref _updateCheckRunning, 0);
            }
        }

        private void StopUpdateTimer()
        {
            lock (_updateGate)
            {
                try { _updateTimer?.Dispose(); }
                catch (Exception) { }
                _updateTimer = null;
            }
        }

        /// <summary>What the panel renders. Cheap, and safe to poll.</summary>
        private string HandleUpdateStatus()
        {
            try
            {
                UpdateService service = Updates;
                StagedUpdate staged = service.Staged;
                return MakeResponse("update_status", new
                {
                    state = service.Status.ToString().ToLowerInvariant(),
                    detail = service.StatusDetail,
                    installed_version = UpdateService.InstalledVersion,
                    available_version = staged?.Release?.Version,
                    notes = staged?.Release?.Notes,
                    size = staged?.Release?.Size ?? 0,
                    progress = service.Progress,
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("Bridge", "HandleUpdateStatus failed", ex);
                return MakeError("读取更新状态失败");
            }
        }

        /// <summary>An explicit "check now" from the panel.</summary>
        private string HandleUpdateCheck()
        {
            try
            {
                UpdateService service = Updates;
                if (!service.IsUsable)
                {
                    return MakeResponse("update_check", new
                    {
                        state = "disabled",
                        detail = service.StatusDetail,
                    });
                }
                UpdateStatus status = RunSync(() =>
                    service.CheckAndStageAsync(null, CancellationToken.None));
                return MakeResponse("update_check", new
                {
                    state = status.ToString().ToLowerInvariant(),
                    detail = service.StatusDetail,
                    available_version = service.Staged?.Release?.Version,
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("Bridge", "HandleUpdateCheck failed", ex);
                return MakeError("检查更新失败：" + ex.Message);
            }
        }

        /// <summary>
        /// Hands the staged update to the updater and asks Excel to close.
        ///
        /// Quitting goes through Excel's own path, so unsaved work still
        /// prompts. If the user cancels that prompt the updater simply times out
        /// waiting and leaves the staged package for next time — nothing here
        /// forces an application with unsaved work to close.
        /// </summary>
        private string HandleUpdateInstall()
        {
            try
            {
                if (!Updates.LaunchInstaller(out string error))
                {
                    return MakeError(error);
                }
                RequestHostShutdown();
                return MakeResponse("update_install", new
                {
                    started = true,
                    message = "关闭 Excel 后将自动安装更新。",
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("Bridge", "HandleUpdateInstall failed", ex);
                return MakeError("启动更新失败：" + ex.Message);
            }
        }

        /// <summary>
        /// Asks the host application to close, and does not insist.
        ///
        /// Goes through Excel's own Quit so unsaved workbooks still prompt, and
        /// the user may well answer no. That is a valid answer: the updater
        /// times out waiting, the package stays staged, and it installs at the
        /// next restart. Nothing here closes an application over the user's
        /// work — this call asks, and the request is not repeated.
        /// </summary>
        private void RequestHostShutdown()
        {
            try
            {
                _excelApp?.Quit();
            }
            catch (Exception ex)
            {
                // COM throws here when the user cancels the save prompt.
                Logger.Instance.Info("Update", "Excel 未关闭，更新将在下次重启时安装：" + ex.Message);
            }
        }
    }
}
