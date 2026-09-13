using System;
using System.Threading;
using System.Threading.Tasks;
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
        private UpdateService _updateService;
        private Task _updateCheck;
        private readonly object _updateGate = new object();

        private UpdateService Updates
        {
            get
            {
                lock (_updateGate)
                {
                    return _updateService ?? (_updateService = new UpdateService());
                }
            }
        }

        /// <summary>
        /// Starts a background check, once per session.
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
                if (!service.IsUsable)
                {
                    Logger.Instance.Debug("Update", "跳过更新检查：" + service.StatusDetail);
                    return;
                }
                lock (_updateGate)
                {
                    if (_updateCheck != null)
                    {
                        return;
                    }
                    _updateCheck = Task.Run(async () =>
                    {
                        try
                        {
                            await service.CheckAndStageAsync(null, CancellationToken.None)
                                .ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Warning("Update", "后台更新检查异常：" + ex.Message);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("Update", "无法启动更新检查：" + ex.Message);
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
