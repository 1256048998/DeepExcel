using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Preview;
using DeepExcel.AddIn.Sidecar;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 等待确认的试跑。试跑用的是某一时刻的副本：用户之后改到了试跑涉及的格，
    /// 面板上的结果就不再代表真实执行会发生什么，要提示重新试跑。
    /// 只在 UI 线程上使用。
    /// </summary>
    public sealed class LabTrialTracker
    {
        private sealed class Pending
        {
            public string WorkbookKey;
            public bool Running = true;
            public bool Stale;
            public List<CellRect> Affected = new List<CellRect>();
            public List<CellRect> EditsWhileRunning = new List<CellRect>();
            public object Context;
        }

        private readonly Dictionary<string, Pending> _pending = new Dictionary<string, Pending>();

        public void Start(string requestId, string workbookKey, object context = null)
        {
            _pending[requestId] = new Pending { WorkbookKey = workbookKey, Context = context };
        }

        public bool IsPending(string requestId) => _pending.ContainsKey(requestId);

        public object ContextOf(string requestId) =>
            _pending.TryGetValue(requestId, out var p) ? p.Context : null;

        /// <summary>试跑出结果。返回 true 表示试跑期间用户已经改到了它涉及的格。</summary>
        public bool Complete(string requestId, IEnumerable<CellRect> affected)
        {
            if (!_pending.TryGetValue(requestId, out var p)) return false;
            p.Running = false;
            p.Affected = affected?.ToList() ?? new List<CellRect>();
            p.Stale = p.EditsWhileRunning.Any(edit => p.Affected.Any(edit.Overlaps));
            p.EditsWhileRunning.Clear();
            return p.Stale;
        }

        /// <summary>用户改了某个区域：返回因此刚刚过期的试跑（每个试跑只报一次）</summary>
        public List<string> RecordUserEdit(string workbookKey, CellRect edit)
        {
            var nowStale = new List<string>();
            foreach (var kv in _pending)
            {
                var p = kv.Value;
                if (!string.Equals(p.WorkbookKey, workbookKey, StringComparison.Ordinal)) continue;
                if (p.Running) { p.EditsWhileRunning.Add(edit); continue; }
                if (p.Stale || !p.Affected.Any(edit.Overlaps)) continue;
                p.Stale = true;
                nowStale.Add(kv.Key);
            }
            return nowStale;
        }

        public void Remove(string requestId) => _pending.Remove(requestId);

        public void RemoveWorkbook(string workbookKey)
        {
            foreach (var id in _pending.Where(kv => kv.Value.WorkbookKey == workbookKey).Select(kv => kv.Key).ToList())
            {
                _pending.Remove(id);
            }
        }
    }

    /// <summary>
    /// 副本试跑接入权限确认：execute_vba 要用户确认时，先在副本上跑一遍，
    /// 确认面板里给出「试跑结果」而不是一句「无法预览」。
    /// 用户允许后，真实工作簿上执行的仍是同一段代码（照常先快照），副本从不拷回去。
    /// </summary>
    public partial class MessageBridge
    {
        private readonly LabTrialTracker _labTrials = new LabTrialTracker();

        private sealed class LabTrialContext
        {
            public PythonSidecar Sender;
            public string Tool;
            public Dictionary<string, object> Args;
        }

        /// <summary>副本存不了的格式：存出来的副本和原文件不是一回事</summary>
        private static readonly HashSet<string> UnsupportedCopyExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".csv", ".txt", ".prn", ".dif", ".slk" };

        /// <summary>
        /// 能试跑就把副本存好、在后台开跑，返回 true：确认面板等结果出来再弹。
        /// 返回 false 时按原来的方式（无法预览 + 快照）确认。
        /// </summary>
        private bool TryStartLabTrial(WorkbookSession session, PythonSidecar sender, string requestId,
            string tool, Dictionary<string, object> args)
        {
            if (!string.Equals(tool, "execute_vba", StringComparison.OrdinalIgnoreCase)) return false;
            if (!(args != null && args.TryGetValue("code", out var codeObj) && codeObj is string code) ||
                string.IsNullOrWhiteSpace(code))
            {
                return false;
            }

            string copyDir = null;
            try
            {
                var wb = _excelApp.ActiveWorkbook;
                if (wb == null) return false;
                var name = wb.Name;
                var ext = Path.GetExtension(name);
                if (UnsupportedCopyExtensions.Contains(ext)) return false;
                if (string.IsNullOrEmpty(ext)) name += wb.HasVBProject ? ".xlsm" : ".xlsx";

                // 副本只在 UI 线程上从真实工作簿取一次；之后试跑线程只碰副本进程
                copyDir = LabProcessRegistry.NewCopyDirectory();
                var copyPath = Path.Combine(copyDir, name);
                var clock = Stopwatch.StartNew();
                wb.SaveCopyAs(copyPath);
                Logger.Instance.Info("MessageBridge", $"Lab copy saved in {clock.ElapsedMilliseconds}ms");

                var request = new LabRequest
                {
                    CopyPath = copyPath,
                    Code = code,
                    UserExcelPid = Process.GetCurrentProcess().Id,
                };
                try { request.ActiveSheet = (_excelApp.ActiveSheet as Excel.Worksheet)?.Name; } catch (Exception) { }
                try { request.Selection = (_excelApp.Selection as Excel.Range)?.Address[false, false]; } catch (Exception) { }
                try { request.Calculation = _excelApp.Calculation; } catch (Exception) { }

                _labTrials.Start(requestId, session.WorkbookKey,
                    new LabTrialContext { Sender = sender, Tool = tool, Args = args });
                SendToSessionUi(session.WorkbookKey, "ui_event",
                    new { v = 1, kind = "status", text = "正在工作簿副本上试跑这段代码…", tool });

                var key = session.WorkbookKey;
                LabRunner.RunAsync(request).ContinueWith(task =>
                {
                    var trial = task.Status == System.Threading.Tasks.TaskStatus.RanToCompletion
                        ? task.Result
                        : new LabTrial { SkippedReason = "试跑出错：" + task.Exception?.GetBaseException().Message };
                    RunOnUiThread(() => FinishLabTrial(key, requestId, trial));
                });
                return true;
            }
            catch (Exception ex)
            {
                // 另存副本失败（受保护的视图、IRM、磁盘满……）：退回普通确认
                Logger.Instance.Warning("MessageBridge", "Lab trial could not start: " + ex.Message);
                _labTrials.Remove(requestId);
                LabProcessRegistry.DeleteCopyDirectory(copyDir);
                return false;
            }
        }

        private void RunOnUiThread(Action action)
        {
            try
            {
                if (_uiControl != null && _uiControl.IsHandleCreated && !_uiControl.IsDisposed)
                {
                    _uiControl.BeginInvoke(action);
                }
                else
                {
                    Logger.Instance.Warning("MessageBridge", "Lab result dropped: UI control unavailable");
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "Lab result marshal failed: " + ex.Message);
            }
        }

        private void FinishLabTrial(string workbookKey, string requestId, LabTrial trial)
        {
            if (!(_labTrials.ContextOf(requestId) is LabTrialContext context)) return;
            if (!_sessions.TryGetValue(workbookKey, out var session) || session.Sidecar != context.Sender)
            {
                // 会话已关闭或侧车已重启：回一个拒绝，侧车还在的话 hook 不用等到超时
                _labTrials.Remove(requestId);
                try { context.Sender?.SendPermissionResponse(requestId, "deny"); } catch (Exception) { }
                return;
            }
            if (!session.IsBusy)
            {
                // 用户在试跑期间停止了任务：别再弹确认；回一个拒绝，免得 hook 等到超时
                _labTrials.Remove(requestId);
                context.Sender.SendPermissionResponse(requestId, "deny");
                return;
            }

            bool staleAlready = _labTrials.Complete(requestId, trial.Diff?.AffectedRects);
            var preview = TrialPreview(context.Tool, trial);
            SendToSessionUi(workbookKey, "ui_event", new { v = 1, kind = "status", text = "", tool = context.Tool });
            SendToSessionUi(workbookKey, "permission_request", new
            {
                request_id = requestId,
                tool = context.Tool,
                args = context.Args,
                preview = PreviewPayload(preview, trial, staleAlready),
            });
        }

        /// <summary>试跑结果换成确认面板用的预览。没跑成的退回「无法预览」，并说明为什么没跑成。</summary>
        internal static ChangePreview TrialPreview(string tool, LabTrial trial)
        {
            if (trial == null || (!trial.Ran && !trial.TimedOut))
            {
                var reason = PreviewPolicy.ReasonFor(tool) ?? "无法预先推算它会改动什么";
                if (!string.IsNullOrEmpty(trial?.SkippedReason)) reason += "（副本试跑未能进行：" + trial.SkippedReason + "）";
                return ChangePreview.NotPreviewable(tool, reason);
            }
            var preview = new ChangePreview
            {
                ToolName = tool,
                Previewable = true,
                Changes = trial.Diff.Changes,
                AffectedCells = trial.Diff.AffectedCells,
            };
            foreach (var sheet in trial.Diff.AddedSheets) preview.Warnings.Add("新建工作表：" + sheet);
            foreach (var sheet in trial.Diff.RemovedSheets) preview.Warnings.Add("删除工作表：" + sheet);
            if (trial.ErrorsAfter > trial.ErrorsBefore)
            {
                preview.Warnings.Add($"公式错误从 {trial.ErrorsBefore} 个增加到 {trial.ErrorsAfter} 个");
            }
            return preview;
        }

        private static object PreviewPayload(ChangePreview preview, LabTrial trial, bool stale)
        {
            bool ran = trial != null && (trial.Ran || trial.TimedOut);
            return new
            {
                previewable = preview.Previewable,
                reason = preview.NotPreviewableReason,
                summary = preview.Summary(),
                affected_cells = preview.AffectedCells,
                truncated = preview.IsTruncated,
                structural = preview.IsStructural,
                formulas_overwritten = preview.FormulasOverwritten,
                deleted_rows = preview.DeletedRows,
                deleted_columns = preview.DeletedColumns,
                warnings = preview.Warnings,
                changes = preview.Changes.Select(c => new
                {
                    address = c.Address,
                    before = c.Before,
                    after = c.After,
                    kind = c.Kind.ToString().ToLowerInvariant(),
                    overwrites_formula = c.OverwritesFormula
                }),
                trial = !ran ? null : new
                {
                    success = trial.Success,
                    error = trial.Error,
                    timed_out = trial.TimedOut,
                    duration_ms = trial.DurationMs,
                    errors_before = trial.ErrorsBefore,
                    errors_after = trial.ErrorsAfter,
                    note = trial.ExecutorNote,
                    not_representative = trial.NotRepresentative,
                    stale,
                },
            };
        }

        /// <summary>SheetChange 里调用：用户改到了试跑涉及的格，面板提示结果已过期</summary>
        private void CheckLabTrialsStale(string workbookKey, CellRect edit)
        {
            foreach (var requestId in _labTrials.RecordUserEdit(workbookKey, edit))
            {
                SendToSessionUi(workbookKey, "permission_preview_stale", new
                {
                    request_id = requestId,
                    message = $"你改动了 {edit.ToA1()}，和试跑改动的区域重叠，试跑结果可能已不准确",
                });
            }
        }

        /// <summary>面板上点「重新试跑」：用当前工作簿重新存副本再跑一遍，结果替换面板里的那份</summary>
        private string HandleRerunTrial(WorkbookSession session, Message msg)
        {
            try
            {
                var requestId = msg.Payload.Value.GetProperty("request_id").GetString();
                if (!(_labTrials.ContextOf(requestId) is LabTrialContext context))
                {
                    return MakeError("这次确认已经结束，无法重新试跑");
                }
                _labTrials.Remove(requestId);
                if (!TryStartLabTrial(session, context.Sender, requestId, context.Tool, context.Args))
                {
                    // 起不来就按原来的方式确认，面板不能一直空着
                    SendToSessionUi(session.WorkbookKey, "permission_request", new
                    {
                        request_id = requestId,
                        tool = context.Tool,
                        args = context.Args,
                        preview = PreviewPayload(TrialPreview(context.Tool, null), null, false),
                    });
                }
                return MakeResponse("ack", new { received = true, kind = "rerun_trial" });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleRerunTrial failed", ex);
                return MakeError("重新试跑失败：" + ex.Message);
            }
        }
    }
}
