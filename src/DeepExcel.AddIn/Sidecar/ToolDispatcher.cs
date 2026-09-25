using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;
using SnapshotScope = DeepExcel.AddIn.Executor.SnapshotScope;
using ExcelTarget = DeepExcel.AddIn.Executor.ExcelTarget;
using SnapshotAttempt = DeepExcel.AddIn.Executor.SnapshotAttempt;
using SnapshotManager = DeepExcel.AddIn.Executor.SnapshotManager;
using Stopwatch = System.Diagnostics.Stopwatch;
using WorkbookIdentity = DeepExcel.AddIn.Executor.WorkbookIdentity;
using DeepExcel.AddIn.Tools;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// 工具执行调度器：接收 sidecar 的 tool_call，在 STA 线程执行 Excel 操作
    /// 从 Orchestrator.ExecuteToolAsync 移植，保留 GetArg/GetArgArray 的 JsonElement 处理
    /// </summary>
    public class ToolDispatcher
    {
        private readonly IExcelActions _excel;
        private readonly Microsoft.Office.Interop.Excel.Application _excelApp;

        /// <summary>
        /// ★ 附件映射（fileName → absolutePath），由 MessageBridge 在创建 session 时注入。
        /// read_attachment 工具通过此映射查找附件文件路径。
        /// 引用 session.Attachments，session 增删附件时自动同步。
        /// </summary>
        public Dictionary<string, string> Attachments { get; set; }

        /// <summary>
        /// read_range 返回的最大行数限制。
        /// Claude Agent SDK 内部消息缓冲区限制 1MB（1048576 bytes），
        /// 超过会导致 "JSON message exceeded maximum buffer size" 崩溃。
        /// 200 行 × 20 列 × 平均 30 字节 ≈ 120KB，留足余量。
        /// </summary>
        private const int MaxReadRangeRows = 200;

        /// <summary>
        /// 共享的 JSON 序列化选项：注册 2D 数组转换器，避免序列化 RangeInfo 时崩溃。
        /// </summary>
        private static readonly JsonSerializerOptions _jsonOptions = BuildJsonOptions();

        private static JsonSerializerOptions BuildJsonOptions()
        {
            var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            opts.Converters.Add(new Object2DArrayConverter());
            opts.Converters.Add(new String2DArrayConverter());
            return opts;
        }

        public ToolDispatcher(IExcelActions excel, Microsoft.Office.Interop.Excel.Application excelApp = null)
        {
            _excel = excel;
            _excelApp = excelApp;
        }

        /// <summary>
        /// 本会话绑定的工作簿 key（由 WorkbookSession 注入，另存为后跟着变）。
        /// 为 null 时不做绑定检查（例如没有会话的调用路径）。
        /// </summary>
        public Func<string> BoundWorkbookKey { get; set; }

        /// <summary>本回合已建的自动备份：工作簿 key → 快照 ID。每条用户消息开始时清空。</summary>
        private readonly Dictionary<string, string> _turnBackups =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 会话内的检查点：每个写入步骤执行前各存一份（工作簿 key → 从旧到新）。之后的写入涉及
        /// 新的表时，扩大所有更早检查点的范围——回到某一步之前时，之后的修改也一起回退。
        /// </summary>
        private sealed class Checkpoint
        {
            public string Id;
            public bool Whole;
            public HashSet<string> Sheets;
        }

        private readonly Dictionary<string, List<Checkpoint>> _checkpoints =
            new Dictionary<string, List<Checkpoint>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>和快照保留数一致：更早的快照文件已被清理，记着也回不去</summary>
        private const int MaxCheckpointsTracked = SnapshotManager.MaxSnapshotsPerWorkbook;

        /// <summary>
        /// 一次备份超过这个时长（大工作簿），本回合后面的步骤不再各存一份，只靠回合备份，
        /// 免得每一步都多等几秒。这些步骤不带 checkpoint_id，面板也就不显示「回到这一步之前」。
        /// </summary>
        internal int SlowCheckpointThresholdMs { get; set; } = 1500;

        private bool _checkpointsSlowThisTurn;

        /// <summary>
        /// 一个回合最多单独存这么多份。快照按工作簿只保留 MaxSnapshotsPerWorkbook 份，一个回合
        /// 存满就会把自己的回合备份挤掉，模型的 rollback 就回不去了。
        /// </summary>
        internal const int MaxCheckpointsPerTurn = 12;
        private int _checkpointsThisTurn;

        /// <summary>新的用户回合开始：下一次写入会重新备份</summary>
        public void BeginTurn()
        {
            _turnBackups.Clear();
            _checkpointsSlowThisTurn = false;
            _checkpointsThisTurn = 0;
        }

        /// <summary>
        /// 同步执行工具（必须在 STA 主线程调用）
        ///
        /// 会改工作簿的工具（见 ToolMutationPolicy）先经过写入守卫：
        /// 每个用户回合第一次写入前整本备份一次，之后的写入只把涉及的表记进同一份快照；
        /// 备份失败、或活动工作簿已不是本会话绑定的那本，就不执行（fail-closed）。
        /// 以前只有 clean_data / execute_vba 会自动备份，其余写入都依赖模型记得调 create_snapshot。
        /// </summary>
        public ToolResult Execute(string toolName, Dictionary<string, object> args)
        {
            Logger.Instance.Info("ToolDispatcher", "Execute: " + toolName + ", args keys=" + (args == null ? "null" : string.Join(",", args.Keys)));

            IsExecuting = true;
            try
            {
                var refusal = ResolveTarget(toolName, out var target);
                if (refusal != null) return refusal;
                using (target)
                {
                    var result = ExecuteGuarded(toolName, args);
                    AttachUserEditNotice(result);
                    return result;
                }
            }
            finally
            {
                IsExecuting = false;
            }
        }

        /// <summary>
        /// 必须在前台窗口里做的工具：VBA / Python 代码里的 ActiveSheet、选区、截图、冻结窗格、
        /// 模拟按键都只认前台。用户切到别的工作簿时这些工具拒绝执行，其余工具照常写会话那本。
        /// </summary>
        internal static readonly HashSet<string> ForegroundOnlyTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "execute_vba", "execute_python", "execute_jsa", "send_keys",
            "read_selection", "screenshot_excel", "freeze_panes",
        };

        /// <summary>不碰工作簿内容的工具：会话那本关了也能用</summary>
        private static readonly HashSet<string> WorkbookIndependentTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "read_attachment", "rollback", "create_snapshot", "clarify_intent", "todo_write",
        };

        /// <summary>
        /// 目标工作簿 = 会话绑定的那本。它就在前台时什么都不用做；用户切走了，就把工具
        /// 重定向到它（ExcelTarget），而不是像以前那样写到前台那本或整体拒绝。
        /// </summary>
        private ToolResult ResolveTarget(string toolName, out IDisposable target)
        {
            target = null;
            var bound = BoundWorkbookKey?.Invoke();
            if (string.IsNullOrEmpty(bound) || WorkbookIndependentTools.Contains(toolName ?? "")) return null;
            var active = _excel.GetActiveWorkbookKey();
            if (WorkbookIdentity.SameKey(active, bound)) return null;

            if (ForegroundOnlyTools.Contains(toolName ?? ""))
            {
                return RefuseWrite(toolName,
                    $"这个操作只能在前台窗口进行，而当前活动窗口是另一个工作簿（本次对话属于「{DisplayName(bound)}」），本次未执行。",
                    $"请告诉用户切回「{DisplayName(bound)}」后再继续；读写单元格的其他工具不受影响，可以继续用。");
            }
            target = _excel.UseTargetWorkbook(bound);
            if (target == null)
            {
                return RefuseWrite(toolName,
                    $"本次对话所属的工作簿「{DisplayName(bound)}」已经关闭，本次未执行。",
                    "请告诉用户重新打开它后再继续，不要改用其他工作簿。");
            }
            return null;
        }

        /// <summary>模型读过 / 写过的区域与用户手动改过的区域（先读后写、读后被改检测）。</summary>
        public ReadLedger Ledger { get; } = new ReadLedger();

        /// <summary>
        /// 正在执行工具：这期间 Excel 的 SheetChange 是我们自己的写入引起的，不算用户改动。
        /// 只在 UI 线程上读写（工具和 SheetChange 都在 UI 线程）。
        /// </summary>
        public bool IsExecuting { get; private set; }

        private ToolResult ExecuteGuarded(string toolName, Dictionary<string, object> args)
        {
            CellRect target = default;
            var hasTarget = false;
            string backupId = null;
            string checkpointId = null;
            if (ToolMutationPolicy.IsMutating(toolName))
            {
                hasTarget = TryResolveWriteTarget(toolName, args, out target);
                if (hasTarget)
                {
                    var ledgerRefusal = CheckLedger(toolName, target);
                    if (ledgerRefusal != null) return ledgerRefusal;
                }

                ToolResult refusal;
                try
                {
                    refusal = EnsureBackupBeforeWrite(toolName, args, out backupId, out checkpointId);
                }
                catch (Exception ex)
                {
                    Logger.Instance.Error("ToolDispatcher", "Write guard failed: " + toolName, ex);
                    refusal = RefuseWrite(toolName, "修改前自动备份失败，为保护数据本次未执行（" + ex.Message + "）", null);
                }
                if (refusal != null) return refusal;
            }

            var healthBefore = WriteCheck.NeedsCheck(toolName) ? SafeCaptureHealth() : null;
            var diffBefore = hasTarget && WriteCheck.NeedsCheck(toolName) ? SafeReadFormulas(target) : null;
            var result = ExecuteCore(toolName, args);
            if (result != null && backupId != null) result.BackupSnapshotId = backupId;
            if (result != null && result.Success && checkpointId != null) result.CheckpointId = checkpointId;
            RecordLedger(toolName, result, hasTarget, target);
            if (healthBefore != null && result != null && result.Success)
            {
                result.Verification = RunWriteCheck(toolName, args, healthBefore, hasTarget, target);
            }
            if (diffBefore != null && result != null && result.Success)
            {
                result.Changes = CellDiff.Compare(target, diffBefore, SafeReadFormulas(target));
            }
            return result;
        }

        /// <summary>内联 diff 用：目标区域的公式 / 值（一次 COM 调用）；区域太大或读不了返回 null。</summary>
        private object[,] SafeReadFormulas(CellRect target)
        {
            try { return _excel.ReadFormulas(target.ToA1(), CellDiff.MaxCells); }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "diff read failed: " + ex.Message);
                return null;
            }
        }

        private const int MaxHealthErrorsCollected = 200;
        private const int MaxSamples = 5;

        private HealthSnapshot SafeCaptureHealth()
        {
            try { return _excel.CaptureHealth(MaxHealthErrorsCollected); }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "health capture failed: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 写后自动体检：和写入前比较公式错误、外部链接；写公式的工具再回读几个计算结果。
        /// 体检本身出错不影响写入结果（返回 null，模型只是少了这份验证）。
        /// </summary>
        private WriteVerification RunWriteCheck(string toolName, Dictionary<string, object> args,
            HealthSnapshot before, bool hasTarget, CellRect target)
        {
            try
            {
                var after = SafeCaptureHealth();
                List<CellSample> samples = null;
                if (hasTarget && WritesFormulas(toolName, args))
                {
                    samples = _excel.SampleCells(target.ToA1(), MaxSamples);
                }
                return WriteCheck.Compare(before, after, samples);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "write check failed: " + ex.Message);
                return null;
            }
        }

        private bool WritesFormulas(string toolName, Dictionary<string, object> args)
        {
            if (WriteCheck.FormulaTools.Contains(toolName)) return true;
            if (toolName != "write_range") return false;
            var values = Extract2DArray(args, "values");
            return values != null && values.Any(row => row != null && row.Any(v => v is string text && text.TrimStart().StartsWith("=")));
        }

        private ToolResult CheckLedger(string toolName, CellRect target)
        {
            var address = target.ToA1();
            var verdict = Ledger.Check(target, () => _excel.RangeHasContent(address), out var changed);
            switch (verdict)
            {
                case LedgerVerdict.StaleRead:
                    return RefuseWrite(toolName,
                        $"你上次读取之后，用户手动改动了 {string.Join("、", changed)}。你手里的是旧内容，本次未写入。",
                        $"先重新 read_range 读取 {address}，确认最新内容后再决定怎么写；不要把用户刚改的内容覆盖掉。");
                case LedgerVerdict.NotRead:
                    return RefuseWrite(toolName,
                        $"目标区域 {address} 已有内容，但你还没有读过它，本次未写入。",
                        $"先用 read_range 读取 {address}（或包含它的区域），确认可以覆盖后再写。");
                default:
                    return null;
            }
        }

        private void RecordLedger(string toolName, ToolResult result, bool hasTarget, CellRect target)
        {
            if (result == null || !result.Success) return;
            try
            {
                if (toolName == "read_range" || toolName == "read_selection")
                {
                    // RangeInfo 或截断后的同形对象：取 Excel 解析后的地址，比模型传的参数可靠
                    var data = result.Data;
                    var address = data?.GetType().GetProperty("Address")?.GetValue(data) as string;
                    var sheet = data?.GetType().GetProperty("WorksheetName")?.GetValue(data) as string;
                    if (CellRect.TryParse(address, sheet ?? _excel.GetActiveSheetName(), out var readRect))
                    {
                        Ledger.RecordRead(readRect);
                    }
                }
                else if (toolName == "rollback")
                {
                    // 整本内容换掉了：之前读到的都不再可信
                    Ledger.ForgetReads();
                }
                else if (hasTarget)
                {
                    Ledger.RecordOwnWrite(target);
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "ledger record failed: " + ex.Message);
            }
        }

        /// <summary>
        /// 用户在模型工作期间手动改了单元格：附在这次工具结果里告诉模型（每处只说一次），
        /// 就像 Claude Code 在文件被外部修改后提醒模型一样。
        /// </summary>
        private void AttachUserEditNotice(ToolResult result)
        {
            if (result == null) return;
            var parts = Ledger.TakeNotices();
            var edits = Ledger.TakeUnreportedUserEdits();
            if (edits.Count > 0)
            {
                parts.Add($"用户刚刚手动修改了 {string.Join("、", edits)}（不是你的操作）。和这些单元格相关的数据请重新读取，不要用之前读到的旧值覆盖。");
            }
            if (parts.Count == 0) return;
            var notice = string.Join(" ", parts);
            result.Warning = string.IsNullOrEmpty(result.Warning) ? notice : result.Warning + " " + notice;
        }

        /// <summary>
        /// 会改单元格内容、且目标区域明确的工具 → 目标区域。格式类（不丢数据）、结构类
        /// （插删行列由确认抽屉把关）、代码类（目标不可知）不检查。地址解析不了（命名区域等）
        /// 也不检查，交给工具本身。
        /// </summary>
        internal bool TryResolveWriteTarget(string toolName, Dictionary<string, object> args, out CellRect rect)
        {
            rect = default;
            var sheet = _excel.GetActiveSheetName();
            string address;
            switch (toolName)
            {
                case "write_value":
                case "write_formula":
                case "merge_cells":
                    address = GetArg<string>(args, "address");
                    return CellRect.TryParse(address, sheet, out rect);
                case "clear_range":
                    if (string.Equals(GetArg<string>(args, "clear_type"), "formats", StringComparison.OrdinalIgnoreCase))
                        return false;
                    return CellRect.TryParse(GetArg<string>(args, "address"), sheet, out rect);
                case "write_range":
                    if (!CellRect.TryParse(GetArg<string>(args, "address"), sheet, out rect)) return false;
                    var values = Extract2DArray(args, "values");
                    var rows = values?.Length ?? 1;
                    var cols = values == null || values.Length == 0 ? 1 : values.Max(r => r?.Length ?? 0);
                    rect = rect.Resize(rows, cols);
                    return true;
                case "fill_formula_down":
                    if (!CellRect.TryParse(GetArg<string>(args, "from_address"), sheet, out rect)) return false;
                    int count;
                    try { count = GetArg<int>(args, "row_count"); } catch { count = 1; }
                    rect = rect.Resize(Math.Max(1, count) + 1, rect.Col2 - rect.Col1 + 1);
                    return true;
                case "copy_range":
                    if (!CellRect.TryParse(GetArg<string>(args, "source_address"), sheet, out var source)) return false;
                    if (!CellRect.TryParse(GetArg<string>(args, "dest_address"), sheet, out rect)) return false;
                    rect = rect.Resize(source.Row2 - source.Row1 + 1, source.Col2 - source.Col1 + 1);
                    return true;
                case "replace_formula":
                case "clean_data":
                case "sort_data":
                case "delete_blank_rows":
                case "split_text_to_columns":
                case "fill_blank_cells":
                case "remove_special_chars":
                case "clean_amount":
                case "merge_columns":
                case "rename_columns":
                case "collapse_spaces":
                    return CellRect.TryParse(GetArg<string>(args, "range_address"), sheet, out rect);
                default:
                    return false;
            }
        }

        private ToolResult EnsureBackupBeforeWrite(string toolName, Dictionary<string, object> args, out string backupId, out string checkpointId)
        {
            backupId = null;
            checkpointId = null;
            // 备份的是工具要写的那本：会话绑定的工作簿（用户切到别的窗口也不变），没有绑定时才用前台那本
            var book = BoundWorkbookKey?.Invoke();
            if (string.IsNullOrEmpty(book)) book = _excel.GetActiveWorkbookKey();
            if (string.IsNullOrEmpty(book))
            {
                return RefuseWrite(toolName, "没有打开的工作簿，无法在修改前备份，本次未执行", null);
            }

            var scope = ToolMutationPolicy.ResolveScope(toolName,
                key => GetArg<string>(args, key),
                () => _excel.GetActiveSheetName());

            // 1. 更早的检查点（含本回合备份）都要覆盖这次写入涉及的表，回退时才不漏
            _turnBackups.TryGetValue(book, out var turnBackup);
            var turnBackupCovered = turnBackup != null;
            var list = CheckpointsOf(book);
            if (turnBackup != null && !list.Any(c => c.Id == turnBackup))
            {
                turnBackupCovered = _excel.ExtendSnapshotScope(turnBackup, scope);
            }
            foreach (var cp in list.ToList())
            {
                if (Covers(cp, scope)) continue;
                if (_excel.ExtendSnapshotScope(cp.Id, scope))
                {
                    Widen(cp, scope);
                    continue;
                }
                // 记不下新涉及的表：回到它会漏掉这张表，不再提供这个检查点
                Logger.Instance.Warning("ToolDispatcher", "ExtendSnapshotScope failed, dropping checkpoint: " + cp.Id);
                list.Remove(cp);
                if (cp.Id == turnBackup) turnBackupCovered = false;
            }

            // 2. 这一步自己的检查点（大工作簿或步骤太多时本回合只靠回合备份）
            if (turnBackupCovered && (_checkpointsSlowThisTurn || _checkpointsThisTurn >= MaxCheckpointsPerTurn))
            {
                backupId = turnBackup;
                return null;
            }

            SnapshotAttempt attempt;
            var watch = Stopwatch.StartNew();
            try
            {
                attempt = _excel.BackupWorkbook(book, $"AI 修改前自动备份（{toolName}）", scope);
            }
            catch (Exception ex) when (turnBackupCovered)
            {
                attempt = new SnapshotAttempt { Error = ex.Message };
            }
            watch.Stop();

            if (attempt == null || !attempt.Success)
            {
                if (turnBackupCovered)
                {
                    // 回合备份仍然保护着这次写入，只是这一步没有单独的检查点
                    Logger.Instance.Warning("ToolDispatcher", "Step checkpoint failed, turn backup still covers it: " + attempt?.Error);
                    backupId = turnBackup;
                    return null;
                }
                return RefuseWrite(toolName,
                    "修改前自动备份失败，为保护数据本次未执行：" + (attempt?.Error ?? "未知原因"),
                    "请告诉用户可能的原因（磁盘空间不足、单元格正处于编辑状态、工作簿受保护），处理后再试；不要换用其他工具绕过。");
            }

            checkpointId = attempt.SnapshotId;
            _checkpointsThisTurn++;
            list.Add(new Checkpoint
            {
                Id = checkpointId,
                Whole = scope.WholeWorkbook,
                Sheets = new HashSet<string>(scope.WholeWorkbook ? Enumerable.Empty<string>() : scope.Sheets, StringComparer.OrdinalIgnoreCase),
            });
            if (list.Count > MaxCheckpointsTracked) list.RemoveRange(0, list.Count - MaxCheckpointsTracked);

            if (watch.ElapsedMilliseconds > SlowCheckpointThresholdMs)
            {
                _checkpointsSlowThisTurn = true;
                Logger.Instance.Info("ToolDispatcher", $"Checkpoint took {watch.ElapsedMilliseconds} ms; later steps this turn share the turn backup");
            }
            if (!turnBackupCovered)
            {
                _turnBackups[book] = checkpointId;
                turnBackup = checkpointId;
                Logger.Instance.Info("ToolDispatcher", $"Turn backup created before {toolName}: {checkpointId}");
            }
            backupId = turnBackup;
            return null;
        }

        private List<Checkpoint> CheckpointsOf(string workbookKey)
        {
            if (!_checkpoints.TryGetValue(workbookKey, out var list))
            {
                list = new List<Checkpoint>();
                _checkpoints[workbookKey] = list;
            }
            return list;
        }

        private static bool Covers(Checkpoint cp, SnapshotScope scope)
        {
            if (cp.Whole) return true;
            if (scope == null) return true;
            if (scope.WholeWorkbook) return false;
            return scope.Sheets.All(s => cp.Sheets.Contains(s));
        }

        private static void Widen(Checkpoint cp, SnapshotScope scope)
        {
            if (scope.WholeWorkbook) { cp.Whole = true; return; }
            foreach (var sheet in scope.Sheets) cp.Sheets.Add(sheet);
        }

        private static ToolResult RefuseWrite(string toolName, string error, string suggestion)
        {
            Logger.Instance.Warning("ToolDispatcher", $"Write refused: tool={toolName}, reason={error}");
            return new ToolResult { Name = toolName, Success = false, Error = error, Suggestion = suggestion };
        }

        private static string DisplayName(string workbookKey)
        {
            try { return Path.GetFileName(workbookKey); }
            catch { return workbookKey; }
        }

        private ToolResult ExecuteCore(string toolName, Dictionary<string, object> args)
        {
            // ★ AI Native 改造后：权限确认由 PreToolUse hook 异步处理，UI 线程不再阻塞，
            // Excel 不会误触发 WorkbookBeforeClose，不再需要 ExecutionGuard 保护标志。
            try
            {
                switch (toolName)
                {
                    // 注：每个 case 的执行结果（含 success/error）在 switch 结束后统一记录
                    case "read_range":
                        var address = GetArg<string>(args, "address");
                        var rangeData = _excel.ReadRange(address);
                        // ★ Claude Agent SDK 内部消息缓冲区限制 1MB（1048576 bytes），
                        // read_range 返回的 Values/Formulas 二维数组序列化后可能超过此限制，
                        // 导致 SDK 抛 "JSON message exceeded maximum buffer size" 崩溃。
                        // 解决：对超过 MaxRows 的数据截断，并附加 truncated 标志告知模型。
                        var truncatedData = TruncateRangeData(rangeData, MaxReadRangeRows);
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = truncatedData,
                            Suggestion = GenerateRangeSuggestion(rangeData),
                            Context = BuildExcelSnapshot(),
                        };

                    case "write_formula":
                        var cellAddress = GetArg<string>(args, "address");
                        var formula = GetArg<string>(args, "formula");
                        Logger.Instance.Info("ToolDispatcher", "write_formula: address=[" + (cellAddress ?? "null") + "], formula=[" + (formula ?? "null") + "]");
                        var result = _excel.WriteFormula(cellAddress, formula);
                        if (!result.Success)
                        {
                            result.Suggestion = GenerateFormulaSuggestion(formula, result.Error);
                        }
                        return result;

                    case "write_value":
                        // ★ 写入纯文本/数字值（不解析为公式）：模型写"张三"应显示张三，而不是 ="张三"
                        var valAddress = GetArg<string>(args, "address");
                        object valRaw = null;
                        if (args.TryGetValue("value", out var valObj) && valObj != null)
                        {
                            // ★ JsonElement 需按 ValueKind 显式转换，否则 GetRawText() 会带引号
                            if (valObj is JsonElement valJe)
                            {
                                switch (valJe.ValueKind)
                                {
                                    case JsonValueKind.String:
                                        valRaw = valJe.GetString();
                                        break;
                                    case JsonValueKind.Number:
                                        if (valJe.TryGetInt32(out int iv)) valRaw = iv;
                                        else valRaw = valJe.GetDouble();
                                        break;
                                    case JsonValueKind.True:
                                        valRaw = true;
                                        break;
                                    case JsonValueKind.False:
                                        valRaw = false;
                                        break;
                                    default:
                                        valRaw = valJe.GetRawText();
                                        break;
                                }
                            }
                            else
                            {
                                valRaw = valObj;
                            }
                        }
                        Logger.Instance.Info("ToolDispatcher", "write_value: address=[" + (valAddress ?? "null") + "], value=[" + (valRaw ?? "null") + "], type=" + (valRaw?.GetType().Name ?? "null"));
                        return _excel.WriteValue(valAddress, valRaw);

                    case "write_range":
                        // ★ 批量写入二维数组：比逐个 write_value 快 100 倍
                        var wrAddress = GetArg<string>(args, "address");
                        object[][] wrValues = Extract2DArray(args, "values");
                        Logger.Instance.Info("ToolDispatcher", "write_range: address=[" + (wrAddress ?? "null") + "], rows=" + (wrValues?.Length ?? 0));
                        return _excel.WriteRange(wrAddress, wrValues);

                    case "read_workbook":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = _excel.ReadWorkbook(),
                            Context = BuildExcelSnapshot(),
                        };

                    case "read_selection":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = _excel.GetSelection(),
                            Context = BuildExcelSnapshot(),
                        };

                    case "fill_formula_down":
                        var fromAddr = GetArg<string>(args, "from_address");
                        var rowCount = GetArg<int>(args, "row_count");
                        return ExecuteFillDown(fromAddr, rowCount);

                    case "replace_formula":
                        var rangeAddr = GetArg<string>(args, "range_address");
                        var find = GetArg<string>(args, "find");
                        var replace = GetArg<string>(args, "replace");
                        return ExecuteReplaceFormula(rangeAddr, find, replace);

                    case "clean_data":
                        var cleanRange = GetArg<string>(args, "range_address");
                        var ops = GetArgArray(args, "operations");
                        return ExecuteCleanData(cleanRange, ops);

                    case "create_chart":
                        var dataRange = GetArg<string>(args, "data_range");
                        var chartType = GetArg<string>(args, "chart_type") ?? "column";
                        var chartTitle = GetArg<string>(args, "title") ?? "";
                        var xLabel = GetArg<string>(args, "x_label") ?? "";
                        var yLabel = GetArg<string>(args, "y_label") ?? "";
                        return ExecuteCreateChart(dataRange, chartType, chartTitle, xLabel, yLabel);

                    case "create_pivot_table":
                        var sourceRange = GetArg<string>(args, "source_range");
                        var destSheet = GetArg<string>(args, "destination_sheet");
                        var pivotName = GetArg<string>(args, "pivot_table_name") ?? "PivotTable1";
                        var rowFields = GetArgArray(args, "row_fields");
                        var colFields = GetArgArray(args, "column_fields");
                        var valFields = GetArgArray(args, "value_fields");
                        var valFunc = GetArg<string>(args, "value_function") ?? "Sum";
                        return ExecuteCreatePivot(sourceRange, destSheet, pivotName, rowFields, colFields, valFields, valFunc);

                    case "delete_blank_rows":
                        return ExecuteDataCleanOperation("delete_blank_rows", GetArg<string>(args, "range_address"));
                    case "split_text_to_columns":
                        return ExecuteSplitText(
                            GetArg<string>(args, "range_address"),
                            GetArg<string>(args, "delimiter") ?? ",");
                    case "fill_blank_cells":
                        return ExecuteDataCleanOperation("fill_blank_cells", GetArg<string>(args, "range_address"));
                    case "highlight_duplicates":
                        return ExecuteDataCleanOperation("highlight_duplicates", GetArg<string>(args, "range_address"));
                    case "remove_special_chars":
                        return ExecuteDataCleanOperation("remove_special_chars", GetArg<string>(args, "range_address"));
                    case "clean_amount":
                        return ExecuteDataCleanOperation("clean_amount", GetArg<string>(args, "range_address"));
                    case "merge_columns":
                        return ExecuteMergeColumns(
                            GetArg<string>(args, "range_address"),
                            GetArg<string>(args, "delimiter") ?? " ",
                            GetArg<string>(args, "target_column") ?? "");
                    case "rename_columns":
                        return ExecuteRenameColumns(
                            GetArg<string>(args, "range_address"),
                            GetArgArray(args, "new_names"));
                    case "collapse_spaces":
                        return ExecuteDataCleanOperation("collapse_spaces", GetArg<string>(args, "range_address"));

                    case "add_data_labels":
                        return ExecuteAddDataLabels(
                            GetArg<string>(args, "chart_name") ?? "",
                            GetArg<string>(args, "position") ?? "outside_end",
                            GetArg<bool?>(args, "show_value") ?? true,
                            GetArg<bool?>(args, "show_category_name") ?? false,
                            GetArg<bool?>(args, "show_percentage") ?? false);
                    case "set_chart_title":
                        return ExecuteSetChartTitle(
                            GetArg<string>(args, "chart_name") ?? "",
                            GetArg<string>(args, "title") ?? "");
                    case "set_chart_colors":
                        return ExecuteSetChartColors(
                            GetArg<string>(args, "chart_name") ?? "",
                            GetArgArray(args, "colors"));
                    case "create_combo_chart":
                        return ExecuteCreateComboChart(
                            GetArg<string>(args, "data_range"),
                            GetArg<string>(args, "title") ?? "",
                            GetArg<string>(args, "x_label") ?? "",
                            GetArg<string>(args, "y_label") ?? "",
                            GetArg<string>(args, "secondary_y_label") ?? "",
                            GetArg<int?>(args, "line_series_index") ?? 2);
                    case "export_chart":
                        return ExecuteExportChart(
                            GetArg<string>(args, "chart_name") ?? "",
                            GetArg<string>(args, "output_path") ?? "",
                            GetArg<string>(args, "format") ?? "png");

                    case "refresh_pivot":
                        return ExecuteRefreshPivot(
                            GetArg<string>(args, "pivot_table_name") ?? "",
                            GetArg<string>(args, "sheet_name") ?? "");
                    case "group_pivot_date":
                        return ExecuteGroupPivotDate(
                            GetArg<string>(args, "pivot_table_name") ?? "",
                            GetArg<string>(args, "field_name"),
                            GetArg<string>(args, "group_by") ?? "month",
                            GetArg<string>(args, "sheet_name") ?? "");
                    case "set_pivot_value_display":
                        return ExecuteSetPivotValueDisplay(
                            GetArg<string>(args, "pivot_table_name") ?? "",
                            GetArg<string>(args, "value_field"),
                            GetArg<string>(args, "display_type") ?? "normal",
                            GetArg<string>(args, "base_field") ?? "",
                            GetArg<string>(args, "sheet_name") ?? "");
                    case "set_pivot_totals":
                        return ExecuteSetPivotTotals(
                            GetArg<string>(args, "pivot_table_name") ?? "",
                            GetArg<bool?>(args, "show_row_totals") ?? true,
                            GetArg<bool?>(args, "show_column_totals") ?? true,
                            GetArg<string>(args, "sheet_name") ?? "");
                    case "add_pivot_slicer":
                        return ExecuteAddPivotSlicer(
                            GetArg<string>(args, "pivot_table_name") ?? "",
                            GetArg<string>(args, "field_name"),
                            GetArg<string>(args, "sheet_name") ?? "");

                    case "execute_vba":
                        var code = GetArg<string>(args, "code");
                        return _excel.ExecuteVBA(code);

                    case "execute_python":
                        var pyCode = GetArg<string>(args, "code");
                        return _excel.ExecutePython(pyCode);

                    case "add_sheet":
                        return _excel.AddSheet(GetArg<string>(args, "name"));

                    case "delete_sheet":
                        return _excel.DeleteSheet(GetArg<string>(args, "name"));

                    case "rename_sheet":
                        return _excel.RenameSheet(
                            GetArg<string>(args, "old_name"),
                            GetArg<string>(args, "new_name"));

                    case "set_number_format":
                        return _excel.SetNumberFormat(
                            GetArg<string>(args, "address"),
                            GetArg<string>(args, "format"));

                    case "set_column_width":
                        return _excel.SetColumnWidth(
                            GetArg<string>(args, "address"),
                            GetArg<double>(args, "width"),
                            GetArg<bool>(args, "auto_fit"));

                    case "sort_data":
                        // ★★★ has_header 是必传参数！AI 漏传会导致表头被当作数据排序。
                        // 策略：has_header 缺失时不拒绝执行，而是返回错误要求 AI 重新调用。
                        // 同时 ExcelActions.SortData 内部有智能检测兜底（has_header=false 时
                        // 检测疑似表头），双重保险。
                        if (!args.ContainsKey("has_header") || args["has_header"] == null)
                        {
                            Logger.Instance.Warning("ToolDispatcher",
                                "sort_data: has_header is missing, rejecting to force AI retry");
                            return new ToolResult
                            {
                                Name = "sort_data",
                                Success = false,
                                Error = "缺少必传参数 has_header。必须根据 read_range 的结果判断第一行是否为列标题后传入。",
                                Suggestion = "请先调用 read_range 查看数据结构：第一行是列标题（如'月份'/'金额'）→ has_header=true；第一行就是数据 → has_header=false。"
                            };
                        }
                        return _excel.SortData(
                            GetArg<string>(args, "range_address"),
                            GetArg<string>(args, "sort_column"),
                            GetArg<bool>(args, "descending"),
                            GetArg<bool>(args, "has_header"));

                    case "filter_data":
                        return _excel.FilterData(
                            GetArg<string>(args, "range_address"),
                            GetArg<int>(args, "column_index"),
                            GetArg<string>(args, "criteria"));

                    case "merge_cells":
                        return _excel.MergeCells(GetArg<string>(args, "address"));

                    case "unmerge_cells":
                        return _excel.UnmergeCells(GetArg<string>(args, "address"));

                    case "set_cell_style":
                        {
                            var scsAddress = GetArg<string>(args, "address");
                            var scsFontName = GetArg<string>(args, "font_name");
                            var scsFontSize = GetArg<double?>(args, "font_size");
                            var scsBold = GetArg<bool?>(args, "bold");
                            var scsItalic = GetArg<bool?>(args, "italic");
                            var scsFontColor = GetArg<string>(args, "font_color");
                            var scsBgColor = GetArg<string>(args, "bg_color");
                            var scsHAlign = GetArg<string>(args, "h_align");
                            var scsVAlign = GetArg<string>(args, "v_align");
                            var scsWrapText = GetArg<bool?>(args, "wrap_text");
                            Logger.Instance.Info("ToolDispatcher",
                                $"set_cell_style params: address={scsAddress}, fontName={scsFontName}, " +
                                $"fontSize={scsFontSize}, bold={scsBold}, italic={scsItalic}, " +
                                $"fontColor={scsFontColor}, bgColor={scsBgColor}, hAlign={scsHAlign}, vAlign={scsVAlign}, wrapText={scsWrapText}");
                            return _excel.SetCellStyle(scsAddress, scsFontName, scsFontSize, scsBold, scsItalic,
                                scsFontColor, scsBgColor, scsHAlign, scsVAlign, scsWrapText);
                        }

                    case "copy_range":
                        return _excel.CopyRange(
                            GetArg<string>(args, "source_address"),
                            GetArg<string>(args, "dest_address"));

                    case "clear_range":
                        return _excel.ClearRange(
                            GetArg<string>(args, "address"),
                            GetArg<string>(args, "clear_type"));

                    case "insert_rows":
                        return _excel.InsertRows(
                            GetArg<int>(args, "row"),
                            GetArg<int>(args, "count"));

                    case "delete_rows":
                        return _excel.DeleteRows(
                            GetArg<int>(args, "row"),
                            GetArg<int>(args, "count"));

                    case "insert_columns":
                        return _excel.InsertColumns(
                            GetArg<int>(args, "column"),
                            GetArg<int>(args, "count"));

                    case "delete_columns":
                        return _excel.DeleteColumns(
                            GetArg<int>(args, "column"),
                            GetArg<int>(args, "count"));

                    case "freeze_panes":
                        return _excel.FreezePanes(GetArg<string>(args, "address"));

                    case "apply_conditional_format":
                        // ★ P-1 修复：rule_args 是 JsonElement，必须显式转 Dictionary<string,object>，
                        // 否则 ApplyConditionalFormat 中的 `is Dictionary<string,object>` 失败 → 静默丢弃 operator/value/n
                        object ruleArgsObj = null;
                        if (args.ContainsKey("rule_args") && args["rule_args"] != null)
                        {
                            var rawArg = args["rule_args"];
                            if (rawArg is JsonElement je && je.ValueKind == JsonValueKind.Object)
                            {
                                var dict = new Dictionary<string, object>();
                                foreach (var prop in je.EnumerateObject())
                                {
                                    dict[prop.Name] = prop.Value.Clone();
                                }
                                ruleArgsObj = dict;
                            }
                            else if (rawArg is Dictionary<string, object> existingDict)
                            {
                                ruleArgsObj = existingDict;
                            }
                            else
                            {
                                ruleArgsObj = rawArg;
                            }
                        }
                        return _excel.ApplyConditionalFormat(
                            GetArg<string>(args, "address"),
                            GetArg<string>(args, "rule_type"),
                            ruleArgsObj);

                    case "write_table":
                        return _excel.WriteTable(
                            GetArg<string>(args, "address"),
                            GetArg<string>(args, "table_name"));

                    case "create_snapshot":
                        return ExecuteCreateSnapshot();

                    case "rollback":
                        return ExecuteRollback(GetArg<string>(args, "snapshot_id"));

                    case "read_attachment":
                        return ExecuteReadAttachment(args);

                    // ★ Computer Use 工具：截图 + 模拟键盘
                    case "screenshot_excel":
                        return ExecuteScreenshotExcel();

                    case "send_keys":
                        return ExecuteSendKeys(GetArg<string>(args, "keys"));

                    default:
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = false,
                            Error = $"未知工具: {toolName}",
                        };
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "Execute failed: " + toolName, ex);
                return new ToolResult
                {
                    Name = toolName,
                    Success = false,
                    Error = ex.Message,
                };
            }
        }

        private ToolResult ExecuteCreateSnapshot()
        {
            var bound = BoundWorkbookKey?.Invoke();
            string snapshotId;
            string error = null;
            if (string.IsNullOrEmpty(bound))
            {
                snapshotId = _excel.CreateSnapshot();
            }
            else
            {
                // 备份会话绑定的那本，而不是恰好处于活动状态的那本
                var attempt = _excel.BackupWorkbook(bound, "手动快照", SnapshotScope.Whole());
                snapshotId = attempt?.SnapshotId;
                error = attempt?.Error;
            }
            return new ToolResult
            {
                Name = "create_snapshot",
                Success = !string.IsNullOrEmpty(snapshotId),
                Data = new { snapshot_id = snapshotId },
                Error = string.IsNullOrEmpty(snapshotId) ? (error ?? "创建快照失败") : null,
            };
        }

        private ToolResult ExecuteRollback(string snapshotId)
        {
            // 模型只能恢复本会话工作簿的快照：快照 ID 来自上下文，传错了不能去改另一本
            var bound = BoundWorkbookKey?.Invoke();
            var meta = _excel.GetSnapshotMeta(snapshotId);
            if (!string.IsNullOrEmpty(bound) && meta != null && !string.IsNullOrEmpty(meta.WorkbookKey)
                && !WorkbookIdentity.SameKey(meta.WorkbookKey, bound))
            {
                return new ToolResult
                {
                    Name = "rollback",
                    Success = false,
                    Error = $"快照 {snapshotId} 属于另一个工作簿「{meta.WorkbookName}」，不能在本次对话中恢复。",
                };
            }

            var r = _excel.Rollback(snapshotId);
            return new ToolResult
            {
                Name = "rollback",
                Success = r != null && r.Success,
                Error = r == null ? "恢复失败" : r.Error,
                Data = r == null ? null : new
                {
                    pre_restore_snapshot_id = r.PreRestoreSnapshotId,
                    restored_sheets = r.RestoredSheets,
                    removed_sheets = r.RemovedSheets,
                    warnings = r.Warnings,
                },
            };
        }

        /// <summary>
        /// 执行后日志记录（在 PythonSidecar.HandleToolCall 调用 Execute 后调用）
        /// </summary>
        public void LogResult(string toolName, ToolResult result)
        {
            if (result == null) return;
            if (result.Success)
            {
                Logger.Instance.Info("ToolDispatcher", $"Execute result: tool={toolName}, success=True");
            }
            else
            {
                Logger.Instance.Warning("ToolDispatcher", $"Execute result: tool={toolName}, success=False, error={result.Error}");
            }
        }

        /// <summary>
        /// ★ 快照缓存：避免每次工具调用都重读工作簿结构。
        /// _snapshotDirty=true 表示工作簿可能已改变，需要重读；
        /// _snapshotDirty=false 表示自上次读取后没有写操作，可复用缓存。
        /// 用户在 Excel 中手动改数据会导致缓存 stale，但 tool_result.context 只是给 AI 的辅助信息，
        /// 不是关键数据（AI 需要最新数据时会主动调 read_range），所以可接受短暂 stale。
        /// </summary>
        private object _snapshotCache;
        private string _snapshotFingerprint;
        private DateTime _snapshotCacheTime;

        /// <summary>
        /// 构建当前 Excel 上下文快照（附在 tool_result.context）
        /// 注意：只放元信息，避免二维数组（Object[,]）导致 System.Text.Json 序列化失败崩溃。
        /// 实际数据由 sidecar 通过 read_selection / read_range 工具按需读取。
        /// ★ 性能优化：基于工作簿指纹（sheet 数量+每个 sheet 的 UsedRange 地址）判断是否变化，
        /// 无变化则复用缓存（只更新 timestamp），避免重复的 COM 调用。
        /// </summary>
        public object BuildExcelSnapshot()
        {
            try
            {
                // 用户切到了别的工作簿：前台那本的上下文不属于这次对话，宁可不给
                var bound = BoundWorkbookKey?.Invoke();
                if (!string.IsNullOrEmpty(bound) && !WorkbookIdentity.SameKey(_excel.GetActiveWorkbookKey(), bound)) return null;

                // 计算工作簿指纹：sheet 数量 + 每个 sheet 的 UsedRange 地址
                string fingerprint = ComputeWorkbookFingerprint();

                // 指纹未变 + 缓存存在：复用缓存（只更新 timestamp）
                if (_snapshotCache != null && fingerprint != null
                    && string.Equals(fingerprint, _snapshotFingerprint, StringComparison.Ordinal))
                {
                    // 复用缓存的 workbook 信息，但重新读取 selection（用户可能切换了选中区域）
                    object freshSelection = ReadSelectionInfo();
                    return new
                    {
                        workbook = ExtractWorkbookFromCache(_snapshotCache),
                        selection = freshSelection ?? new { address = "" },
                        timestamp = DateTime.Now.ToString("o"),
                        cached = true,
                    };
                }

                // 指纹变化或无缓存：重新读取完整快照
                object selectionInfo = ReadSelectionInfo();

                var snapshot = new
                {
                    workbook = _excel.ReadWorkbook(),
                    selection = selectionInfo ?? new { address = "" },
                    timestamp = DateTime.Now.ToString("o"),
                };

                _snapshotCache = snapshot;
                _snapshotFingerprint = fingerprint;
                _snapshotCacheTime = DateTime.Now;

                return snapshot;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "BuildExcelSnapshot failed", ex);
                return new { error = ex.Message };
            }
        }

        /// <summary>计算工作簿指纹：sheet 数量 + 每个 sheet 的 UsedRange 地址 + 活动 sheet 名</summary>
        private string ComputeWorkbookFingerprint()
        {
            try
            {
                if (_excelApp == null || _excelApp.ActiveWorkbook == null) return null;
                var wb = _excelApp.ActiveWorkbook;
                var sb = new System.Text.StringBuilder();
                sb.Append("sheets=").Append(wb.Worksheets.Count).Append("|");
                try { sb.Append("active=").Append((wb.ActiveSheet as Worksheet)?.Name ?? "").Append("|"); }
                catch { sb.Append("active=err|"); }
                foreach (Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        var used = ws.UsedRange;
                        sb.Append(ws.Name).Append(":")
                          .Append(used.Address).Append(",")
                          .Append(used.Rows.Count).Append("x")
                          .Append(used.Columns.Count).Append(";");
                    }
                    catch
                    {
                        sb.Append(ws.Name).Append(":err;");
                    }
                }
                return sb.ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>读取当前选中区域信息（无选中返回 null）</summary>
        private object ReadSelectionInfo()
        {
            try
            {
                if (_excelApp != null)
                {
                    var sel = _excelApp.Selection;
                    if (sel is Range range)
                    {
                        return new
                        {
                            address = range.Address,
                            worksheet = SafeGetWorksheetName(range),
                            rowCount = range.Rows.Count,
                            columnCount = range.Columns.Count,
                        };
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>从缓存对象中提取 workbook 字段（反射）</summary>
        private static object ExtractWorkbookFromCache(object cache)
        {
            if (cache == null) return null;
            try
            {
                var prop = cache.GetType().GetProperty("workbook");
                return prop?.GetValue(cache);
            }
            catch { return null; }
        }

        private static string SafeGetWorksheetName(Range range)
        {
            try { return range.Worksheet.Name; }
            catch { return ""; }
        }

        /// <summary>
        /// 截断 RangeInfo 的二维数组到 MaxRows 行，避免 SDK 1MB 消息缓冲区限制。
        /// 返回一个动态对象（包含截断标志和原始行数），供序列化发送给 sidecar。
        /// </summary>
        private object TruncateRangeData(object rangeData, int maxRows)
        {
            if (rangeData == null) return null;

            try
            {
                var type = rangeData.GetType();

                // 通过反射读取字段（避免直接依赖 RangeInfo 类型）
                var addressProp = type.GetProperty("Address");
                var wsProp = type.GetProperty("WorksheetName");
                var rowProp = type.GetProperty("RowCount");
                var colProp = type.GetProperty("ColumnCount");
                var valuesProp = type.GetProperty("Values");
                var formulasProp = type.GetProperty("Formulas");
                var numFmtProp = type.GetProperty("NumberFormats");

                int originalRows = (int)(rowProp?.GetValue(rangeData) ?? 0);
                int cols = (int)(colProp?.GetValue(rangeData) ?? 0);

                // 不需要截断
                if (originalRows <= maxRows)
                {
                    return rangeData;
                }

                Logger.Instance.Info("ToolDispatcher",
                    $"TruncateRangeData: originalRows={originalRows}, maxRows={maxRows}, truncating");

                // 截断二维数组
                object[,] origValues = valuesProp?.GetValue(rangeData) as object[,];
                string[,] origFormulas = formulasProp?.GetValue(rangeData) as string[,];
                string[,] origNumFmt = numFmtProp?.GetValue(rangeData) as string[,];

                int rowStart = origValues?.GetLowerBound(0) ?? 0;
                int colStart = origValues?.GetLowerBound(1) ?? 0;

                var newValues = new object[maxRows, cols];
                var newFormulas = origFormulas != null ? new string[maxRows, cols] : null;
                var newNumFmt = origNumFmt != null ? new string[maxRows, cols] : null;

                for (int i = 0; i < maxRows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        if (origValues != null)
                            newValues[i, j] = origValues[rowStart + i, colStart + j];
                        if (newFormulas != null)
                            newFormulas[i, j] = origFormulas[rowStart + i, colStart + j];
                        if (newNumFmt != null)
                            newNumFmt[i, j] = origNumFmt[rowStart + i, colStart + j];
                    }
                }

                // 返回截断后的匿名对象（保持与 RangeInfo 相同的字段名）
                return new
                {
                    Address = addressProp?.GetValue(rangeData),
                    WorksheetName = wsProp?.GetValue(rangeData),
                    RowCount = maxRows,
                    ColumnCount = cols,
                    Values = newValues,
                    Formulas = newFormulas,
                    NumberFormats = newNumFmt,
                    // 截断标志：告知模型数据被截断，原始行数
                    truncated = true,
                    original_row_count = originalRows,
                    truncation_hint = $"数据超过 {maxRows} 行被截断。如需查看后续数据，请用更精确的地址（如 A201:B400）再次调用 read_range。",
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "TruncateRangeData failed, returning original", ex);
                return rangeData;
            }
        }

        /// <summary>
        /// read_range 返回数据的类型检测：检测到 text/mixed 时返回澄清提示。
        /// ★ 直接从 RangeInfo 对象读取，不序列化整个对象（避免 2D 数组序列化崩溃）。
        /// </summary>
        private string GenerateRangeSuggestion(object rangeData)
        {
            try
            {
                // 尝试读取 RangeInfo.Values 进行类型检测
                if (rangeData == null) return null;

                // 通过反射读取 Values 属性（避免直接依赖 RangeInfo 类型）
                var valuesProp = rangeData.GetType().GetProperty("Values");
                if (valuesProp == null) return null;
                var values = valuesProp.GetValue(rangeData) as object[,];
                if (values == null) return null;

                // 1-based 数组安全访问
                int rowStart = values.GetLowerBound(0);
                int rowEnd = values.GetUpperBound(0);
                int colStart = values.GetLowerBound(1);
                int colEnd = values.GetUpperBound(1);

                bool hasNum = false, hasText = false, hasDate = false;
                for (int i = rowStart; i <= rowEnd; i++)
                {
                    for (int j = colStart; j <= colEnd; j++)
                    {
                        var item = values[i, j];
                        if (item == null) continue;
                        if (item is double || item is int || item is long || item is decimal || item is float)
                            hasNum = true;
                        else if (item is DateTime)
                            hasDate = true;
                        else if (item is string)
                            hasText = true;
                        else if (item is bool)
                            continue; // bool 不计入 text/num
                        else
                            hasText = true; // 其他类型按文本处理
                    }
                }

                if (hasDate && !hasNum && !hasText)
                    return "该列是日期，求和无意义。建议用 COUNTA 计数，或按日期分组。";
                if (hasNum && hasText)
                    return "该列同时包含数字和文本，无法直接求和。建议：求和(忽略文本)、计数(全部 COUNTA)、计数(仅数字 COUNT)";
                if (hasText && !hasNum)
                    return "该列是文本，无法直接求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "GenerateRangeSuggestion failed", ex);
            }
            return null;
        }

        /// <summary>
        /// write_formula 失败时根据公式和错误生成澄清提示
        /// </summary>
        private string GenerateFormulaSuggestion(string formula, string error)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(error)) return null;

            if (formula.Contains("SUM") && (error.Contains("类型") || error.Contains("type") || error.Contains("mismatch")))
                return "目标区域包含文本，无法求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
            if (formula.Contains("AVERAGE") && (error.Contains("除数为零") || error.Contains("zero") || error.Contains("divisor")))
                return "目标区域没有数字，无法计算平均值。建议检查数据源。";
            if (error.Contains("范围") || error.Contains("range"))
                return "目标范围无效。请检查地址格式，例如 A1、B1:B10、A:A。";
            return null;
        }

        // ============= 从 Orchestrator 移植的 5 个私有方法 =============

        private ToolResult ExecuteFillDown(string fromAddress, int rowCount)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteFillDown: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "fill_formula_down", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var app = _excelApp;
                var fromRange = ExcelTarget.Range(app, fromAddress);
                var toRange = fromRange.Resize[rowCount, fromRange.Columns.Count];
                fromRange.AutoFill(toRange, XlAutoFillType.xlFillDefault);
                return new ToolResult { Name = "fill_formula_down", Success = true };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteFillDown failed", ex);
                return new ToolResult { Name = "fill_formula_down", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteReplaceFormula(string rangeAddress, string find, string replace)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteReplaceFormula: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "replace_formula", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var app = _excelApp;
                var range = ExcelTarget.Range(app, rangeAddress);
                int count = 0;

                foreach (Range cell in range.Cells)
                {
                    if ((bool)cell.HasFormula)
                    {
                        var formula = cell.Formula.ToString();
                        if (formula.Contains(find))
                        {
                            cell.Formula = formula.Replace(find, replace);
                            count++;
                        }
                    }
                }

                return new ToolResult { Name = "replace_formula", Success = true, Data = new { replacedCount = count } };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteReplaceFormula failed", ex);
                return new ToolResult { Name = "replace_formula", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCleanData(string rangeAddress, string[] operations)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteCleanData: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "clean_data", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                // 修改前的备份由 Execute 的写入守卫统一完成（备份失败不会走到这里）

                var cleaner = new DataCleaner(_excelApp);
                var results = new List<string>();

                foreach (var op in operations)
                {
                    switch (op.ToLower())
                    {
                        case "unify_date":
                            var r1 = cleaner.UnifyDateFormats(rangeAddress);
                            results.Add($"unify_date: {(r1.Success ? (r1.Data as dynamic)?.convertedCount + "个" : "失败-" + r1.Error)}");
                            break;
                        case "remove_duplicates":
                            var r2 = cleaner.RemoveDuplicates(rangeAddress);
                            results.Add($"remove_duplicates: {(r2.Success ? (r2.Data as dynamic)?.removedCount + "行" : "失败-" + r2.Error)}");
                            break;
                        case "highlight_missing":
                            var r3 = cleaner.HighlightMissingValues(rangeAddress);
                            results.Add($"highlight_missing: {(r3.Success ? (r3.Data as dynamic)?.highlightedCount + "个" : "失败-" + r3.Error)}");
                            break;
                        case "trim_spaces":
                            var r4 = cleaner.TrimSpaces(rangeAddress);
                            results.Add($"trim_spaces: {(r4.Success ? (r4.Data as dynamic)?.trimmedCount + "个" : "失败-" + r4.Error)}");
                            break;
                        case "text_to_number":
                            var r5 = cleaner.TextToNumber(rangeAddress);
                            results.Add($"text_to_number: {(r5.Success ? (r5.Data as dynamic)?.convertedCount + "个" : "失败-" + r5.Error)}");
                            break;
                        case "delete_blank_rows":
                            var r6 = cleaner.DeleteBlankRows(rangeAddress);
                            results.Add($"delete_blank_rows: {(r6.Success ? (r6.Data as dynamic)?.deletedCount + "行" : "失败-" + r6.Error)}");
                            break;
                        case "fill_blank_cells":
                            var r7 = cleaner.FillBlankCells(rangeAddress);
                            results.Add($"fill_blank_cells: {(r7.Success ? (r7.Data as dynamic)?.filledCount + "个" : "失败-" + r7.Error)}");
                            break;
                        case "highlight_duplicates":
                            var r8 = cleaner.HighlightDuplicates(rangeAddress);
                            results.Add($"highlight_duplicates: {(r8.Success ? (r8.Data as dynamic)?.highlightedCount + "个" : "失败-" + r8.Error)}");
                            break;
                        case "remove_special_chars":
                            var r9 = cleaner.RemoveSpecialChars(rangeAddress);
                            results.Add($"remove_special_chars: {(r9.Success ? (r9.Data as dynamic)?.cleanedCount + "个" : "失败-" + r9.Error)}");
                            break;
                        case "clean_amount":
                            var r10 = cleaner.CleanAmount(rangeAddress);
                            results.Add($"clean_amount: {(r10.Success ? (r10.Data as dynamic)?.cleanedCount + "个" : "失败-" + r10.Error)}");
                            break;
                        case "collapse_spaces":
                            var r11 = cleaner.CollapseSpaces(rangeAddress);
                            results.Add($"collapse_spaces: {(r11.Success ? (r11.Data as dynamic)?.cleanedCount + "个" : "失败-" + r11.Error)}");
                            break;
                    }
                }

                return new ToolResult
                {
                    Name = "clean_data",
                    Success = true,
                    Data = new { operations = results }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCleanData failed", ex);
                return new ToolResult { Name = "clean_data", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteDataCleanOperation(string operation, string rangeAddress)
        {
            if (_excelApp == null)
                return new ToolResult { Name = operation, Success = false, Error = "excelApp not available" };
            try
            {
                var cleaner = new DataCleaner(_excelApp);
                return operation switch
                {
                    "delete_blank_rows" => cleaner.DeleteBlankRows(rangeAddress),
                    "fill_blank_cells" => cleaner.FillBlankCells(rangeAddress),
                    "highlight_duplicates" => cleaner.HighlightDuplicates(rangeAddress),
                    "remove_special_chars" => cleaner.RemoveSpecialChars(rangeAddress),
                    "clean_amount" => cleaner.CleanAmount(rangeAddress),
                    "collapse_spaces" => cleaner.CollapseSpaces(rangeAddress),
                    _ => new ToolResult { Name = operation, Success = false, Error = "Unknown operation: " + operation }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", $"ExecuteDataCleanOperation({operation}) failed", ex);
                return new ToolResult { Name = operation, Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteSplitText(string rangeAddress, string delimiter)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "split_text_to_columns", Success = false, Error = "excelApp not available" };
            try
            {
                var cleaner = new DataCleaner(_excelApp);
                return cleaner.SplitTextToColumns(rangeAddress, delimiter);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteSplitText failed", ex);
                return new ToolResult { Name = "split_text_to_columns", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteMergeColumns(string rangeAddress, string delimiter, string targetColumn)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "merge_columns", Success = false, Error = "excelApp not available" };
            try
            {
                var cleaner = new DataCleaner(_excelApp);
                return cleaner.MergeColumns(rangeAddress, delimiter, targetColumn);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteMergeColumns failed", ex);
                return new ToolResult { Name = "merge_columns", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteRenameColumns(string rangeAddress, string[] newNames)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "rename_columns", Success = false, Error = "excelApp not available" };
            try
            {
                var cleaner = new DataCleaner(_excelApp);
                return cleaner.RenameColumns(rangeAddress, newNames);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteRenameColumns failed", ex);
                return new ToolResult { Name = "rename_columns", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteAddDataLabels(string chartName, string position, bool showValue, bool showCategoryName, bool showPercentage)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "add_data_labels", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.AddDataLabels(chartName, position, showValue, showCategoryName, showPercentage);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteAddDataLabels failed", ex);
                return new ToolResult { Name = "add_data_labels", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteSetChartTitle(string chartName, string title)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "set_chart_title", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.SetChartTitle(chartName, title);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteSetChartTitle failed", ex);
                return new ToolResult { Name = "set_chart_title", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteSetChartColors(string chartName, string[] colors)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "set_chart_colors", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.SetChartColors(chartName, colors);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteSetChartColors failed", ex);
                return new ToolResult { Name = "set_chart_colors", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCreateComboChart(string dataRange, string title, string xLabel, string yLabel, string secondaryYLabel, int lineSeriesIndex)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "create_combo_chart", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.CreateComboChart(dataRange, title, xLabel, yLabel, secondaryYLabel, lineSeriesIndex);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCreateComboChart failed", ex);
                return new ToolResult { Name = "create_combo_chart", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteExportChart(string chartName, string outputPath, string format)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "export_chart", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.ExportChartAsImage(chartName, outputPath, format);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteExportChart failed", ex);
                return new ToolResult { Name = "export_chart", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteRefreshPivot(string pivotTableName, string sheetName)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "refresh_pivot", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.RefreshPivotTable(pivotTableName, sheetName);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteRefreshPivot failed", ex);
                return new ToolResult { Name = "refresh_pivot", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteGroupPivotDate(string pivotTableName, string fieldName, string groupBy, string sheetName)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "group_pivot_date", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.GroupPivotDateField(pivotTableName, fieldName, groupBy, sheetName);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteGroupPivotDate failed", ex);
                return new ToolResult { Name = "group_pivot_date", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteSetPivotValueDisplay(string pivotTableName, string valueField, string displayType, string baseField, string sheetName)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "set_pivot_value_display", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.SetPivotValueDisplay(pivotTableName, valueField, displayType, baseField, sheetName);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteSetPivotValueDisplay failed", ex);
                return new ToolResult { Name = "set_pivot_value_display", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteSetPivotTotals(string pivotTableName, bool showRowTotals, bool showColumnTotals, string sheetName)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "set_pivot_totals", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.SetPivotTotals(pivotTableName, showRowTotals, showColumnTotals, sheetName);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteSetPivotTotals failed", ex);
                return new ToolResult { Name = "set_pivot_totals", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteAddPivotSlicer(string pivotTableName, string fieldName, string sheetName)
        {
            if (_excelApp == null)
                return new ToolResult { Name = "add_pivot_slicer", Success = false, Error = "excelApp not available" };
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.AddPivotSlicer(pivotTableName, fieldName, sheetName);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteAddPivotSlicer failed", ex);
                return new ToolResult { Name = "add_pivot_slicer", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCreateChart(string dataRange, string chartType, string title, string xLabel, string yLabel)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteCreateChart: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "create_chart", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.CreateChart(dataRange, chartType, title, xLabel, yLabel);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCreateChart failed", ex);
                return new ToolResult { Name = "create_chart", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// ★ AI Native 权限确认已迁移至 Python sidecar 的 PreToolUse hook，
        /// 此处不再需要同步 MessageBox 确认方法。
        /// </summary>

        private ToolResult ExecuteCreatePivot(
            string sourceRange,
            string destSheet,
            string pivotName,
            string[] rowFields,
            string[] colFields,
            string[] valFields,
            string valFunc)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteCreatePivot: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "create_pivot_table", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.CreatePivotTable(sourceRange, destSheet, pivotName,
                    rowFields?.Length > 0 ? rowFields : null,
                    colFields?.Length > 0 ? colFields : null,
                    valFields?.Length > 0 ? valFields : null,
                    valFunc);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCreatePivot failed", ex);
                return new ToolResult { Name = "create_pivot_table", Success = false, Error = ex.Message };
            }
        }

        // ============= 参数提取（从 Orchestrator 移植）=============

        protected internal T GetArg<T>(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null) return default;
            try
            {
                var val = args[key];

                if (val is JsonElement je)
                {
                    if (je.ValueKind == JsonValueKind.Null || je.ValueKind == JsonValueKind.Undefined)
                        return default;

                    if (typeof(T) == typeof(string))
                        return (T)(object)je.GetString();

                    if (typeof(T) == typeof(int))
                        return (T)(object)je.GetInt32();

                    if (typeof(T) == typeof(bool))
                        return (T)(object)je.GetBoolean();

                    if (typeof(T) == typeof(double))
                        return (T)(object)je.GetDouble();

                    // ★ 支持 Nullable<double> (double?)
                    if (typeof(T) == typeof(double?))
                        return (T)(object)(double?)je.GetDouble();

                    // ★ 支持 Nullable<bool> (bool?)
                    if (typeof(T) == typeof(bool?))
                        return (T)(object)(bool?)je.GetBoolean();

                    // ★ 支持 Nullable<int> (int?)
                    if (typeof(T) == typeof(int?))
                        return (T)(object)(int?)je.GetInt32();

                    return (T)Convert.ChangeType(je.GetRawText(), typeof(T));
                }

                // 处理字符串值转 nullable 类型（模型可能传 "true"/"false" 字符串）
                if (val is string s)
                {
                    if (typeof(T) == typeof(double?))
                    {
                        if (double.TryParse(s, out double d)) return (T)(object)(double?)d;
                        return default;
                    }
                    if (typeof(T) == typeof(bool?))
                    {
                        if (bool.TryParse(s, out bool b)) return (T)(object)(bool?)b;
                        return default;
                    }
                    if (typeof(T) == typeof(int?))
                    {
                        if (int.TryParse(s, out int i)) return (T)(object)(int?)i;
                        return default;
                    }
                }

                return (T)Convert.ChangeType(val, typeof(T));
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", $"GetArg<{typeof(T).Name}>(\"{key}\") failed: " + ex.Message);
                return default;
            }
        }

        protected internal string[] GetArgArray(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key)) return new string[0];
            try
            {
                if (args[key] is JsonElement je)
                {
                    var list = new List<string>();
                    foreach (var item in je.EnumerateArray())
                    {
                        list.Add(item.GetString());
                    }
                    return list.ToArray();
                }
                if (args[key] is object[] arr)
                {
                    return Array.ConvertAll(arr, x => x?.ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Debug("ToolDispatcher", $"GetArgArray(\"{key}\") failed: " + ex.Message);
            }
            return new string[0];
        }

        /// <summary>
        /// ★ 从 args 中提取二维数组参数（write_range 用）。
        /// AI 传入的 values 是 JSON 二维数组，反序列化后是 JsonElement（ValueKind=Array）。
        /// 需要转换为 object[][] 供 MessageBridge.WriteRange 使用。
        /// </summary>
        protected internal object[][] Extract2DArray(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null) return null;
            try
            {
                if (args[key] is JsonElement je)
                {
                    if (je.ValueKind != JsonValueKind.Array) return null;
                    var rows = new List<object[]>();
                    foreach (var rowEl in je.EnumerateArray())
                    {
                        if (rowEl.ValueKind != JsonValueKind.Array)
                        {
                            // 单行非数组（如 [1,2,3] 被解析为单个值），包装为单元素行
                            rows.Add(new object[] { ConvertJsonElement(rowEl) });
                        }
                        else
                        {
                            var cols = new List<object>();
                            foreach (var cellEl in rowEl.EnumerateArray())
                            {
                                cols.Add(ConvertJsonElement(cellEl));
                            }
                            rows.Add(cols.ToArray());
                        }
                    }
                    return rows.ToArray();
                }
                if (args[key] is object[][] arr2d) return arr2d;
                if (args[key] is object[] arr1d)
                {
                    // 一维数组包装成二维
                    return new object[][] { arr1d };
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", $"Extract2DArray(\"{key}\") failed: " + ex.Message);
            }
            return null;
        }

        /// <summary>把 JsonElement 转换为合适的 CLR 类型</summary>
        private static object ConvertJsonElement(JsonElement je)
        {
            switch (je.ValueKind)
            {
                case JsonValueKind.String:
                    return je.GetString();
                case JsonValueKind.Number:
                    if (je.TryGetInt32(out int iv)) return iv;
                    return je.GetDouble();
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
                default:
                    return je.GetRawText();
            }
        }

        /// <summary>
        /// ★ read_attachment 工具：读取用户上传的附件文件内容。
        /// 支持读取 xlsx/xls/csv/txt 等格式。
        /// xlsx/xls 用 Excel COM 以 ReadOnly 方式打开读取（不修改当前工作簿），
        /// csv/txt 等文本文件直接读取内容。
        /// 返回 { fileName, sheets: [{ name, rowCount, columnCount, values }] }。
        /// </summary>
        private ToolResult ExecuteReadAttachment(Dictionary<string, object> args)
        {
            var fileName = GetArg<string>(args, "file_name");
            if (string.IsNullOrEmpty(fileName))
            {
                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = false,
                    Error = "file_name 是必传参数",
                    Suggestion = "请传入用户上传的附件文件名（可在对话上下文的附件列表中找到）",
                };
            }

            // 从 Attachments 映射中查找文件路径
            string filePath = null;
            if (Attachments != null && Attachments.Count > 0)
            {
                if (Attachments.TryGetValue(fileName, out var directPath))
                {
                    filePath = directPath;
                }
                else
                {
                    // 模糊匹配：用户可能只传了部分文件名
                    var match = Attachments.FirstOrDefault(kvp =>
                        kvp.Key.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        fileName.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (match.Value != null) filePath = match.Value;
                }
            }

            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
            {
                var available = Attachments != null && Attachments.Count > 0
                    ? string.Join(", ", Attachments.Keys)
                    : "无";
                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = false,
                    Error = $"附件 '{fileName}' 不存在",
                    Suggestion = $"已上传的附件: {available}。请用正确的文件名重新调用 read_attachment",
                };
            }

            Logger.Instance.Info("ToolDispatcher", $"read_attachment: fileName={fileName}, path={filePath}");

            var ext = System.IO.Path.GetExtension(fileName).ToLowerInvariant();

            // 文本文件：直接读取内容
            if (ext == ".txt" || ext == ".md" || ext == ".csv" || ext == ".json" || ext == ".xml" ||
                ext == ".log" || ext == ".py" || ext == ".cs" || ext == ".js" || ext == ".ts")
            {
                try
                {
                    var content = System.IO.File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                    // 截断过长的文本（避免 SDK 消息缓冲区溢出）
                    if (content.Length > 50000)
                    {
                        content = content.Substring(0, 50000) + "\n... [文件过长，已截断，共 " + content.Length + " 字符]";
                    }
                    return new ToolResult
                    {
                        Name = "read_attachment",
                        Success = true,
                        Data = new { fileName, type = "text", content },
                    };
                }
                catch (Exception ex)
                {
                    return new ToolResult
                    {
                        Name = "read_attachment",
                        Success = false,
                        Error = "读取文本附件失败: " + ex.Message,
                    };
                }
            }

            // Excel 文件：用 Excel COM 以 ReadOnly 方式打开读取
            if (ext == ".xlsx" || ext == ".xls" || ext == ".xlsm")
            {
                return ReadAttachmentExcel(filePath, fileName);
            }

            // ★ 图片附件：已通过 vision 方式直接发送给 AI，无需调用此工具
            if (ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".gif" ||
                ext == ".bmp" || ext == ".webp")
            {
                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = false,
                    Error = $"图片附件 {fileName} 已直接以 vision 方式发送给您，无需调用 read_attachment。",
                    Suggestion = "请直接根据图片内容回答用户问题，或用 write_range 把图片中的数据写入 Excel。",
                };
            }

            // ★ PDF 附件：已通过 document block 直接发送给 AI，无需调用此工具
            if (ext == ".pdf")
            {
                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = false,
                    Error = $"PDF 附件 {fileName} 已直接以 document 方式发送给您，无需调用 read_attachment。",
                    Suggestion = "请直接根据 PDF 内容回答用户问题，或用 write_range 把 PDF 中的表格数据写入 Excel。",
                };
            }

            // ★ Word 文档：Claude 不直接支持，提示用户转换
            if (ext == ".doc" || ext == ".docx")
            {
                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = false,
                    Error = $"Word 文档 {fileName} 暂不支持直接读取。",
                    Suggestion = "请提示用户把 Word 另存为 PDF 后重新上传，PDF 可以直接被 AI 识别。",
                };
            }

            return new ToolResult
            {
                Name = "read_attachment",
                Success = false,
                Error = $"不支持的附件格式: {ext}",
                Suggestion = "read_attachment 支持 xlsx/xls/csv/txt/json 等格式；图片和 PDF 会直接发送给 AI。",
            };
        }

        /// <summary>
        /// 用 Excel COM 以 ReadOnly 方式打开附件 xlsx，读取所有 sheet 数据后关闭。
        /// 临时关闭 ScreenUpdating 避免界面闪烁。
        /// </summary>
        private ToolResult ReadAttachmentExcel(string filePath, string fileName)
        {
            Workbook wb = null;
            bool prevScreenUpdating = true;
            try
            {
                prevScreenUpdating = _excelApp.ScreenUpdating;
                _excelApp.ScreenUpdating = false;

                // ReadOnly 打开，不更新链接，只读
                wb = _excelApp.Workbooks.Open(
                    Filename: filePath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    Origin: XlPlatform.xlWindows,
                    Editable: false,
                    Notify: false,
                    AddToMru: false);

                var sheetsData = new List<object>();
                foreach (Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        var usedRange = ws.UsedRange;
                        if (usedRange == null)
                        {
                            sheetsData.Add(new { name = ws.Name, rowCount = 0, columnCount = 0, values = new object[0][] });
                            continue;
                        }

                        int rowCount = usedRange.Rows.Count;
                        int colCount = usedRange.Columns.Count;

                        // 截断：每个 sheet 最多 200 行（与 read_range 一致，避免 SDK 缓冲区溢出）
                        bool truncated = false;
                        int actualRows = rowCount;
                        if (rowCount > MaxReadRangeRows)
                        {
                            actualRows = MaxReadRangeRows;
                            truncated = true;
                        }

                        // 读取数据（Value2 返回 2D 数组，1-based）
                        var rangeToRead = actualRows < rowCount
                            ? usedRange.Resize[actualRows, colCount]
                            : usedRange;
                        var rawValues = rangeToRead.Value2;

                        // 转换为 0-based 2D 数组（JSON 友好）
                        object[][] valuesArr = null;
                        if (rawValues is object[,] raw2D)
                        {
                            int r = raw2D.GetLength(0);
                            int c = raw2D.GetLength(1);
                            valuesArr = new object[r][];
                            for (int i = 0; i < r; i++)
                            {
                                valuesArr[i] = new object[c];
                                for (int j = 0; j < c; j++)
                                {
                                    valuesArr[i][j] = raw2D[i + 1, j + 1];
                                }
                            }
                        }
                        else
                        {
                            // 单个单元格
                            valuesArr = new object[][] { new object[] { rawValues } };
                        }

                        var sheetData = new Dictionary<string, object>
                        {
                            ["name"] = ws.Name,
                            ["rowCount"] = rowCount,
                            ["columnCount"] = colCount,
                            ["values"] = valuesArr,
                        };
                        if (truncated)
                        {
                            sheetData["truncated"] = true;
                            sheetData["original_row_count"] = rowCount;
                        }
                        sheetsData.Add(sheetData);
                    }
                    catch (Exception ex)
                    {
                        Logger.Instance.Warning("ToolDispatcher", $"read_attachment: sheet {ws.Name} read failed: {ex.Message}");
                        sheetsData.Add(new { name = ws.Name, error = ex.Message });
                    }
                }

                Logger.Instance.Info("ToolDispatcher", $"read_attachment: read {sheetsData.Count} sheets from {fileName}");

                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = true,
                    Data = new { fileName, type = "excel", sheets = sheetsData },
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", $"read_attachment: Open Excel failed: {filePath}", ex);
                return new ToolResult
                {
                    Name = "read_attachment",
                    Success = false,
                    Error = "打开附件 Excel 文件失败: " + ex.Message,
                    Suggestion = "附件文件可能被其他程序锁定或格式损坏",
                };
            }
            finally
            {
                // 关闭附件工作簿（不保存）
                if (wb != null)
                {
                    try { wb.Close(SaveChanges: false); }
                    catch (Exception ex) { Logger.Instance.Warning("ToolDispatcher", "read_attachment: close wb failed: " + ex.Message); }
                }
                try { _excelApp.ScreenUpdating = prevScreenUpdating; } catch { }
            }
        }

        // ============================ Computer Use 工具 ============================

        /// <summary>
        /// ★ Win32 API：获取窗口矩形（含标题栏和边框）
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;

        /// <summary>
        /// ★ 截图 Excel 主窗口，返回 base64 PNG（image content block）。
        /// 用于 Computer Use：AI 看不到 Excel 界面时，可调用此工具截图，
        /// 根据截图判断当前状态（如对话框、错误提示、图表效果等）。
        /// </summary>
        private ToolResult ExecuteScreenshotExcel()
        {
            try
            {
                if (_excelApp == null)
                {
                    return new ToolResult
                    {
                        Name = "screenshot_excel",
                        Success = false,
                        Error = "Excel Application 不可用",
                    };
                }

                // 获取 Excel 主窗口句柄
                IntPtr hwnd;
                try
                {
                    hwnd = new IntPtr(_excelApp.Hwnd);
                }
                catch (Exception ex)
                {
                    return new ToolResult
                    {
                        Name = "screenshot_excel",
                        Success = false,
                        Error = "获取 Excel 窗口句柄失败: " + ex.Message,
                    };
                }

                if (hwnd == IntPtr.Zero)
                {
                    return new ToolResult
                    {
                        Name = "screenshot_excel",
                        Success = false,
                        Error = "Excel 窗口句柄为 0",
                    };
                }

                // 最小化状态先恢复
                if (IsIconic(hwnd))
                {
                    ShowWindow(hwnd, SW_RESTORE);
                    Thread.Sleep(200);
                }

                // 获取窗口矩形
                RECT rect;
                if (!GetWindowRect(hwnd, out rect))
                {
                    return new ToolResult
                    {
                        Name = "screenshot_excel",
                        Success = false,
                        Error = "GetWindowRect 失败",
                    };
                }

                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                if (width <= 0 || height <= 0)
                {
                    return new ToolResult
                    {
                        Name = "screenshot_excel",
                        Success = false,
                        Error = $"窗口尺寸无效: width={width}, height={height}",
                    };
                }

                Logger.Instance.Info("ToolDispatcher", $"screenshot_excel: hwnd={hwnd}, rect=({rect.Left},{rect.Top},{rect.Right},{rect.Bottom}), size={width}x{height}");

                // 截图
                using (var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // ★ 体积优化：缩放 + JPEG 压缩，避免 base64 过大触发 SDK 1MB 缓冲区限制
                    // 全屏 PNG 通常 2-5MB，base64 后更大。目标 < 300KB。
                    var (scaledBmp, scaledW, scaledH) = ScaleToFit(bmp, 1280);
                    string base64;
                    int quality = 85;
                    using (var ms = new MemoryStream())
                    {
                        // JPEG 编码 + 指定质量
                        var jpegCodec = ImageCodecInfo.GetImageEncoders()
                            .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
                        var encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
                        scaledBmp.Save(ms, jpegCodec, encoderParams);
                        long len = ms.Length;

                        // 如果仍超过 300KB，逐步降低质量
                        while (len > 300 * 1024 && quality > 30)
                        {
                            quality -= 15;
                            ms.SetLength(0);
                            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
                            scaledBmp.Save(ms, jpegCodec, encoderParams);
                            len = ms.Length;
                        }

                        base64 = Convert.ToBase64String(ms.ToArray());
                    }

                    Logger.Instance.Info("ToolDispatcher", $"screenshot_excel: success, orig={width}x{height}, scaled={scaledW}x{scaledH}, jpeg q={quality}, base64={base64.Length}");

                    return new ToolResult
                    {
                        Name = "screenshot_excel",
                        Success = true,
                        Data = new
                        {
                            image_base64 = base64,
                            media_type = "image/jpeg",
                            width = scaledW,
                            height = scaledH,
                            original_width = width,
                            original_height = height,
                            jpeg_quality = quality,
                        },
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "screenshot_excel failed", ex);
                return new ToolResult
                {
                    Name = "screenshot_excel",
                    Success = false,
                    Error = "截图失败: " + ex.Message,
                };
            }
        }

        /// <summary>
        /// ★ 等比缩放图片到最大宽度（保持宽高比）。
        /// 如果原图宽度 ≤ maxWidth，返回原图（不放大）。
        /// 返回的 Bitmap 调用方负责 Dispose。
        /// </summary>
        private (Bitmap bmp, int width, int height) ScaleToFit(Bitmap src, int maxWidth)
        {
            if (src.Width <= maxWidth)
            {
                // 不缩放，但转 Format24bppRgb（JPEG 不支持 32bppArgb，会黑屏）
                var copy = new Bitmap(src.Width, src.Height, PixelFormat.Format24bppRgb);
                using (var g = Graphics.FromImage(copy))
                {
                    g.DrawImage(src, 0, 0, src.Width, src.Height);
                }
                return (copy, src.Width, src.Height);
            }

            int newW = maxWidth;
            int newH = (int)((long)src.Height * maxWidth / src.Width);
            var scaled = new Bitmap(newW, newH, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, newW, newH);
            }
            return (scaled, newW, newH);
        }

        /// <summary>
        /// ★ 模拟键盘输入到 Excel 窗口（Computer Use）。
        /// 使用 SendKeys.SendWait，支持特殊键语法：
        ///   {ENTER} {ESC} {TAB} {BACKSPACE} {DELETE} {UP} {DOWN} {LEFT} {RIGHT}
        ///   {F1}~{F12} {HOME} {END} {PGUP} {PGDN}
        ///   + = Shift, ^ = Ctrl, % = Alt（如 ^c 复制, ^v 粘贴, %{F4} Alt+F4）
        /// 用于：操作 Excel 对话框、触发快捷键、关闭弹窗等。
        /// </summary>
        private ToolResult ExecuteSendKeys(string keys)
        {
            if (string.IsNullOrEmpty(keys))
            {
                return new ToolResult
                {
                    Name = "send_keys",
                    Success = false,
                    Error = "keys 是必传参数",
                    Suggestion = "传入要模拟的按键，如 \"{ENTER}\"、\"^c\"（Ctrl+C）、\"%{F4}\"（Alt+F4）",
                };
            }

            try
            {
                // ★ 关键：发送按键前必须把焦点从 WebView2 面板转到 Excel 工作表
                // 否则 Ctrl+C/V/Z/S 等快捷键会作用于面板的输入框，而非 Excel 单元格
                if (_excelApp != null)
                {
                    try
                    {
                        // 1. 激活 Excel 主窗口到前台
                        IntPtr hwnd = new IntPtr(_excelApp.Hwnd);
                        if (hwnd != IntPtr.Zero)
                        {
                            SetForegroundWindow(hwnd);
                            Thread.Sleep(50);
                        }

                        // 2. 激活当前工作表的 ActiveCell，把焦点从 WebView2 转到工作表
                        //    这是跨进程转移焦点的可靠方式（SetFocus 跨进程无效）
                        var activeSheet = _excelApp.ActiveSheet as Worksheet;
                        if (activeSheet != null)
                        {
                            try
                            {
                                // 激活工作表，再激活一个单元格，强制 Excel 工作区获取焦点
                                activeSheet.Activate();
                                var activeCell = _excelApp.ActiveCell;
                                if (activeCell != null)
                                {
                                    activeCell.Activate();
                                    Thread.Sleep(50);
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Instance.Warning("ToolDispatcher", "send_keys: activate cell failed: " + ex.Message);
                            }
                        }
                    }
                    catch { /* 忽略，继续尝试 SendKeys */ }
                }

                Logger.Instance.Info("ToolDispatcher", $"send_keys: keys=[{keys}]");

                // ★ SendKeys 必须在 STA 线程，ToolDispatcher.Execute 已在 STA 线程
                SendKeys.SendWait(keys);

                // 等待 Excel 处理按键
                Thread.Sleep(200);

                return new ToolResult
                {
                    Name = "send_keys",
                    Success = true,
                    Data = new
                    {
                        keys_sent = keys,
                        note = "按键已发送到 Excel 工作表（已自动转移焦点）",
                    },
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "send_keys failed: " + keys, ex);
                return new ToolResult
                {
                    Name = "send_keys",
                    Success = false,
                    Error = "模拟键盘输入失败: " + ex.Message,
                    Suggestion = "keys 参数语法：{ENTER}/{TAB}/{ESC} 等特殊键用大括号；+Shift ^Ctrl %Alt 修饰键前缀",
                };
            }
        }
    }
}
