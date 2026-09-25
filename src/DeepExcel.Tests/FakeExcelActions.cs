using System;
using System.Collections.Generic;
using DeepExcel.AddIn.Bridge;

namespace DeepExcel.Tests
{
    /// <summary>
    /// ★ 共享的手写 IExcelActions mock —— 不引入 Moq 依赖（环境无 Moq 包）。
    ///
    /// 为什么抽成独立文件：之前 ToolDispatcherTests 和 PythonSidecarTests 各自维护一份
    /// 私有 FakeExcelActions，IExcelActions 每加一个方法（行列操作、冻结窗格、条件格式、
    /// 快照列表等）两份 mock 都会编译失败，测试套件因此长期跑不起来。
    /// 现在只在这里跟随接口演进一次。
    ///
    /// 用法：未配置的方法走默认实现（success / 空对象），不会抛异常；
    /// 需要断言的行为用 XxxFn 覆盖，用 XxxCalls 读取调用记录。
    /// </summary>
    internal class FakeExcelActions : IExcelActions
    {
        // ============ 可配置行为 ============
        public Func<object> GetSelectionFn { get; set; } = () => null;
        public Func<string, object> ReadRangeFn { get; set; } = _ => new { };
        public Func<object> ReadWorkbookFn { get; set; } = () => new { };
        public Func<string, object> ReadWorksheetFn { get; set; } = _ => new { };
        public Func<string, string, ToolResult> ExecuteVBAFn { get; set; } = (c, m) => new ToolResult { Success = true };
        public Func<string, ToolResult> ExecutePythonFn { get; set; } = _ => new ToolResult { Success = true };
        public Func<string, string, ToolResult> WriteFormulaFn { get; set; } = (a, f) => new ToolResult { Success = true };
        public Func<string, object, ToolResult> WriteValueFn { get; set; } = (a, v) => new ToolResult { Success = true };
        public Func<string, object[][], ToolResult> WriteRangeFn { get; set; } = (a, v) => new ToolResult { Success = true };
        public Func<string> CreateSnapshotFn { get; set; } = () => "snap-1";
        public Func<string, DeepExcel.AddIn.Executor.RollbackResult> RollbackFn { get; set; }
            = id => new DeepExcel.AddIn.Executor.RollbackResult { Success = true, SnapshotId = id };
        public Func<string, string, DeepExcel.AddIn.Executor.SnapshotScope, DeepExcel.AddIn.Executor.SnapshotAttempt> BackupWorkbookFn { get; set; }
        public Func<string, DeepExcel.AddIn.Executor.SnapshotScope, bool> ExtendSnapshotScopeFn { get; set; } = (id, scope) => true;
        public Func<string, DeepExcel.AddIn.Executor.SnapshotMeta> GetSnapshotMetaFn { get; set; } = _ => null;
        public string ActiveWorkbookKey { get; set; } = @"C:\data\book.xlsx";
        public string ActiveSheetName { get; set; } = "Sheet1";
        /// <summary>默认所有区域都是空的：不做先读后写检查的测试不受影响</summary>
        public Func<string, bool> RangeHasContentFn { get; set; } = _ => false;
        /// <summary>默认 null：不做体检（和没有工作簿时一样），不关心体检的测试不受影响</summary>
        public Func<DeepExcel.AddIn.Sidecar.HealthSnapshot> CaptureHealthFn { get; set; } = () => null;
        public Func<string, int, List<DeepExcel.AddIn.Sidecar.CellSample>> SampleCellsFn { get; set; }
            = (_, __) => new List<DeepExcel.AddIn.Sidecar.CellSample>();
        public List<string> SampleCellsCalls { get; } = new List<string>();
        /// <summary>默认 null：不做内联 diff</summary>
        public Func<string, object[,]> ReadFormulasFn { get; set; } = _ => null;
        /// <summary>重定向到的工作簿；OpenWorkbooks 里没有就返回 null（已关闭）</summary>
        public List<string> TargetCalls { get; } = new List<string>();
        public HashSet<string> OpenWorkbooks { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { @"C:\data\book.xlsx" };
        public Func<List<DeepExcel.AddIn.Executor.SnapshotMeta>> ListSnapshotsFn { get; set; }
            = () => new List<DeepExcel.AddIn.Executor.SnapshotMeta>();
        public Func<string, bool> DeleteSnapshotFn { get; set; } = _ => true;

        // ============ 调用记录 ============
        public int GetSelectionCalls { get; private set; }
        public int ReadWorkbookCalls { get; private set; }
        public List<(string, string)> WriteFormulaCalls { get; } = new List<(string, string)>();
        public List<(string, string)> ExecuteVBACalls { get; } = new List<(string, string)>();
        public List<string> ExecutePythonCalls { get; } = new List<string>();
        public int CreateSnapshotCalls { get; private set; }
        public List<string> RollbackCalls { get; } = new List<string>();
        /// <summary>(workbookKey, reason, scope)</summary>
        public List<(string Key, string Reason, DeepExcel.AddIn.Executor.SnapshotScope Scope)> BackupCalls { get; }
            = new List<(string, string, DeepExcel.AddIn.Executor.SnapshotScope)>();
        public List<(string Id, DeepExcel.AddIn.Executor.SnapshotScope Scope)> ExtendScopeCalls { get; }
            = new List<(string, DeepExcel.AddIn.Executor.SnapshotScope)>();
        /// <summary>按顺序记录"备份"和"执行"，用于断言备份发生在写入之前</summary>
        public List<string> Timeline { get; } = new List<string>();
        /// <summary>其余"只要能调通就行"的方法统一记到这里：(方法名, 主要参数)</summary>
        public List<(string Method, string Arg)> OtherCalls { get; } = new List<(string, string)>();

        // ============ 感知类 ============
        public object GetSelection()
        {
            GetSelectionCalls++;
            return GetSelectionFn();
        }

        public object ReadRange(string address) => ReadRangeFn(address);

        public object ReadWorkbook()
        {
            ReadWorkbookCalls++;
            return ReadWorkbookFn();
        }

        public object ReadWorksheet(string name) => ReadWorksheetFn(name);

        // ============ 执行类 ============
        public ToolResult ExecuteVBA(string code, string macroName = null)
        {
            ExecuteVBACalls.Add((code, macroName));
            Timeline.Add("execute_vba");
            return ExecuteVBAFn(code, macroName);
        }

        public ToolResult ExecutePython(string code)
        {
            ExecutePythonCalls.Add(code);
            Timeline.Add("execute_python");
            return ExecutePythonFn(code);
        }

        public ToolResult WriteFormula(string address, string formula)
        {
            WriteFormulaCalls.Add((address, formula));
            Timeline.Add("write_formula");
            return WriteFormulaFn(address, formula);
        }

        public ToolResult WriteValue(string address, object value)
        {
            Timeline.Add("write_value");
            return WriteValueFn(address, value);
        }

        public ToolResult WriteRange(string address, object[][] values)
        {
            Timeline.Add("write_range");
            return WriteRangeFn(address, values);
        }

        // ============ Sheet 管理 ============
        public ToolResult AddSheet(string name) => Record("add_sheet", name);
        public ToolResult DeleteSheet(string name) => Record("delete_sheet", name);
        public ToolResult RenameSheet(string oldName, string newName) => Record("rename_sheet", oldName + "->" + newName);

        // ============ 格式化 ============
        public ToolResult SetNumberFormat(string address, string format) => Record("set_number_format", address);
        public ToolResult SetColumnWidth(string address, double width, bool autoFit) => Record("set_column_width", address);

        // ============ 数据操作 ============
        public ToolResult SortData(string rangeAddress, string sortColumn, bool descending, bool hasHeader = false)
            => Record("sort_data", rangeAddress);

        public ToolResult FilterData(string rangeAddress, int columnIndex, string criteria)
            => Record("filter_data", rangeAddress);

        // ============ 单元格操作 ============
        public ToolResult MergeCells(string address) => Record("merge_cells", address);
        public ToolResult UnmergeCells(string address) => Record("unmerge_cells", address);

        public ToolResult SetCellStyle(string address, string fontName, double? fontSize, bool? bold, bool? italic,
            string fontColor, string bgColor, string hAlign, string vAlign, bool? wrapText)
            => Record("set_cell_style", address);

        public ToolResult CopyRange(string sourceAddress, string destAddress)
            => Record("copy_range", sourceAddress + "->" + destAddress);

        public ToolResult ClearRange(string address, string clearType) => Record("clear_range", address);

        // ============ 行列操作 ============
        public ToolResult InsertRows(int row, int count) => Record("insert_rows", row + "+" + count);
        public ToolResult DeleteRows(int row, int count) => Record("delete_rows", row + "+" + count);
        public ToolResult InsertColumns(int column, int count) => Record("insert_columns", column + "+" + count);
        public ToolResult DeleteColumns(int column, int count) => Record("delete_columns", column + "+" + count);

        // ============ 视图 / 高级 ============
        public ToolResult FreezePanes(string address) => Record("freeze_panes", address);

        public ToolResult ApplyConditionalFormat(string address, string ruleType, object ruleArgs)
            => Record("apply_conditional_format", address);

        public ToolResult WriteTable(string address, string tableName) => Record("write_table", address);

        // ============ 安全类 ============
        public string CreateSnapshot()
        {
            CreateSnapshotCalls++;
            return CreateSnapshotFn();
        }

        public DeepExcel.AddIn.Executor.RollbackResult Rollback(string snapshotId)
        {
            RollbackCalls.Add(snapshotId);
            return RollbackFn(snapshotId);
        }

        private int _backupSeq;

        public DeepExcel.AddIn.Executor.SnapshotAttempt BackupWorkbook(
            string workbookKey, string reason, DeepExcel.AddIn.Executor.SnapshotScope scope)
        {
            BackupCalls.Add((workbookKey, reason, scope));
            Timeline.Add("backup");
            if (BackupWorkbookFn != null) return BackupWorkbookFn(workbookKey, reason, scope);
            return new DeepExcel.AddIn.Executor.SnapshotAttempt { SnapshotId = "backup-" + (++_backupSeq) };
        }

        public bool ExtendSnapshotScope(string snapshotId, DeepExcel.AddIn.Executor.SnapshotScope scope)
        {
            ExtendScopeCalls.Add((snapshotId, scope));
            return ExtendSnapshotScopeFn(snapshotId, scope);
        }

        public DeepExcel.AddIn.Executor.SnapshotMeta GetSnapshotMeta(string snapshotId) => GetSnapshotMetaFn(snapshotId);

        public string GetActiveWorkbookKey() => ActiveWorkbookKey;

        public string GetActiveSheetName() => ActiveSheetName;

        public bool RangeHasContent(string address) => RangeHasContentFn(address);

        public DeepExcel.AddIn.Sidecar.HealthSnapshot CaptureHealth(int maxCollected) => CaptureHealthFn();

        public object[,] ReadFormulas(string address, int maxCells) => ReadFormulasFn(address);

        public List<(string Query, bool InFormulas, IList<string> Sheets, bool WholeCell, int Max)> FindCalls { get; }
            = new List<(string, bool, IList<string>, bool, int)>();
        public Func<string, object> FindCellsFn { get; set; } = q => new { query = q, total = 0, matches = new object[0] };

        public object FindCells(string query, bool inFormulas, IList<string> sheets, bool wholeCell, int maxResults)
        {
            FindCalls.Add((query, inFormulas, sheets, wholeCell, maxResults));
            return FindCellsFn(query);
        }

        public List<(string Sheet, int MaxCells)> SnapshotCalls { get; } = new List<(string, int)>();
        public Func<string, object> SheetSnapshotFn { get; set; } = s => new { sheet = s ?? "Sheet1", cells = new object[0] };
        public object SheetSnapshot(string sheetName, int maxCells)
        {
            SnapshotCalls.Add((sheetName, maxCells));
            return SheetSnapshotFn(sheetName);
        }

        public List<string> ListCalls { get; } = new List<string>();
        public Func<string, object> ListObjectsFn { get; set; } = k => new { kind = k, count = 0, items = new object[0] };

        public object ListObjects(string kind)
        {
            ListCalls.Add(kind);
            return ListObjectsFn(kind);
        }

        public List<(string Address, int Offset, int? Limit)> ReadPageCalls { get; } = new List<(string, int, int?)>();

        /// <summary>默认沿用 ReadRangeFn（不关心分页的测试不受影响）</summary>
        public Func<string, int, int?, object> ReadRangePageFn { get; set; }

        public object ReadRangePage(string address, int offset, int? limit)
        {
            ReadPageCalls.Add((address, offset, limit));
            return ReadRangePageFn != null ? ReadRangePageFn(address, offset, limit) : ReadRangeFn(address);
        }

        public IDisposable UseTargetWorkbook(string workbookKey)
        {
            TargetCalls.Add(workbookKey);
            return OpenWorkbooks.Contains(workbookKey) ? new NoopScope() : null;
        }

        private sealed class NoopScope : IDisposable { public void Dispose() { } }

        public List<DeepExcel.AddIn.Sidecar.CellSample> SampleCells(string address, int max)
        {
            SampleCellsCalls.Add(address);
            return SampleCellsFn(address, max);
        }

        public List<DeepExcel.AddIn.Executor.SnapshotMeta> ListSnapshots() => ListSnapshotsFn();

        public bool DeleteSnapshot(string snapshotId) => DeleteSnapshotFn(snapshotId);

        private ToolResult Record(string method, string arg)
        {
            OtherCalls.Add((method, arg));
            Timeline.Add(method);
            return new ToolResult { Name = method, Success = true };
        }
    }
}
