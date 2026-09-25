using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Excel操作接口 - 由执行引擎/感知层实现
    /// Bridge层只负责协议转换，不关心具体业务
    /// </summary>
    public interface IExcelActions
    {
        // 感知类
        object GetSelection();
        object ReadRange(string address);
        object ReadWorkbook();
        object ReadWorksheet(string name);

        // 执行类
        ToolResult ExecuteVBA(string code, string macroName = null);
        ToolResult ExecutePython(string code);
        ToolResult WriteFormula(string address, string formula);
        ToolResult WriteValue(string address, object value);
        /// <summary>★ 批量写入二维数组到指定起始单元格（比逐个 write_value 快 100 倍）</summary>
        ToolResult WriteRange(string address, object[][] values);

        // Sheet 管理
        ToolResult AddSheet(string name);
        ToolResult DeleteSheet(string name);
        ToolResult RenameSheet(string oldName, string newName);

        // 格式化
        ToolResult SetNumberFormat(string address, string format);
        ToolResult SetColumnWidth(string address, double width, bool autoFit);

        // 数据操作
        ToolResult SortData(string rangeAddress, string sortColumn, bool descending, bool hasHeader = false);
        ToolResult FilterData(string rangeAddress, int columnIndex, string criteria);

        // 单元格操作
        ToolResult MergeCells(string address);
        ToolResult UnmergeCells(string address);
        ToolResult SetCellStyle(string address, string fontName, double? fontSize, bool? bold, bool? italic, string fontColor, string bgColor, string hAlign, string vAlign, bool? wrapText);
        ToolResult CopyRange(string sourceAddress, string destAddress);
        ToolResult ClearRange(string address, string clearType);

        // 行列操作
        ToolResult InsertRows(int row, int count);
        ToolResult DeleteRows(int row, int count);
        ToolResult InsertColumns(int column, int count);
        ToolResult DeleteColumns(int column, int count);

        // 视图
        ToolResult FreezePanes(string address);

        // 高级
        ToolResult ApplyConditionalFormat(string address, string ruleType, object ruleArgs);
        ToolResult WriteTable(string address, string tableName);

        // 安全类
        /// <summary>给当前活动工作簿做整本快照（模型的 create_snapshot 工具）；失败返回 null</summary>
        string CreateSnapshot();
        /// <summary>给指定 key 的已打开工作簿做快照，失败时 Error 说明原因（写入前的自动备份用）</summary>
        DeepExcel.AddIn.Executor.SnapshotAttempt BackupWorkbook(string workbookKey, string reason, DeepExcel.AddIn.Executor.SnapshotScope scope);
        /// <summary>扩大快照的覆盖范围；false 表示没记下，调用方应视为备份失败</summary>
        bool ExtendSnapshotScope(string snapshotId, DeepExcel.AddIn.Executor.SnapshotScope scope);
        DeepExcel.AddIn.Executor.SnapshotMeta GetSnapshotMeta(string snapshotId);
        /// <summary>恢复到快照所属的已打开工作簿（不看 ActiveWorkbook，不改写磁盘原文件）</summary>
        DeepExcel.AddIn.Executor.RollbackResult Rollback(string snapshotId);

        /// <summary>当前活动工作簿的会话 key（规则见 WorkbookIdentity）；没有时返回 null</summary>
        string GetActiveWorkbookKey();
        /// <summary>当前活动工作表名；没有时返回 null</summary>
        string GetActiveSheetName();
        /// <summary>区域里有没有非空单元格（先读后写检查用）；地址无效时返回 false，交给工具本身报错</summary>
        bool RangeHasContent(string address);
        /// <summary>写后体检：整本的公式错误单元格（最多收集 maxCollected 个）、外部链接、计算模式；失败返回 null</summary>
        DeepExcel.AddIn.Sidecar.HealthSnapshot CaptureHealth(int maxCollected);
        /// <summary>写后体检：回读区域里的几个单元格（前几个 + 最后一个）的显示值和公式</summary>
        List<DeepExcel.AddIn.Sidecar.CellSample> SampleCells(string address, int max);
        /// <summary>内联 diff：区域的 Formula（0 起始 [行, 列]）；超过 maxCells 格或地址无效返回 null</summary>
        object[,] ReadFormulas(string address, int maxCells);
        /// <summary>
        /// 在返回的范围内，所有读写都落到这本工作簿（按 key 找已打开的），而不是前台那本；
        /// 找不到返回 null。
        /// </summary>
        IDisposable UseTargetWorkbook(string workbookKey);
        /// <summary>
        /// read_range 的一页：先裁到已用区域，再从第 offset 行起最多读 limit 行（见 RangePaging）。
        /// 失败返回带 error / suggestion 的对象。
        /// </summary>
        object ReadRangePage(string address, int offset, int? limit);
        /// <summary>find：在目标工作簿里按值或公式搜索（Range.Find），返回命中位置和每表计数</summary>
        object FindCells(string query, bool inFormulas, IList<string> sheets, bool wholeCell, int maxResults);
        /// <summary>list：列出 sheets / names / tables / pivots / charts</summary>
        object ListObjects(string kind);
        /// <summary>
        /// sheet_snapshot：一次批量读出一张表的有界快照（值、R1C1 公式、合并、溢出区域、对象），
        /// 给侧车 perception 包做结构分析。sheetName 为空时用目标工作簿的活动表。
        /// </summary>
        object SheetSnapshot(string sheetName, int maxCells);

        // ★ 新增：历史版本管理（供前端 UI 调用）
        System.Collections.Generic.List<DeepExcel.AddIn.Executor.SnapshotMeta> ListSnapshots();
        bool DeleteSnapshot(string snapshotId);
    }
}
