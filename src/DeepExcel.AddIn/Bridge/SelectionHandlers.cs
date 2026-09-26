using System;
using DeepExcel.AddIn.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 选区条：面板输入框上方显示「Sheet1!A1:D20 · 80 格」，让用户看到 AI 这一轮会看哪块。
    /// 模型每轮本来就拿到选区（BuildContext），这里只是把它推给面板；用户点 × 后那一条消息不带选区。
    /// </summary>
    public partial class MessageBridge
    {
        /// <summary>选区的轻量描述：只取地址和行列数，不读单元格</summary>
        public static object DescribeSelection(Excel.Range range)
        {
            if (range == null) return new { address = "" };
            var sheet = (range.Worksheet as Excel.Worksheet)?.Name ?? "";
            long rows = 0, cols = 0;
            try
            {
                // 行列数取第一块；多块选区（Ctrl 点选）看总格数和地址。整列 / 整行选择会是很大的数，照实给
                var first = range.Areas[1];
                rows = Convert.ToInt64(first.Rows.CountLarge);
                cols = Convert.ToInt64(first.Columns.CountLarge);
            }
            catch (Exception) { }
            long cells = 0;
            try { cells = Convert.ToInt64(range.CountLarge); } catch (Exception) { }
            return new
            {
                sheet,
                address = range.Address[false, false],
                rows,
                cols,
                cells,
            };
        }

        /// <summary>ThisAddIn 的 SheetSelectionChange（已防抖）调用；我们自己的工具执行期间不推</summary>
        public void OnSelectionChanged(string workbookKey, Excel.Range target)
        {
            if (string.IsNullOrEmpty(workbookKey) || target == null) return;
            try
            {
                if (!_sessions.TryGetValue(workbookKey, out var session) || session == null) return;
                if (session.Sidecar?.Dispatcher?.IsExecuting == true) return;
                SendToSessionUi(workbookKey, "selection_brief", DescribeSelection(target));
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "OnSelectionChanged failed: " + ex.Message);
            }
        }

        /// <summary>面板打开时要一次当前选区</summary>
        private string HandleGetSelectionBrief()
        {
            try
            {
                return MakeResponse("selection_brief", DescribeSelection(_excelApp.Selection as Excel.Range));
            }
            catch (Exception)
            {
                return MakeResponse("selection_brief", new { address = "" });
            }
        }
    }
}
