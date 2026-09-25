using System;
using System.Diagnostics;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Perception;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Supplies the workbook structure summary attached to every request.
    ///
    /// Replaces the previous approach, where the model discovered structure by
    /// calling read_range and inferring it from raw rows. On a real workbook that
    /// fails three ways at once: it cannot read 8000 rows, the rows it does read
    /// dominate the token budget, and the structure it infers is frequently
    /// wrong -- which is the actual cause of "AI 又改错地方了".
    /// </summary>
    public partial class MessageBridge
    {
        private readonly IndexCache _indexCache = new IndexCache();

        /// <summary>
        /// Structure summary for the active workbook, cached until it changes.
        ///
        /// Returns null rather than throwing on any failure. A missing index
        /// costs accuracy; an exception here would break the message the user
        /// just sent.
        /// </summary>
        internal string GetSemanticIndex(string workbookKey)
        {
            try
            {
                var cached = _indexCache.TryGet(workbookKey, DateTime.UtcNow);
                if (cached != null)
                {
                    return cached;
                }

                // 按会话绑定的工作簿找，而不是取活动工作簿：用户切到别的工作簿时，
                // 摘要不能描述另一本
                var workbook = FindOpenWorkbook(workbookKey);
                if (workbook == null)
                {
                    return StaleIndex(workbookKey, "工作簿不在当前 Excel 里");
                }

                var clock = Stopwatch.StartNew();
                WorkbookIndex index;
                try
                {
                    index = new SheetProfiler().Build(workbook, _excelApp);
                }
                catch (Exception ex)
                {
                    // 最常见的是用户正在编辑单元格（COM 返回 0x800AC472）
                    Logger.Instance.Warning("MessageBridge", "Semantic index rebuild failed: " + ex.Message);
                    return StaleIndex(workbookKey, DescribeComFailure(ex));
                }
                _indexCache.Store(workbookKey, index, DateTime.UtcNow);

                var rendered = _indexCache.TryGet(workbookKey, DateTime.UtcNow);
                Logger.Instance.Info("MessageBridge",
                    $"Semantic index built: sheets={index.Sheets.Count}, " +
                    $"~{index.EstimateTokens()} tokens, {clock.ElapsedMilliseconds}ms" +
                    (string.IsNullOrEmpty(index.Incomplete) ? "" : $", incomplete: {index.Incomplete}"));
                return rendered;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "Semantic index failed: " + ex.Message);
                return null;
            }
        }

        private string StaleIndex(string workbookKey, string reason)
        {
            if (!_indexCache.TryGetStale(workbookKey, out var rendered, out var builtUtc))
            {
                return null;
            }
            Logger.Instance.Info("MessageBridge", $"Semantic index: serving stale cache ({reason})");
            return IndexCache.MarkStale(rendered, DateTime.UtcNow - builtUtc, reason);
        }

        private Workbook FindOpenWorkbook(string workbookKey)
        {
            if (_excelApp == null)
            {
                return null;
            }
            var active = _excelApp.ActiveWorkbook;
            if (string.IsNullOrEmpty(workbookKey))
            {
                return active;
            }
            if (active != null && string.Equals(GetWorkbookKey(active), workbookKey, StringComparison.OrdinalIgnoreCase))
            {
                return active;
            }
            foreach (Workbook wb in _excelApp.Workbooks)
            {
                if (string.Equals(GetWorkbookKey(wb), workbookKey, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
            }
            return null;
        }

        internal static string DescribeComFailure(Exception ex)
        {
            var hr = (ex as System.Runtime.InteropServices.COMException)?.ErrorCode ?? ex.HResult;
            switch (unchecked((uint)hr))
            {
                case 0x800AC472: return "Excel 正在编辑单元格或有对话框打开";
                case 0x8001010A: return "Excel 正忙";
                case 0x80010001: return "Excel 拒绝了调用";
                default: return "Excel 暂时无响应";
            }
        }

        /// <summary>
        /// Called from Excel's SheetChange. Must stay trivial: it fires once per
        /// changed range, including inside a loop writing thousands of cells.
        /// </summary>
        internal void InvalidateSemanticIndex(string workbookKey)
        {
            _indexCache.Invalidate(workbookKey);
        }

        internal void ForgetSemanticIndex(string workbookKey)
        {
            _indexCache.Remove(workbookKey);
        }
    }
}
