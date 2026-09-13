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

                var workbook = _excelApp?.ActiveWorkbook;
                if (workbook == null)
                {
                    return null;
                }

                var clock = Stopwatch.StartNew();
                var index = new SheetProfiler().Build(workbook, _excelApp);
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
