using System;
using System.Collections.Generic;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>
    /// Previews for the tools beyond plain writes, clears and deletes.
    ///
    /// Each one mirrors what its executor actually does, quirks included. Where
    /// the executor's behaviour looks like a bug, the preview still reports it
    /// faithfully: the point is to show what will happen, not what should.
    /// </summary>
    public sealed partial class PreviewBuilder
    {
        /// <summary>
        /// replace_formula: plain string replacement inside formula text
        /// (ToolDispatcher.ExecuteReplaceFormula). Case-sensitive, formulas only.
        /// </summary>
        private ChangePreview BuildReplaceFormula(string toolName, IDictionary<string, object> args)
        {
            var target = ResolveTarget(args, "range_address", "address", "range");
            if (target == null)
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }
            var find = GetString(args, "find");
            if (string.IsNullOrEmpty(find))
            {
                return ChangePreview.NotPreviewable(toolName, "缺少要查找的文本");
            }
            var replace = GetString(args, "replace") ?? "";

            var failed = ReadGrid(toolName, target, out var grid);
            if (failed != null) return failed;

            var preview = new ChangePreview { ToolName = toolName };
            foreach (var cell in grid.Cells)
            {
                if (!cell.HasFormula || cell.Formula.IndexOf(find, StringComparison.Ordinal) < 0)
                {
                    continue;
                }
                preview.Changes.Add(new CellChange
                {
                    Address = cell.Address,
                    Before = cell.Formula,
                    After = cell.Formula.Replace(find, replace),
                    Kind = ChangeKind.Overwrite,
                    OverwritesFormula = true
                });
            }
            Finalize(preview);
            return preview;
        }
    }
}
