using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DeepExcel.AddIn.Sidecar;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>
    /// sort_data (ExcelActions.SortData): rows of the range reordered by one
    /// column with Range.Sort -- pinyin order for text, header row optional.
    /// </summary>
    public sealed partial class PreviewBuilder
    {
        private static readonly CompareInfo PinyinOrder = CultureInfo.GetCultureInfo("zh-CN").CompareInfo;

        private ChangePreview BuildSortData(string toolName, IDictionary<string, object> args)
        {
            var target = ResolveTarget(args, "range_address", "address", "range");
            if (target == null)
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }

            var preview = new ChangePreview { ToolName = toolName };
            var hasHeaderArg = GetBool(args, "has_header");
            if (hasHeaderArg == null)
            {
                // ToolDispatcher refuses the call and asks the model to retry.
                preview.Warnings.Add("缺少必传参数 has_header，执行器会拒绝本次排序，不会改动任何格");
                Finalize(preview);
                return preview;
            }
            var descending = GetBool(args, "descending") ?? false;

            var failed = ReadGrid(toolName, target, out var grid);
            if (failed != null) return failed;

            // sort_column: a number is an absolute sheet column, otherwise letters.
            var sortColumn = (GetString(args, "sort_column") ?? "").Trim();
            int absolute;
            if (!int.TryParse(sortColumn, out absolute))
            {
                absolute = CellRect.TryParse(sortColumn + "1", null, out var col) ? col.Col1 : 0;
            }
            var key = absolute - grid.Rect.Col1;
            if (key < 0 || key >= grid.Cols)
            {
                preview.Warnings.Add($"排序列「{sortColumn}」不在区域 {LocalAddress(grid.Rect)} 内，执行器会报错，不会改动任何格");
                Finalize(preview);
                return preview;
            }

            var hasHeader = hasHeaderArg.Value;
            if (!hasHeader && LooksLikeHeader(grid))
            {
                hasHeader = true;
                preview.Warnings.Add("has_header=false，但第一行是文本、下面是数字，像表头：执行器会自动按有表头排序，第一行不参与");
            }

            var first = hasHeader ? 1 : 0;
            if (grid.Rows - first < 2)
            {
                Finalize(preview);
                return preview;
            }

            var order = Enumerable.Range(first, grid.Rows - first)
                .Select(r => new { Row = r, Key = RawValue(grid[r, key]) })
                .ToList();
            // LINQ ordering is stable, as Excel's sort is: ties keep their order.
            var sorted = order
                .OrderBy(x => x.Key == null ? 1 : 0)
                .ThenBy(x => x.Key, new ExcelSortComparer(descending))
                .Select(x => x.Row)
                .ToList();

            var formulas = false;
            for (var i = 0; i < sorted.Count; i++)
            {
                var destRow = first + i;
                var fromRow = sorted[i];
                for (var c = 0; c < grid.Cols; c++)
                {
                    var origin = grid[fromRow, c];
                    var cell = grid[destRow, c];
                    formulas |= origin.HasFormula;
                    var after = origin.HasFormula
                        ? FormulaShift.ShiftRows(origin.Formula, destRow - fromRow)
                        : Describe(origin);
                    var before = Describe(cell);
                    if (string.Equals(before, after, StringComparison.Ordinal)) continue;
                    // A formula moved elsewhere is not a formula lost, so this
                    // is not flagged as an overwrite.
                    preview.Changes.Add(new CellChange
                    {
                        Address = cell.Address,
                        Before = before,
                        After = after,
                        Kind = after.Length == 0 ? ChangeKind.Clear : cell.IsEmpty ? ChangeKind.Add : ChangeKind.Overwrite
                    });
                }
            }
            preview.Changes = preview.Changes
                .OrderBy(c => RowOf(c.Address)).ThenBy(c => ColumnOf(c.Address)).ToList();

            if (preview.Changes.Count > 0)
            {
                var firstRow = grid.Rect.Row1 + first;
                preview.Warnings.Insert(0,
                    $"按 {CellRect.ColumnName(absolute)} 列{(descending ? "降序" : "升序")}重排第 {firstRow}–{grid.Rect.Row2} 行" +
                    (hasHeader ? $"（第 {grid.Rect.Row1} 行为表头，不参与）" : ""));
            }
            if (formulas && preview.Changes.Count > 0)
            {
                preview.Warnings.Add("区域里有公式：排序后相对引用随行调整，引用其他行的公式结果可能改变");
            }
            Finalize(preview);
            return preview;
        }

        /// <summary>
        /// The executor's has_header correction: every one of the first ten
        /// columns has text (not numeric text) in row 1, and some later row
        /// has a number in one of them.
        /// </summary>
        private static bool LooksLikeHeader(Grid grid)
        {
            if (grid.Rows < 2) return false;
            var laterNumber = false;
            for (var c = 0; c < Math.Min(grid.Cols, 10); c++)
            {
                var head = RawValue(grid[0, c]);
                if (!(head is string text) || text.Length == 0 || double.TryParse(text, out _)) return false;
                for (var r = 1; r < grid.Rows && !laterNumber; r++)
                {
                    var value = RawValue(grid[r, c]);
                    if (value is double || (value is string s && double.TryParse(s, out _))) laterNumber = true;
                }
            }
            return laterNumber;
        }

        /// <summary>
        /// Excel's ascending order: numbers, then text (pinyin, case-blind),
        /// then logical values, then errors. Descending reverses it. Blanks
        /// are handled by the caller: they go last either way.
        /// </summary>
        private sealed class ExcelSortComparer : IComparer<object>
        {
            private readonly int _sign;

            public ExcelSortComparer(bool descending)
            {
                _sign = descending ? -1 : 1;
            }

            public int Compare(object x, object y)
            {
                if (x == null || y == null) return 0;
                var rank = Rank(x).CompareTo(Rank(y));
                if (rank != 0) return _sign * rank;
                int result;
                if (x is double a && y is double b) result = a.CompareTo(b);
                else if (x is string s && y is string t) result = PinyinOrder.Compare(s, t, CompareOptions.IgnoreCase);
                else if (x is bool p && y is bool q) result = p.CompareTo(q);
                else result = 0;
                return _sign * result;
            }

            private static int Rank(object value)
            {
                if (value is double) return 0;
                if (value is string) return 1;
                if (value is bool) return 2;
                return 3;
            }
        }

        private static bool? GetBool(IDictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out var value) || value == null) return null;
            if (value is bool b) return b;
            if (value is double d) return d != 0;
            if (value is int i) return i != 0;
            return bool.TryParse(value.ToString(), out var parsed) ? parsed : (bool?)null;
        }
    }
}
