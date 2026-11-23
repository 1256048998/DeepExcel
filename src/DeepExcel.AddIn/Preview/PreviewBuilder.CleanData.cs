using System;
using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Sidecar;
using DeepExcel.AddIn.Tools;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>
    /// clean_data: several cleaning operations run one after another
    /// (ToolDispatcher.ExecuteCleanData).
    ///
    /// Each operation works on what the previous one left, so the preview
    /// simulates the chain on a copy of the range instead of previewing each
    /// operation against the original: trimming can turn rows into duplicates
    /// that remove_duplicates then deletes. And because every operation
    /// re-resolves the same address, a row deletion pulls the rows below the
    /// range up into it, where the next operation acts on them -- so those
    /// rows are loaded too.
    /// </summary>
    public sealed partial class PreviewBuilder
    {
        private sealed class SimCell
        {
            public CellSnapshot Original;
            public object Value;
            public string Formula;
            public bool IsDate;
            public string NumberFormat;
            public string Fill;

            public bool FormatChanged;
            public bool ValueChanged;

            public void Write(object value)
            {
                Value = value;
                Formula = null;
                ValueChanged = true;
            }
        }

        private sealed class SimRow
        {
            public int OriginalRow;
            /// <summary>False for rows past the used range, which were empty.</summary>
            public bool HadData;
            public SimCell[] Cells;
        }

        private static readonly HashSet<string> CleanOperations = new HashSet<string>(StringComparer.Ordinal)
        {
            "unify_date", "remove_duplicates", "highlight_missing", "trim_spaces", "text_to_number",
            "delete_blank_rows", "fill_blank_cells", "highlight_duplicates", "remove_special_chars",
            "clean_amount", "collapse_spaces"
        };

        private ChangePreview BuildCleanData(string toolName, IDictionary<string, object> args)
        {
            var target = ResolveTarget(args, "range_address", "address", "range");
            if (target == null)
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }

            // The executor lower-cases each name and ignores unknown ones; a
            // null name throws there and stops the run, so nothing after it runs.
            var operations = new List<string>();
            foreach (var name in GetStringArray(args, "operations"))
            {
                if (name == null) break;
                var op = name.ToLower();
                if (CleanOperations.Contains(op)) operations.Add(op);
            }

            var failed = ReadGrid(toolName, target, out var grid);
            if (failed != null) return failed;

            var rows = new List<SimRow>();
            for (var r = 0; r < grid.Rows; r++)
            {
                rows.Add(ToSimRow(grid.Rect.Row1 + r, Enumerable.Range(0, grid.Cols).Select(c => grid[r, c])));
            }

            var below = new RowsBelow(this, target, grid);
            var deleted = new List<int>();
            var leadingZeros = new List<CellSnapshot>();
            var longDigits = new List<CellSnapshot>();
            var notDates = new List<CellSnapshot>();
            var duplicatesRemoved = false;
            foreach (var op in operations)
            {
                if (!below.Fill(rows, grid.Rows)) return Unreadable(toolName);
                var view = rows.Take(grid.Rows).ToList();
                switch (op)
                {
                    case "trim_spaces": EachCell(view, v => CleaningRules.TrimSpaces(v)); break;
                    case "collapse_spaces": EachCell(view, v => CleaningRules.CollapseSpaces(v)); break;
                    case "remove_special_chars": EachCell(view, v => CleaningRules.RemoveSpecialChars(v)); break;
                    case "text_to_number":
                        foreach (var cell in view.SelectMany(r => r.Cells))
                        {
                            if (!(cell.Value is string s) || CleaningRules.TextToNumber(s) == null) continue;
                            if (CleaningRules.HasLeadingZero(s)) leadingZeros.Add(cell.Original);
                            if (CleaningRules.LosesPrecision(s)) longDigits.Add(cell.Original);
                        }
                        EachCell(view, v => CleaningRules.TextToNumber(v), cell => cell.NumberFormat = "General");
                        break;
                    case "clean_amount":
                        EachCell(view, v => CleaningRules.CleanAmount(v), cell => cell.NumberFormat = CleaningRules.AmountFormat);
                        break;
                    case "unify_date":
                        notDates.AddRange(view.SelectMany(r => r.Cells)
                            .Where(c => c.Value is double d && !c.IsDate && (d < PlausibleDateFrom || d >= PlausibleDateTo))
                            .Select(c => c.Original));
                        SimUnifyDate(view);
                        break;
                    case "fill_blank_cells": SimFillBlanks(view); break;
                    case "highlight_missing": SimHighlightMissing(view); break;
                    case "highlight_duplicates": SimHighlightDuplicates(view); break;
                    case "delete_blank_rows":
                        DeleteRows(rows, grid.Rows, deleted,
                            row => row.Cells.All(c => CleaningRules.IsBlank(c.Value)));
                        break;
                    case "remove_duplicates":
                        var seen = new HashSet<string>(StringComparer.Ordinal);
                        var before = deleted.Count;
                        DeleteRows(rows, grid.Rows, deleted,
                            row => !seen.Add(CleaningRules.RowKey(row.Cells.Select(c => c.Value))));
                        duplicatesRemoved |= deleted.Count > before;
                        break;
                }
            }

            var preview = new ChangePreview { ToolName = toolName };
            foreach (var row in rows)
            {
                foreach (var cell in row.Cells)
                {
                    var change = Diff(cell);
                    if (change != null) preview.Changes.Add(change);
                }
            }
            preview.Changes = preview.Changes.OrderBy(c => RowOf(c.Address)).ThenBy(c => ColumnOf(c.Address)).ToList();
            deleted.Sort();
            preview.DeletedRows.AddRange(deleted);

            // The same warnings each operation gives on its own.
            if (duplicatesRemoved)
            {
                preview.Warnings.Add("重复的行保留最下面的一行，删除上面的（与 Excel 自带的「删除重复项」相反）");
            }
            if (leadingZeros.Count > 0)
            {
                preview.Warnings.Add($"{leadingZeros.Count} 个格以 0 开头（{Examples(leadingZeros)}），转成数字后前导零会丢失");
                preview.RequiresConfirmation = true;
            }
            if (longDigits.Count > 0)
            {
                preview.Warnings.Add($"{longDigits.Count} 个格超过 15 位数字（{Examples(longDigits)}，如身份证号），转成数字后末几位会变成 0");
                preview.RequiresConfirmation = true;
            }
            if (notDates.Count > 0)
            {
                preview.Warnings.Add($"{notDates.Count} 个数字不像日期（{Examples(notDates)}），套上日期格式后会显示成日期，数值本身不变");
                preview.RequiresConfirmation = true;
            }
            if (deleted.Count > 0)
            {
                var outside = DataOutsideDeletedRows(target, grid, deleted);
                if (outside != null) preview.Warnings.Add(outside);
                if (preview.Changes.Count > 0)
                {
                    preview.Warnings.Add("有行被删除，单元格地址按执行前的位置标注");
                }
                preview.Warnings.Add("删除行会使下方所有行上移，引用这些行的公式可能失效");
            }
            Finalize(preview);
            return preview;
        }

        /// <summary>
        /// Supplies the rows that move up into the range after a deletion:
        /// first the sheet's own rows below the range, then empty ones.
        /// </summary>
        private sealed class RowsBelow
        {
            private readonly PreviewBuilder _owner;
            private readonly Target _target;
            private readonly Grid _grid;
            private bool _loaded;
            private int _nextEmpty;

            public RowsBelow(PreviewBuilder owner, Target target, Grid grid)
            {
                _owner = owner;
                _target = target;
                _grid = grid;
            }

            public bool Fill(List<SimRow> rows, int needed)
            {
                if (rows.Count >= needed) return true;
                if (!_loaded)
                {
                    _loaded = true;
                    var used = _owner._reader.UsedRowCount(_target.Sheet);
                    _nextEmpty = Math.Max(used, _grid.Rect.Row2) + 1;
                    if (used > _grid.Rect.Row2)
                    {
                        var block = new CellRect(_grid.Rect.Sheet, _grid.Rect.Row2 + 1, _grid.Rect.Col1, used, _grid.Rect.Col2);
                        var count = (long)(block.Row2 - block.Row1 + 1) * _grid.Cols;
                        if (count + (long)_grid.Rows * _grid.Cols > ExcelRangeReader.MaxCellsRead) return false;
                        var cells = _owner._reader.Read(_target.Sheet, LocalAddress(block));
                        if (cells.Count < count) return false;
                        for (var r = 0; r < block.Row2 - block.Row1 + 1; r++)
                        {
                            rows.Add(ToSimRow(block.Row1 + r, cells.Skip(r * _grid.Cols).Take(_grid.Cols)));
                        }
                    }
                }
                while (rows.Count < needed)
                {
                    var row = _nextEmpty++;
                    var empty = Enumerable.Range(0, _grid.Cols).Select(c => new CellSnapshot
                    {
                        Address = CellRect.ColumnName(_grid.Rect.Col1 + c) + row,
                        Display = "",
                        Formula = ""
                    });
                    var sim = ToSimRow(row, empty);
                    sim.HadData = false;
                    rows.Add(sim);
                }
                return true;
            }
        }

        private static SimRow ToSimRow(int originalRow, IEnumerable<CellSnapshot> cells)
        {
            return new SimRow
            {
                OriginalRow = originalRow,
                HadData = true,
                Cells = cells.Select(c => new SimCell
                {
                    Original = c,
                    Value = RawValue(c),
                    Formula = c.HasFormula ? c.Formula : null,
                    IsDate = c.IsDate
                }).ToArray()
            };
        }

        private static void EachCell(IEnumerable<SimRow> view, Func<object, object> rule, Action<SimCell> format = null)
        {
            foreach (var cell in view.SelectMany(r => r.Cells))
            {
                var next = rule(cell.Value);
                if (next == null) continue;
                cell.Write(next);
                cell.IsDate = false;
                if (format != null)
                {
                    format(cell);
                    cell.FormatChanged = true;
                }
            }
        }

        private static void SimUnifyDate(IEnumerable<SimRow> view)
        {
            foreach (var cell in view.SelectMany(r => r.Cells))
            {
                switch (CleaningRules.UnifyDate(cell.Value, out var date))
                {
                    case CleaningRules.DateAction.FormatOnly:
                        if (cell.Value is double && !cell.IsDate)
                        {
                            cell.IsDate = true;
                            cell.FormatChanged = true;
                        }
                        break;
                    case CleaningRules.DateAction.Convert:
                        cell.Write(date.ToOADate());
                        cell.IsDate = true;
                        break;
                }
            }
        }

        private static void SimFillBlanks(List<SimRow> view)
        {
            if (view.Count == 0) return;
            for (var c = 0; c < view[0].Cells.Length; c++)
            {
                object last = null;
                foreach (var row in view)
                {
                    var cell = row.Cells[c];
                    if (!CleaningRules.IsBlank(cell.Value))
                    {
                        last = cell.Value;
                    }
                    else if (last != null)
                    {
                        cell.Write(last);
                    }
                }
            }
        }

        private static void SimHighlightMissing(IEnumerable<SimRow> view)
        {
            foreach (var cell in view.SelectMany(r => r.Cells).Where(c => CleaningRules.IsBlank(c.Value)))
            {
                cell.Fill = "黄色底色";
                cell.FormatChanged = true;
            }
        }

        private static void SimHighlightDuplicates(IEnumerable<SimRow> view)
        {
            var groups = view.SelectMany(r => r.Cells)
                .Where(c => !string.IsNullOrEmpty(c.Value?.ToString()))
                .GroupBy(c => c.Value.ToString(), StringComparer.Ordinal)
                .Where(g => g.Count() > 1);
            foreach (var cell in groups.SelectMany(g => g))
            {
                cell.Fill = "浅红底色";
                cell.FormatChanged = true;
            }
        }

        /// <summary>
        /// Bottom-up over the range's rows, as the executor deletes. Rows that
        /// move up from below during the pass are not looked at again.
        /// </summary>
        private static void DeleteRows(List<SimRow> rows, int rangeRows, List<int> deleted, Func<SimRow, bool> doomed)
        {
            for (var i = Math.Min(rangeRows, rows.Count) - 1; i >= 0; i--)
            {
                if (!doomed(rows[i])) continue;
                // An empty row past the used range vanishing changes nothing.
                if (rows[i].HadData) deleted.Add(rows[i].OriginalRow);
                rows.RemoveAt(i);
            }
        }

        private static CellChange Diff(SimCell cell)
        {
            if (!cell.ValueChanged && !cell.FormatChanged) return null;
            var original = cell.Original;
            var before = Describe(original);
            var after = ShowSim(cell);
            var valueSame = Equals(cell.Value, RawValue(original)) && cell.Formula == (original.HasFormula ? original.Formula : null);

            if (valueSame && string.Equals(before, after, StringComparison.Ordinal) && cell.Fill == null)
            {
                return null;
            }

            if (valueSame || (!cell.ValueChanged))
            {
                return new CellChange
                {
                    Address = original.Address,
                    Before = before,
                    After = cell.Fill != null ? (after.Length > 0 ? after + "（" + cell.Fill + "）" : cell.Fill) : after,
                    Kind = ChangeKind.Format
                };
            }

            // Whitespace and invisible characters only: show them.
            if (cell.Value is string text && RawValue(original) is string old && Invisible(old) == Invisible(text))
            {
                before = ShowSpecialChars(old);
                after = ShowSpecialChars(text);
            }
            return new CellChange
            {
                Address = original.Address,
                Before = before,
                After = cell.Fill != null ? after + "（" + cell.Fill + "）" : after,
                Kind = original.IsEmpty ? ChangeKind.Add : ChangeKind.Overwrite,
                OverwritesFormula = original.HasFormula && cell.Formula == null
            };
        }

        private static string Invisible(string text)
        {
            var collapsed = CleaningRules.CollapseSpaces(text) ?? text;
            return CleaningRules.RemoveSpecialChars(collapsed) ?? collapsed;
        }

        private static string ShowSim(SimCell cell)
        {
            if (cell.Formula != null) return cell.Formula;
            if (cell.Value is double number)
            {
                if (cell.IsDate) return SerialToDateText(number);
                if (cell.NumberFormat == CleaningRules.AmountFormat) return ShowAmount(number);
            }
            return Stringify(cell.Value);
        }

        private static int RowOf(string address)
        {
            return CellRect.TryParse(address, null, out var rect) ? rect.Row1 : 0;
        }

        private static int ColumnOf(string address)
        {
            return CellRect.TryParse(address, null, out var rect) ? rect.Col1 : 0;
        }
    }
}
