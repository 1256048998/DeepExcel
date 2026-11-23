using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DeepExcel.AddIn.Tools;

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

        /// <summary>
        /// A cleaning operation that rewrites cells one at a time with a rule
        /// from <see cref="CleaningRules"/> -- the same rule DataCleaner runs.
        ///
        /// The rule sees the raw value, not the display text: the executor
        /// reads Value2, so a formula whose result needs cleaning is rewritten
        /// as a literal and loses its formula. That is reported as a formula
        /// overwrite, with the formula text as the "before".
        /// </summary>
        private ChangePreview BuildPerCell(
            string toolName, IDictionary<string, object> args,
            Func<object, object> rule, Func<object, string> show,
            Action<Grid, ChangePreview> annotate = null)
        {
            var target = ResolveTarget(args, "range_address", "address", "range");
            if (target == null)
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }
            var failed = ReadGrid(toolName, target, out var grid);
            if (failed != null) return failed;

            var preview = new ChangePreview { ToolName = toolName };
            foreach (var cell in grid.Cells)
            {
                var raw = RawValue(cell);
                var next = rule(raw);
                if (next == null) continue;
                preview.Changes.Add(new CellChange
                {
                    Address = cell.Address,
                    Before = cell.HasFormula ? cell.Formula : show(raw),
                    After = show(next),
                    Kind = ChangeKind.Overwrite,
                    OverwritesFormula = cell.HasFormula
                });
            }
            annotate?.Invoke(grid, preview);
            Finalize(preview);
            return preview;
        }

        /// <summary>
        /// fill_blank_cells: each column top to bottom, a blank cell takes the
        /// last non-blank value above it (DataCleaner.FillBlankCells). Blanks
        /// with nothing above stay blank. A formula showing "" counts as blank
        /// and is replaced by a literal.
        /// </summary>
        private ChangePreview BuildFillBlankCells(string toolName, IDictionary<string, object> args)
        {
            var target = ResolveTarget(args, "range_address", "address", "range");
            if (target == null)
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }
            var failed = ReadGrid(toolName, target, out var grid);
            if (failed != null) return failed;

            var preview = new ChangePreview { ToolName = toolName };
            var changes = new List<(int Row, int Col, CellChange Change)>();
            for (var c = 0; c < grid.Cols; c++)
            {
                object last = null;
                for (var r = 0; r < grid.Rows; r++)
                {
                    var cell = grid[r, c];
                    var raw = RawValue(cell);
                    if (!CleaningRules.IsBlank(raw))
                    {
                        last = raw;
                        continue;
                    }
                    if (last == null) continue;
                    changes.Add((r, c, new CellChange
                    {
                        Address = cell.Address,
                        Before = Describe(cell),
                        After = Stringify(last),
                        Kind = cell.IsEmpty ? ChangeKind.Add : ChangeKind.Overwrite,
                        OverwritesFormula = cell.HasFormula
                    }));
                }
            }
            // Listed row by row, the way the user reads the sheet.
            preview.Changes.AddRange(changes.OrderBy(x => x.Row).ThenBy(x => x.Col).Select(x => x.Change));
            Finalize(preview);
            return preview;
        }

        /// <summary>
        /// rename_columns: new_names written left to right into the range's
        /// first row (DataCleaner.RenameColumns); names past the last column
        /// are ignored. A null name clears the header.
        /// </summary>
        private ChangePreview BuildRenameColumns(string toolName, IDictionary<string, object> args)
        {
            var target = ResolveTarget(args, "range_address", "address", "range");
            if (target == null)
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }
            var names = GetStringArray(args, "new_names");
            var failed = ReadGrid(toolName, target, out var grid);
            if (failed != null) return failed;

            var preview = new ChangePreview { ToolName = toolName };
            for (var c = 0; c < names.Count && c < grid.Cols; c++)
            {
                AddChangeIfDifferent(preview, grid[0, c], names[c] ?? "");
            }
            Finalize(preview);
            return preview;
        }

        /// <summary>
        /// Text that text_to_number converts at a cost the user would not
        /// guess from "turn text into numbers": codes with leading zeros, and
        /// long digit strings such as ID numbers. Both are irreversible once
        /// saved, so they make the preview interrupt however few cells there are.
        /// </summary>
        private static void NumbersThatLoseDigits(Grid grid, ChangePreview preview)
        {
            var converted = grid.Cells
                .Where(c => RawValue(c) is string s && CleaningRules.TextToNumber(s) != null)
                .ToList();
            var zeros = converted.Where(c => CleaningRules.HasLeadingZero((string)RawValue(c))).ToList();
            var longDigits = converted.Where(c => CleaningRules.LosesPrecision((string)RawValue(c))).ToList();
            if (zeros.Count > 0)
            {
                preview.Warnings.Add($"{zeros.Count} 个格以 0 开头（{Examples(zeros)}），转成数字后前导零会丢失");
                preview.RequiresConfirmation = true;
            }
            if (longDigits.Count > 0)
            {
                preview.Warnings.Add($"{longDigits.Count} 个格超过 15 位数字（{Examples(longDigits)}，如身份证号），转成数字后末几位会变成 0");
                preview.RequiresConfirmation = true;
            }
        }

        private static string Examples(IList<CellSnapshot> cells)
        {
            return string.Join("、", cells.Take(3).Select(c => c.Address + "「" + Clip(RawValue(c) as string ?? Stringify(RawValue(c))) + "」"));
        }

        private static string ShowAmount(object value)
        {
            return value is double number
                ? number.ToString(CleaningRules.AmountFormat, CultureInfo.InvariantCulture)
                : Stringify(value);
        }

        /// <summary>
        /// Text clean_amount will leave as it is. Usually a header, sometimes
        /// an amount written in a way the rule does not read ("约 5 万") --
        /// worth naming, since the user asked for the column to be cleaned.
        /// </summary>
        private static void AmountsLeftAsText(Grid grid, ChangePreview preview)
        {
            var skipped = grid.Cells
                .Where(c => RawValue(c) is string s && s.Length > 0 && CleaningRules.CleanAmount(s) == null)
                .ToList();
            if (skipped.Count == 0) return;
            preview.Warnings.Add($"{skipped.Count} 个文本格不是可识别的金额，保持原样（{Examples(skipped)}）");
        }

        /// <summary>
        /// Text with its edges visible. Whitespace changes are otherwise
        /// invisible in the panel: "  张三 " → "张三" would read as no change.
        /// </summary>
        internal static string Quoted(object value)
        {
            return "\"" + Clip(value as string ?? Stringify(value)) + "\"";
        }

        /// <summary>
        /// Quoted, with each character remove_special_chars strips spelled out
        /// as [U+XXXX] -- they are invisible by definition.
        /// </summary>
        internal static string ShowSpecialChars(object value)
        {
            var text = value as string ?? Stringify(value);
            var shown = new System.Text.StringBuilder();
            foreach (var c in text)
            {
                if (CleaningRules.IsSpecialChar(c))
                {
                    shown.Append("[U+").Append(((int)c).ToString("X4")).Append(']');
                }
                else
                {
                    shown.Append(c);
                }
            }
            return Quoted(shown.ToString());
        }

        internal static string Clip(string text)
        {
            const int max = 60;
            if (text == null) return "";
            return text.Length > max ? text.Substring(0, max) + "…" : text;
        }
    }
}
