using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DeepExcel.AddIn.Preview;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// A worksheet held in memory, read the way <see cref="ExcelRangeReader"/>
    /// reads Excel: whole rectangular blocks in row-major order, raw values next
    /// to their display text.
    ///
    /// Like the real reader it does not understand "Sheet2!A1" -- a range on
    /// another sheet has to arrive as (sheet, local address), or nothing is read.
    /// </summary>
    internal sealed class GridSheet : IRangeReader
    {
        private readonly Dictionary<string, CellSnapshot> _cells =
            new Dictionary<string, CellSnapshot>(StringComparer.OrdinalIgnoreCase);

        public string ActiveSheet { get; set; } = "Sheet1";

        /// <summary>Caps cells returned per read, like ExcelRangeReader.MaxCellsRead.</summary>
        public int MaxCells { get; set; } = int.MaxValue;

        public List<string> Reads { get; } = new List<string>();

        public GridSheet Set(string address, object value, string formula = null)
        {
            var sheet = ActiveSheet;
            var bang = address.IndexOf('!');
            if (bang >= 0)
            {
                sheet = address.Substring(0, bang);
                address = address.Substring(bang + 1);
            }
            ParseCell(address, out var row, out var col);
            _cells[Key(sheet, row, col)] = new CellSnapshot
            {
                Address = address.ToUpperInvariant(),
                Value = value,
                Display = PreviewBuilder.Stringify(value),
                Formula = formula ?? (value == null ? "" : PreviewBuilder.Stringify(value))
            };
            return this;
        }

        /// <summary>Fills a block starting at a cell, one array per row.</summary>
        public GridSheet Rows(string topLeft, params object[][] rows)
        {
            ParseCell(topLeft, out var row, out var col);
            for (var r = 0; r < rows.Length; r++)
            {
                for (var c = 0; c < rows[r].Length; c++)
                {
                    Set(Letter(col + c) + (row + r), rows[r][c]);
                }
            }
            return this;
        }

        public IReadOnlyList<CellSnapshot> Read(string sheet, string address)
        {
            Reads.Add((sheet ?? "") + "|" + address);
            var cells = new List<CellSnapshot>();
            if (string.IsNullOrEmpty(address) || address.Contains("!"))
            {
                return cells;
            }
            var parts = address.Replace("$", "").Split(':');
            ParseCell(parts[0], out var r1, out var c1);
            ParseCell(parts.Length > 1 ? parts[1] : parts[0], out var r2, out var c2);
            var name = string.IsNullOrEmpty(sheet) ? ActiveSheet : sheet;
            for (var r = r1; r <= r2; r++)
            {
                for (var c = c1; c <= c2; c++)
                {
                    if (cells.Count >= MaxCells) return cells;
                    cells.Add(_cells.TryGetValue(Key(name, r, c), out var cell)
                        ? cell
                        : new CellSnapshot { Address = Letter(c) + r, Display = "", Formula = "" });
                }
            }
            return cells;
        }

        public int UsedRowCount(string sheet)
        {
            var name = string.IsNullOrEmpty(sheet) ? ActiveSheet : sheet;
            var prefix = name.ToUpperInvariant() + "|";
            return _cells.Where(kv => kv.Key.StartsWith(prefix, StringComparison.Ordinal))
                .Select(kv => int.Parse(kv.Key.Split('|')[1]))
                .DefaultIfEmpty(0).Max();
        }

        private static string Key(string sheet, int row, int col) =>
            sheet.ToUpperInvariant() + "|" + row + "|" + col;

        private static void ParseCell(string text, out int row, out int col)
        {
            var m = Regex.Match(text.Replace("$", ""), @"^([A-Za-z]+)(\d+)$");
            if (!m.Success) throw new ArgumentException("bad cell: " + text);
            col = 0;
            foreach (var ch in m.Groups[1].Value.ToUpperInvariant()) col = col * 26 + (ch - 'A' + 1);
            row = int.Parse(m.Groups[2].Value);
        }

        internal static string Letter(int col)
        {
            var name = "";
            while (col > 0)
            {
                var rem = (col - 1) % 26;
                name = (char)('A' + rem) + name;
                col = (col - 1) / 26;
            }
            return name;
        }
    }

    /// <summary>
    /// Previews for the write tools beyond plain writes, clears and deletes.
    ///
    /// Each tool's preview has to mirror what its executor actually does,
    /// quirks included: a preview that reports fewer changes than really happen
    /// is worse than none, because a small or empty change set is auto-allowed
    /// without asking the user at all.
    /// </summary>
    public class PreviewToolTests
    {
        private static Dictionary<string, object> Args(params object[] pairs)
        {
            var args = new Dictionary<string, object>();
            for (var i = 0; i + 1 < pairs.Length; i += 2)
            {
                args[(string)pairs[i]] = pairs[i + 1];
            }
            return args;
        }

        private static CellChange At(ChangePreview preview, string address) =>
            preview.Changes.SingleOrDefault(c => c.Address == address);

        // ------------------------------------------------------------------
        // Shared plumbing
        // ------------------------------------------------------------------

        [Fact]
        public void An_unreadable_range_is_not_previewable_rather_than_a_no_op()
        {
            // A no-op is auto-allowed without asking. Failing to read the target
            // must never be mistaken for "nothing will change".
            var reader = new EmptyReader();
            var preview = new PreviewBuilder(reader).Build("clear_range", Args("address", "A1:B5"));

            Assert.False(preview.Previewable);
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void A_range_on_another_sheet_is_read_from_that_sheet()
        {
            var sheet = new GridSheet().Set("Data!A1", "x").Set("Data!A2", "y");
            var preview = new PreviewBuilder(sheet).Build("clear_range", Args("address", "Data!A1:A2"));

            Assert.True(preview.Previewable);
            Assert.Equal(2, preview.AffectedCells);
            Assert.Contains("Data|A1:A2", sheet.Reads);
        }

        [Fact]
        public void A_quoted_sheet_name_is_unquoted_before_reading()
        {
            var sheet = new GridSheet().Set("It's mine!B2", "x");
            var preview = new PreviewBuilder(sheet).Build("clear_range", Args("address", "'It''s mine'!B2"));

            Assert.Equal(1, preview.AffectedCells);
            Assert.Contains("It's mine|B2", sheet.Reads);
        }

        private sealed class EmptyReader : IRangeReader
        {
            public IReadOnlyList<CellSnapshot> Read(string sheet, string address) => new List<CellSnapshot>();
            public int UsedRowCount(string sheet) => 0;
        }

        // ------------------------------------------------------------------
        // replace_formula
        // ------------------------------------------------------------------

        [Fact]
        public void Replace_formula_lists_each_formula_that_contains_the_text()
        {
            var sheet = new GridSheet()
                .Set("C1", 10.0, "=SUM(A1:B1)")
                .Set("C2", 20.0, "=SUM(A2:B2)")
                .Set("C3", 5.0, "=A3*2");
            var preview = new PreviewBuilder(sheet).Build("replace_formula",
                Args("range_address", "C1:C3", "find", "SUM", "replace", "AVERAGE"));

            Assert.True(preview.Previewable);
            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("=SUM(A1:B1)", At(preview, "C1").Before);
            Assert.Equal("=AVERAGE(A1:B1)", At(preview, "C1").After);
            Assert.Null(At(preview, "C3"));
        }

        [Fact]
        public void Replace_formula_leaves_plain_values_alone_even_when_they_match()
        {
            // The executor only touches cells with HasFormula.
            var sheet = new GridSheet().Set("A1", "SUM of sales").Set("A2", 3.0, "=SUM(B1:B2)");
            var preview = new PreviewBuilder(sheet).Build("replace_formula",
                Args("range_address", "A1:A2", "find", "SUM", "replace", "MAX"));

            Assert.Equal(1, preview.AffectedCells);
            Assert.NotNull(At(preview, "A2"));
        }

        [Fact]
        public void Replace_formula_is_case_sensitive_like_the_executor()
        {
            var sheet = new GridSheet().Set("A1", 1.0, "=sum(B1:B2)");
            var preview = new PreviewBuilder(sheet).Build("replace_formula",
                Args("range_address", "A1", "find", "SUM", "replace", "MAX"));

            Assert.True(preview.IsNoOp);
        }

        [Fact]
        public void Rewriting_a_formula_still_interrupts_but_is_not_called_a_formula_loss()
        {
            // The cell keeps calculating, so "will no longer recalculate" would
            // be false. It is still a formula edit, which always interrupts.
            var sheet = new GridSheet().Set("A1", 1.0, "=B1+C1");
            var preview = new PreviewBuilder(sheet).Build("replace_formula",
                Args("range_address", "A1", "find", "C1", "replace", "D1"));

            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
            Assert.DoesNotContain(preview.Warnings, w => w.Contains("不再自动计算"));
            Assert.Contains(preview.Warnings, w => w.Contains("1 个公式会被改写"));
        }

        [Fact]
        public void Replace_formula_without_find_text_is_not_previewable()
        {
            var sheet = new GridSheet().Set("A1", 1.0, "=B1");
            var preview = new PreviewBuilder(sheet).Build("replace_formula",
                Args("range_address", "A1", "find", "", "replace", "x"));

            Assert.False(preview.Previewable);
        }

        [Fact]
        public void A_range_larger_than_the_read_cap_is_not_previewed_cell_by_cell()
        {
            // Changes past the cap would go uncounted, and an undercount can
            // turn into an auto-allow.
            var sheet = new GridSheet { MaxCells = 4 };
            for (var r = 1; r <= 6; r++) sheet.Set("A" + r, 1.0, "=B" + r);
            var preview = new PreviewBuilder(sheet).Build("replace_formula",
                Args("range_address", "A1:A6", "find", "B", "replace", "C"));

            Assert.False(preview.Previewable);
        }

        // ------------------------------------------------------------------
        // trim_spaces
        // ------------------------------------------------------------------

        [Fact]
        public void Trim_spaces_lists_padded_text_with_the_padding_visible()
        {
            var sheet = new GridSheet().Set("A1", "  张三 ").Set("A2", "李四").Set("A3", 12.0);
            var preview = new PreviewBuilder(sheet).Build("trim_spaces", Args("range_address", "A1:A3"));

            Assert.True(preview.Previewable);
            Assert.Equal(1, preview.AffectedCells);
            // Spaces are invisible in the panel, so both sides are quoted.
            Assert.Equal("\"  张三 \"", At(preview, "A1").Before);
            Assert.Equal("\"张三\"", At(preview, "A1").After);
        }

        [Fact]
        public void Trim_spaces_flags_a_formula_whose_result_is_padded()
        {
            // The executor writes the trimmed result over the cell, so the
            // formula is lost.
            var sheet = new GridSheet().Set("B1", " x ", "=\" x \"");
            var preview = new PreviewBuilder(sheet).Build("trim_spaces", Args("range_address", "B1"));

            Assert.Equal(1, preview.FormulasOverwritten);
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void Trim_spaces_on_clean_text_is_a_no_op()
        {
            var sheet = new GridSheet().Set("A1", "ok").Set("A2", null);
            var preview = new PreviewBuilder(sheet).Build("trim_spaces", Args("range_address", "A1:A2"));

            Assert.True(preview.IsNoOp);
        }

        // ------------------------------------------------------------------
        // collapse_spaces
        // ------------------------------------------------------------------

        [Fact]
        public void Collapse_spaces_also_trims_the_edges_like_the_executor()
        {
            var sheet = new GridSheet().Set("A1", " 北京   朝阳 ").Set("A2", "上海 浦东").Set("A3", "");
            var preview = new PreviewBuilder(sheet).Build("collapse_spaces", Args("range_address", "A1:A3"));

            Assert.Equal(1, preview.AffectedCells);
            Assert.Equal("\"北京 朝阳\"", At(preview, "A1").After);
        }

        [Fact]
        public void Collapse_spaces_leaves_numbers_alone()
        {
            var sheet = new GridSheet().Set("A1", 3.0);
            Assert.True(new PreviewBuilder(sheet).Build("collapse_spaces", Args("range_address", "A1")).IsNoOp);
        }

        // ------------------------------------------------------------------
        // remove_special_chars
        // ------------------------------------------------------------------

        [Fact]
        public void Remove_special_chars_names_the_invisible_characters_it_removes()
        {
            var sheet = new GridSheet()
                .Set("A1", "单价\u00A0100")
                .Set("A2", "备注\u0007")
                .Set("A3", "换行\n保留");
            var preview = new PreviewBuilder(sheet).Build("remove_special_chars", Args("range_address", "A1:A3"));

            // Tabs and line breaks are kept by the executor.
            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("\"单价[U+00A0]100\"", At(preview, "A1").Before);
            Assert.Equal("\"单价100\"", At(preview, "A1").After);
            Assert.Equal("\"备注[U+0007]\"", At(preview, "A2").Before);
        }

        // ------------------------------------------------------------------
        // clean_amount
        // ------------------------------------------------------------------

        [Fact]
        public void Clean_amount_shows_the_number_as_it_will_be_displayed()
        {
            var sheet = new GridSheet()
                .Set("B1", "金额")
                .Set("B2", "¥1,234.5")
                .Set("B3", "2,000元")
                .Set("B4", 88.0);
            var preview = new PreviewBuilder(sheet).Build("clean_amount", Args("range_address", "B1:B4"));

            // The executor applies #,##0.00 to what it converts.
            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("¥1,234.5", At(preview, "B2").Before);
            Assert.Equal("1,234.50", At(preview, "B2").After);
            Assert.Equal("2,000.00", At(preview, "B3").After);
            Assert.Null(At(preview, "B4"));
        }

        [Fact]
        public void Clean_amount_says_which_text_it_could_not_read_as_money()
        {
            var sheet = new GridSheet().Set("B1", "金额").Set("B2", "¥5").Set("B3", "待定");
            var preview = new PreviewBuilder(sheet).Build("clean_amount", Args("range_address", "B1:B3"));

            Assert.Contains(preview.Warnings, w => w.Contains("2 个文本格") && w.Contains("B1"));
        }

        // ------------------------------------------------------------------
        // text_to_number
        // ------------------------------------------------------------------

        [Fact]
        public void Text_to_number_converts_numeric_text_only()
        {
            var sheet = new GridSheet().Set("A1", "42").Set("A2", 7.0).Set("A3", "abc").Set("A4", "3.5");
            var preview = new PreviewBuilder(sheet).Build("text_to_number", Args("range_address", "A1:A4"));

            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("42", At(preview, "A1").After);
            Assert.Equal("3.5", At(preview, "A4").After);
            Assert.Null(At(preview, "A2"));
        }

        [Fact]
        public void Text_to_number_warns_that_leading_zeros_will_be_lost()
        {
            // Employee numbers and postcodes stored as text are the usual
            // victims: 001234 becomes 1234.
            var sheet = new GridSheet().Set("A1", "001234").Set("A2", "0");
            var preview = new PreviewBuilder(sheet).Build("text_to_number", Args("range_address", "A1:A2"));

            Assert.Equal("1234", At(preview, "A1").After);
            Assert.Contains(preview.Warnings, w => w.Contains("前导零") && w.Contains("A1") && !w.Contains("A2"));
        }

        [Fact]
        public void Text_to_number_warns_that_long_digit_strings_lose_precision()
        {
            // An ID card number has 18 digits; Excel keeps 15.
            var sheet = new GridSheet().Set("A1", "110101199003071234");
            var preview = new PreviewBuilder(sheet).Build("text_to_number", Args("range_address", "A1"));

            Assert.Contains(preview.Warnings, w => w.Contains("15 位"));
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        // ------------------------------------------------------------------
        // fill_blank_cells
        // ------------------------------------------------------------------

        [Fact]
        public void Fill_blank_cells_copies_the_nearest_value_above_in_each_column()
        {
            var sheet = new GridSheet()
                .Rows("A1",
                    new object[] { "华东", 1.0 },
                    new object[] { null, null },
                    new object[] { "  ", 3.0 },
                    new object[] { "华南", null });
            var preview = new PreviewBuilder(sheet).Build("fill_blank_cells", Args("range_address", "A1:B4"));

            Assert.Equal(4, preview.AffectedCells);
            Assert.Equal("华东", At(preview, "A2").After);
            Assert.Equal(ChangeKind.Add, At(preview, "A2").Kind);
            // Whitespace counts as blank and is overwritten.
            Assert.Equal("华东", At(preview, "A3").After);
            Assert.Equal(ChangeKind.Overwrite, At(preview, "A3").Kind);
            Assert.Equal("1", At(preview, "B2").After);
            Assert.Equal("3", At(preview, "B4").After);
        }

        [Fact]
        public void Fill_blank_cells_leaves_blanks_with_nothing_above_them()
        {
            var sheet = new GridSheet().Rows("A1", new object[] { null }, new object[] { "x" });
            var preview = new PreviewBuilder(sheet).Build("fill_blank_cells", Args("range_address", "A1:A2"));

            Assert.True(preview.IsNoOp);
        }

        [Fact]
        public void Fill_blank_cells_overwrites_a_formula_that_shows_nothing()
        {
            var sheet = new GridSheet().Set("A1", "x").Set("A2", "", "=IF(B2>0,B2,\"\")");
            var preview = new PreviewBuilder(sheet).Build("fill_blank_cells", Args("range_address", "A1:A2"));

            Assert.Equal(1, preview.FormulasOverwritten);
            Assert.Equal("=IF(B2>0,B2,\"\")", At(preview, "A2").Before);
        }
    }
}
