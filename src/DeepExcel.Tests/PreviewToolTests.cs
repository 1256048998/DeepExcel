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

        /// <summary>A date-formatted cell: Value2 is the serial number.</summary>
        public GridSheet Date(string address, DateTime date)
        {
            Set(address, date.ToOADate());
            var cell = _cells.Values.Single(c => c.Address == address.ToUpperInvariant());
            cell.IsDate = true;
            cell.Display = PreviewBuilder.FormatDate(date);
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

        public int UsedColumnCount(string sheet)
        {
            var name = string.IsNullOrEmpty(sheet) ? ActiveSheet : sheet;
            var prefix = name.ToUpperInvariant() + "|";
            return _cells.Where(kv => kv.Key.StartsWith(prefix, StringComparison.Ordinal))
                .Select(kv => int.Parse(kv.Key.Split('|')[2]))
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
            public int UsedColumnCount(string sheet) => 0;
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

        // ------------------------------------------------------------------
        // rename_columns
        // ------------------------------------------------------------------

        [Fact]
        public void Rename_columns_shows_each_header_that_changes()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "name", "amt", "date" },
                new object[] { "张三", 5.0, "2026-01-01" });
            // Arrays arrive from the sidecar as raw JSON text.
            var preview = new PreviewBuilder(sheet).Build("rename_columns",
                Args("range_address", "A1:C2", "new_names", "[\"姓名\", \"amt\", \"日期\"]"));

            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("name", At(preview, "A1").Before);
            Assert.Equal("姓名", At(preview, "A1").After);
            Assert.Null(At(preview, "B1"));
            Assert.Equal("日期", At(preview, "C1").After);
        }

        [Fact]
        public void Rename_columns_ignores_names_beyond_the_range()
        {
            var sheet = new GridSheet().Rows("A1", new object[] { "a", "b" });
            var preview = new PreviewBuilder(sheet).Build("rename_columns",
                Args("range_address", "A1:B1", "new_names", new List<object> { "x", "y", "z" }));

            Assert.Equal(2, preview.AffectedCells);
            Assert.Null(At(preview, "C1"));
        }

        [Fact]
        public void Rename_columns_with_a_non_text_name_renames_nothing_like_the_executor()
        {
            // The executor reads new_names with JsonElement.GetString(), which
            // throws on a number; the whole list is then dropped.
            var sheet = new GridSheet().Rows("A1", new object[] { "a", "b" });
            var preview = new PreviewBuilder(sheet).Build("rename_columns",
                Args("range_address", "A1:B1", "new_names", "[\"x\", 2]"));

            Assert.True(preview.IsNoOp);
        }

        // ------------------------------------------------------------------
        // merge_columns
        // ------------------------------------------------------------------

        [Fact]
        public void Merge_columns_writes_into_the_first_column_by_default()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "北京", "朝阳", "望京" },
                new object[] { "上海", null, "浦东" });
            var preview = new PreviewBuilder(sheet).Build("merge_columns",
                Args("range_address", "A1:C2", "delimiter", "-"));

            // Blank parts are skipped; the other columns are left as they are.
            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("北京-朝阳-望京", At(preview, "A1").After);
            Assert.Equal("上海-浦东", At(preview, "A2").After);
        }

        [Fact]
        public void Merge_columns_into_a_column_outside_the_range_reads_that_column()
        {
            var sheet = new GridSheet()
                .Rows("A1", new object[] { "a", "b" })
                .Set("E1", "old", "=X1");
            var preview = new PreviewBuilder(sheet).Build("merge_columns",
                Args("range_address", "A1:B1", "target_column", "E"));

            // Default delimiter is a space.
            Assert.Equal("a b", At(preview, "E1").After);
            Assert.Equal(1, preview.FormulasOverwritten);
            Assert.Null(At(preview, "A1"));
        }

        // ------------------------------------------------------------------
        // split_text_to_columns
        // ------------------------------------------------------------------

        [Fact]
        public void Split_text_writes_parts_rightwards_from_the_first_column()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "张三,销售,北京" },
                new object[] { "李四,财务" });
            var preview = new PreviewBuilder(sheet).Build("split_text_to_columns",
                Args("range_address", "A1:A2", "delimiter", ","));

            Assert.Equal("张三", At(preview, "A1").After);
            Assert.Equal("销售", At(preview, "B1").After);
            Assert.Equal("北京", At(preview, "C1").After);
            Assert.Equal("财务", At(preview, "B2").After);
            // C2 is written "" but was already empty: not a change.
            Assert.Null(At(preview, "C2"));
            Assert.Equal(5, preview.AffectedCells);
        }

        [Fact]
        public void Split_text_calls_out_data_it_overwrites_to_the_right()
        {
            // The parts land in whatever columns follow, range or not.
            var sheet = new GridSheet().Rows("A1",
                new object[] { "a;b;c", "keep?", "and me?" },
                new object[] { "d", "x", null });
            var preview = new PreviewBuilder(sheet).Build("split_text_to_columns",
                Args("range_address", "A1:A2", "delimiter", ";"));

            Assert.Equal("b", At(preview, "B1").After);
            Assert.Equal("keep?", At(preview, "B1").Before);
            // Shorter rows are padded with "": B2 is cleared.
            Assert.Equal(ChangeKind.Clear, At(preview, "B2").Kind);
            Assert.Contains(preview.Warnings, w => w.Contains("B 列") && w.Contains("C 列") && w.Contains("3 个"));
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void Split_text_with_nothing_to_split_is_a_no_op()
        {
            var sheet = new GridSheet().Rows("A1", new object[] { "abc" }, new object[] { "def" });
            var preview = new PreviewBuilder(sheet).Build("split_text_to_columns",
                Args("range_address", "A1:A2", "delimiter", ","));

            Assert.True(preview.IsNoOp);
        }

        // ------------------------------------------------------------------
        // delete_blank_rows
        // ------------------------------------------------------------------

        [Fact]
        public void Delete_blank_rows_lists_the_rows_that_are_blank_across_the_range()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "姓名", "金额" },
                new object[] { null, null },
                new object[] { "张三", null },
                new object[] { " ", "" },
                new object[] { "李四", 5.0 });
            var preview = new PreviewBuilder(sheet).Build("delete_blank_rows", Args("range_address", "A1:B5"));

            Assert.Equal(new[] { 2, 4 }, preview.DeletedRows.ToArray());
            Assert.True(preview.IsStructural);
            Assert.Contains("删除 2 行", preview.Summary());
            Assert.Contains(preview.Warnings, w => w.Contains("上移"));
        }

        [Fact]
        public void Delete_blank_rows_warns_when_a_row_has_data_outside_the_range()
        {
            // The executor deletes the entire row, not just the range's part.
            var sheet = new GridSheet().Rows("A1",
                new object[] { "x", "y", null, null },
                new object[] { null, null, null, "备注" });
            var preview = new PreviewBuilder(sheet).Build("delete_blank_rows", Args("range_address", "A1:B2"));

            Assert.Equal(new[] { 2 }, preview.DeletedRows.ToArray());
            Assert.Contains(preview.Warnings, w => w.Contains("第 2 行") && w.Contains("区域外"));
        }

        [Fact]
        public void Delete_blank_rows_with_no_blank_rows_is_a_no_op()
        {
            var sheet = new GridSheet().Rows("A1", new object[] { "a" }, new object[] { "b" });
            Assert.True(new PreviewBuilder(sheet).Build("delete_blank_rows", Args("range_address", "A1:A2")).IsNoOp);
        }

        // ------------------------------------------------------------------
        // remove_duplicates
        // ------------------------------------------------------------------

        [Fact]
        public void Remove_duplicates_keeps_the_last_copy_like_the_executor()
        {
            // The executor walks bottom-up and deletes rows it has already
            // seen, so the lowest copy survives -- the opposite of Excel's own
            // Remove Duplicates, which is worth saying out loud.
            var sheet = new GridSheet().Rows("A1",
                new object[] { "姓名", "金额" },
                new object[] { "张三", 5.0 },
                new object[] { "李四", 6.0 },
                new object[] { "张三", 5.0 },
                new object[] { "张三", 7.0 });
            var preview = new PreviewBuilder(sheet).Build("remove_duplicates", Args("range_address", "A1:B5"));

            Assert.Equal(new[] { 2 }, preview.DeletedRows.ToArray());
            Assert.Contains(preview.Warnings, w => w.Contains("最下面"));
        }

        [Fact]
        public void Remove_duplicates_treats_repeated_blank_rows_as_duplicates()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "a" }, new object[] { null }, new object[] { null }, new object[] { null });
            var preview = new PreviewBuilder(sheet).Build("remove_duplicates", Args("range_address", "A1:A4"));

            Assert.Equal(new[] { 2, 3 }, preview.DeletedRows.ToArray());
        }

        // ------------------------------------------------------------------
        // unify_date
        // ------------------------------------------------------------------

        [Fact]
        public void Unify_date_converts_date_text_to_real_dates()
        {
            var sheet = new GridSheet().Set("A1", "2026/3/5").Set("A2", "备注");
            var preview = new PreviewBuilder(sheet).Build("unify_date", Args("range_address", "A1:A2"));

            Assert.Equal(1, preview.AffectedCells);
            Assert.Equal("2026/3/5", At(preview, "A1").Before);
            Assert.Equal("2026-03-05", At(preview, "A1").After);
            Assert.Equal(ChangeKind.Overwrite, At(preview, "A1").Kind);
        }

        [Fact]
        public void Unify_date_reformats_plain_numbers_as_dates_and_says_when_they_are_not_dates()
        {
            // The executor gives every number in the range a date format. A
            // serial number becomes its date; an amount becomes nonsense.
            var sheet = new GridSheet()
                .Set("A1", new DateTime(2026, 3, 5).ToOADate())
                .Set("A2", 1500.0);
            var preview = new PreviewBuilder(sheet).Build("unify_date", Args("range_address", "A1:A2"));

            Assert.Equal("2026-03-05", At(preview, "A1").After);
            Assert.Equal(ChangeKind.Format, At(preview, "A1").Kind);
            Assert.Equal("1904-02-08", At(preview, "A2").After);
            Assert.Contains(preview.Warnings, w => w.Contains("不像日期") && w.Contains("A2") && !w.Contains("A1"));
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void Unify_date_leaves_cells_that_are_already_dates()
        {
            var sheet = new GridSheet().Date("A1", new DateTime(2026, 3, 5));
            Assert.True(new PreviewBuilder(sheet).Build("unify_date", Args("range_address", "A1")).IsNoOp);
        }

        [Fact]
        public void Unify_date_keeps_a_formula_when_it_only_reformats_it()
        {
            var sheet = new GridSheet().Set("A1", 46000.0, "=TODAY()");
            var preview = new PreviewBuilder(sheet).Build("unify_date", Args("range_address", "A1"));

            Assert.Equal(0, preview.FormulasOverwritten);
            Assert.Equal(ChangeKind.Format, At(preview, "A1").Kind);
        }

        [Fact]
        public void Serial_numbers_before_march_1900_follow_excels_calendar()
        {
            // Excel counts a 29 February 1900 that never existed.
            Assert.Equal("1900-01-01", PreviewBuilder.SerialToDateText(1));
            Assert.Equal("1900-02-28", PreviewBuilder.SerialToDateText(59));
            Assert.Equal("1900-03-01", PreviewBuilder.SerialToDateText(61));
        }

        // ------------------------------------------------------------------
        // clean_data
        // ------------------------------------------------------------------

        [Fact]
        public void Clean_data_runs_operations_in_order_on_the_result_of_the_last()
        {
            // Trimming first turns two rows into duplicates; the second
            // operation then deletes one. Previewing each operation against
            // the original sheet would miss the deletion.
            var sheet = new GridSheet().Rows("A1",
                new object[] { "姓名" },
                new object[] { "张三 " },
                new object[] { "张三" });
            var preview = new PreviewBuilder(sheet).Build("clean_data",
                Args("range_address", "A1:A3", "operations", "[\"trim_spaces\", \"remove_duplicates\"]"));

            Assert.True(preview.Previewable);
            Assert.Equal(new[] { 2 }, preview.DeletedRows.ToArray());
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void Clean_data_combines_value_changes_from_several_operations()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { " ¥1,200 " },
                new object[] { "  北京   朝阳" });
            var preview = new PreviewBuilder(sheet).Build("clean_data",
                Args("range_address", "A1:A2", "operations", new List<object> { "collapse_spaces", "clean_amount" }));

            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal("1,200.00", At(preview, "A1").After);
            // A whitespace-only change is quoted, as in collapse_spaces itself.
            Assert.Equal("\"北京 朝阳\"", At(preview, "A2").After);
        }

        [Fact]
        public void Clean_data_follows_rows_that_shift_up_into_the_range()
        {
            // delete_blank_rows pulls row 4 (outside the range) up into it;
            // the following text_to_number then converts it. Addresses are
            // given as they are before anything runs.
            var sheet = new GridSheet().Rows("A1",
                new object[] { "1" },
                new object[] { null },
                new object[] { "2" },
                new object[] { "3" });
            var preview = new PreviewBuilder(sheet).Build("clean_data",
                Args("range_address", "A1:A3", "operations", "[\"delete_blank_rows\", \"text_to_number\"]"));

            Assert.Equal(new[] { 2 }, preview.DeletedRows.ToArray());
            Assert.Equal("3", At(preview, "A4").After);
            Assert.NotNull(At(preview, "A1"));
            Assert.Contains(preview.Warnings, w => w.Contains("执行前的位置"));
        }

        [Fact]
        public void Clean_data_reports_highlighting_as_a_format_change()
        {
            var sheet = new GridSheet().Rows("A1", new object[] { "a" }, new object[] { null });
            var preview = new PreviewBuilder(sheet).Build("clean_data",
                Args("range_address", "A1:A2", "operations", "[\"highlight_missing\", \"no_such_op\"]"));

            Assert.Equal(1, preview.AffectedCells);
            Assert.Equal(ChangeKind.Format, At(preview, "A2").Kind);
            Assert.Contains("黄色", At(preview, "A2").After);
        }

        [Fact]
        public void Clean_data_without_operations_changes_nothing()
        {
            var sheet = new GridSheet().Set("A1", " x ");
            var preview = new PreviewBuilder(sheet).Build("clean_data",
                Args("range_address", "A1", "operations", "[]"));

            Assert.True(preview.IsNoOp);
        }

        [Fact]
        public void Clean_data_carries_the_warnings_of_the_operations_it_runs()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "001", "x" },
                new object[] { "001", "x" },
                new object[] { 1500.0, null });
            var preview = new PreviewBuilder(sheet).Build("clean_data",
                Args("range_address", "A1:B3", "operations", "[\"remove_duplicates\", \"text_to_number\", \"unify_date\"]"));

            Assert.Contains(preview.Warnings, w => w.Contains("最下面"));
            Assert.Contains(preview.Warnings, w => w.Contains("前导零"));
            Assert.Contains(preview.Warnings, w => w.Contains("不像日期"));
            Assert.True(preview.RequiresConfirmation);
        }

        // ------------------------------------------------------------------
        // merge_cells
        // ------------------------------------------------------------------

        [Fact]
        public void Merge_cells_lists_everything_but_the_top_left_as_lost()
        {
            var sheet = new GridSheet().Rows("A1",
                new object[] { "标题", "副标题" },
                new object[] { null, 3.0 });
            var preview = new PreviewBuilder(sheet).Build("merge_cells", Args("address", "A1:B2"));

            Assert.Equal(2, preview.AffectedCells);
            Assert.Equal(ChangeKind.Clear, At(preview, "B1").Kind);
            Assert.Equal("副标题", At(preview, "B1").Before);
            Assert.Null(At(preview, "A1"));
            Assert.Contains(preview.Warnings, w => w.Contains("A1"));
        }

        [Fact]
        public void Merging_cells_that_only_hold_a_top_left_value_loses_nothing()
        {
            var sheet = new GridSheet().Set("A1", "标题");
            var preview = new PreviewBuilder(sheet).Build("merge_cells", Args("address", "A1:D1"));

            Assert.True(preview.IsNoOp);
        }
    }
}
