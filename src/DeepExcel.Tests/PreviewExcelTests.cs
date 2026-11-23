using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Preview;
using DeepExcel.AddIn.Tools;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// A hidden Excel instance owned by one test. Disposing closes its
    /// workbooks without saving, quits, and kills the process if a stray add-in
    /// keeps it alive -- only this instance's process, never the user's Excel.
    /// </summary>
    internal sealed class HiddenExcel : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private readonly int _pid;

        public HiddenExcel()
        {
            App = new Excel.Application { Visible = false, DisplayAlerts = false };
            try { GetWindowThreadProcessId(new IntPtr(App.Hwnd), out _pid); } catch { }
        }

        public Excel.Application App { get; private set; }

        /// <summary>A new workbook whose sheets are named as given; the first is active.</summary>
        public Workbook NewBook(params string[] sheetNames)
        {
            if (sheetNames.Length == 0) sheetNames = new[] { "Sheet1" };
            var wb = App.Workbooks.Add();
            while (wb.Worksheets.Count < sheetNames.Length) wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            for (var i = 0; i < sheetNames.Length; i++) ((Worksheet)wb.Worksheets[i + 1]).Name = sheetNames[i];
            ((Worksheet)wb.Worksheets[1]).Activate();
            return wb;
        }

        public void Dispose()
        {
            if (App != null)
            {
                try
                {
                    foreach (Workbook wb in App.Workbooks.Cast<Workbook>().ToList())
                    {
                        try { wb.Close(false); } catch { }
                    }
                    App.Quit();
                }
                catch { }
                try { Marshal.FinalReleaseComObject(App); } catch { }
                App = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (_pid != 0)
            {
                try
                {
                    var process = System.Diagnostics.Process.GetProcessById(_pid);
                    if (!process.WaitForExit(5000)) process.Kill();
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// The change preview against real Excel: what ExcelRangeReader actually
    /// gets back from COM, and whether a preview matches what the executor
    /// then really does to the same workbook. The unit tests prove the diff
    /// logic against a fake sheet; only these prove the fake is faithful.
    /// </summary>
    public class PreviewExcelTests : IDisposable
    {
        private readonly HiddenExcel _excel = new HiddenExcel();

        public void Dispose() => _excel.Dispose();

        private static Dictionary<string, object> Args(params object[] pairs)
        {
            var args = new Dictionary<string, object>();
            for (var i = 0; i + 1 < pairs.Length; i += 2) args[(string)pairs[i]] = pairs[i + 1];
            return args;
        }

        // ------------------------------------------------------------------
        // ExcelRangeReader
        // ------------------------------------------------------------------

        [ExcelFact]
        public void Reader_returns_the_block_row_major_with_raw_values_and_formulas()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["B2"].Value2 = "001";
            ws.Range["B2"].NumberFormat = "@";
            ws.Range["B2"].Value2 = "001";
            ws.Range["C2"].Value2 = 1.5;
            ws.Range["B3"].Formula = "=C2*2";
            ws.Range["C3"].Value2 = true;

            var cells = new ExcelRangeReader(_excel.App).Read(null, "B2:C3");

            Assert.Equal(new[] { "B2", "C2", "B3", "C3" }, cells.Select(c => c.Address).ToArray());
            // Text stays text, so text_to_number and trim_spaces can tell.
            Assert.IsType<string>(cells[0].Value);
            Assert.Equal("001", cells[0].Value);
            Assert.IsType<double>(cells[1].Value);
            Assert.Equal("=C2*2", cells[2].Formula);
            Assert.True(cells[2].HasFormula);
            Assert.Equal(3.0, cells[2].Value);
            Assert.Equal("TRUE", cells[3].Display);
            Assert.False(cells[1].HasFormula);
        }

        [ExcelFact]
        public void Reader_recognises_date_cells_and_shows_them_as_dates()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["A1"].Value2 = new DateTime(2026, 3, 5).ToOADate();
            ws.Range["A1"].NumberFormat = "yyyy/m/d";
            ws.Range["A2"].Value2 = 46086.0;

            var cells = new ExcelRangeReader(_excel.App).Read(null, "A1:A2");

            Assert.True(cells[0].IsDate);
            Assert.Equal("2026-03-05", cells[0].Display);
            Assert.IsType<double>(cells[0].Value);
            // The same number without a date format is just a number.
            Assert.False(cells[1].IsDate);
            Assert.Equal("46086", cells[1].Display);
        }

        [ExcelFact]
        public void Reader_truncates_the_display_but_keeps_the_whole_value()
        {
            var wb = _excel.NewBook("Data");
            var text = new string('长', 60);
            ((Worksheet)wb.Worksheets["Data"]).Range["A1"].Value2 = text;

            var cell = new ExcelRangeReader(_excel.App).Read(null, "A1").Single();

            Assert.Equal(text, cell.Value);
            Assert.EndsWith("…", cell.Display);
            Assert.Equal(41, cell.Display.Length);
        }

        [ExcelFact]
        public void Reader_reads_a_named_sheet_while_another_is_active()
        {
            var wb = _excel.NewBook("Front", "Data");
            ((Worksheet)wb.Worksheets["Front"]).Range["A1"].Value2 = "front";
            ((Worksheet)wb.Worksheets["Data"]).Range["A1"].Value2 = "data";
            ((Worksheet)wb.Worksheets["Front"]).Activate();

            var reader = new ExcelRangeReader(_excel.App);

            Assert.Equal("data", reader.Read("Data", "A1").Single().Value);
            Assert.Equal("front", reader.Read(null, "A1").Single().Value);
            // A sheet that does not exist reads nothing rather than the active one.
            Assert.Empty(reader.Read("Nope", "A1"));
        }

        [ExcelFact]
        public void Reader_reports_the_used_extent_and_which_sheets_exist()
        {
            var wb = _excel.NewBook("Data", "Empty");
            ((Worksheet)wb.Worksheets["Data"]).Range["D7"].Value2 = 1;
            ((Worksheet)wb.Worksheets["Data"]).Range["B2"].Value2 = 1;

            var reader = new ExcelRangeReader(_excel.App);

            Assert.Equal(7, reader.UsedRowCount("Data"));
            Assert.Equal(4, reader.UsedColumnCount("Data"));
            Assert.True(reader.SheetExists("Data"));
            Assert.True(reader.SheetExists("Empty"));
            Assert.False(reader.SheetExists("Nope"));
            // An empty sheet's used range is A1.
            Assert.Equal(1, reader.UsedRowCount("Empty"));
            Assert.True(reader.Read("Empty", "A1").Single().IsEmpty);
        }

        [ExcelFact]
        public void Reader_stops_at_the_cell_cap_so_the_preview_can_refuse()
        {
            var wb = _excel.NewBook("Data");
            ((Worksheet)wb.Worksheets["Data"]).Range["A1:B4000"].Value2 = "x";

            var cells = new ExcelRangeReader(_excel.App).Read(null, "A1:B4000");
            Assert.True(cells.Count <= ExcelRangeReader.MaxCellsRead);

            var preview = new PreviewBuilder(new ExcelRangeReader(_excel.App))
                .Build("trim_spaces", Args("range_address", "A1:B4000"));
            Assert.False(preview.Previewable);
        }

        // ------------------------------------------------------------------
        // Preview against what the executor then does
        // ------------------------------------------------------------------

        /// <summary>
        /// Previews a tool, runs it, and checks every listed change landed as
        /// described and no cell outside the list changed.
        /// </summary>
        private void AssertPreviewMatchesExecution(Worksheet ws, string block, ChangePreview preview, System.Action execute)
        {
            Assert.True(preview.Previewable, preview.NotPreviewableReason);
            Assert.False(preview.IsTruncated);
            var reader = new ExcelRangeReader(_excel.App);
            var before = reader.Read(ws.Name, block).ToDictionary(c => c.Address, PreviewBuilder.Describe);

            execute();

            var after = reader.Read(ws.Name, block).ToDictionary(c => c.Address, PreviewBuilder.Describe);
            var listed = preview.Changes.ToDictionary(c => c.Address);
            var picture = "\n预览：" + string.Join("，", preview.Changes.Select(c => c.Address + "=" + c.After)) +
                          "\n实际：" + string.Join("，", after.Select(kv => kv.Key + "=" + kv.Value));
            foreach (var address in before.Keys)
            {
                if (listed.TryGetValue(address, out var change))
                {
                    Assert.True(change.Kind == ChangeKind.Format || before[address] != after[address],
                        $"{address} was listed as changing but did not" + picture);
                }
                else
                {
                    Assert.True(before[address] == after[address],
                        $"{address} changed ({before[address]} → {after[address]}) but the preview did not list it" + picture);
                }
            }
        }

        [ExcelFact]
        public void Trim_and_clean_amount_previews_match_what_the_cleaner_writes()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["A1"].Value2 = "  张三 ";
            ws.Range["A2"].Value2 = "¥1,234.5";
            ws.Range["A3"].Value2 = 88;
            ws.Range["A4"].Value2 = "备注";

            var builder = new PreviewBuilder(new ExcelRangeReader(_excel.App));
            var cleaner = new DataCleaner(_excel.App);

            var trim = builder.Build("trim_spaces", Args("range_address", "A1:A4"));
            AssertPreviewMatchesExecution(ws, "A1:A4", trim, () => cleaner.TrimSpaces("A1:A4"));
            Assert.Equal("张三", ws.Range["A1"].Value2);

            var amount = builder.Build("clean_amount", Args("range_address", "A1:A4"));
            Assert.Equal("1,234.50", amount.Changes.Single(c => c.Address == "A2").After);
            AssertPreviewMatchesExecution(ws, "A1:A4", amount, () => cleaner.CleanAmount("A1:A4"));
            // The panel showed what Excel now displays.
            Assert.Equal("1,234.50", ws.Range["A2"].Text);
        }

        [ExcelFact]
        public void Delete_blank_rows_preview_names_the_rows_excel_removes()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["A1"].Value2 = "a";
            ws.Range["A3"].Value2 = "b";
            ws.Range["A5"].Value2 = "c";

            var preview = new PreviewBuilder(new ExcelRangeReader(_excel.App))
                .Build("delete_blank_rows", Args("range_address", "A1:A5"));
            Assert.Equal(new[] { 2, 4 }, preview.DeletedRows.ToArray());

            new DataCleaner(_excel.App).DeleteBlankRows("A1:A5");
            Assert.Equal("a", ws.Range["A1"].Value2);
            Assert.Equal("b", ws.Range["A2"].Value2);
            Assert.Equal("c", ws.Range["A3"].Value2);
        }

        private static void AssertOk(ToolResult result)
        {
            Assert.True(result.Success, result.Error);
        }

        [ExcelFact]
        public void Sort_preview_matches_excels_own_sort_including_pinyin_and_blanks()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            var rows = new object[,]
            {
                { "姓名", "金额" },
                { "张三", 30.0 },
                { "王五", null },
                { "李四", 10.0 },
                { "赵六", "待定" },
                { "陈七", 20.0 }
            };
            ws.Range["A1:B6"].Value2 = rows;
            var actions = new ExcelActionsImpl(_excel.App, new DeepExcel.AddIn.Perception.WorkbookAnalyzer(_excel.App),
                new DeepExcel.AddIn.Perception.RangeAnalyzer(), null, null, null);

            var builder = new PreviewBuilder(new ExcelRangeReader(_excel.App));
            var byName = builder.Build("sort_data",
                Args("range_address", "A1:B6", "sort_column", "A", "descending", false, "has_header", true));
            AssertPreviewMatchesExecution(ws, "A1:B6", byName, () => AssertOk(actions.SortData("A1:B6", "A", false, true)));
            foreach (var change in byName.Changes)
            {
                Assert.Equal(change.After, PreviewBuilder.Describe(new ExcelRangeReader(_excel.App).Read(null, change.Address).Single()));
            }

            var byAmount = builder.Build("sort_data",
                Args("range_address", "A1:B6", "sort_column", "B", "descending", true, "has_header", true));
            AssertPreviewMatchesExecution(ws, "A1:B6", byAmount, () => AssertOk(actions.SortData("A1:B6", "B", true, true)));
            foreach (var change in byAmount.Changes)
            {
                Assert.Equal(change.After, PreviewBuilder.Describe(new ExcelRangeReader(_excel.App).Read(null, change.Address).Single()));
            }
        }

        [ExcelFact]
        public void Fill_formula_down_preview_matches_autofill()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["A1:B4"].Value2 = new object[,] { { 1, 2 }, { 3, 4 }, { 5, 6 }, { 7, 8 } };
            ws.Range["C1"].Formula = "=A1*B1+$A$1";
            ws.Range["C3"].Value2 = "旧值";

            var preview = new PreviewBuilder(new ExcelRangeReader(_excel.App))
                .Build("fill_formula_down", Args("from_address", "C1", "row_count", 4));

            ws.Range["C1"].AutoFill(ws.Range["C1"].Resize[4, 1], XlAutoFillType.xlFillDefault);

            foreach (var change in preview.Changes)
            {
                Assert.Equal(change.After, (string)ws.Range[change.Address].Formula);
            }
            Assert.Equal(3, preview.AffectedCells);
        }
    }
}
