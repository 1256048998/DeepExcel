using System;
using System.Linq;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Bridge;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 写后体检的 COM 部分在真 Excel 里是不是这样：SpecialCells 找错误格、LinkSources 找外部链接、
    /// 引用不存在的表会变成外部链接。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行；独立的隐藏实例。
    /// </summary>
    public class WriteCheckExcelTests : IDisposable
    {
        private Excel.Application _app;
        private int _excelPid;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private ExcelActionsImpl Actions(out Workbook wb)
        {
            _app = new Excel.Application { Visible = false, DisplayAlerts = false };
            try { GetWindowThreadProcessId(new IntPtr(_app.Hwnd), out _excelPid); } catch { }
            wb = _app.Workbooks.Add();
            while (wb.Worksheets.Count < 2) wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            ((Worksheet)wb.Worksheets[1]).Name = "Data";
            ((Worksheet)wb.Worksheets[2]).Name = "Sum";
            ((Worksheet)wb.Worksheets[1]).Activate();
            return new ExcelActionsImpl(_app, null, null, null, null, null);
        }

        public void Dispose()
        {
            if (_app != null)
            {
                try
                {
                    foreach (Workbook wb in _app.Workbooks.Cast<Workbook>().ToList())
                    {
                        try { wb.Close(false); } catch { }
                    }
                    _app.Quit();
                }
                catch { }
                try { Marshal.FinalReleaseComObject(_app); } catch { }
                _app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (_excelPid != 0)
            {
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(_excelPid);
                    if (!p.WaitForExit(5000)) p.Kill();
                }
                catch { }
            }
        }

        [ExcelFact]
        public void Finds_error_cells_on_every_sheet()
        {
            var actions = Actions(out var wb);
            Assert.Equal(0, actions.CaptureHealth(200).ErrorCount);

            var data = (Worksheet)wb.Worksheets["Data"];
            data.Range["A1"].Value2 = 10;
            data.Range["B1"].Formula = "=A1/0";
            ((Worksheet)wb.Worksheets["Sum"]).Range["C3"].Formula = "=NA()";

            var health = actions.CaptureHealth(200);
            Assert.Equal(2, health.ErrorCount);
            Assert.Contains(health.Errors, e => e.Sheet == "Data" && e.Address == "B1" && e.Text == "#DIV/0!");
            Assert.Contains(health.Errors, e => e.Sheet == "Sum" && e.Address == "C3" && e.Text == "#N/A");
            Assert.False(health.Truncated);
        }

        [ExcelFact]
        public void Stops_collecting_at_the_cap_but_keeps_the_true_count()
        {
            var actions = Actions(out var wb);
            ((Worksheet)wb.Worksheets["Data"]).Range["A1:A50"].Formula = "=1/0";

            var health = actions.CaptureHealth(10);
            Assert.Equal(50, health.ErrorCount);
            Assert.Equal(10, health.Errors.Count);
            Assert.True(health.Truncated);
        }

        [ExcelFact]
        public void A_formula_naming_a_missing_sheet_becomes_an_external_link()
        {
            var actions = Actions(out var wb);
            var before = actions.CaptureHealth(200);

            // DisplayAlerts=false 时 Excel 不弹「更新值」对话框，把 NoSuchSheet 当成外部文件
            ((Worksheet)wb.Worksheets["Data"]).Range["D1"].Formula = "=NoSuchSheet!A1";

            var after = actions.CaptureHealth(200);
            var v = DeepExcel.AddIn.Sidecar.WriteCheck.Compare(before, after, null);
            Assert.False(v.Ok);
            Assert.NotNull(v.NewExternalLinks);
        }

        [ExcelFact]
        public void Samples_the_computed_values_of_a_filled_column()
        {
            var actions = Actions(out var wb);
            var data = (Worksheet)wb.Worksheets["Data"];
            data.Range["A1:A20"].Formula = "=ROW()*2";

            var samples = actions.SampleCells("Data!A1:A20", 5);

            Assert.Equal(5, samples.Count);
            Assert.Equal("Data!A1", samples[0].Address);
            Assert.Equal("2", samples[0].Value);
            Assert.Equal("=ROW()*2", samples[0].Formula);
            Assert.Equal("Data!A20", samples.Last().Address);
            Assert.Equal("40", samples.Last().Value);
        }
    }
}
