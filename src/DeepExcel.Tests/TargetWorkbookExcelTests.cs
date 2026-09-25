using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 目标工作簿可指定，在真 Excel 里验证：用户切到另一本工作簿后，工具仍然读写会话那本，
    /// 前台那本不受影响、窗口也不被切走。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。
    /// </summary>
    public class TargetWorkbookExcelTests : IDisposable
    {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "DeepExcelTarget_" + Guid.NewGuid().ToString("N"));
        private Excel.Application _app;
        private int _excelPid;
        private Workbook _bound;
        private Workbook _other;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private ToolDispatcher Setup()
        {
            Directory.CreateDirectory(_root);
            _app = new Excel.Application { Visible = false, DisplayAlerts = false };
            try { GetWindowThreadProcessId(new IntPtr(_app.Hwnd), out _excelPid); } catch { }

            _bound = _app.Workbooks.Add();
            while (_bound.Worksheets.Count < 2) _bound.Worksheets.Add(After: _bound.Worksheets[_bound.Worksheets.Count]);
            ((Worksheet)_bound.Worksheets[1]).Name = "Data";
            ((Worksheet)_bound.Worksheets[2]).Name = "Sum";
            ((Worksheet)_bound.Worksheets["Data"]).Activate();
            _bound.SaveAs(Path.Combine(_root, "bound.xlsx"));

            _other = _app.Workbooks.Add();
            ((Worksheet)_other.Worksheets[1]).Name = "Data";
            _other.SaveAs(Path.Combine(_root, "other.xlsx"));
            _other.Activate();  // 用户切到了另一本

            var actions = new ExcelActionsImpl(_app, new WorkbookAnalyzer(_app), new RangeAnalyzer(), null, null,
                new SnapshotManager(_app, Path.Combine(_root, "snapshots")));
            var boundKey = WorkbookIdentity.KeyOf(_bound);
            return new ToolDispatcher(actions, _app) { BoundWorkbookKey = () => boundKey };
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
            try { Directory.Delete(_root, true); } catch { }
        }

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        private string Cell(Workbook wb, string sheet, string address)
            => Convert.ToString(((Worksheet)wb.Worksheets[sheet]).Range[address].Value2);

        [ExcelFact]
        public void Writes_land_in_the_bound_workbook_while_another_is_in_front()
        {
            var d = Setup();

            var qualified = d.Execute("write_value", Args(("address", "Sum!B2"), ("value", "季度合计")));
            var unqualified = d.Execute("write_value", Args(("address", "C3"), ("value", "明细")));

            Assert.True(qualified.Success, qualified.Error);
            Assert.True(unqualified.Success, unqualified.Error);
            Assert.Equal("季度合计", Cell(_bound, "Sum", "B2"));
            Assert.Equal("明细", Cell(_bound, "Data", "C3"));   // 会话那本自己的活动表
            Assert.Null(((Worksheet)_other.Worksheets["Data"]).Range["C3"].Value2);
            Assert.Equal("other.xlsx", _app.ActiveWorkbook.Name);  // 没把用户的窗口切走
            Assert.NotNull(qualified.BackupSnapshotId);
        }

        [ExcelFact]
        public void Reads_come_from_the_bound_workbook()
        {
            var d = Setup();
            ((Worksheet)_bound.Worksheets["Data"]).Range["A1"].Value2 = "BOUND_BOOK";
            ((Worksheet)_other.Worksheets["Data"]).Range["A1"].Value2 = "FRONT_BOOK";

            var r = d.Execute("read_range", Args(("address", "Data!A1")));

            Assert.True(r.Success, r.Error);
            var json = PythonSidecar.BuildToolResultJson("c1", true, r.Data, null, null, null, null, null);
            Assert.Contains("BOUND_BOOK", json);
            Assert.DoesNotContain("FRONT_BOOK", json);
        }

        [ExcelFact]
        public void Cleaning_a_range_on_a_background_sheet_writes_back_to_that_sheet()
        {
            // 以前逐格写回用 _app.Cells，永远落在前台活动表上
            var d = Setup();
            var sum = (Worksheet)_bound.Worksheets["Sum"];
            sum.Range["A1"].Value2 = "张三,销售";
            sum.Range["A2"].Value2 = "李四,财务";

            Assert.True(d.Execute("read_range", Args(("address", "Sum!A1:A2"))).Success);  // 先读后写
            var r = d.Execute("split_text_to_columns", Args(("range_address", "Sum!A1:A2"), ("delimiter", ",")));

            Assert.True(r.Success, r.Error);
            Assert.Equal("销售", Cell(_bound, "Sum", "B1"));
            Assert.Null(((Worksheet)_bound.Worksheets["Data"]).Range["B1"].Value2);
            Assert.Null(((Worksheet)_other.Worksheets["Data"]).Range["B1"].Value2);
        }

        [ExcelFact]
        public void Vba_is_refused_rather_than_run_against_the_wrong_workbook()
        {
            var d = Setup();
            var r = d.Execute("execute_vba", Args(("code", "Sub A()\nActiveSheet.Range(\"A1\").Value = 1\nEnd Sub"), ("macro_name", "A")));
            Assert.False(r.Success);
            Assert.Contains("前台", r.Error);
            Assert.Null(((Worksheet)_other.Worksheets["Data"]).Range["A1"].Value2);
        }
    }
}
