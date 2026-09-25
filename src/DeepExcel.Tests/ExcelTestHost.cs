using System;
using System.Linq;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Perception;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 真 Excel 测试的公共部分：起一个隐藏的独立实例、建一本带 Data / Sum 两张表的工作簿，
    /// 结束时关掉并确保进程退出（本机其他加载项可能让 Quit 之后的进程挂着）。
    /// 测试方法用 [ExcelFact]，默认跳过。
    /// </summary>
    public abstract class ExcelTestHost : IDisposable
    {
        protected Excel.Application App { get; private set; }
        private int _excelPid;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private protected ExcelActionsImpl NewBook(out Workbook wb)
        {
            App = new Excel.Application { Visible = false, DisplayAlerts = false };
            try { GetWindowThreadProcessId(new IntPtr(App.Hwnd), out _excelPid); } catch { }
            wb = App.Workbooks.Add();
            while (wb.Worksheets.Count < 2) wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            ((Worksheet)wb.Worksheets[1]).Name = "Data";
            ((Worksheet)wb.Worksheets[2]).Name = "Sum";
            ((Worksheet)wb.Worksheets[1]).Activate();
            return new ExcelActionsImpl(App, new WorkbookAnalyzer(App), new RangeAnalyzer(), null, null, null);
        }

        protected static Worksheet Sheet(Workbook wb, string name) => (Worksheet)wb.Worksheets[name];

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
    }
}
