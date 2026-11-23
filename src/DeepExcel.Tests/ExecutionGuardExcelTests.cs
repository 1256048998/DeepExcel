using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The execution guardrails against real Excel, where their unit tests
    /// can only assume what COM does:
    ///
    /// - read-before-write: the addresses Excel's SheetChange hands over
    ///   (single cells, multi-area edits, whole deleted rows) must parse into
    ///   the ledger, or a user's edit goes unnoticed and gets overwritten;
    /// - AppStateGuard: switches VBA left changed really are put back,
    ///   including the status bar, which only clears through MESSAGE(FALSE);
    /// - DialogGuard: a modal dialog in Excel's process is really dismissed,
    ///   and the blocked call returns.
    ///
    /// Each test owns a hidden Excel instance and kills only that process.
    /// </summary>
    public class ExecutionGuardExcelTests : IDisposable
    {
        private readonly HiddenExcel _excel = new HiddenExcel();

        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder name, int max);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int max);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        /// <summary>Visible top-level windows of a process, for a failure message.</summary>
        private static string WindowsOf(int processId)
        {
            var found = new List<string>();
            EnumWindows((hwnd, _) =>
            {
                GetWindowThreadProcessId(hwnd, out var pid);
                if (pid == processId && IsWindowVisible(hwnd))
                {
                    var cls = new System.Text.StringBuilder(256);
                    var title = new System.Text.StringBuilder(256);
                    GetClassName(hwnd, cls, 256);
                    GetWindowText(hwnd, title, 256);
                    found.Add(cls + "「" + title + "」");
                }
                return true;
            }, IntPtr.Zero);
            return found.Count == 0 ? "（无可见窗口）" : string.Join("；", found);
        }

        public void Dispose() => _excel.Dispose();

        // ------------------------------------------------------------------
        // Read-before-write: SheetChange → ReadLedger
        // ------------------------------------------------------------------

        /// <summary>
        /// What ThisAddIn.OnSheetChangeInvalidateIndex does with the event:
        /// the address without $ signs, split on commas, parsed per sheet.
        /// </summary>
        private sealed class LedgerHook
        {
            private readonly object _lock = new object();
            public readonly ReadLedger Ledger = new ReadLedger();
            public readonly List<string> Addresses = new List<string>();

            public void OnSheetChange(object sheet, Range target)
            {
                var worksheet = (Worksheet)sheet;
                var address = target.Address[false, false];
                lock (_lock) Addresses.Add(address);
                foreach (var part in address.Split(','))
                {
                    if (CellRect.TryParse(part, worksheet.Name, out var rect)) Ledger.RecordUserEdit(rect);
                }
            }

            public void WaitFor(int count)
            {
                var clock = Stopwatch.StartNew();
                while (clock.ElapsedMilliseconds < 5000)
                {
                    lock (_lock) if (Addresses.Count >= count) return;
                    Thread.Sleep(20);
                }
            }
        }

        private static CellRect Rect(string address)
        {
            Assert.True(CellRect.TryParse(address, null, out var rect), address);
            return rect;
        }

        [ExcelFact]
        public void A_user_edit_after_the_model_read_is_caught_before_the_model_writes()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["A1:C10"].Value2 = "原值";
            var hook = new LedgerHook();
            _excel.App.SheetChange += hook.OnSheetChange;

            hook.Ledger.RecordRead(Rect("Data!A1:C10"));
            ws.Range["B2"].Value2 = "用户改的";
            hook.WaitFor(1);

            Assert.Equal("B2", hook.Addresses.Single());
            var verdict = hook.Ledger.Check(Rect("Data!B1:B5"), () => true, out var changed);
            Assert.Equal(LedgerVerdict.StaleRead, verdict);
            Assert.Equal(new[] { "Data!B2" }, changed.ToArray());
            // A write elsewhere in what the model read is still fine.
            Assert.Equal(LedgerVerdict.Allowed, hook.Ledger.Check(Rect("Data!C1:C5"), () => true, out _));
        }

        [ExcelFact]
        public void Multi_area_edits_and_deleted_rows_reach_the_ledger_in_a_form_it_parses()
        {
            var wb = _excel.NewBook("Data");
            var ws = (Worksheet)wb.Worksheets["Data"];
            ws.Range["A1:F10"].Value2 = 1;
            var hook = new LedgerHook();
            _excel.App.SheetChange += hook.OnSheetChange;
            hook.Ledger.RecordRead(Rect("Data!A1:F10"));

            ws.Range["D1,F3:F4"].Value2 = 2;
            hook.WaitFor(1);
            ((Range)ws.Rows[8]).Delete();
            hook.WaitFor(2);

            Assert.Equal("D1,F3:F4", hook.Addresses[0]);
            Assert.Equal("8:8", hook.Addresses[1]);
            Assert.Equal(LedgerVerdict.StaleRead, hook.Ledger.Check(Rect("Data!F4"), () => true, out _));
            // The whole deleted row counts as edited, whatever the column.
            Assert.Equal(LedgerVerdict.StaleRead, hook.Ledger.Check(Rect("Data!A8"), () => true, out var changed));
            Assert.Contains("Data!A8:XFD8", changed);
        }

        // ------------------------------------------------------------------
        // AppStateGuard
        // ------------------------------------------------------------------

        [ExcelFact]
        public void App_state_left_changed_is_restored_and_reported()
        {
            _excel.NewBook("Data");
            var app = _excel.App;
            app.Calculation = XlCalculation.xlCalculationAutomatic;
            app.EnableEvents = true;
            var guard = AppStateGuard.Capture(app);

            // What a VBA macro that errors out half way leaves behind.
            app.Calculation = XlCalculation.xlCalculationManual;
            app.EnableEvents = false;
            app.StatusBar = "正在处理……";

            var restored = guard.Restore();

            Assert.Equal(XlCalculation.xlCalculationAutomatic, app.Calculation);
            Assert.True(app.EnableEvents);
            // Cleared means handed back to Excel, which reads as False -- not
            // the text "FALSE".
            Assert.Equal(false, app.StatusBar);
            Assert.Equal(3, restored.Count);
            Assert.Contains(restored, r => r.StartsWith("Calculation"));
            Assert.Contains(restored, r => r.StartsWith("EnableEvents"));
            Assert.Contains(restored, r => r.StartsWith("StatusBar"));
        }

        [ExcelFact]
        public void App_state_untouched_reports_nothing_and_no_workbook_is_not_an_error()
        {
            // With no workbook open, Calculation cannot even be read.
            var guard = AppStateGuard.Capture(_excel.App);
            Assert.Empty(guard.Restore());
            Assert.Null(AppStateGuard.Describe(guard.Restore()));
        }

        // ------------------------------------------------------------------
        // DialogGuard
        // ------------------------------------------------------------------

        [ExcelFact]
        public void A_modal_dialog_in_excel_is_dismissed_and_the_blocked_call_returns()
        {
            _excel.NewBook("Data");
            Assert.NotEqual(0, _excel.ProcessId);

            using (var guard = DialogGuard.Start(_excel.ProcessId, 100))
            {
                Exception failure = null;
                var call = new Thread(() =>
                {
                    // ALERT shows an OK-only message box and blocks until it closes.
                    try { _excel.App.ExecuteExcel4Macro("ALERT(\"DeepExcel 弹窗测试\",2)"); }
                    catch (Exception ex) { failure = ex; }
                }) { IsBackground = true };
                call.Start();

                var returned = call.Join(TimeSpan.FromSeconds(20));
                var windows = returned ? "" : WindowsOf(_excel.ProcessId);
                if (!returned)
                {
                    // Never leave a hung Excel behind; Dispose kills the process too.
                    try { Process.GetProcessById(_excel.ProcessId).Kill(); } catch { }
                }

                Assert.True(returned, "the dialog was not dismissed; the call stayed blocked. Windows: " + windows);
                Assert.Null(failure);
                // Excel's own alert is an Office NUIDialog, not a #32770 box:
                // its text and buttons are not Win32 windows, so it is closed
                // (the same as Esc) and reported by its title.
                var handled = guard.Handled.Single();
                Assert.Equal("Microsoft Excel", handled.Title);
                Assert.Contains("关闭", handled.Clicked);
            }
        }
    }
}
