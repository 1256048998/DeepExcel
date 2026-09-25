using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>COM 瞬时错误重试、硬错误提示、同错换策略。</summary>
    public class ComErrorTests
    {
        private static COMException Com(uint hr) => new COMException("boom", unchecked((int)hr));

        [Fact]
        public void Only_the_three_call_rejected_codes_are_transient()
        {
            Assert.True(ComErrors.IsTransient(Com(0x80010001)));
            Assert.True(ComErrors.IsTransient(Com(0x8001010A)));
            Assert.True(ComErrors.IsTransient(Com(0x800AC472)));
            Assert.True(ComErrors.IsTransient(new InvalidOperationException("wrap", Com(0x800AC472))));
            Assert.False(ComErrors.IsTransient(Com(0x800A03EC)));
            Assert.False(ComErrors.IsTransient(new InvalidOperationException("x")));
        }

        [Fact]
        public void Hard_errors_come_with_a_next_step()
        {
            Assert.Contains("read_range", ComErrors.Hint(Com(0x800A03EC)));
            Assert.Contains("不要继续重试", ComErrors.Hint(Com(0x800706BA)));
            Assert.Contains("Esc", ComErrors.Hint(Com(0x800AC472)));
            Assert.Null(ComErrors.Hint(new InvalidOperationException("随便什么")));
        }

        private static ToolDispatcher Dispatcher(FakeExcelActions fake) =>
            new ToolDispatcher(fake, null) { TransientRetryDelayMs = _ => 0 };

        [Fact]
        public void Read_tools_are_retried_through_transient_errors()
        {
            var calls = 0;
            var fake = new FakeExcelActions
            {
                ReadRangeFn = a =>
                {
                    if (++calls < 3) throw Com(0x800AC472);
                    return new { cells = new[] { "ok" } };
                },
            };
            var result = Dispatcher(fake).Execute("read_range", new Dictionary<string, object> { { "address", "A1" } });
            Assert.True(result.Success, result.Error);
            Assert.Equal(3, calls);
        }

        [Fact]
        public void Retries_are_bounded_and_end_with_a_hint()
        {
            var calls = 0;
            var fake = new FakeExcelActions { ReadRangeFn = a => { calls++; throw Com(0x80010001); } };
            var result = Dispatcher(fake).Execute("read_range", new Dictionary<string, object> { { "address", "A1" } });
            Assert.False(result.Success);
            Assert.Equal(ComErrors.MaxAttempts, calls);
            Assert.Contains("Esc", result.Suggestion);
        }

        [Fact]
        public void Write_tools_are_not_rerun_on_a_transient_error()
        {
            var calls = 0;
            var fake = new FakeExcelActions { WriteFormulaFn = (a, f) => { calls++; throw Com(0x800AC472); } };
            var result = Dispatcher(fake).Execute("write_formula",
                new Dictionary<string, object> { { "address", "A1" }, { "formula", "=1" } });
            Assert.False(result.Success);
            Assert.Equal(1, calls);
            Assert.Contains("Esc", result.Suggestion);
        }

        [Fact]
        public void The_same_error_three_times_in_a_row_asks_for_a_different_strategy()
        {
            var row = 0;
            var fail = true;
            var fake = new FakeExcelActions
            {
                ReadRangeFn = a =>
                {
                    if (fail) throw new InvalidOperationException($"第 {++row} 行不存在");
                    return new { cells = new[] { "ok" } };
                },
            };
            var dispatcher = Dispatcher(fake);
            var args = new Dictionary<string, object> { { "address", "A1" } };
            Assert.DoesNotContain("换一种做法", dispatcher.Execute("read_range", args).Suggestion ?? "");
            Assert.DoesNotContain("换一种做法", dispatcher.Execute("read_range", args).Suggestion ?? "");
            Assert.Contains("换一种做法", dispatcher.Execute("read_range", args).Suggestion);  // 行号不同也算同一个错

            fail = false;
            Assert.True(dispatcher.Execute("read_range", args).Success);
            fail = true;
            Assert.DoesNotContain("换一种做法", dispatcher.Execute("read_range", args).Suggestion ?? "");  // 成功一次就重新计数
        }
    }

    /// <summary>弹窗兜底：在测试进程自己身上弹 Win32 MessageBox，看守护线程点了哪个按钮。</summary>
    public class DialogGuardTests
    {
        private const uint MB_OK = 0x0, MB_OKCANCEL = 0x1, MB_YESNOCANCEL = 0x3, MB_YESNO = 0x4, MB_RETRYCANCEL = 0x5;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr hwnd, string text, string caption, uint type);

        private static int ShowGuarded(uint type, out IReadOnlyList<HandledDialog> handled)
        {
            var answer = 0;
            using (var guard = DialogGuard.Start(Process.GetCurrentProcess().Id, pollMs: 50))
            {
                var thread = new Thread(() => answer = MessageBox(IntPtr.Zero, "要覆盖现有文件吗？", "DeepExcel 测试", type));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                Assert.True(thread.Join(10000), "对话框没有被自动处理");
                handled = guard.Handled;
            }
            return answer;
        }

        [Fact]
        public void Questions_are_answered_no_never_yes()
        {
            Assert.Equal(7, ShowGuarded(MB_YESNO, out var handled));  // IDNO
            var dialog = Assert.Single(handled);
            Assert.Equal("DeepExcel 测试", dialog.Title);
            Assert.Contains("要覆盖现有文件吗", dialog.Text);
            Assert.Contains("点了", dialog.ToString());
        }

        [Fact]
        public void Cancel_wins_over_no_and_retry_is_never_pressed()
        {
            Assert.Equal(2, ShowGuarded(MB_YESNOCANCEL, out _));  // IDCANCEL
            Assert.Equal(2, ShowGuarded(MB_OKCANCEL, out _));
            Assert.Equal(2, ShowGuarded(MB_RETRYCANCEL, out _));
        }

        [Fact]
        public void Notices_are_acknowledged_with_ok()
        {
            Assert.Equal(1, ShowGuarded(MB_OK, out var handled));  // IDOK
            Assert.Contains("执行期间弹出了 1 个对话框", DialogGuard.Describe(handled));
        }

        [Fact]
        public void Nothing_to_report_when_no_dialog_appeared()
        {
            using (var guard = DialogGuard.Start(Process.GetCurrentProcess().Id, pollMs: 20))
            {
                Thread.Sleep(100);
                Assert.Null(DialogGuard.Describe(guard.Handled));
            }
            using (var idle = DialogGuard.Start(0)) Assert.Empty(idle.Handled);
        }
    }

    /// <summary>真 Excel：VBA 改了全局开关没改回来会被恢复并回报；编译错误弹窗被点掉、原因交回模型。默认跳过。</summary>
    public class VbaExecutionGuardExcelTests : ExcelTestHost
    {
        [ExcelFact]
        public void Global_switches_left_changed_by_the_code_are_restored_and_reported()
        {
            NewBook(out _);
            App.Calculation = XlCalculation.xlCalculationAutomatic;
            App.EnableEvents = true;
            var result = new VBAExecutor(App).Execute(string.Join("\n",
                "Sub Work()",
                "    Application.Calculation = xlCalculationManual",
                "    Application.EnableEvents = False",
                "    Application.StatusBar = \"processing\"",
                "    Range(\"A1\").Value = 1",
                "End Sub"));
            Assert.True(result.Success, result.Error);
            Assert.Contains("Calculation", result.Warning);
            Assert.Contains("EnableEvents", result.Warning);
            Assert.Contains("StatusBar", result.Warning);
            Assert.Equal(XlCalculation.xlCalculationAutomatic, App.Calculation);
            Assert.True(App.EnableEvents);
            Assert.Equal(false, (object)App.StatusBar);
            Assert.True(App.ScreenUpdating);
        }

        [ExcelFact]
        public void Tidy_code_gets_no_restore_notice()
        {
            NewBook(out _);
            var result = new VBAExecutor(App).Execute(string.Join("\n",
                "Sub Work()",
                "    Application.EnableEvents = False",
                "    Range(\"A1\").Value = 1",
                "    Application.EnableEvents = True",
                "End Sub"));
            Assert.True(result.Success, result.Error);
            Assert.Null(result.Warning);
        }

        [ExcelFact]
        public void A_compile_error_dialog_is_dismissed_and_its_message_returned()
        {
            NewBook(out var wb);
            var result = new VBAExecutor(App).Execute(string.Join("\n",
                "Option Explicit",
                "Sub Work()",
                "    undeclared = 1",
                "End Sub"));
            Assert.False(result.Success);
            Assert.Contains("VBA 编译错误", result.Error);
            Assert.Contains("undeclared = 1", result.Error);  // 出错的那一行
            Assert.False(App.VBE.MainWindow.Visible);
            // 事后 Excel 还能正常用
            ((Worksheet)wb.Worksheets[1]).Range["A1"].Value2 = 5;
            Assert.Equal(5.0, ((Worksheet)wb.Worksheets[1]).Range["A1"].Value2);
        }
    }
}
