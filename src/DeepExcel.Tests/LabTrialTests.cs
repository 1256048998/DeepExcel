using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Executor;
using CellRect = DeepExcel.AddIn.Sidecar.CellRect;
using DeepExcel.AddIn.Preview;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    public class LabDiffTests
    {
        private static SheetContent Sheet(string name, int row1, int col1, string[,] cells) =>
            new SheetContent { Name = name, Row1 = row1, Col1 = col1, Cells = cells };

        [Fact]
        public void Reports_changed_cells_and_their_bounding_rect()
        {
            var before = new[] { Sheet("Data", 1, 1, new[,] { { "姓名", "金额" }, { "张三", "100" } }) };
            var after = new[] { Sheet("Data", 1, 1, new[,] { { "姓名", "金额", "税" }, { "张三", "100", "=B2*0.1" } }) };

            var diff = LabDiff.Compare(before, after);

            Assert.Equal(2, diff.AffectedCells);
            Assert.Equal(new[] { "Data!C1", "Data!C2" }, diff.Changes.Select(c => c.Address).ToArray());
            Assert.All(diff.Changes, c => Assert.Equal(ChangeKind.Add, c.Kind));
            var rect = Assert.Single(diff.AffectedRects);
            Assert.Equal("Data!C1:C2", rect.ToA1());
        }

        [Fact]
        public void Flags_a_formula_replaced_by_a_constant_and_cleared_cells()
        {
            var before = new[] { Sheet("S", 2, 2, new[,] { { "=A2*2", "x" } }) };
            var after = new[] { Sheet("S", 2, 2, new[,] { { "84", "" } }) };

            var diff = LabDiff.Compare(before, after);

            var overwrite = diff.Changes.Single(c => c.Address == "S!B2");
            Assert.True(overwrite.OverwritesFormula);
            Assert.Equal(ChangeKind.Overwrite, overwrite.Kind);
            Assert.Equal(ChangeKind.Clear, diff.Changes.Single(c => c.Address == "S!C2").Kind);
        }

        [Fact]
        public void Unchanged_workbook_is_a_no_op_and_lists_are_capped()
        {
            var same = new[] { Sheet("S", 1, 1, new[,] { { "1", "2" } }) };
            Assert.Equal(0, LabDiff.Compare(same, same).AffectedCells);
            Assert.Empty(LabDiff.Compare(same, same).AffectedRects);

            var wide = new string[1, 80];
            for (int i = 0; i < 80; i++) wide[0, i] = "v" + i;
            var diff = LabDiff.Compare(new[] { Sheet("S", 1, 1, new string[0, 0]) }, new[] { Sheet("S", 1, 1, wide) }, maxListed: 50);
            Assert.Equal(80, diff.AffectedCells);
            Assert.Equal(50, diff.Changes.Count);
        }

        [Fact]
        public void Added_and_removed_sheets_are_named()
        {
            var before = new[] { Sheet("Old", 1, 1, new[,] { { "a" } }) };
            var after = new[] { Sheet("汇总", 1, 1, new[,] { { "b" } }) };

            var diff = LabDiff.Compare(before, after);

            Assert.Equal(new[] { "汇总" }, diff.AddedSheets);
            Assert.Equal(new[] { "Old" }, diff.RemovedSheets);
            Assert.Equal("汇总!A1", diff.Changes.Single().Address);
        }
    }

    public class LabCodeCaveatTests
    {
        [Theory]
        [InlineData("Workbooks(\"预算.xlsx\").Sheets(1).Range(\"A1\").Value = 1", "其他工作簿")]
        [InlineData("Set wb = Workbooks.Open(\"C:\\a.xlsx\")", "其他工作簿")]
        [InlineData("p = ThisWorkbook.Path & \"\\out\"", "路径")]
        [InlineData("ActiveWorkbook.Save", "保存")]
        [InlineData("Range(\"A1\").Value = Now", "时间或随机数")]
        [InlineData("x = Int(Rnd * 10)", "时间或随机数")]
        [InlineData("Range(\"A1\").Value = Date", "时间或随机数")]
        public void Spots_code_the_copy_cannot_represent(string code, string expected)
        {
            Assert.Contains(LabRunner.CodeCaveats("Sub T()\n" + code + "\nEnd Sub"), n => n.Contains(expected));
        }

        [Theory]
        [InlineData("Dim d As Date\nd = DateSerial(2024, 1, 1)")]
        [InlineData("Range(\"A1\").Formula = \"=TODAY()+Now\"")]
        [InlineData("' 用 Now 取时间会不稳定，这里不用\nRange(\"A1\").Value = 1")]
        [InlineData("ActiveSheet.Range(\"B2\").Value = ThisWorkbook.Name")]
        public void Ignores_strings_comments_and_type_names(string code)
        {
            Assert.Empty(LabRunner.CodeCaveats("Sub T()\n" + code + "\nEnd Sub"));
        }
    }

    public class LabTrialTrackerTests
    {
        private static CellRect R(string a1) { CellRect.TryParse(a1, "Data", out var r); return r; }

        [Fact]
        public void Edit_overlapping_the_trial_marks_it_stale_once()
        {
            var t = new LabTrialTracker();
            t.Start("r1", "wb");
            Assert.False(t.Complete("r1", new[] { R("Data!B1:C3") }));

            Assert.Empty(t.RecordUserEdit("wb", R("Data!E5")));
            Assert.Empty(t.RecordUserEdit("other", R("Data!B2")));
            Assert.Equal(new[] { "r1" }, t.RecordUserEdit("wb", R("Data!C2")));
            Assert.Empty(t.RecordUserEdit("wb", R("Data!C3")));
        }

        [Fact]
        public void Edits_made_while_the_trial_runs_are_checked_when_it_finishes()
        {
            var t = new LabTrialTracker();
            t.Start("r1", "wb");
            Assert.Empty(t.RecordUserEdit("wb", R("Data!A1")));
            Assert.True(t.Complete("r1", new[] { R("Data!A1:A5") }));

            t.Start("r2", "wb");
            t.RecordUserEdit("wb", R("Data!Z9"));
            Assert.False(t.Complete("r2", new[] { R("Data!A1:A5") }));
        }

        [Fact]
        public void Removed_trials_are_forgotten()
        {
            var t = new LabTrialTracker();
            t.Start("r1", "wb", "ctx");
            Assert.Equal("ctx", t.ContextOf("r1"));
            t.RemoveWorkbook("wb");
            Assert.False(t.IsPending("r1"));
            Assert.Null(t.ContextOf("r1"));
        }
    }

    public class LabTrialPreviewTests
    {
        [Fact]
        public void Skipped_trial_falls_back_to_not_previewable_with_the_reason()
        {
            var preview = MessageBridge.TrialPreview("execute_vba", new LabTrial { SkippedReason = "没能启动独立的副本 Excel" });
            Assert.False(preview.Previewable);
            Assert.Contains("没能启动独立的副本 Excel", preview.NotPreviewableReason);
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void Trial_diff_becomes_the_change_list_with_new_errors_called_out()
        {
            var diff = new LabDiffResult { AffectedCells = 1 };
            diff.Changes.Add(new CellChange { Address = "Data!C2", Before = "", After = "=A2/0", Kind = ChangeKind.Add });
            diff.AddedSheets.Add("汇总");
            var trial = new LabTrial { Ran = true, Success = true, ErrorsBefore = 0, ErrorsAfter = 1, Diff = diff };

            var preview = MessageBridge.TrialPreview("execute_vba", trial);

            Assert.True(preview.Previewable);
            Assert.Equal(1, preview.AffectedCells);
            Assert.Contains(preview.Warnings, w => w.Contains("汇总"));
            Assert.Contains(preview.Warnings, w => w.Contains("从 0 个增加到 1 个"));
        }
    }

    [Collection("LabProcessRegistry")]
    public class LabProcessRegistryTests : IDisposable
    {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "DeepExcelLabReg_" + Guid.NewGuid().ToString("N"));
        private readonly string _savedRoot = LabProcessRegistry.Root;

        public LabProcessRegistryTests() { LabProcessRegistry.Root = _root; }

        public void Dispose()
        {
            LabProcessRegistry.Root = _savedRoot;
            try { Directory.Delete(_root, true); } catch { }
        }

        [Fact]
        public void Live_owner_keeps_its_lab_process()
        {
            // 用当前测试进程充当「副本进程」：主人（也是自己）活着，所以不能被回收
            var self = Process.GetCurrentProcess().Id;
            LabProcessRegistry.Register(self);
            Assert.Equal(0, LabProcessRegistry.Reap());
            Assert.Contains(self.ToString(), File.ReadAllText(Path.Combine(_root, "processes.json")));
            LabProcessRegistry.Unregister(self);
            Assert.DoesNotContain("\"Pid\":" + self, File.ReadAllText(Path.Combine(_root, "processes.json")));
        }

        [Fact]
        public void Reused_pid_is_not_the_same_process()
        {
            var self = Process.GetCurrentProcess();
            var started = self.StartTime.ToUniversalTime().Ticks;
            Assert.True(LabProcessRegistry.IsSameProcess(self.Id, started));
            Assert.False(LabProcessRegistry.IsSameProcess(self.Id, started - TimeSpan.TicksPerMinute));
        }

        [Fact]
        public void Copy_directories_are_only_deleted_inside_the_lab_root()
        {
            var dir = LabProcessRegistry.NewCopyDirectory();
            Assert.StartsWith(_root, dir);
            var outside = Path.Combine(Path.GetTempPath(), "DeepExcelLabOutside_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(outside);
            try
            {
                LabProcessRegistry.DeleteCopyDirectory(outside);
                Assert.True(Directory.Exists(outside));
                LabProcessRegistry.DeleteCopyDirectory(dir);
                Assert.False(Directory.Exists(dir));
            }
            finally
            {
                Directory.Delete(outside, true);
            }
        }
    }

    /// <summary>在真的隐藏 Excel 里试跑。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。</summary>
    [Collection("LabProcessRegistry")]
    public class LabRunnerExcelTests : IDisposable
    {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "DeepExcelLab_" + Guid.NewGuid().ToString("N"));
        private readonly string _savedRoot = LabProcessRegistry.Root;
        private Excel.Application _user;
        private int _userPid;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        public LabRunnerExcelTests() { LabProcessRegistry.Root = _root; }

        private Workbook UserBook(out string copyPath)
        {
            _user = new Excel.Application { Visible = false, DisplayAlerts = false };
            GetWindowThreadProcessId(new IntPtr(_user.Hwnd), out _userPid);
            var wb = _user.Workbooks.Add();
            var ws = (Worksheet)wb.Worksheets[1];
            ws.Name = "Data";
            ws.Range["A1"].Value2 = "金额";
            ws.Range["A2"].Value2 = 100;
            ws.Range["A3"].Value2 = 250;
            ws.Range["B2"].Formula = "=A2*2";
            var dir = LabProcessRegistry.NewCopyDirectory();
            copyPath = Path.Combine(dir, "账本.xlsx");
            wb.SaveCopyAs(copyPath);
            return wb;
        }

        private LabRequest Request(string copy, string code, TimeSpan? timeout = null) => new LabRequest
        {
            CopyPath = copy,
            Code = code,
            ActiveSheet = "Data",
            Selection = "A1",
            Calculation = XlCalculation.xlCalculationAutomatic,
            UserExcelPid = _userPid,
            Timeout = timeout ?? LabRunner.DefaultTimeout,
        };

        public void Dispose()
        {
            if (_user != null)
            {
                try
                {
                    foreach (Workbook wb in _user.Workbooks.Cast<Workbook>().ToList()) { try { wb.Close(false); } catch { } }
                    _user.Quit();
                }
                catch { }
                try { Marshal.FinalReleaseComObject(_user); } catch { }
                _user = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (_userPid != 0)
            {
                try { var p = Process.GetProcessById(_userPid); if (!p.WaitForExit(5000)) p.Kill(); } catch { }
            }
            LabProcessRegistry.Root = _savedRoot;
            try { Directory.Delete(_root, true); } catch { }
        }

        /// <summary>结果先返回、副本进程随后才关：等它退干净（别的测试类并行开着 Excel，不能数进程总数）</summary>
        private static bool ExitsWithin(int pid, TimeSpan limit)
        {
            try
            {
                using (var p = Process.GetProcessById(pid)) return p.WaitForExit((int)limit.TotalMilliseconds);
            }
            catch (ArgumentException)
            {
                return true;  // 已经没有这个进程
            }
        }

        [ExcelFact]
        public void Trial_reports_changes_without_touching_the_real_workbook()
        {
            var wb = UserBook(out var copy);

            var trial = LabRunner.RunAsync(Request(copy,
                "Sub T()\r\n  Range(\"C1\").Value = \"税\"\r\n  Range(\"C2:C3\").Formula = \"=A2*0.1\"\r\n  Range(\"B2\").Value = 1\r\nEnd Sub")).Result;

            Assert.True(trial.Ran, trial.SkippedReason);
            Assert.True(trial.Success, trial.Error);
            Assert.Equal(4, trial.Diff.AffectedCells);
            Assert.Contains(trial.Diff.Changes, c => c.Address == "Data!B2" && c.OverwritesFormula);
            Assert.Contains(trial.Diff.Changes, c => c.Address == "Data!C3" && c.After == "=A3*0.1");
            Assert.Equal("Data!B1:C3", Assert.Single(trial.Diff.AffectedRects).ToA1());

            // 真实工作簿一格没动
            var ws = (Worksheet)wb.Worksheets["Data"];
            Assert.Null(ws.Range["C1"].Value2);
            Assert.Equal("=A2*2", ws.Range["B2"].Formula);
            // 副本进程已经退出、登记已清掉
            Assert.NotEqual(0, trial.ProcessId);
            Assert.True(ExitsWithin(trial.ProcessId, TimeSpan.FromSeconds(15)), "副本 Excel 应已退出");
            System.Threading.Thread.Sleep(500);
            Assert.DoesNotContain("\"Pid\"", File.ReadAllText(Path.Combine(_root, "processes.json")));
            Assert.False(File.Exists(copy), "副本应在试跑后删除");
        }

        [ExcelFact]
        public void Trial_counts_formula_errors_introduced_by_the_code()
        {
            UserBook(out var copy);
            var trial = LabRunner.RunAsync(Request(copy,
                "Sub T()\r\n  Range(\"D2\").Formula = \"=A2/0\"\r\n  Range(\"D3\").Formula = \"=VLOOKUP(99,A1:A3,2,0)\"\r\nEnd Sub")).Result;

            Assert.True(trial.Success, trial.Error);
            Assert.Equal(0, trial.ErrorsBefore);
            Assert.Equal(2, trial.ErrorsAfter);
        }

        [ExcelFact]
        public void Runtime_error_is_reported_as_a_failed_trial()
        {
            UserBook(out var copy);
            var trial = LabRunner.RunAsync(Request(copy,
                "Sub T()\r\n  Worksheets(\"不存在\").Range(\"A1\").Value = 1\r\nEnd Sub")).Result;

            Assert.True(trial.Ran);
            Assert.False(trial.Success);
            Assert.False(string.IsNullOrEmpty(trial.Error));
            Assert.Equal(0, trial.Diff.AffectedCells);
        }

        [ExcelFact]
        public void Endless_loop_is_killed_at_the_timeout()
        {
            UserBook(out var copy);
            var clock = Stopwatch.StartNew();

            var trial = LabRunner.RunAsync(Request(copy,
                "Sub T()\r\n  Dim i As Long\r\n  Do\r\n    i = i + 1\r\n    If i < 0 Then Exit Do\r\n  Loop\r\nEnd Sub", TimeSpan.FromSeconds(8))).Result;

            Assert.True(trial.TimedOut);
            Assert.False(trial.Success);
            Assert.True(clock.Elapsed < TimeSpan.FromSeconds(20), clock.Elapsed.ToString());
            Assert.NotEqual(0, trial.ProcessId);
            Assert.True(ExitsWithin(trial.ProcessId, TimeSpan.FromSeconds(5)), "超时的副本 Excel 应已被结束");
        }
    }
}
