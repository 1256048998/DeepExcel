using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Executor;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 需要真实 Excel 的测试：默认跳过，设置 DEEPEXCEL_EXCEL_TESTS=1 才运行。
    /// 单元测试证明不了"COM 那边真的是这样"，回滚的正确性只能在真 Excel 里验证。
    /// 每个测试启动独立的隐藏 Excel 实例，不会碰用户正开着的 Excel。
    /// </summary>
    public sealed class ExcelFactAttribute : FactAttribute
    {
        public ExcelFactAttribute()
        {
            if (Environment.GetEnvironmentVariable("DEEPEXCEL_EXCEL_TESTS") != "1")
            {
                Skip = "需要真实 Excel：设置 DEEPEXCEL_EXCEL_TESTS=1 后运行";
            }
        }
    }

    public class SnapshotManagerExcelTests : IDisposable
    {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "DeepExcelSnapExcel_" + Guid.NewGuid().ToString("N"));
        private readonly string _snapDir;
        private readonly string _docDir;
        private Excel.Application _app;

        public SnapshotManagerExcelTests()
        {
            _snapDir = Path.Combine(_root, "snapshots");
            _docDir = Path.Combine(_root, "docs");
            Directory.CreateDirectory(_snapDir);
            Directory.CreateDirectory(_docDir);
        }

        private int _excelPid;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private Excel.Application App()
        {
            if (_app == null)
            {
                _app = new Excel.Application { Visible = false, DisplayAlerts = false };
                try { GetWindowThreadProcessId(new IntPtr(_app.Hwnd), out _excelPid); } catch { }
            }
            return _app;
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
            // 本机装的其他加载项可能让 Quit 之后的进程继续挂着；只结束本测试自己启动的那个
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

        private static byte[] ReadShared(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private Workbook NewBook(int sheets = 2)
        {
            var wb = App().Workbooks.Add();
            while (wb.Worksheets.Count < sheets) wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            for (int i = 1; i <= sheets; i++) ((Worksheet)wb.Worksheets[i]).Name = "Sheet" + i;
            return wb;
        }

        private static Worksheet Ws(Workbook wb, string name) => (Worksheet)wb.Worksheets[name];
        private static object V(Workbook wb, string sheet, string addr) => Ws(wb, sheet).Range[addr].Value2;

        [ExcelFact]
        public void Rollback_targets_the_snapshots_workbook_even_when_another_one_is_active()
        {
            var target = NewBook();
            var path = Path.Combine(_docDir, "target.xlsx");
            Ws(target, "Sheet1").Range["A1"].Value2 = "original";
            target.SaveAs(path);
            var bytesOnDisk = ReadShared(path);

            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(target, "test", SnapshotScope.ForSheets("Sheet1"));
            Assert.True(snap.Success, snap.Error);

            Ws(target, "Sheet1").Range["A1"].Value2 = "changed by AI";
            Ws(target, "Sheet1").Range["B1"].Value2 = "unsaved user edit";

            // 用户切到另一个工作簿
            var other = NewBook();
            Ws(other, "Sheet1").Range["A1"].Value2 = "other unsaved";
            other.Activate();

            var r = mgr.Rollback(snap.SnapshotId);

            Assert.True(r.Success, r.Error);
            Assert.Equal("original", V(target, "Sheet1", "A1"));
            // 两本都还开着，另一本没被关、没被动
            Assert.Equal(2, App().Workbooks.Count);
            Assert.Equal("other unsaved", V(other, "Sheet1", "A1"));
            // 不改写磁盘上的原文件
            Assert.Equal(bytesOnDisk, ReadShared(path));
            Assert.False(target.Saved);
            // 恢复前的状态（含未保存编辑）另存了一份
            Assert.NotNull(r.PreRestoreSnapshotId);
            Assert.Equal(".xlsx", mgr.GetMeta(r.PreRestoreSnapshotId).FileExtension);
        }

        [ExcelFact]
        public void Pre_restore_backup_lets_the_restore_itself_be_undone()
        {
            var wb = NewBook();
            Ws(wb, "Sheet1").Range["A1"].Value2 = "v1";
            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole());
            Ws(wb, "Sheet1").Range["A1"].Value2 = "v2";

            var r = mgr.Rollback(snap.SnapshotId);
            Assert.Equal("v1", V(wb, "Sheet1", "A1"));

            var undo = mgr.Rollback(r.PreRestoreSnapshotId);
            Assert.True(undo.Success, undo.Error);
            Assert.Equal("v2", V(wb, "Sheet1", "A1"));
        }

        [ExcelFact]
        public void Macro_workbooks_are_snapshotted_as_xlsm_and_can_be_restored()
        {
            var wb = NewBook();
            var path = Path.Combine(_docDir, "宏.xlsm");
            Ws(wb, "Sheet1").Range["A1"].Value2 = 1;
            wb.SaveAs(path, XlFileFormat.xlOpenXMLWorkbookMacroEnabled);

            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole());
            Assert.True(snap.Success, snap.Error);
            Assert.True(File.Exists(Path.Combine(_snapDir, snap.SnapshotId + ".xlsm")));

            Ws(wb, "Sheet1").Range["A1"].Value2 = 2;
            var r = mgr.Rollback(snap.SnapshotId);

            Assert.True(r.Success, r.Error);
            Assert.Equal(1.0, V(wb, "Sheet1", "A1"));
            Assert.Equal((int)XlFileFormat.xlOpenXMLWorkbookMacroEnabled, (int)wb.FileFormat);
        }

        [ExcelFact]
        public void Scoped_restore_keeps_edits_on_other_sheets()
        {
            var wb = NewBook();
            Ws(wb, "Sheet1").Range["A1"].Value2 = "s1";
            Ws(wb, "Sheet2").Range["A1"].Value2 = "s2";
            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.ForSheets("Sheet1"));

            Ws(wb, "Sheet1").Range["A1"].Value2 = "AI wrote this";
            Ws(wb, "Sheet2").Range["A1"].Value2 = "user wrote this";

            var r = mgr.Rollback(snap.SnapshotId);

            Assert.True(r.Success, r.Error);
            Assert.Equal("s1", V(wb, "Sheet1", "A1"));
            Assert.Equal("user wrote this", V(wb, "Sheet2", "A1"));
            Assert.Equal(new[] { "Sheet1" }, r.RestoredSheets);
        }

        [ExcelFact]
        public void Cross_sheet_formulas_come_back_pointing_at_the_workbook_itself()
        {
            // 跨工作簿复制单元格会把 =Sheet2!A1 变成指向快照文件的外部链接
            foreach (var saved in new[] { true, false })
            {
                var wb = NewBook();
                if (saved) wb.SaveAs(Path.Combine(_docDir, "links.xlsx"));
                Ws(wb, "Sheet2").Range["A1"].Value2 = 21;
                Ws(wb, "Sheet1").Range["A1"].Formula = "=Sheet2!A1*2";
                var mgr = new SnapshotManager(App(), _snapDir);
                var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.ForSheets("Sheet1"));

                Ws(wb, "Sheet1").Range["A1"].Value2 = 0;
                var r = mgr.Rollback(snap.SnapshotId);

                Assert.True(r.Success, r.Error);
                Assert.Equal("=Sheet2!A1*2", (string)Ws(wb, "Sheet1").Range["A1"].Formula);
                Assert.Equal(42.0, V(wb, "Sheet1", "A1"));
                Assert.Null(wb.LinkSources(XlLink.xlExcelLinks));
                Assert.Empty(r.Warnings.Where(w => w.Contains("快照文件")));
                wb.Close(false);
            }
        }

        [ExcelFact]
        public void References_from_other_sheets_survive_an_in_place_restore()
        {
            // 原地恢复而不是删表重建：别的表对这张表的引用不会变成 #REF!
            var wb = NewBook();
            Ws(wb, "Sheet1").Range["A1"].Value2 = 5;
            Ws(wb, "Sheet2").Range["A1"].Formula = "=Sheet1!A1+1";
            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.ForSheets("Sheet1"));

            Ws(wb, "Sheet1").Range["A1"].Value2 = 100;
            var r = mgr.Rollback(snap.SnapshotId);

            Assert.True(r.Success, r.Error);
            Assert.Equal("=Sheet1!A1+1", (string)Ws(wb, "Sheet2").Range["A1"].Formula);
            Assert.Equal(6.0, V(wb, "Sheet2", "A1"));
        }

        [ExcelFact]
        public void Whole_workbook_restore_removes_added_sheets_and_brings_back_deleted_ones()
        {
            var wb = NewBook(3);
            Ws(wb, "Sheet3").Range["A1"].Value2 = "keep me";
            Ws(wb, "Sheet1").Range["C3"].Interior.Color = 255;
            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole());

            App().DisplayAlerts = false;
            Ws(wb, "Sheet3").Delete();
            var added = (Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            added.Name = "AI新建";
            Ws(wb, "Sheet1").Range["C3"].Interior.ColorIndex = XlColorIndex.xlColorIndexNone;

            var r = mgr.Rollback(snap.SnapshotId);

            Assert.True(r.Success, r.Error);
            var names = wb.Worksheets.Cast<Worksheet>().Select(w => w.Name).ToArray();
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, names);
            Assert.Equal("keep me", V(wb, "Sheet3", "A1"));
            Assert.Equal(255.0, Convert.ToDouble(Ws(wb, "Sheet1").Range["C3"].Interior.Color));
            Assert.Contains("AI新建", r.RemovedSheets);
        }

        [ExcelFact]
        public void Restoring_the_oldest_snapshot_at_the_quota_limit_still_works()
        {
            // 恢复前的备份会让配额超 1；若先清理，被删掉的正是要恢复的这份
            var wb = NewBook();
            var mgr = new SnapshotManager(App(), _snapDir);
            Ws(wb, "Sheet1").Range["A1"].Value2 = "oldest";
            var oldest = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole());
            Ws(wb, "Sheet1").Range["A1"].Value2 = "newer";
            for (int i = 1; i < SnapshotManager.MaxSnapshotsPerWorkbook; i++)
            {
                System.Threading.Thread.Sleep(5);
                Assert.True(mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole()).Success);
            }

            var r = mgr.Rollback(oldest.SnapshotId);

            Assert.True(r.Success, r.Error);
            Assert.Equal("oldest", V(wb, "Sheet1", "A1"));
            Assert.True(mgr.ListSnapshots().Count <= SnapshotManager.MaxSnapshotsPerWorkbook);
            Assert.NotNull(mgr.GetMeta(r.PreRestoreSnapshotId));
        }

        [ExcelFact]
        public void Rollback_refuses_when_the_owner_workbook_is_not_open()
        {
            var wb = NewBook();
            wb.SaveAs(Path.Combine(_docDir, "closed.xlsx"));
            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole());
            wb.Close(false);
            var other = NewBook();
            Ws(other, "Sheet1").Range["A1"].Value2 = "untouched";

            var r = mgr.Rollback(snap.SnapshotId);

            Assert.False(r.Success);
            Assert.Contains("closed.xlsx", r.Error);
            Assert.Equal("untouched", V(other, "Sheet1", "A1"));
            Assert.Equal(1, App().Workbooks.Count);
        }

        [ExcelFact]
        public void Restore_leaves_excel_state_as_it_found_it()
        {
            var wb = NewBook();
            var other = NewBook();
            other.Activate();
            var mgr = new SnapshotManager(App(), _snapDir);
            var snap = mgr.TryCreateSnapshot(wb, "test", SnapshotScope.Whole());
            App().EnableEvents = true;

            mgr.Rollback(snap.SnapshotId);

            Assert.True(App().EnableEvents);
            Assert.Equal(other.Name, App().ActiveWorkbook.Name);
            Assert.Equal(2, App().Workbooks.Count); // 隐藏打开的快照已关掉
        }
    }
}
