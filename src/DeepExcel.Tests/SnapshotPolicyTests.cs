using System;
using System.IO;
using DeepExcel.AddIn.Executor;
using Xunit;

namespace DeepExcel.Tests
{
    public class SnapshotFormatTests
    {
        [Theory]
        [InlineData(@"C:\d\宏.xlsm", 52, ".xlsm")]
        [InlineData(@"C:\d\a.XLSM", 52, ".xlsm")]
        [InlineData(@"C:\d\a.xlsb", 50, ".xlsb")]
        [InlineData(@"C:\d\a.xls", 56, ".xls")]
        [InlineData(@"C:\d\a.xlsx", 51, ".xlsx")]
        [InlineData(@"C:\d\a.csv", 6, ".csv")]
        public void Snapshot_keeps_the_source_extension(string fullName, int format, string expected)
        {
            // 以前一律 .xlsx：xlsm 工作簿的快照内容是 xlsm、名字是 xlsx，Excel 拒绝打开
            Assert.Equal(expected, SnapshotFormats.ChooseExtension(fullName, format));
        }

        [Theory]
        [InlineData("工作簿1", 51, ".xlsx")]
        [InlineData("工作簿1", 52, ".xlsm")]
        [InlineData("Book1", 50, ".xlsb")]
        [InlineData("Book1", null, ".xlsx")]
        [InlineData(@"C:\d\odd.data", 52, ".xlsm")]
        public void Unsaved_or_unusual_names_fall_back_to_the_file_format(string fullName, int? format, string expected)
        {
            Assert.Equal(expected, SnapshotFormats.ChooseExtension(fullName, format));
        }
    }

    public class RestorePlanTests
    {
        private static readonly string[] Snap = { "Sheet1", "Sheet2", "旧表" };
        private static readonly string[] Now = { "Sheet1", "sheet2", "新表" };

        [Fact]
        public void Whole_workbook_restores_shared_sheets_in_place_and_reconciles_the_rest()
        {
            var plan = RestorePlan.Build(Snap, Now, SnapshotScope.Whole());

            Assert.Equal(new[] { "Sheet1", "Sheet2" }, plan.RestoreInPlace);
            Assert.Equal(new[] { "旧表" }, plan.CopyFromSnapshot);
            Assert.Equal(new[] { "新表" }, plan.RemoveFromTarget);
        }

        [Fact]
        public void Scoped_restore_leaves_other_sheets_alone()
        {
            // 用户同时在 Sheet2 上的编辑不应被恢复冲掉
            var plan = RestorePlan.Build(Snap, Now, SnapshotScope.ForSheets("Sheet1"));

            Assert.Equal(new[] { "Sheet1" }, plan.RestoreInPlace);
            Assert.Empty(plan.CopyFromSnapshot);
            Assert.Empty(plan.RemoveFromTarget);
        }

        [Fact]
        public void Scoped_restore_of_an_added_sheet_removes_it_and_of_a_deleted_sheet_brings_it_back()
        {
            var plan = RestorePlan.Build(Snap, Now, SnapshotScope.ForSheets("新表", "旧表", "不存在"));

            Assert.Empty(plan.RestoreInPlace);
            Assert.Equal(new[] { "旧表" }, plan.CopyFromSnapshot);
            Assert.Equal(new[] { "新表" }, plan.RemoveFromTarget);
        }

        [Fact]
        public void Empty_scope_means_whole_workbook()
        {
            Assert.True(SnapshotScope.ForSheets().WholeWorkbook);
            Assert.True(SnapshotScope.ForSheets(null, "").WholeWorkbook);
        }

        [Fact]
        public void Legacy_metadata_without_a_scope_restores_the_whole_workbook()
        {
            Assert.True(SnapshotManager.ScopeOf(new SnapshotMeta()).WholeWorkbook);
            Assert.True(SnapshotManager.ScopeOf(null).WholeWorkbook);
            var scoped = SnapshotManager.ScopeOf(new SnapshotMeta { AffectedSheets = new System.Collections.Generic.List<string> { "A" } });
            Assert.Equal(new[] { "A" }, scoped.Sheets);
        }
    }

    public class WorkbookIdentityTests
    {
        [Fact]
        public void Saved_workbooks_use_the_full_path_unsaved_use_the_name()
        {
            Assert.Equal(@"C:\d\a.xlsx", WorkbookIdentity.KeyFrom(@"C:\d\a.xlsx", "a.xlsx"));
            Assert.Equal("工作簿1", WorkbookIdentity.KeyFrom("工作簿1", "工作簿1"));
        }

        [Fact]
        public void Null_keys_never_match()
        {
            Assert.False(WorkbookIdentity.SameKey(null, null));
            Assert.False(WorkbookIdentity.SameKey("", ""));
            Assert.True(WorkbookIdentity.SameKey(@"C:\A.xlsx", @"c:\a.XLSX"));
        }
    }

    /// <summary>不需要 Excel 的 SnapshotManager 行为：ID 校验、文件定位、元数据、清理</summary>
    public class SnapshotStoreTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "DeepExcelSnapTest_" + Guid.NewGuid().ToString("N"));

        public SnapshotStoreTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { }
        }

        [Theory]
        [InlineData(@"..\..\Windows\win")]
        [InlineData("abc")]
        [InlineData("")]
        [InlineData(null)]
        public void Ids_that_are_not_guids_are_rejected_before_touching_the_disk(string id)
        {
            var m = new SnapshotManager(null, _dir);
            Assert.False(SnapshotManager.IsValidSnapshotId(id));
            Assert.False(m.Rollback(id).Success);
            Assert.False(m.DeleteSnapshot(id));
            Assert.Null(m.GetMeta(id));
        }

        [Fact]
        public void Rollback_without_an_owner_refuses_instead_of_guessing()
        {
            var id = Guid.NewGuid().ToString("N");
            File.WriteAllText(Path.Combine(_dir, id + ".xlsx"), "x");
            var m = new SnapshotManager(null, _dir);

            var r = m.Rollback(id);

            Assert.False(r.Success);
            Assert.Contains("未做任何修改", r.Error);
        }

        [Fact]
        public void Finds_macro_snapshots_and_lists_them()
        {
            var id = Guid.NewGuid().ToString("N");
            File.WriteAllText(Path.Combine(_dir, id + ".xlsm"), "x");
            File.WriteAllText(Path.Combine(_dir, id + ".meta.json"),
                "{\"Id\":\"" + id + "\",\"WorkbookName\":\"宏.xlsm\",\"FileExtension\":\".xlsm\",\"CreatedAt\":\"2026-09-24T10:00:00\"}");
            var m = new SnapshotManager(null, _dir);

            Assert.EndsWith(".xlsm", m.FindSnapshotFile(id, m.GetMeta(id)));
            var list = m.ListSnapshots();
            Assert.Single(list);
            Assert.Equal("宏.xlsm", list[0].WorkbookName);

            Assert.True(m.DeleteSnapshot(id));
            Assert.Empty(Directory.GetFiles(_dir));
        }
    }
}
