using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 写入守卫：会改工作簿的工具在执行前自动备份，备份失败就不执行。
    ///
    /// 以前只有 clean_data / execute_vba 会自动备份（VBA 那次还是失败照样执行），
    /// 其余 40 多个写入工具都依赖模型记得先调 create_snapshot——改坏了就无从恢复。
    /// </summary>
    public class WriteGuardTests
    {
        private const string Book = @"C:\data\book.xlsx";

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        private static ToolDispatcher Dispatcher(FakeExcelActions fake, string bound = Book)
            => new ToolDispatcher(fake, null) { BoundWorkbookKey = () => bound };

        [Fact]
        public void First_write_of_a_turn_is_backed_up_before_it_runs()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            var r = d.Execute("write_formula", Args(("address", "B2"), ("formula", "=1")));

            Assert.True(r.Success);
            Assert.Equal(new[] { "backup", "write_formula" }, fake.Timeline);
            Assert.Equal(Book, fake.BackupCalls[0].Key);
            Assert.Equal("backup-1", r.BackupSnapshotId);
        }

        [Fact]
        public void Failed_backup_means_the_write_never_runs()
        {
            var fake = new FakeExcelActions
            {
                BackupWorkbookFn = (k, reason, s) => new SnapshotAttempt { Error = "磁盘已满" },
            };
            var d = Dispatcher(fake);

            var r = d.Execute("write_formula", Args(("address", "B2"), ("formula", "=1")));

            Assert.False(r.Success);
            Assert.Contains("备份失败", r.Error);
            Assert.Contains("磁盘已满", r.Error);
            Assert.Empty(fake.WriteFormulaCalls);
        }

        [Fact]
        public void Failed_backup_blocks_vba_too()
        {
            // 以前 VBA 的快照失败时照样执行（fail-open）
            var fake = new FakeExcelActions
            {
                BackupWorkbookFn = (k, reason, s) => new SnapshotAttempt { Error = "x" },
            };
            var r = Dispatcher(fake).Execute("execute_vba", Args(("code", "Sub A()\nEnd Sub")));

            Assert.False(r.Success);
            Assert.Empty(fake.ExecuteVBACalls);
        }

        [Fact]
        public void A_throwing_backup_is_treated_as_a_failed_backup()
        {
            var fake = new FakeExcelActions
            {
                BackupWorkbookFn = (k, reason, s) => throw new System.Runtime.InteropServices.COMException("busy"),
            };
            var r = Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));

            Assert.False(r.Success);
            Assert.DoesNotContain("write_value", fake.Timeline);
        }

        [Fact]
        public void Later_writes_in_the_same_turn_keep_the_turn_backup_and_extend_its_scope()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            d.Execute("write_value", Args(("address", "A1"), ("value", "x")));
            var second = d.Execute("write_value", Args(("address", "Sheet2!A1"), ("value", "y")));

            // 第二步另存了自己的检查点，但给模型的回合备份仍是第一份
            Assert.Equal(2, fake.BackupCalls.Count);
            Assert.Equal("backup-1", second.BackupSnapshotId);
            Assert.Equal("backup-2", second.CheckpointId);
            Assert.Single(fake.ExtendScopeCalls);
            Assert.Equal("backup-1", fake.ExtendScopeCalls[0].Id);
            Assert.Equal(new[] { "Sheet2" }, fake.ExtendScopeCalls[0].Scope.Sheets);
        }

        [Fact]
        public void A_new_turn_takes_a_new_backup()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            d.Execute("write_value", Args(("address", "A1"), ("value", "x")));
            d.BeginTurn();
            var r = d.Execute("write_value", Args(("address", "A1"), ("value", "y")));

            Assert.Equal(2, fake.BackupCalls.Count);
            Assert.Equal("backup-2", r.BackupSnapshotId);
        }

        [Fact]
        public void If_the_scope_cannot_be_recorded_a_fresh_backup_is_taken()
        {
            // 记不下新涉及的表，回滚就会漏掉它——不能当作"已备份"
            var fake = new FakeExcelActions { ExtendSnapshotScopeFn = (id, s) => false };
            var d = Dispatcher(fake);

            d.Execute("write_value", Args(("address", "A1"), ("value", "x")));
            var r = d.Execute("write_value", Args(("address", "Sheet2!A1"), ("value", "y")));

            Assert.True(r.Success);
            Assert.Equal(2, fake.BackupCalls.Count);
            Assert.Equal("backup-2", r.BackupSnapshotId);
        }

        [Fact]
        public void Writes_go_to_the_bound_workbook_after_the_user_switches_to_another()
        {
            // 以前整体拒绝；现在工具被重定向到会话那本，备份的也是那本
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"C:\data\other.xlsx" };
            var r = Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));

            Assert.True(r.Success);
            Assert.Equal(new[] { Book }, fake.TargetCalls);
            Assert.Equal(Book, fake.BackupCalls.Single().Key);
            Assert.Contains("write_value", fake.Timeline);
        }

        [Theory]
        [InlineData("execute_vba")]
        [InlineData("execute_python")]
        [InlineData("read_selection")]
        [InlineData("freeze_panes")]
        [InlineData("send_keys")]
        public void Foreground_only_tools_are_refused_after_the_user_switches(string tool)
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"C:\data\other.xlsx" };
            var r = Dispatcher(fake).Execute(tool, Args(("address", "A1"), ("code", "x"), ("macro_name", "A"), ("keys", "{F9}")));

            Assert.False(r.Success);
            Assert.Contains("book.xlsx", r.Error);
            Assert.Contains("前台", r.Error);
            Assert.Empty(fake.BackupCalls);
            Assert.Empty(fake.TargetCalls);
        }

        [Fact]
        public void Nothing_is_redirected_while_the_bound_workbook_is_in_front()
        {
            var fake = new FakeExcelActions();
            Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));
            Assert.Empty(fake.TargetCalls);
        }

        [Fact]
        public void Workbook_keys_compare_case_insensitively()
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"c:\DATA\BOOK.xlsx" };
            var r = Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));
            Assert.True(r.Success);
            Assert.Empty(fake.TargetCalls);
        }

        [Fact]
        public void A_closed_bound_workbook_refuses_instead_of_writing_elsewhere()
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"C:\data\other.xlsx" };
            fake.OpenWorkbooks.Clear();
            var r = Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));

            Assert.False(r.Success);
            Assert.Contains("已经关闭", r.Error);
            Assert.DoesNotContain("write_value", fake.Timeline);
            Assert.Empty(fake.BackupCalls);
        }

        [Fact]
        public void No_open_workbook_refuses_the_write()
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = null };
            fake.OpenWorkbooks.Clear();
            var r = Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));
            Assert.False(r.Success);
            Assert.DoesNotContain("write_value", fake.Timeline);
        }

        [Fact]
        public void Attachments_can_be_read_even_if_the_bound_workbook_was_closed()
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"C:\data\other.xlsx" };
            fake.OpenWorkbooks.Clear();
            var r = Dispatcher(fake).Execute("read_attachment", Args(("file_name", "a.csv")));
            Assert.DoesNotContain("已经关闭", r.Error ?? "");
        }

        [Theory]
        [InlineData("read_range")]
        [InlineData("read_workbook")]
        [InlineData("read_selection")]
        [InlineData("execute_python")]
        [InlineData("export_chart")]
        public void Read_only_tools_never_back_up(string tool)
        {
            var fake = new FakeExcelActions();
            Dispatcher(fake).Execute(tool, Args(("address", "A1"), ("code", "print(1)")));
            Assert.Empty(fake.BackupCalls);
        }

        [Fact]
        public void Read_only_tools_still_work_after_switching_workbooks()
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"C:\data\other.xlsx" };
            var r = Dispatcher(fake).Execute("read_range", Args(("address", "A1")));
            Assert.True(r.Success);
        }

        [Fact]
        public void Tools_nobody_classified_are_protected_by_default()
        {
            // 白名单是只读工具；新增的写入工具忘了登记也会先备份
            Assert.True(ToolMutationPolicy.IsMutating("some_future_tool"));
            var fake = new FakeExcelActions();
            Dispatcher(fake).Execute("some_future_tool", Args());
            Assert.Single(fake.BackupCalls);
        }

        [Theory]
        [InlineData("delete_rows")]
        [InlineData("insert_columns")]
        [InlineData("delete_sheet")]
        [InlineData("rename_sheet")]
        [InlineData("execute_vba")]
        [InlineData("clean_data")]
        public void Tools_that_can_touch_other_sheets_back_up_the_whole_workbook(string tool)
        {
            var fake = new FakeExcelActions();
            Dispatcher(fake).Execute(tool, Args(("row", 1), ("count", 1), ("name", "Sheet1"), ("code", "x"), ("range_address", "A1:B2")));
            Assert.True(fake.BackupCalls[0].Scope.WholeWorkbook);
        }

        [Fact]
        public void Unqualified_addresses_resolve_to_the_active_sheet()
        {
            var fake = new FakeExcelActions { ActiveSheetName = "明细" };
            Dispatcher(fake).Execute("write_value", Args(("address", "A1"), ("value", "x")));
            Assert.Equal(new[] { "明细" }, fake.BackupCalls[0].Scope.Sheets);
        }

        [Fact]
        public void Model_rollback_cannot_touch_another_workbooks_snapshot()
        {
            var fake = new FakeExcelActions
            {
                GetSnapshotMetaFn = id => new SnapshotMeta { Id = id, WorkbookKey = @"C:\data\other.xlsx", WorkbookName = "other.xlsx" },
            };
            var r = Dispatcher(fake).Execute("rollback", Args(("snapshot_id", "abc")));

            Assert.False(r.Success);
            Assert.Empty(fake.RollbackCalls);
        }

        [Fact]
        public void Model_rollback_reports_where_the_pre_restore_state_went()
        {
            var fake = new FakeExcelActions
            {
                GetSnapshotMetaFn = id => new SnapshotMeta { Id = id, WorkbookKey = Book },
                RollbackFn = id =>
                {
                    var rr = new RollbackResult { Success = true, SnapshotId = id, PreRestoreSnapshotId = "pre-1" };
                    rr.RestoredSheets.Add("Sheet1");
                    return rr;
                },
            };
            var r = Dispatcher(fake).Execute("rollback", Args(("snapshot_id", "abc")));

            Assert.True(r.Success);
            Assert.Equal("pre-1", (string)((dynamic)r.Data).pre_restore_snapshot_id);
            Assert.Empty(fake.BackupCalls);
        }

        [Fact]
        public void Create_snapshot_backs_up_the_bound_workbook_not_the_active_one()
        {
            var fake = new FakeExcelActions { ActiveWorkbookKey = @"C:\data\other.xlsx" };
            var r = Dispatcher(fake).Execute("create_snapshot", Args());

            Assert.True(r.Success);
            Assert.Equal(Book, fake.BackupCalls.Single().Key);
            Assert.Equal(0, fake.CreateSnapshotCalls);
        }

        [Fact]
        public void Without_a_bound_workbook_the_guard_still_backs_up_the_active_one()
        {
            var fake = new FakeExcelActions();
            var r = new ToolDispatcher(fake, null).Execute("write_value", Args(("address", "A1"), ("value", "x")));

            Assert.True(r.Success);
            Assert.Equal(fake.ActiveWorkbookKey, fake.BackupCalls.Single().Key);
        }
    }

    public class ToolMutationPolicyTests
    {
        [Theory]
        [InlineData("A1", null, false)]
        [InlineData("A1:B20", null, false)]
        [InlineData("$A$1:$C$3", null, false)]
        [InlineData("A:C", null, false)]
        [InlineData("3:5", null, false)]
        [InlineData("Sheet2!A1", "Sheet2", true)]
        [InlineData("'My Sheet'!A1:B2", "My Sheet", true)]
        [InlineData("'It''s'!A1", "It's", true)]
        [InlineData("明细!B:B", "明细", true)]
        public void Parses_single_area_a1_addresses(string address, string sheet, bool qualified)
        {
            Assert.True(ToolMutationPolicy.TryGetSheetOfAddress(address, out var s, out var q));
            Assert.Equal(sheet, s);
            Assert.Equal(qualified, q);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("SalesData")]          // 名称：指向哪张表不知道
        [InlineData("A1,B2")]              // 多区域
        [InlineData("[Other.xlsx]Sheet1!A1")] // 另一个工作簿
        [InlineData("!A1")]
        public void Anything_else_is_not_trusted_as_a_single_sheet(string address)
        {
            Assert.False(ToolMutationPolicy.TryGetSheetOfAddress(address, out _, out _));
        }

        [Fact]
        public void Unknown_addresses_widen_the_scope_to_the_whole_workbook()
        {
            var scope = ToolMutationPolicy.ResolveScope("write_value", k => "SalesData", () => "Sheet1");
            Assert.True(scope.WholeWorkbook);
        }

        [Theory]
        [InlineData("send_keys")]
        [InlineData("execute_vba")]
        [InlineData("delete_blank_rows")]
        [InlineData("create_pivot_table")]
        public void Tools_with_unknowable_reach_are_whole_workbook(string tool)
        {
            var scope = ToolMutationPolicy.ResolveScope(tool, k => "A1:B2", () => "Sheet1");
            Assert.True(scope.WholeWorkbook);
        }

        [Fact]
        public void Add_sheet_scopes_to_the_new_sheet_name()
        {
            var scope = ToolMutationPolicy.ResolveScope("add_sheet", k => k == "name" ? "汇总" : null, () => "Sheet1");
            Assert.Equal(new[] { "汇总" }, scope.Sheets);
        }

        [Fact]
        public void Copy_range_scopes_to_the_destination()
        {
            var scope = ToolMutationPolicy.ResolveScope("copy_range",
                k => k == "dest_address" ? "Out!A1" : "In!A1:B2", () => "Sheet1");
            Assert.Equal(new[] { "Out" }, scope.Sheets);
        }
    }
}
