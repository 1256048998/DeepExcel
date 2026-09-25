using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Executor;
using ToolResult = DeepExcel.AddIn.Bridge.ToolResult;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 每步检查点：每个写入步骤执行前各存一份，面板上「回到这一步之前」就是字面意思；
    /// 之后的写入涉及新的表时扩大更早检查点的范围，回退时之后的修改一起回退。
    /// </summary>
    public class CheckpointTests
    {
        private const string Book = @"C:\data\book.xlsx";

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        private static ToolDispatcher Dispatcher(FakeExcelActions fake)
            => new ToolDispatcher(fake, null) { BoundWorkbookKey = () => Book, SlowCheckpointThresholdMs = int.MaxValue };

        private static ToolResult Write(ToolDispatcher d, string address)
            => d.Execute("write_value", Args(("address", address), ("value", "x")));

        [Fact]
        public void Every_write_step_gets_its_own_checkpoint_taken_before_it_runs()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            var first = Write(d, "A1");
            var second = Write(d, "A2");

            Assert.Equal("backup-1", first.CheckpointId);
            Assert.Equal("backup-2", second.CheckpointId);
            Assert.Equal(new[] { "backup", "write_value", "backup", "write_value" },
                fake.Timeline.Where(t => t == "backup" || t == "write_value"));
        }

        [Fact]
        public void Read_only_steps_have_no_checkpoint()
        {
            var fake = new FakeExcelActions();
            Assert.Null(Dispatcher(fake).Execute("read_workbook", Args()).CheckpointId);
        }

        [Fact]
        public void A_failed_step_offers_no_checkpoint()
        {
            var fake = new FakeExcelActions { WriteValueFn = (a, v) => new ToolResult { Success = false, Error = "工作表受保护" } };
            var r = Write(Dispatcher(fake), "A1");
            Assert.False(r.Success);
            Assert.Null(r.CheckpointId);
        }

        [Fact]
        public void Later_steps_on_new_sheets_widen_every_earlier_checkpoint()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            Write(d, "A1");                 // backup-1: Sheet1
            Write(d, "Sheet2!A1");          // backup-2: Sheet2；backup-1 扩到 Sheet2
            d.BeginTurn();
            Write(d, "Sheet3!A1");          // backup-3: Sheet3；backup-1、backup-2 都扩到 Sheet3

            var widened = fake.ExtendScopeCalls.Select(c => (c.Id, string.Join(",", c.Scope.Sheets))).ToList();
            Assert.Equal(new[] { ("backup-1", "Sheet2"), ("backup-1", "Sheet3"), ("backup-2", "Sheet3") }, widened);
        }

        [Fact]
        public void Writes_to_sheets_a_checkpoint_already_covers_do_not_rewrite_its_metadata()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            Write(d, "A1");
            Write(d, "A2");
            Write(d, "B7");

            Assert.Empty(fake.ExtendScopeCalls);
        }

        [Fact]
        public void A_checkpoint_whose_scope_cannot_be_widened_is_dropped_but_writing_continues()
        {
            var fake = new FakeExcelActions { ExtendSnapshotScopeFn = (id, s) => id != "backup-2" };
            var d = Dispatcher(fake);

            Write(d, "A1");                 // backup-1（回合备份）
            Write(d, "Sheet2!A1");          // backup-2
            var r = Write(d, "Sheet3!A1");  // backup-2 记不下 Sheet3 → 不再扩它
            Write(d, "Sheet4!A1");

            Assert.True(r.Success);
            Assert.Equal("backup-1", r.BackupSnapshotId);
            Assert.DoesNotContain(fake.ExtendScopeCalls.Skip(3), c => c.Id == "backup-2");
        }

        [Fact]
        public void A_slow_backup_makes_the_rest_of_the_turn_share_the_turn_backup()
        {
            var fake = new FakeExcelActions();
            var d = new ToolDispatcher(fake, null) { BoundWorkbookKey = () => Book, SlowCheckpointThresholdMs = -1 };

            var first = Write(d, "A1");
            var second = Write(d, "A2");

            Assert.Equal("backup-1", first.CheckpointId);
            Assert.Null(second.CheckpointId);           // 面板不给这一步「回到这一步之前」
            Assert.Equal("backup-1", second.BackupSnapshotId);
            Assert.Single(fake.BackupCalls);

            d.BeginTurn();
            Assert.Equal("backup-2", Write(d, "A3").CheckpointId);
        }

        [Fact]
        public void A_long_turn_stops_taking_checkpoints_before_it_evicts_its_own_turn_backup()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);

            var results = Enumerable.Range(1, ToolDispatcher.MaxCheckpointsPerTurn + 3).Select(i => Write(d, "A" + i)).ToList();

            Assert.Equal(ToolDispatcher.MaxCheckpointsPerTurn, results.Count(r => r.CheckpointId != null));
            Assert.All(results, r => Assert.Equal("backup-1", r.BackupSnapshotId));
            Assert.True(ToolDispatcher.MaxCheckpointsPerTurn < SnapshotManager.MaxSnapshotsPerWorkbook);
        }

        [Fact]
        public void A_failed_step_checkpoint_does_not_block_a_write_the_turn_backup_protects()
        {
            var calls = 0;
            var fake = new FakeExcelActions
            {
                BackupWorkbookFn = (k, r, s) => ++calls == 1
                    ? new SnapshotAttempt { SnapshotId = "turn" }
                    : new SnapshotAttempt { Error = "磁盘满" },
            };
            var d = Dispatcher(fake);

            Write(d, "A1");
            var r2 = Write(d, "A2");

            Assert.True(r2.Success);
            Assert.Null(r2.CheckpointId);
            Assert.Equal("turn", r2.BackupSnapshotId);
        }

        [Fact]
        public void The_checkpoint_id_reaches_the_sidecar()
        {
            var json = PythonSidecar.BuildToolResultJson("c1", true, null, null, null, null, "turn", null, null, "step-3");
            Assert.Contains("\"checkpoint_id\":\"step-3\"", json);
            Assert.Contains("\"backup_snapshot_id\":\"turn\"", json);
        }
    }
}
