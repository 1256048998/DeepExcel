using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Preview;
using DeepExcel.AddIn.Sidecar;
using Xunit;
using ToolResult = DeepExcel.AddIn.Bridge.ToolResult;

namespace DeepExcel.Tests
{
    /// <summary>
    /// write_cells（侧车 tools_v2，DEEPEXCEL_TOOLS_V2）：一次调用多步写入，只做一个检查点、
    /// 只弹一个合并预览；任何一步失败就回到那个检查点。
    /// </summary>
    public class WriteCellsBatchTests
    {
        private const string Book = @"C:\data\book.xlsx";

        private static ToolDispatcher Dispatcher(FakeExcelActions fake)
            => new ToolDispatcher(fake, null) { BoundWorkbookKey = () => Book, SlowCheckpointThresholdMs = int.MaxValue };

        /// <summary>Arguments as the sidecar's tool_call delivers them: JsonElement values.</summary>
        private static Dictionary<string, object> Ops(string opsJson)
        {
            using var doc = JsonDocument.Parse("{\"ops\":" + opsJson + "}");
            return new Dictionary<string, object> { ["ops"] = doc.RootElement.GetProperty("ops").Clone() };
        }

        private const string ThreeSteps =
            "[{\"type\":\"text\",\"address\":\"A1\",\"value\":\"月份\"}," +
            " {\"type\":\"formula\",\"address\":\"B1\",\"formula\":\"=SUM(C:C)\"}," +
            " {\"type\":\"values\",\"address\":\"C1\",\"values\":[[1],[2]]}]";

        [Fact]
        public void A_batch_takes_one_checkpoint_before_the_first_write()
        {
            var fake = new FakeExcelActions();
            var r = Dispatcher(fake).Execute("write_cells", Ops(ThreeSteps));

            Assert.True(r.Success, r.Error);
            Assert.Equal("backup-1", r.CheckpointId);
            Assert.Equal(new[] { "backup", "write_value", "write_formula", "write_range" },
                fake.Timeline.Where(t => t == "backup" || t.StartsWith("write_")));
        }

        [Fact]
        public void A_failed_step_rolls_the_whole_batch_back_to_its_checkpoint()
        {
            var fake = new FakeExcelActions
            {
                WriteFormulaFn = (a, f) => new ToolResult { Success = false, Error = "公式写错了", Suggestion = "用英文逗号" },
            };
            var r = Dispatcher(fake).Execute("write_cells", Ops(ThreeSteps));

            Assert.False(r.Success);
            Assert.Contains("第 2 步", r.Error);
            Assert.Contains("公式写错了", r.Error);
            // the step's own suggestion travels with the batch failure (write_formula derives one from the formula)
            var single = Dispatcher(new FakeExcelActions { WriteFormulaFn = fake.WriteFormulaFn })
                .Execute("write_formula", new Dictionary<string, object> { ["address"] = "B1", ["formula"] = "=SUM(C:C)" });
            Assert.Equal(single.Suggestion, r.Suggestion);
            Assert.Equal(new[] { "backup-1" }, fake.RollbackCalls);
            Assert.DoesNotContain("write_range", fake.Timeline);   // step 3 never ran
        }

        [Fact]
        public void A_step_that_would_overwrite_unread_content_stops_the_batch_before_anything_is_written()
        {
            var fake = new FakeExcelActions { RangeHasContentFn = address => address.StartsWith("C") || address.Contains("!C") };
            var r = Dispatcher(fake).Execute("write_cells", Ops(ThreeSteps));

            Assert.False(r.Success);
            Assert.Contains("第 3 步", r.Error);
            Assert.DoesNotContain(fake.Timeline, t => t == "backup" || t.StartsWith("write_"));
            Assert.Empty(fake.RollbackCalls);
        }

        [Theory]
        [InlineData("[]")]
        [InlineData("[{\"type\":\"explode\",\"address\":\"A1\"}]")]
        [InlineData("[{\"address\":\"A1\"}]")]
        public void Bad_batches_are_refused_without_a_backup(string ops)
        {
            var fake = new FakeExcelActions();
            var r = Dispatcher(fake).Execute("write_cells", Ops(ops));

            Assert.False(r.Success);
            Assert.DoesNotContain("backup", fake.Timeline);
        }

        [Fact]
        public void The_checkpoint_covers_just_the_sheets_the_steps_touch()
        {
            SnapshotScopeSeen = null;
            var fake = new FakeExcelActions
            {
                BackupWorkbookFn = (book, label, scope) =>
                {
                    SnapshotScopeSeen = scope;
                    return new DeepExcel.AddIn.Executor.SnapshotAttempt { SnapshotId = "backup-1" };
                },
            };
            Dispatcher(fake).Execute("write_cells", Ops(
                "[{\"type\":\"add_sheet\",\"name\":\"汇总\"}," +
                " {\"type\":\"values\",\"address\":\"汇总!A1\",\"values\":[[1]]}," +
                " {\"type\":\"text\",\"address\":\"明细!B2\",\"value\":\"x\"}]"));

            Assert.False(SnapshotScopeSeen.WholeWorkbook);
            Assert.Equal(new[] { "汇总", "明细" }, SnapshotScopeSeen.Sheets);
        }

        [Fact]
        public void An_unclear_step_makes_the_checkpoint_cover_the_whole_workbook()
        {
            SnapshotScopeSeen = null;
            var fake = new FakeExcelActions
            {
                BackupWorkbookFn = (book, label, scope) =>
                {
                    SnapshotScopeSeen = scope;
                    return new DeepExcel.AddIn.Executor.SnapshotAttempt { SnapshotId = "backup-1" };
                },
            };
            Dispatcher(fake).Execute("write_cells", Ops(
                "[{\"type\":\"text\",\"address\":\"明细!B2\",\"value\":\"x\"}," +
                " {\"type\":\"values\",\"address\":\"销售区域\",\"values\":[[1]]}]"));

            Assert.True(SnapshotScopeSeen.WholeWorkbook);
        }

        private static DeepExcel.AddIn.Executor.SnapshotScope SnapshotScopeSeen;

        // ---- one combined preview ------------------------------------------------

        [Fact]
        public void The_preview_adds_up_every_step()
        {
            var sheet = new GridSheet().Set("A1", "旧").Set("B1", 5, "=A2*2").Set("C1", 1).Set("C2", 9);
            var preview = new PreviewBuilder(sheet).Build("write_cells", new Dictionary<string, object>
            {
                // The permission request carries nested arguments as JSON text.
                ["ops"] = "[{\"type\":\"text\",\"address\":\"A1\",\"value\":\"新\"}," +
                          " {\"type\":\"values\",\"address\":\"C1\",\"values\":[[1],[2]]}," +
                          " {\"type\":\"style\",\"address\":\"A1\",\"bold\":true}]",
            });

            Assert.True(preview.Previewable, preview.NotPreviewableReason);
            Assert.Equal("write_cells", preview.ToolName);
            Assert.Equal(2, preview.AffectedCells);          // A1 旧→新, C2 9→2; C1 unchanged; style has no cell change
            Assert.Contains(preview.Changes, c => c.Address == "A1");
            Assert.Contains(preview.Changes, c => c.Address == "C2");
        }

        [Fact]
        public void A_step_that_cannot_be_previewed_is_named_in_the_combined_preview()
        {
            var sheet = new GridSheet().Set("A1", "旧");
            var preview = new PreviewBuilder(sheet).Build("write_cells", new Dictionary<string, object>
            {
                ["ops"] = "[{\"type\":\"text\",\"address\":\"A1\",\"value\":\"新\"}," +
                          " {\"type\":\"values\",\"values\":[[1]]}]",
            });

            Assert.False(preview.Previewable);
            Assert.Contains("第 2 步", preview.NotPreviewableReason);
        }

        [Fact]
        public void Write_cells_always_goes_through_the_preview_policy()
        {
            Assert.Equal(ConfirmationLevel.Preview, PreviewPolicy.LevelFor("write_cells"));
        }
    }
}
