using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 先读后写 + 读后被改检测（Claude Code 的「编辑前必须先 Read」「读取后文件被改过」）。
    /// </summary>
    public class ReadLedgerTests
    {
        // ---------------- 地址解析 ----------------

        [Theory]
        [InlineData("A1", "Sheet1", "Sheet1", 1, 1, 1, 1)]
        [InlineData("$B$2:$D$9", "Sheet1", "Sheet1", 2, 2, 9, 4)]
        [InlineData("Data!C3:A1", "Sheet1", "Data", 1, 1, 3, 3)]
        [InlineData("'My Sheet'!AA10", "Sheet1", "My Sheet", 10, 27, 10, 27)]
        [InlineData("'O''Brien'!A1", "Sheet1", "O'Brien", 1, 1, 1, 1)]
        [InlineData("B:C", "Sheet1", "Sheet1", 1, 2, CellRect.MaxRow, 3)]
        [InlineData("2:5", "Sheet1", "Sheet1", 2, 1, 5, CellRect.MaxColumn)]
        public void Parses_a1_addresses(string address, string defaultSheet, string sheet, int r1, int c1, int r2, int c2)
        {
            Assert.True(CellRect.TryParse(address, defaultSheet, out var rect));
            Assert.Equal(sheet, rect.Sheet);
            Assert.Equal((r1, c1, r2, c2), (rect.Row1, rect.Col1, rect.Row2, rect.Col2));
        }

        [Theory]
        [InlineData("")]
        [InlineData("SalesTotal")]
        [InlineData("A1,B2")]
        [InlineData("R1C1")]
        public void Does_not_guess_at_names_or_multi_areas(string address)
        {
            Assert.False(CellRect.TryParse(address, "Sheet1", out _));
        }

        [Fact]
        public void Formats_back_to_a1()
        {
            Assert.Equal("Sheet1!B2:D9", new CellRect("Sheet1", 2, 2, 9, 4).ToA1());
            Assert.Equal("'My Sheet'!AA10", new CellRect("My Sheet", 10, 27, 10, 27).ToA1());
        }

        // ---------------- 账本 ----------------

        private static CellRect R(string a) { CellRect.TryParse(a, "Sheet1", out var r); return r; }

        [Fact]
        public void Writing_into_empty_cells_needs_no_read()
        {
            var ledger = new ReadLedger();
            Assert.Equal(LedgerVerdict.Allowed, ledger.Check(R("D2"), () => false, out _));
        }

        [Fact]
        public void Overwriting_unseen_content_is_refused()
        {
            var ledger = new ReadLedger();
            Assert.Equal(LedgerVerdict.NotRead, ledger.Check(R("A1:C4"), () => true, out _));
        }

        [Fact]
        public void Reading_any_overlapping_part_is_enough()
        {
            var ledger = new ReadLedger();
            ledger.RecordRead(R("A1:C200"));
            Assert.Equal(LedgerVerdict.Allowed, ledger.Check(R("C150:C5000"), () => true, out _));
            // 别的表上同样的地址不算读过
            Assert.Equal(LedgerVerdict.NotRead, ledger.Check(R("Other!A1"), () => true, out _));
        }

        [Fact]
        public void A_user_edit_after_the_read_makes_it_stale_until_re_read()
        {
            var ledger = new ReadLedger();
            ledger.RecordRead(R("A1:C4"));
            ledger.RecordUserEdit(R("B3"));
            Assert.Equal(LedgerVerdict.StaleRead, ledger.Check(R("A1:C4"), () => true, out var changed));
            Assert.Equal(new[] { "Sheet1!B3" }, changed);

            ledger.RecordRead(R("A1:C4"));
            Assert.Equal(LedgerVerdict.Allowed, ledger.Check(R("A1:C4"), () => true, out _));
        }

        [Fact]
        public void A_user_edit_elsewhere_does_not_block()
        {
            var ledger = new ReadLedger();
            ledger.RecordRead(R("A1:C4"));
            ledger.RecordUserEdit(R("Z99"));
            Assert.Equal(LedgerVerdict.Allowed, ledger.Check(R("A1:C4"), () => true, out _));
        }

        [Fact]
        public void The_models_own_write_counts_as_knowing_the_content()
        {
            var ledger = new ReadLedger();
            ledger.RecordOwnWrite(R("D2:D10"));
            Assert.Equal(LedgerVerdict.Allowed, ledger.Check(R("D2"), () => true, out _));
        }

        [Fact]
        public void User_edits_are_reported_once()
        {
            var ledger = new ReadLedger();
            ledger.RecordUserEdit(R("B3"));
            ledger.RecordUserEdit(R("Data!A1:A5"));
            Assert.Equal(new[] { "Sheet1!B3", "Data!A1:A5" }, ledger.TakeUnreportedUserEdits());
            Assert.Empty(ledger.TakeUnreportedUserEdits());
        }

        [Fact]
        public void Forgetting_reads_requires_reading_again()
        {
            var ledger = new ReadLedger();
            ledger.RecordRead(R("A1:C4"));
            ledger.ForgetReads();
            Assert.Equal(LedgerVerdict.NotRead, ledger.Check(R("A1"), () => true, out _));
        }

        // ---------------- 接到 ToolDispatcher 上 ----------------

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        private static ToolDispatcher Dispatcher(FakeExcelActions fake)
            => new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };

        private static FakeExcelActions BookWithData() => new FakeExcelActions
        {
            RangeHasContentFn = _ => true,
            ReadRangeFn = a => new RangeInfo { Address = "$A$1:$C$4", WorksheetName = "Sheet1", RowCount = 4, ColumnCount = 3 },
        };

        [Fact]
        public void Dispatcher_refuses_to_overwrite_content_it_never_read()
        {
            var fake = BookWithData();
            var d = Dispatcher(fake);

            var r = d.Execute("write_value", Args(("address", "B2"), ("value", "x")));

            Assert.False(r.Success);
            Assert.Contains("还没有读过", r.Error);
            Assert.Contains("read_range", r.Suggestion);
            Assert.DoesNotContain("write_value", fake.Timeline);
            Assert.Empty(fake.BackupCalls);  // 拒绝发生在备份之前
        }

        [Fact]
        public void Dispatcher_allows_the_write_after_a_read()
        {
            var fake = BookWithData();
            var d = Dispatcher(fake);

            Assert.True(d.Execute("read_range", Args(("address", "A1:C4"))).Success);
            var r = d.Execute("write_value", Args(("address", "B2"), ("value", "x")));

            Assert.True(r.Success);
            Assert.Contains("write_value", fake.Timeline);
        }

        [Fact]
        public void Dispatcher_refuses_a_write_based_on_a_stale_read_and_tells_the_model()
        {
            var fake = BookWithData();
            var d = Dispatcher(fake);
            d.Execute("read_range", Args(("address", "A1:C4")));

            // 用户在 Excel 里改了 B3（不在工具执行期间）
            d.Ledger.RecordUserEdit(new CellRect("Sheet1", 3, 2, 3, 2));
            var r = d.Execute("write_range", Args(("address", "A2"), ("values", new object[][] { new object[] { 1, 2, 3 }, new object[] { 4, 5, 6 } })));

            Assert.False(r.Success);
            Assert.Contains("Sheet1!B3", r.Error);
            Assert.Contains("用户手动改动", r.Error);
            Assert.DoesNotContain("write_range", fake.Timeline);
        }

        [Fact]
        public void User_edits_are_mentioned_in_the_next_tool_result()
        {
            var fake = BookWithData();
            var d = Dispatcher(fake);
            d.Ledger.RecordUserEdit(new CellRect("Sheet1", 7, 1, 7, 1));

            var r = d.Execute("read_workbook", Args());

            Assert.Contains("Sheet1!A7", r.Warning);
            Assert.Null(d.Execute("read_workbook", Args()).Warning);
        }

        [Fact]
        public void Notices_such_as_a_panel_rollback_reach_the_model_once()
        {
            var fake = BookWithData();
            var d = Dispatcher(fake);
            d.Ledger.AddNotice("用户在面板上把工作簿回退到了 10:02:03 的检查点");

            Assert.Contains("10:02:03 的检查点", d.Execute("read_workbook", Args()).Warning);
            Assert.Null(d.Execute("read_workbook", Args()).Warning);
        }

        [Fact]
        public void Formatting_does_not_require_a_read()
        {
            var fake = BookWithData();
            var d = Dispatcher(fake);
            Assert.True(d.Execute("set_number_format", Args(("address", "A1:C4"), ("format", "0.00"))).Success);
            Assert.True(d.Execute("clear_range", Args(("address", "A1:C4"), ("clear_type", "formats"))).Success);
        }

        [Fact]
        public void Write_range_target_covers_the_whole_array()
        {
            var d = Dispatcher(new FakeExcelActions());
            Assert.True(d.TryResolveWriteTarget("write_range",
                Args(("address", "Data!B2"), ("values", new object[][] { new object[] { 1, 2, 3 }, new object[] { 4, 5 } })),
                out var rect));
            Assert.Equal("Data!B2:D3", rect.ToA1());
        }
    }
}
