using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>read_range 分页：先裁到已用区域，每页有上限，结果里写明下一页怎么读。</summary>
    public class RangePagingTests
    {
        [Fact]
        public void A_whole_column_is_clipped_to_the_used_range()
        {
            var clip = RangePaging.Clip(new SheetBox(1, 1, 1048576, 1), new SheetBox(1, 1, 450, 6));
            Assert.Equal((1, 1, 450, 1), (clip.Value.Row1, clip.Value.Col1, clip.Value.Row2, clip.Value.Col2));
        }

        [Fact]
        public void A_region_outside_the_used_range_is_empty()
        {
            Assert.Null(RangePaging.Clip(new SheetBox(900, 1, 950, 3), new SheetBox(1, 1, 450, 6)));
        }

        [Fact]
        public void The_first_page_is_200_rows_and_points_to_the_next()
        {
            var plan = RangePaging.Plan(450, 6, 0, null);
            Assert.Equal(200, plan.Rows);
            Assert.Equal(200, plan.NextOffset);
            Assert.Contains("下一页：read_range(address=\"Data!A:F\", offset=200)", RangePaging.Hint("Data!A:F", 450, 6, plan));
        }

        [Fact]
        public void The_last_page_has_no_next_offset_and_no_hint()
        {
            var plan = RangePaging.Plan(450, 6, 400, null);
            Assert.Equal(50, plan.Rows);
            Assert.Null(plan.NextOffset);
            Assert.Null(RangePaging.Hint("Data!A:F", 450, 6, plan));
        }

        [Fact]
        public void Wide_regions_get_fewer_rows_per_page_and_say_so()
        {
            var plan = RangePaging.Plan(1000, 400, 0, 500);
            Assert.Equal(RangePaging.MaxColumns, plan.Columns);
            Assert.True(plan.Columns * plan.Rows <= RangePaging.MaxCells);
            Assert.True(plan.ColumnsTruncated);
            Assert.Contains("只返回了前 256 列", RangePaging.Hint("S!A1:OJ1000", 1000, 400, plan));
        }

        [Fact]
        public void Limit_is_capped()
        {
            Assert.Equal(RangePaging.MaxRows, RangePaging.Plan(5000, 2, 0, 100000).Rows);
            Assert.Equal(RangePaging.DefaultRows, RangePaging.Plan(5000, 2, 0, 0).Rows);
        }

        [Fact]
        public void An_offset_past_the_end_is_an_error_that_says_how_many_rows_there_are()
        {
            var plan = RangePaging.Plan(450, 6, 450, null);
            Assert.Contains("一共 450 行", plan.Error);
        }

        [Fact]
        public void The_page_serializes_with_its_paging_block()
        {
            var page = new RangePage
            {
                Address = "$A$1:$B$2", WorksheetName = "S", RowCount = 2, ColumnCount = 2,
                Paging = new PagingInfo { Requested = "S!A:B", TotalRows = 450, TotalColumns = 2, ReturnedRows = 2, NextOffset = 2 },
            };
            var json = PythonSidecar.BuildToolResultJson("c", true, page, null, null, null, null, null);
            using (var doc = JsonDocument.Parse(json))
            {
                var paging = doc.RootElement.GetProperty("data").GetProperty("paging");
                Assert.Equal(450, paging.GetProperty("total_rows").GetInt32());
                Assert.Equal(2, paging.GetProperty("next_offset").GetInt32());
            }
        }

        // ---------------- 接到 ToolDispatcher 上 ----------------

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        [Fact]
        public void Dispatcher_passes_offset_and_limit_through()
        {
            var fake = new FakeExcelActions();
            var d = new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };

            d.Execute("read_range", Args(("address", "A:A")));
            using (var doc = JsonDocument.Parse("{\"offset\": 200, \"limit\": \"50\"}"))
            {
                d.Execute("read_range", new Dictionary<string, object>
                {
                    ["address"] = "A:A",
                    ["offset"] = doc.RootElement.GetProperty("offset").Clone(),
                    ["limit"] = doc.RootElement.GetProperty("limit").Clone(),
                });
            }

            Assert.Equal(("A:A", 0, (int?)null), fake.ReadPageCalls[0]);
            Assert.Equal(("A:A", 200, (int?)50), fake.ReadPageCalls[1]);
        }

        [Fact]
        public void Reading_a_page_marks_only_that_page_as_read()
        {
            var fake = new FakeExcelActions
            {
                RangeHasContentFn = _ => true,
                ReadRangePageFn = (a, o, l) => new RangePage { Address = "$A$1:$A$200", WorksheetName = "Sheet1", RowCount = 200, ColumnCount = 1 },
            };
            var d = new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };

            d.Execute("read_range", Args(("address", "A:A")));

            Assert.True(d.Execute("write_value", Args(("address", "A10"), ("value", "x"))).Success);
            Assert.False(d.Execute("write_value", Args(("address", "A300"), ("value", "x"))).Success);
        }
    }
}
