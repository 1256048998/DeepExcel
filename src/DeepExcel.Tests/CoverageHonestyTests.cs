using System;
using System.Runtime.InteropServices;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Perception;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 覆盖诚实：结构摘要里采样得出的结论标成「样本」；读不到 Excel 时沿用缓存，并说清是多久以前的。
    /// </summary>
    public class CoverageHonestyTests
    {
        private static ColumnProfile Amount() => new ColumnProfile
        {
            Letter = "D",
            Header = "数量",
            Kind = ColumnKind.Number,
            Confidence = 1,
            NonEmptyCount = 500,
            MinDisplay = "1",
            MaxDisplay = "500",
            EmptyCount = 3,
        };

        [Fact]
        public void Sampled_sheets_label_their_statistics_as_samples()
        {
            var column = Amount();
            Assert.Equal("D  数量  number  1 ~ 500  3 个空值", column.Render());
            Assert.Equal("D  数量  number  样本: 1 ~ 500  3 个空值", column.Render(sampled: true));
        }

        [Fact]
        public void The_formula_sample_itself_is_not_labelled()
        {
            var column = new ColumnProfile
            {
                Letter = "F", Header = "金额", Kind = ColumnKind.Formula, Confidence = 1,
                FormulaSample = "=D2*E2", EmptyCount = 2,
            };
            Assert.Equal("F  金额  formula  =D2*E2  样本: 2 个空值", column.Render(sampled: true));
            column.EmptyCount = 0;
            Assert.Equal("F  金额  formula  =D2*E2", column.Render(sampled: true));
        }

        [Fact]
        public void Workbook_render_passes_the_sampling_flag_to_each_column()
        {
            var index = new WorkbookIndex { WorkbookName = "book.xlsx" };
            var sheet = new SheetIndex { Name = "Data", FirstRow = 2, LastRow = 90000, Sampled = true, SampledRows = 2000, HasHeader = true, HeaderRow = 1 };
            sheet.Columns.Add(Amount());
            index.Sheets.Add(sheet);
            var text = index.Render();
            Assert.Contains("采样 2000 行", text);
            Assert.Contains("样本: 1 ~ 500", text);
        }

        [Fact]
        public void A_stale_render_says_how_old_it_is_right_under_the_heading()
        {
            var rendered = "## 工作簿结构：book.xlsx\n\n### Data (2..10 行)\n";
            var stale = IndexCache.MarkStale(rendered, TimeSpan.FromMinutes(12.4), "Excel 正在编辑单元格或有对话框打开");
            var lines = stale.Split('\n');
            Assert.Equal("## 工作簿结构：book.xlsx", lines[0]);
            Assert.Equal("注意：这次读不到 Excel（Excel 正在编辑单元格或有对话框打开），以下沿用 12 分钟前的缓存，" +
                         "之后的修改没有反映；涉及具体数值或结构时请先 read_range 实读。", lines[1]);
            Assert.Contains("### Data", stale);
            Assert.Contains("以下沿用 1 分钟前", IndexCache.MarkStale(rendered, TimeSpan.FromSeconds(5), null));
        }

        [Fact]
        public void Stale_entries_are_available_even_after_invalidation_and_expiry()
        {
            var cache = new IndexCache();
            var built = new DateTime(2026, 9, 25, 8, 0, 0, DateTimeKind.Utc);
            Assert.False(cache.TryGetStale("book", out _, out _));
            cache.Store("book", new WorkbookIndex { WorkbookName = "book.xlsx" }, built);
            cache.Invalidate("book");
            Assert.Null(cache.TryGet("book", built.AddMinutes(1)));
            Assert.True(cache.TryGetStale("book", out var rendered, out var builtUtc));
            Assert.StartsWith("## 工作簿结构：book.xlsx", rendered);
            Assert.Equal(built, builtUtc);
        }

        [Theory]
        [InlineData(unchecked((int)0x800AC472), "Excel 正在编辑单元格或有对话框打开")]
        [InlineData(unchecked((int)0x8001010A), "Excel 正忙")]
        [InlineData(unchecked((int)0x80004005), "Excel 暂时无响应")]
        public void Com_failures_are_described_in_plain_words(int hresult, string text)
        {
            Assert.Equal(text, MessageBridge.DescribeComFailure(new COMException("x", hresult)));
        }
    }
}
