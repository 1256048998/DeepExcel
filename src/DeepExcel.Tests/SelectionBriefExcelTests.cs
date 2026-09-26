using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>选区条的选区描述在真 Excel 里的行为。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。</summary>
    public class SelectionBriefExcelTests : ExcelTestHost
    {
        private static JsonElement Brief(Range range) =>
            JsonDocument.Parse(JsonSerializer.Serialize(MessageBridge.DescribeSelection(range))).RootElement;

        [ExcelFact]
        public void Describes_a_block_without_reading_cells()
        {
            NewBook(out var wb);
            var sheet = Sheet(wb, "Data");
            var brief = Brief(sheet.Range["A1:D20"]);
            Assert.Equal("Data", brief.GetProperty("sheet").GetString());
            Assert.Equal("A1:D20", brief.GetProperty("address").GetString());
            Assert.Equal(20, brief.GetProperty("rows").GetInt64());
            Assert.Equal(4, brief.GetProperty("cols").GetInt64());
            Assert.Equal(80, brief.GetProperty("cells").GetInt64());
        }

        [ExcelFact]
        public void Multi_area_and_whole_column_selections()
        {
            NewBook(out var wb);
            var sheet = Sheet(wb, "Data");
            var multi = Brief(sheet.Range["A1:B2,D5"]);
            Assert.Equal("A1:B2,D5", multi.GetProperty("address").GetString());
            Assert.Equal(5, multi.GetProperty("cells").GetInt64());
            // 整列：1048576 格，超过 int 也不能溢出
            var column = Brief(sheet.Range["C:D"]);
            Assert.Equal(2L * 1048576, column.GetProperty("cells").GetInt64());
        }

        [Fact]
        public void No_range_means_no_selection()
        {
            var brief = JsonDocument.Parse(JsonSerializer.Serialize(MessageBridge.DescribeSelection(null))).RootElement;
            Assert.Equal("", brief.GetProperty("address").GetString());
        }
    }
}
