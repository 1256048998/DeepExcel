using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace DeepExcel.Tests
{
    public class StarterOutlineTests
    {
        [Fact]
        public void Cells_become_json_friendly_values()
        {
            Assert.Null(StarterOutline.ToCell(null));
            Assert.Null(StarterOutline.ToCell(""));
            Assert.Equal(42.0, StarterOutline.ToCell(42.0));
            Assert.Equal("#N/A", StarterOutline.ToCell(-2146826246));
            Assert.Equal("#DIV/0!", StarterOutline.ToCell(-2146826281));
            Assert.Equal(true, StarterOutline.ToCell(true));
        }

        [Fact]
        public void Sheet_names_never_collide_and_stay_within_31_chars()
        {
            Assert.Equal("示例-销售明细", StarterOutline.UniqueSheetName("示例-销售明细", new[] { "Sheet1" }));
            Assert.Equal("示例-销售明细(2)", StarterOutline.UniqueSheetName("示例-销售明细", new[] { "示例-销售明细" }));
            Assert.Equal("示例-销售明细(3)", StarterOutline.UniqueSheetName("示例-销售明细", new[] { "示例-销售明细", "示例-销售明细(2)" }));
            var longName = new string('表', 40);
            var unique = StarterOutline.UniqueSheetName(longName, new[] { new string('表', 31) });
            Assert.True(unique.Length <= 31);
            Assert.EndsWith("(2)", unique);
            Assert.Equal("ab", StarterOutline.UniqueSheetName("a/b", new string[0]));
        }

        [Fact]
        public void Rows_are_parsed_with_formulas_neutralised_and_limits_enforced()
        {
            var rows = JsonDocument.Parse("[[\"日期\",\"金额\"],[45474,\"'1,280\"],[45475,\"=SUM(A1)\"],[null,true]]").RootElement;
            var data = StarterOutline.ParseRows(rows, out var error);
            Assert.Null(error);
            Assert.Equal(45474.0, data[1, 0]);
            Assert.Equal("'1,280", data[1, 1]);
            Assert.Equal("'=SUM(A1)", data[2, 1]);
            Assert.Null(data[3, 0]);

            var tooWide = JsonDocument.Parse("[[" + string.Join(",", Enumerable.Range(0, 31)) + "]]").RootElement;
            Assert.Null(StarterOutline.ParseRows(tooWide, out error));
            Assert.NotNull(error);
            Assert.Null(StarterOutline.ParseRows(JsonDocument.Parse("[]").RootElement, out _));
        }
    }

    /// <summary>示例写入 + 结构读取在真 Excel 里的行为。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。</summary>
    public class StarterExcelTests : ExcelTestHost
    {
        [ExcelFact]
        public void Sample_goes_into_a_new_sheet_and_reads_back_with_its_planted_problems()
        {
            NewBook(out var wb);
            Sheet(wb, "Data").Range["A1"].Value2 = "用户的数据";
            var rows = JsonDocument.Parse(
                "[[\"订单日期\",\"区域\",\"金额\"],[45474,\"华东\",100],[45475,\"华南\",\"'1,280\"],[45476,\"\",300]]").RootElement;
            var data = StarterOutline.ParseRows(rows, out _);

            var first = StarterOutline.InsertSample(wb, "示例-销售明细", data, new Dictionary<int, string> { [0] = "yyyy-mm-dd" });
            var second = StarterOutline.InsertSample(wb, "示例-销售明细", data, null);

            Assert.Equal("示例-销售明细", first);
            Assert.Equal("示例-销售明细(2)", second);
            // 用户的表没动，示例排在最后
            Assert.Equal("用户的数据", Sheet(wb, "Data").Range["A1"].Value2);
            Assert.Equal(second, ((Worksheet)wb.Worksheets[wb.Worksheets.Count]).Name);
            var sample = Sheet(wb, first);
            Assert.Equal("1,280", sample.Range["C3"].Value2);   // 撇号让它留在文本
            Assert.Null(sample.Range["B4"].Value2);

            var outline = JsonDocument.Parse(JsonSerializer.Serialize(StarterOutline.Describe(wb))).RootElement;
            var sheet = outline.GetProperty("sheets").EnumerateArray().Single(s => s.GetProperty("name").GetString() == first);
            Assert.Equal(4, sheet.GetProperty("rows").GetInt32());
            Assert.True(sheet.GetProperty("date_columns")[0].GetBoolean());
            Assert.False(sheet.GetProperty("date_columns")[2].GetBoolean());
            Assert.Equal("1,280", sheet.GetProperty("grid")[2][2].GetString());
            Assert.Equal(45474, sheet.GetProperty("grid")[1][0].GetDouble());
        }

        [ExcelFact]
        public void Protected_structure_refuses_the_sample()
        {
            NewBook(out var wb);
            wb.Protect(Structure: true);
            var data = StarterOutline.ParseRows(JsonDocument.Parse("[[\"a\"]]").RootElement, out _);
            var ex = Assert.Throws<InvalidOperationException>(() => StarterOutline.InsertSample(wb, "示例", data, null));
            Assert.Contains("受保护", ex.Message);
        }
    }
}
