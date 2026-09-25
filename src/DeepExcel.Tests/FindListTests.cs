using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;
using ToolResult = DeepExcel.AddIn.Bridge.ToolResult;

namespace DeepExcel.Tests
{
    /// <summary>find / list：Grep 和 Glob 的工作簿版。</summary>
    public class FindListTests
    {
        private static ToolDispatcher Dispatcher(FakeExcelActions fake)
            => new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };

        private static Dictionary<string, object> Json(string json)
        {
            using (var doc = JsonDocument.Parse(json))
            {
                return doc.RootElement.EnumerateObject().ToDictionary(p => p.Name, p => (object)p.Value.Clone());
            }
        }

        [Fact]
        public void Find_parses_its_arguments()
        {
            var fake = new FakeExcelActions();
            var r = Dispatcher(fake).Execute("find", Json("{\"query\":\"应收\",\"scope\":\"formulas\",\"match\":\"exact\",\"sheets\":[\"Data\",\"汇总\"],\"max_results\":999}"));

            Assert.True(r.Success);
            var call = fake.FindCalls.Single();
            Assert.Equal("应收", call.Query);
            Assert.True(call.InFormulas);
            Assert.True(call.WholeCell);
            Assert.Equal(new[] { "Data", "汇总" }, call.Sheets);
            Assert.Equal(200, call.Max);
        }

        [Fact]
        public void Find_defaults_to_values_contains_across_all_sheets()
        {
            var fake = new FakeExcelActions();
            Dispatcher(fake).Execute("find", Json("{\"query\":\"x\"}"));
            var call = fake.FindCalls.Single();
            Assert.False(call.InFormulas);
            Assert.False(call.WholeCell);
            Assert.Empty(call.Sheets);
            Assert.Equal(50, call.Max);
        }

        [Fact]
        public void Find_without_a_query_fails_cleanly()
        {
            var r = Dispatcher(new FakeExcelActions()).Execute("find", Json("{\"query\":\"\"}"));
            Assert.False(r.Success);
        }

        [Fact]
        public void Find_and_list_never_back_up_the_workbook()
        {
            var fake = new FakeExcelActions();
            var d = Dispatcher(fake);
            d.Execute("find", Json("{\"query\":\"x\"}"));
            d.Execute("list", Json("{\"kind\":\"names\"}"));
            Assert.Empty(fake.BackupCalls);
            Assert.Equal(new[] { "names" }, fake.ListCalls);
        }

        [Fact]
        public void Host_errors_become_failed_results_with_the_suggestion()
        {
            var fake = new FakeExcelActions { ListObjectsFn = k => new { error = "kind 不对", suggestion = "用 tables" } };
            var r = Dispatcher(fake).Execute("list", Json("{\"kind\":\"widgets\"}"));
            Assert.False(r.Success);
            Assert.Equal("kind 不对", r.Error);
            Assert.Equal("用 tables", r.Suggestion);
        }

        [Fact]
        public void Sheet_lists_also_accept_comma_separated_text()
        {
            Assert.Equal(new[] { "A", "B" }, ToolDispatcher.GetStringList(new Dictionary<string, object> { ["s"] = "A, B" }, "s"));
            Assert.Equal(new[] { "A", "B" }, ToolDispatcher.GetStringList(new Dictionary<string, object> { ["s"] = "A，B" }, "s"));
        }
    }

    /// <summary>find / list 在真 Excel 里的行为。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。</summary>
    public class FindListExcelTests : ExcelTestHost
    {
        private static string Json(object data) =>
            PythonSidecar.BuildToolResultJson("c", true, data, null, null, null, null, null);

        [ExcelFact]
        public void Find_searches_values_and_formulas_on_every_sheet()
        {
            var actions = NewBook(out var wb);
            Sheet(wb, "Data").Range["B3"].Value2 = "应收账款";
            Sheet(wb, "Data").Range["B9"].Value2 = "其他应收款";
            Sheet(wb, "Sum").Range["C2"].Formula = "=SUMIF(Data!B:B,\"应收账款\",Data!C:C)";

            var values = JsonDocument.Parse(Json(actions.FindCells("应收", false, null, false, 50))).RootElement.GetProperty("data");
            Assert.Equal(2, values.GetProperty("total").GetInt32());
            var addresses = values.GetProperty("matches").EnumerateArray().Select(m => m.GetProperty("sheet").GetString() + "!" + m.GetProperty("address").GetString()).ToList();
            Assert.Contains("Data!B3", addresses);
            Assert.Contains("Data!B9", addresses);

            var exact = JsonDocument.Parse(Json(actions.FindCells("应收账款", false, null, true, 50))).RootElement.GetProperty("data");
            Assert.Equal(1, exact.GetProperty("total").GetInt32());

            var formulas = JsonDocument.Parse(Json(actions.FindCells("SUMIF", true, null, false, 50))).RootElement.GetProperty("data");
            Assert.Equal("Sum", formulas.GetProperty("matches")[0].GetProperty("sheet").GetString());
            Assert.Contains("SUMIF", formulas.GetProperty("matches")[0].GetProperty("formula").GetString());

            var onlySum = JsonDocument.Parse(Json(actions.FindCells("应收", false, new List<string> { "Sum" }, false, 50))).RootElement.GetProperty("data");
            Assert.Equal(0, onlySum.GetProperty("total").GetInt32());
        }

        [ExcelFact]
        public void Find_caps_the_listed_matches_but_counts_all_of_them()
        {
            var actions = NewBook(out var wb);
            Sheet(wb, "Data").Range["A1:A30"].Value2 = "重复";
            var data = JsonDocument.Parse(Json(actions.FindCells("重复", false, null, false, 5))).RootElement.GetProperty("data");
            Assert.Equal(30, data.GetProperty("total").GetInt32());
            Assert.Equal(5, data.GetProperty("returned").GetInt32());
            Assert.True(data.GetProperty("truncated").GetBoolean());
        }

        [ExcelFact]
        public void List_reports_sheets_names_tables_and_charts()
        {
            var actions = NewBook(out var wb);
            var data = Sheet(wb, "Data");
            data.Range["A1:C1"].Value2 = new object[,] { { "日期", "客户", "金额" } };
            data.Range["A2:C4"].Value2 = new object[,] { { 1, "甲", 10 }, { 2, "乙", 20 }, { 3, "丙", 30 } };
            data.ListObjects.Add(XlListObjectSourceType.xlSrcRange, data.Range["A1:C4"], XlListObjectHasHeaders: XlYesNoGuess.xlYes).Name = "销售表";
            wb.Names.Add("税率", "=Sum!$B$1");
            Sheet(wb, "Sum").Visible = XlSheetVisibility.xlSheetHidden;
            var chart = ((ChartObjects)data.ChartObjects()).Add(200, 10, 300, 200);
            chart.Chart.SetSourceData(data.Range["B1:C4"]);

            var sheets = JsonDocument.Parse(Json(actions.ListObjects("sheets"))).RootElement.GetProperty("data").GetProperty("items");
            Assert.Equal("hidden", sheets.EnumerateArray().Single(s => s.GetProperty("name").GetString() == "Sum").GetProperty("visibility").GetString());

            var names = JsonDocument.Parse(Json(actions.ListObjects("names"))).RootElement.GetProperty("data").GetProperty("items");
            Assert.Contains(names.EnumerateArray(), n => n.GetProperty("name").GetString() == "税率");

            var tables = JsonDocument.Parse(Json(actions.ListObjects("tables"))).RootElement.GetProperty("data").GetProperty("items");
            var table = tables.EnumerateArray().Single();
            Assert.Equal("销售表", table.GetProperty("name").GetString());
            Assert.Equal(3, table.GetProperty("rows").GetInt32());
            Assert.Equal("金额", table.GetProperty("columns")[2].GetString());

            var charts = JsonDocument.Parse(Json(actions.ListObjects("charts"))).RootElement.GetProperty("data").GetProperty("items");
            Assert.Equal("Data", charts.EnumerateArray().Single().GetProperty("sheet").GetString());

            Assert.NotNull(actions.ListObjects("widgets").GetType().GetProperty("error"));
        }
    }
}
