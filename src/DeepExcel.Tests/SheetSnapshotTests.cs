using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>sheet_snapshot 的纯逻辑：窗口、值编码、错误码、溢出识别。</summary>
    public class SnapshotEncoderTests
    {
        [Fact]
        public void Window_caps_columns_then_rows_by_cell_budget()
        {
            Assert.Equal((450, 3), SnapshotEncoder.PlanWindow(450, 3, 0));
            Assert.Equal((600, 100), SnapshotEncoder.PlanWindow(90000, 300, 0));
            Assert.Equal((1000, 100), SnapshotEncoder.PlanWindow(90000, 300, 999999));  // 硬上限 10 万格
            Assert.Equal((0, 0), SnapshotEncoder.PlanWindow(0, 5, 0));
        }

        [Fact]
        public void Values_are_encoded_by_their_com_type()
        {
            Assert.Null(SnapshotEncoder.EncodeValue(null));
            Assert.Null(SnapshotEncoder.EncodeValue(""));
            Assert.Equal(3.5, SnapshotEncoder.EncodeValue(3.5));
            Assert.Equal(true, SnapshotEncoder.EncodeValue(true));
            var date = Assert.IsType<Dictionary<string, string>>(SnapshotEncoder.EncodeValue(new DateTime(2024, 1, 31)));
            Assert.Equal("2024-01-31", date["d"]);
            var stamp = Assert.IsType<Dictionary<string, string>>(SnapshotEncoder.EncodeValue(new DateTime(2024, 1, 31, 9, 30, 0)));
            Assert.Equal("2024-01-31 09:30", stamp["d"]);
            var error = Assert.IsType<Dictionary<string, string>>(SnapshotEncoder.EncodeValue(-2146826281));
            Assert.Equal("#DIV/0!", error["e"]);
            Assert.Equal(61, ((string)SnapshotEncoder.EncodeValue(new string('x', 100))).Length);
        }

        [Theory]
        [InlineData(-2146826288, "#NULL!")]
        [InlineData(-2146826246, "#N/A")]
        [InlineData(-2146826265, "#REF!")]
        [InlineData(-2146826243, "#SPILL!")]
        [InlineData(-2146826238, "#CALC!")]
        [InlineData(12345, "#ERROR")]
        public void Error_codes_map_to_their_text(int code, string text)
        {
            Assert.Equal(text, SnapshotEncoder.ErrorText(code));
        }

        [Fact]
        public void Only_dynamic_array_formulas_are_asked_about_spills()
        {
            Assert.True(SnapshotEncoder.MaySpill("=FILTER(R2C1:R9C3,R2C2:R9C2>0)"));
            Assert.True(SnapshotEncoder.MaySpill("=_xlfn.UNIQUE(C[-1])"));
            Assert.False(SnapshotEncoder.MaySpill("=RC[-2]*RC[-1]"));
            Assert.False(SnapshotEncoder.MaySpill("=SORTED_NAME+1"));
        }

        [Fact]
        public void Dispatcher_passes_sheet_and_budget_and_never_backs_up()
        {
            var fake = new FakeExcelActions();
            var d = new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };
            var r = d.Execute("sheet_snapshot", new Dictionary<string, object> { ["sheet"] = "Data", ["max_cells"] = 5000 });
            Assert.True(r.Success);
            Assert.Equal(("Data", 5000), fake.SnapshotCalls.Single());
            Assert.Empty(fake.BackupCalls);
            Assert.False(ToolMutationPolicy.IsMutating("sheet_snapshot"));
        }
    }

    /// <summary>sheet_snapshot 在真 Excel 里读出来的东西。默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。</summary>
    public class SheetSnapshotExcelTests : ExcelTestHost
    {
        private static JsonElement Json(object data)
        {
            var text = PythonSidecar.BuildToolResultJson("c", true, data, null, null, null, null, null);
            return JsonDocument.Parse(text).RootElement.GetProperty("data");
        }

        /// <summary>
        /// 标题（合并）+ 两级表头（合并）+ 日期 + 公式列 + 错误值 + 合计行 + 表格。
        /// 快照同时写到 %TEMP%\deepexcel_snapshot_payroll.json，侧车测试的 fixture 就是它。
        /// </summary>
        [ExcelFact]
        public void Payroll_sheet_snapshot_has_values_formulas_merges_and_objects()
        {
            var actions = NewBook(out var wb);
            var ws = Sheet(wb, "Data");
            ws.Range["A1"].Value2 = "2024年9月工资表";
            ws.Range["A1:F1"].Merge();
            ws.Range["A2"].Value2 = "姓名";
            ws.Range["B2"].Value2 = "入职日期";
            ws.Range["C2"].Value2 = "扣款";
            ws.Range["F2"].Value2 = "实发";
            ws.Range["C3"].Value2 = "养老";
            ws.Range["D3"].Value2 = "医疗";
            ws.Range["E3"].Value2 = "小计";
            ws.Range["A2:A3"].Merge();
            ws.Range["B2:B3"].Merge();
            ws.Range["C2:E2"].Merge();
            ws.Range["F2:F3"].Merge();
            for (var i = 0; i < 6; i++)
            {
                var row = 4 + i;
                ws.Cells[row, 1] = "员工" + (i + 1);
                ws.Cells[row, 2] = new DateTime(2020 + i, 3, 1);
                ws.Cells[row, 3] = 100 + i;
                ws.Cells[row, 4] = 50 + i;
                ((Range)ws.Cells[row, 5]).FormulaR1C1 = "=RC[-2]+RC[-1]";
                ((Range)ws.Cells[row, 6]).FormulaR1C1 = "=8000-RC[-1]";
            }
            ws.Range["E7"].Value2 = 999;                       // 被改成死值
            ws.Range["F8"].Formula = "=1/0";                   // 错误值
            ws.Range["A10"].Value2 = "合计";
            ws.Range["C10"].FormulaR1C1 = "=SUM(R[-6]C:R[-2]C)";  // 漏了第 9 行
            ws.Range["D10"].FormulaR1C1 = "=SUM(R[-6]C:R[-1]C)";

            var other = Sheet(wb, "Sum");
            other.Range["A1:B1"].Value2 = new object[,] { { "k", "v" } };
            other.Range["A2:B3"].Value2 = new object[,] { { "a", 1 }, { "b", 2 } };
            other.ListObjects.Add(XlListObjectSourceType.xlSrcRange, other.Range["A1:B3"], XlListObjectHasHeaders: XlYesNoGuess.xlYes).Name = "Lookup";

            var data = Json(actions.SheetSnapshot("data", 0));  // 表名不区分大小写
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "deepexcel_snapshot_payroll.json"),
                data.GetRawText(), new System.Text.UTF8Encoding(false));

            Assert.Equal("Data", data.GetProperty("sheet").GetString());
            Assert.Equal("A1:F10", data.GetProperty("used").GetString());
            Assert.False(data.GetProperty("truncated").GetBoolean());
            var cells = data.GetProperty("cells");
            Assert.Equal(10, cells.GetArrayLength());
            Assert.Equal("2024年9月工资表", cells[0][0].GetString());
            Assert.Equal(JsonValueKind.Null, cells[0][1].ValueKind);
            Assert.Equal("2020-03-01", cells[3][1].GetProperty("d").GetString());
            Assert.Equal(150, cells[3][4].GetDouble());
            Assert.Equal("#DIV/0!", cells[7][5].GetProperty("e").GetString());

            var formulas = data.GetProperty("formulas").EnumerateArray()
                .ToDictionary(f => (f[0].GetInt32(), f[1].GetInt32()), f => f[2].GetString());
            Assert.Equal("=RC[-2]+RC[-1]", formulas[(3, 4)]);
            Assert.False(formulas.ContainsKey((6, 4)));        // 死值那格没有公式
            Assert.Equal("=SUM(R[-6]C:R[-2]C)", formulas[(9, 2)]);

            var merges = data.GetProperty("merges").EnumerateArray()
                .Select(m => string.Join(",", m.EnumerateArray().Select(v => v.GetInt32()))).ToList();
            Assert.Contains("0,0,0,5", merges);
            Assert.Contains("1,0,2,0", merges);
            Assert.Contains("1,2,1,4", merges);
            Assert.Equal(5, merges.Count);

            Assert.Empty(data.GetProperty("objects").GetProperty("tables").EnumerateArray());
            var sum = Json(actions.SheetSnapshot("Sum", 0));
            Assert.Equal("Lookup", sum.GetProperty("objects").GetProperty("tables")[0].GetProperty("name").GetString());
        }

        [ExcelFact]
        public void Big_sheets_are_read_up_to_the_budget_and_marked_truncated()
        {
            var actions = NewBook(out var wb);
            Sheet(wb, "Data").Range["A1:C3000"].Value2 = 1;
            var data = Json(actions.SheetSnapshot("Data", 3000));
            Assert.True(data.GetProperty("truncated").GetBoolean());
            Assert.Equal(1000, data.GetProperty("cells").GetArrayLength());
            Assert.Equal(3000, data.GetProperty("total_rows").GetInt32());
        }

        [ExcelFact]
        public void Empty_sheets_and_missing_sheets()
        {
            var actions = NewBook(out var wb);
            var empty = Json(actions.SheetSnapshot("Sum", 0));
            Assert.Equal(0, empty.GetProperty("cells").GetArrayLength());
            Assert.Equal(0, empty.GetProperty("total_rows").GetInt32());

            var missing = actions.SheetSnapshot("NoSuchSheet", 0);
            Assert.NotNull(missing.GetType().GetProperty("error"));
        }
    }
}
