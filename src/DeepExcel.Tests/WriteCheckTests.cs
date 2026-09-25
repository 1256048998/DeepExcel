using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 写后自动体检：模型拿到的是「做完了，这是验证结果」，而不是「调用成功」。
    /// </summary>
    public class WriteCheckTests
    {
        private static HealthSnapshot Health(params string[] errors) => new HealthSnapshot
        {
            ErrorCount = errors.Length,
            Errors = errors.Select(e =>
            {
                var parts = e.Split(' ');
                var addr = parts[0].Split('!');
                return new ErrorCell { Sheet = addr[0], Address = addr[1], Text = parts[1] };
            }).ToList(),
        };

        [Fact]
        public void A_clean_write_passes()
        {
            var v = WriteCheck.Compare(Health(), Health(), null);
            Assert.True(v.Ok);
            Assert.Equal("体检通过：无公式错误，无新增外部链接", v.Summary);
            Assert.Null(v.NewErrors);
        }

        [Fact]
        public void Pre_existing_errors_are_not_blamed_on_the_write()
        {
            var v = WriteCheck.Compare(Health("Sheet1!A1 #N/A"), Health("Sheet1!A1 #N/A"), null);
            Assert.True(v.Ok);
            Assert.Contains("原有 1 个", v.Summary);
        }

        [Fact]
        public void New_errors_are_listed_and_the_model_is_told_to_fix_them()
        {
            var after = Health("Sheet1!A1 #N/A", "Sheet1!D5 #DIV/0!");
            after.Errors.Add(new ErrorCell { Sheet = "My Data", Address = "B2", Text = "#REF!" });
            after.ErrorCount = 3;
            var v = WriteCheck.Compare(Health("Sheet1!A1 #N/A"), after, null);
            Assert.False(v.Ok);
            Assert.Equal(2, v.NewErrors.Count);
            Assert.Contains("Sheet1!D5 #DIV/0!", v.Summary);
            Assert.Contains("'My Data'!B2 #REF!", v.Summary);
            Assert.Contains("不要直接宣布完成", v.Summary);
        }

        [Fact]
        public void Fixed_errors_are_reported_as_progress()
        {
            var v = WriteCheck.Compare(Health("Sheet1!A1 #N/A", "Sheet1!A2 #N/A"), Health(), null);
            Assert.True(v.Ok);
            Assert.Contains("公式错误 2→0", v.Summary);
        }

        [Fact]
        public void A_reference_to_a_missing_sheet_shows_up_as_a_new_external_link()
        {
            var before = Health();
            var after = Health();
            after.ExternalLinks.Add(@"C:\Users\me\Documents\销售明细");
            var v = WriteCheck.Compare(before, after, null);
            Assert.False(v.Ok);
            Assert.Contains("不存在的工作表", v.Summary);
            Assert.Single(v.NewExternalLinks);
        }

        [Fact]
        public void When_the_error_list_was_truncated_the_count_difference_decides()
        {
            var before = Health("Sheet1!A1 #N/A");
            before.ErrorCount = 900; before.Truncated = true;
            var after = Health("Sheet1!A9 #N/A");
            after.ErrorCount = 903; after.Truncated = true;
            var v = WriteCheck.Compare(before, after, null);
            Assert.False(v.Ok);
            Assert.Contains("新增 3 个公式错误", v.Summary);
            // 写入前的名单不全：A9 可能早就在了，不点名
            Assert.Null(v.NewErrors);
        }

        [Fact]
        public void Manual_calculation_is_flagged()
        {
            var after = Health();
            after.CalculationManual = true;
            Assert.Contains("手动", WriteCheck.Compare(Health(), after, null).Summary);
        }

        [Fact]
        public void Cosmetic_tools_and_reads_are_not_checked()
        {
            Assert.False(WriteCheck.NeedsCheck("set_number_format"));
            Assert.False(WriteCheck.NeedsCheck("read_range"));
            Assert.True(WriteCheck.NeedsCheck("write_formula"));
            Assert.True(WriteCheck.NeedsCheck("delete_columns"));
            Assert.True(WriteCheck.NeedsCheck("execute_vba"));
        }

        [Fact]
        public void Serialized_verification_omits_empty_lists()
        {
            var json = JsonSerializer.Serialize(WriteCheck.Compare(Health(), Health(), null));
            Assert.DoesNotContain("new_errors", json);
            Assert.DoesNotContain("samples", json);
            Assert.Contains("\"ok\":true", json);
        }

        // ---------------- 接到 ToolDispatcher 上 ----------------

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        private static ToolDispatcher Dispatcher(FakeExcelActions fake)
            => new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };

        [Fact]
        public void A_formula_write_comes_back_with_the_check_and_computed_samples()
        {
            var calls = 0;
            var fake = new FakeExcelActions
            {
                // 写入前干净，写入后 D2 出现 #DIV/0!
                CaptureHealthFn = () => calls++ == 0 ? Health() : Health("Sheet1!D2 #DIV/0!"),
                SampleCellsFn = (a, n) => new List<CellSample> { new CellSample { Address = "Sheet1!D2", Value = "#DIV/0!", Formula = "=B2/C2" } },
            };
            var d = Dispatcher(fake);

            var r = d.Execute("write_formula", Args(("address", "D2"), ("formula", "=B2/C2")));

            Assert.True(r.Success);
            var v = Assert.IsType<WriteVerification>(r.Verification);
            Assert.False(v.Ok);
            Assert.Contains("Sheet1!D2 #DIV/0!", v.Summary);
            Assert.Equal(new[] { "Sheet1!D2" }, fake.SampleCellsCalls);
            Assert.Equal("=B2/C2", v.Samples[0].Formula);
        }

        [Fact]
        public void Plain_value_writes_are_checked_but_not_read_back()
        {
            var fake = new FakeExcelActions { CaptureHealthFn = () => Health() };
            var d = Dispatcher(fake);

            var r = d.Execute("write_value", Args(("address", "D2"), ("value", "x")));

            Assert.IsType<WriteVerification>(r.Verification);
            Assert.Empty(fake.SampleCellsCalls);
        }

        [Fact]
        public void A_write_range_containing_formulas_is_read_back()
        {
            var fake = new FakeExcelActions { CaptureHealthFn = () => Health() };
            var d = Dispatcher(fake);

            d.Execute("write_range", Args(("address", "A1"), ("values", new object[][] { new object[] { "合计", "=SUM(B2:B9)" } })));

            Assert.Equal(new[] { "Sheet1!A1:B1" }, fake.SampleCellsCalls);
        }

        [Fact]
        public void No_health_means_no_verification_and_the_write_still_succeeds()
        {
            var fake = new FakeExcelActions();  // CaptureHealth 返回 null
            var r = Dispatcher(fake).Execute("write_value", Args(("address", "D2"), ("value", "x")));
            Assert.True(r.Success);
            Assert.Null(r.Verification);
        }

        [Fact]
        public void The_verification_reaches_the_model_in_the_tool_result_json()
        {
            var json = PythonSidecar.BuildToolResultJson("c1", true, null, null, null, null, null, null,
                WriteCheck.Compare(Health(), Health("Sheet1!D2 #DIV/0!"), null));
            using (var doc = JsonDocument.Parse(json))
            {
                var v = doc.RootElement.GetProperty("verification");
                Assert.False(v.GetProperty("ok").GetBoolean());
                Assert.Equal("Sheet1!D2 #DIV/0!", v.GetProperty("new_errors")[0].GetString());
            }
        }
    }
}
