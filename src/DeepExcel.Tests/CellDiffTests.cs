using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>内联 diff：这一步实际改了哪些格，原值 → 新值。</summary>
    public class CellDiffTests
    {
        private static readonly CellRect B2 = new CellRect("Sheet1", 2, 2, 3, 3);  // B2:C3

        [Fact]
        public void Lists_only_cells_whose_content_changed()
        {
            var before = new object[,] { { "a", 1.0 }, { null, "=A1" } };
            var after = new object[,] { { "a", 2.0 }, { "新", "=A1" } };

            var diff = CellDiff.Compare(B2, before, after);

            Assert.Equal(2, diff.Changed);
            Assert.Equal(4, diff.Cells);
            Assert.Equal(new[] { ("C2", "1", "2"), ("B3", "", "新") },
                diff.Samples.Select(s => (s.Address, s.Before, s.After)));
        }

        [Fact]
        public void Formulas_are_compared_as_formulas()
        {
            var diff = CellDiff.Compare(new CellRect("S", 1, 4, 1, 4), new object[,] { { "" } }, new object[,] { { "=B1/C1" } });
            Assert.Equal("D1", diff.Samples.Single().Address);
            Assert.Equal("=B1/C1", diff.Samples.Single().After);
        }

        [Fact]
        public void Samples_are_capped_but_the_count_is_not()
        {
            var before = new object[100, 1];
            var after = new object[100, 1];
            for (var i = 0; i < 100; i++) after[i, 0] = "x";

            var diff = CellDiff.Compare(new CellRect("S", 1, 1, 100, 1), before, after);

            Assert.Equal(100, diff.Changed);
            Assert.Equal(CellDiff.MaxSamples, diff.Samples.Count);
        }

        [Fact]
        public void Long_text_is_clipped()
        {
            var diff = CellDiff.Compare(new CellRect("S", 1, 1, 1, 1), new object[,] { { "" } }, new object[,] { { new string('长', 200) } });
            Assert.True(diff.Samples[0].After.Length <= 81);
        }

        [Fact]
        public void Nothing_to_compare_gives_no_diff()
        {
            Assert.Null(CellDiff.Compare(B2, null, new object[,] { { 1 } }));
        }

        // ---------------- 接到 ToolDispatcher 上 ----------------

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        [Fact]
        public void A_targeted_write_comes_back_with_what_it_changed()
        {
            var reads = 0;
            var fake = new FakeExcelActions
            {
                ReadFormulasFn = a => reads++ == 0 ? new object[,] { { "旧" } } : new object[,] { { "新" } },
            };
            var d = new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };

            var r = d.Execute("write_value", Args(("address", "B2"), ("value", "新")));

            var changes = Assert.IsType<CellChanges>(r.Changes);
            Assert.Equal(1, changes.Changed);
            Assert.Equal(("B2", "旧", "新"), (changes.Samples[0].Address, changes.Samples[0].Before, changes.Samples[0].After));
        }

        [Fact]
        public void Writes_without_a_known_target_have_no_diff()
        {
            var fake = new FakeExcelActions { ReadFormulasFn = a => new object[,] { { "x" } } };
            var d = new ToolDispatcher(fake, null) { BoundWorkbookKey = () => fake.ActiveWorkbookKey };
            Assert.Null(d.Execute("execute_vba", Args(("code", "Sub A()\nEnd Sub"), ("macro_name", "A"))).Changes);
            Assert.Null(d.Execute("set_number_format", Args(("address", "A1"), ("format", "0.00"))).Changes);
        }

        [Fact]
        public void The_diff_reaches_the_sidecar()
        {
            var diff = CellDiff.Compare(B2, new object[,] { { 1.0 } }, new object[,] { { 2.0 } });
            var json = PythonSidecar.BuildToolResultJson("c1", true, null, null, null, null, null, null, null, null, diff);
            using (var doc = JsonDocument.Parse(json))
            {
                var changes = doc.RootElement.GetProperty("changes");
                Assert.Equal(1, changes.GetProperty("changed").GetInt32());
                Assert.Equal("B2", changes.GetProperty("samples")[0].GetProperty("address").GetString());
            }
        }
    }
}
