using System;
using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Preview;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Pre-execution change preview.
    ///
    /// The product currently executes first and offers a rollback afterwards,
    /// which is why people will not point it at a workbook that matters. These
    /// tests cover the two things that decide whether the preview is
    /// trustworthy: that it reports exactly what will change, and that it admits
    /// when it cannot know.
    /// </summary>
    public class PreviewTests
    {
        private sealed class FakeSheet : IRangeReader
        {
            private readonly Dictionary<string, CellSnapshot> _cells =
                new Dictionary<string, CellSnapshot>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, string[]> _ranges =
                new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

            public int UsedRows { get; set; } = 100;

            public FakeSheet Cell(string address, string display, string formula = null)
            {
                _cells[address] = new CellSnapshot { Address = address, Display = display, Formula = formula };
                return this;
            }

            /// <summary>Declares which addresses a range expands to.</summary>
            public FakeSheet Range(string address, params string[] cellAddresses)
            {
                _ranges[address] = cellAddresses;
                return this;
            }

            public IReadOnlyList<CellSnapshot> Read(string sheet, string address)
            {
                if (_ranges.TryGetValue(address, out var members))
                {
                    return members.Select(Get).ToList();
                }
                return new[] { Get(address) };
            }

            private CellSnapshot Get(string address)
            {
                return _cells.TryGetValue(address, out var cell)
                    ? cell
                    : new CellSnapshot { Address = address, Display = "", Formula = null };
            }

            public int UsedRowCount(string sheet) => UsedRows;
        }

        private static Dictionary<string, object> Args(params object[] pairs)
        {
            var args = new Dictionary<string, object>();
            for (var i = 0; i + 1 < pairs.Length; i += 2)
            {
                args[(string)pairs[i]] = pairs[i + 1];
            }
            return args;
        }

        // ------------------------------------------------------------------
        // Policy
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("read_range", ConfirmationLevel.None)]
        [InlineData("read_workbook", ConfirmationLevel.None)]
        [InlineData("create_chart", ConfirmationLevel.None)]
        [InlineData("write_range", ConfirmationLevel.Preview)]
        [InlineData("delete_rows", ConfirmationLevel.Preview)]
        [InlineData("clear_range", ConfirmationLevel.Preview)]
        [InlineData("execute_vba", ConfirmationLevel.SnapshotAndWarn)]
        [InlineData("execute_python", ConfirmationLevel.SnapshotAndWarn)]
        public void OperationsGetTheRightConfirmationLevel(string tool, ConfirmationLevel expected)
        {
            Assert.Equal(expected, PreviewPolicy.LevelFor(tool));
        }

        [Fact]
        public void ReadOnlyToolsNeverPrompt()
        {
            // Prompting for reads would be noise, and noise is what makes people
            // stop reading prompts at all.
            foreach (var tool in new[] { "read_range", "read_selection", "list_snapshots", "read_attachment" })
            {
                Assert.Equal(ConfirmationLevel.None, PreviewPolicy.LevelFor(tool));
            }
        }

        [Fact]
        public void CodeExecutionIsNeverClaimedToBePreviewable()
        {
            // Pretending to have simulated VBA would be worse than admitting we
            // cannot: the user would trust a preview that does not cover what runs.
            var preview = new PreviewBuilder(new FakeSheet())
                .Build("execute_vba", Args("code", "Sub Foo()\nCells.Clear\nEnd Sub"));

            Assert.False(preview.Previewable);
            Assert.Contains("VBA", preview.NotPreviewableReason);
            Assert.Contains("快照", preview.Render());
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        // ------------------------------------------------------------------
        // Diff computation
        // ------------------------------------------------------------------

        [Fact]
        public void OverwritingAValueIsReportedWithBothSides()
        {
            var sheet = new FakeSheet().Cell("C2", "1200");
            var preview = new PreviewBuilder(sheet)
                .Build("write_value", Args("address", "C2", "value", "1500"));

            Assert.True(preview.Previewable);
            Assert.Equal(1, preview.AffectedCells);
            var change = preview.Changes.Single();
            Assert.Equal("1200", change.Before);
            Assert.Equal("1500", change.After);
            Assert.Equal(ChangeKind.Overwrite, change.Kind);
        }

        [Fact]
        public void WritingIntoAnEmptyCellIsAnAddNotAnOverwrite()
        {
            var preview = new PreviewBuilder(new FakeSheet())
                .Build("write_value", Args("address", "C3", "value", "0"));

            Assert.Equal(ChangeKind.Add, preview.Changes.Single().Kind);
            Assert.Equal("", preview.Changes.Single().Before);
        }

        [Fact]
        public void WritingTheSameValueIsNotCountedAsAChange()
        {
            // Counting no-ops would inflate the number the user is approving,
            // which is the number they use to decide.
            var sheet = new FakeSheet().Cell("C2", "1200");
            var preview = new PreviewBuilder(sheet)
                .Build("write_value", Args("address", "C2", "value", "1200"));

            Assert.Equal(0, preview.AffectedCells);
            Assert.True(preview.IsNoOp);
            // A no-op is not worth interrupting for, but it is worth reporting.
            Assert.False(PreviewPolicy.ShouldInterrupt(preview));
            Assert.Contains("不会改变任何内容", preview.Summary());
        }

        [Fact]
        public void OverwritingAFormulaIsCalledOutSpecifically()
        {
            // The edit users least expect and most regret: the number looks
            // right immediately and stops updating forever.
            var sheet = new FakeSheet().Cell("F2", "125", "=D2*E2");
            var preview = new PreviewBuilder(sheet)
                .Build("write_value", Args("address", "F2", "value", "125"));

            Assert.Equal(1, preview.FormulasOverwritten);
            Assert.True(preview.Changes.Single().OverwritesFormula);
            // The formula text, not its current result, is what is being lost.
            Assert.Equal("=D2*E2", preview.Changes.Single().Before);
            Assert.Contains(preview.Warnings, w => w.Contains("公式"));
        }

        [Fact]
        public void FormulaOverwriteAlwaysInterruptsEvenForOneCell()
        {
            var sheet = new FakeSheet().Cell("F2", "125", "=D2*E2");
            var preview = new PreviewBuilder(sheet)
                .Build("write_value", Args("address", "F2", "value", "999"));

            Assert.Equal(1, preview.AffectedCells);
            Assert.True(preview.AffectedCells < PreviewPolicy.SmallChangeThreshold);
            // Small but expensive to discover later.
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void SmallOrdinaryEditsDoNotInterrupt()
        {
            var sheet = new FakeSheet().Cell("A1", "x");
            var preview = new PreviewBuilder(sheet)
                .Build("write_value", Args("address", "A1", "value", "y"));

            Assert.False(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void RangeWriteDiffsEachCell()
        {
            var sheet = new FakeSheet()
                .Range("A1:A3", "A1", "A2", "A3")
                .Cell("A1", "old1")
                .Cell("A2", "same")
                .Cell("A3", "", "=B3");

            var preview = new PreviewBuilder(sheet).Build("write_range",
                Args("address", "A1:A3", "values", new List<object> { "new1", "same", "literal" }));

            Assert.Equal(2, preview.AffectedCells); // A2 unchanged
            Assert.Equal(1, preview.FormulasOverwritten);
            Assert.DoesNotContain(preview.Changes, c => c.Address == "A2");
        }

        [Fact]
        public void NarrowerDataDoesNotClaimToClearTheRemainingCells()
        {
            // Writing 2 values into a 4-cell range leaves the tail alone;
            // reporting those as changed would overstate the blast radius.
            var sheet = new FakeSheet()
                .Range("A1:A4", "A1", "A2", "A3", "A4")
                .Cell("A3", "keep me")
                .Cell("A4", "keep me too");

            var preview = new PreviewBuilder(sheet).Build("write_range",
                Args("address", "A1:A4", "values", new List<object> { "x", "y" }));

            Assert.Equal(2, preview.AffectedCells);
            Assert.DoesNotContain(preview.Changes, c => c.Address == "A3" || c.Address == "A4");
        }

        [Fact]
        public void NestedRowArraysAreFlattenedInRowMajorOrder()
        {
            var values = new List<object>
            {
                new List<object> { "a", "b" },
                new List<object> { "c", "d" }
            };
            var flat = PreviewBuilder.FlattenValues(Args("values", values));
            Assert.Equal(new[] { "a", "b", "c", "d" }, flat.ToArray());
        }

        [Fact]
        public void ClearingReportsOnlyCellsThatActuallyHoldSomething()
        {
            var sheet = new FakeSheet()
                .Range("A1:A3", "A1", "A2", "A3")
                .Cell("A1", "value")
                .Cell("A3", "10", "=SUM(B:B)");

            var preview = new PreviewBuilder(sheet).Build("clear_range", Args("address", "A1:A3"));

            Assert.Equal(2, preview.AffectedCells); // A2 was already empty
            Assert.All(preview.Changes, c => Assert.Equal(ChangeKind.Clear, c.Kind));
            Assert.Equal(1, preview.FormulasOverwritten);
        }

        // ------------------------------------------------------------------
        // Structural changes
        // ------------------------------------------------------------------

        [Fact]
        public void DeletingRowsIsStructuralAndAlwaysInterrupts()
        {
            var preview = new PreviewBuilder(new FakeSheet())
                .Build("delete_rows", Args("start_row", 5, "count", 3));

            Assert.True(preview.IsStructural);
            Assert.Equal(new[] { 5, 6, 7 }, preview.DeletedRows.ToArray());
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
            // Shifting rows silently breaks formulas that referenced them.
            Assert.Contains(preview.Warnings, w => w.Contains("上移"));
        }

        [Fact]
        public void DeletingBeyondTheDataIsFlaggedAsALikelyMiscount()
        {
            var sheet = new FakeSheet { UsedRows = 20 };
            var preview = new PreviewBuilder(sheet)
                .Build("delete_rows", Args("start_row", 18, "count", 10));

            Assert.Contains(preview.Warnings, w => w.Contains("超出数据区"));
        }

        [Fact]
        public void DeletingColumnsListsThem()
        {
            var preview = new PreviewBuilder(new FakeSheet())
                .Build("delete_columns", Args("columns", "C, E , G"));

            Assert.Equal(new[] { "C", "E", "G" }, preview.DeletedColumns.ToArray());
            Assert.True(preview.IsStructural);
        }

        [Fact]
        public void InvalidParametersDoNotProduceAFalseEmptyPreview()
        {
            // An empty change set here would read as "this is safe".
            var builder = new PreviewBuilder(new FakeSheet());
            Assert.False(builder.Build("delete_rows", Args("start_row", 0)).Previewable);
            Assert.False(builder.Build("write_value", Args("value", "x")).Previewable);
            Assert.False(builder.Build("delete_columns", Args()).Previewable);
        }

        [Fact]
        public void UnimplementedPreviewsAdmitItRatherThanReportingNoChanges()
        {
            // sort_data is simulable in principle but not implemented. Claiming
            // an empty change set would be a lie the user would act on.
            var preview = new PreviewBuilder(new FakeSheet())
                .Build("sort_data", Args("address", "A1:F100"));

            Assert.False(preview.Previewable);
            Assert.True(PreviewPolicy.ShouldInterrupt(preview));
        }

        [Fact]
        public void ToolsWithoutAPolicyReturnNoPreviewAtAll()
        {
            Assert.Null(new PreviewBuilder(new FakeSheet()).Build("read_range", Args("address", "A1")));
        }

        // ------------------------------------------------------------------
        // Rendering
        // ------------------------------------------------------------------

        [Fact]
        public void LargeChangeSetsAreTruncatedButTheCountIsHonest()
        {
            var sheet = new FakeSheet();
            var addresses = new List<string>();
            for (var row = 1; row <= 300; row++)
            {
                var address = "A" + row;
                addresses.Add(address);
                sheet.Cell(address, "old" + row);
            }
            sheet.Range("A1:A300", addresses.ToArray());

            var values = Enumerable.Range(1, 300).Select(i => (object)("new" + i)).ToList();
            var preview = new PreviewBuilder(sheet)
                .Build("write_range", Args("address", "A1:A300", "values", values));

            // The list is capped so the panel stays usable...
            Assert.Equal(ChangePreview.MaxListedChanges, preview.Changes.Count);
            // ...but the number the user approves is the real one.
            Assert.Equal(300, preview.AffectedCells);
            Assert.True(preview.IsTruncated);
            Assert.Contains("300", preview.Summary());
            Assert.Contains("其余", preview.Render());
        }

        [Fact]
        public void RenderedPreviewLeadsWithTheSummary()
        {
            var sheet = new FakeSheet().Cell("C2", "1200");
            var preview = new PreviewBuilder(sheet)
                .Build("write_value", Args("address", "C2", "value", "1500"));

            var lines = preview.Render().Split('\n');
            Assert.Contains("1 个单元格", lines[0]);
            Assert.Contains("C2", preview.Render());
        }
    }
}
