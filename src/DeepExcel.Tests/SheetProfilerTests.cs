using System;
using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Perception;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The parts of sheet profiling that do not need Excel: sampling strategy,
    /// OLE date conversion, and number-format detection.
    ///
    /// These are worth isolating because both have failure modes that look
    /// correct. An OLE serial profiles as a plausible number, and a sampling bug
    /// produces an index that is merely incomplete rather than obviously broken.
    /// </summary>
    public class SheetProfilerTests
    {
        // ------------------------------------------------------------------
        // Sampling
        // ------------------------------------------------------------------

        [Fact]
        public void SmallSheetsAreReadInFull()
        {
            var blocks = SheetProfiler.PlanSample(firstRow: 1, rowCount: 40);
            Assert.Single(blocks);
            Assert.Equal(1, blocks[0].Item1);
            Assert.Equal(40, blocks[0].Item2);
        }

        [Fact]
        public void MediumSheetsAreStillReadInFull()
        {
            // Below the sampling threshold there is no reason to approximate.
            var blocks = SheetProfiler.PlanSample(firstRow: 1, rowCount: 500);
            Assert.Equal(500, blocks.Sum(b => b.Item2));
        }

        [Fact]
        public void LargeSheetsAreSampledWithinBudget()
        {
            var blocks = SheetProfiler.PlanSample(firstRow: 1, rowCount: 8420);
            var sampled = blocks.Sum(b => b.Item2);

            // The entire premise is not reading 8420 rows.
            Assert.True(sampled < 8420, "sampled " + sampled);
            Assert.True(sampled <= SheetProfiler.HeadRows + SheetProfiler.BodySampleRows + 25,
                "sampled " + sampled);
            // The head must always be covered: that is where structure lives.
            Assert.Equal(1, blocks[0].Item1);
            Assert.Equal(SheetProfiler.HeadRows, blocks[0].Item2);
        }

        [Fact]
        public void SamplingIsDeterministic()
        {
            // A random sample would make an index differ between two runs on an
            // unchanged sheet, which makes any bug report unreproducible.
            var first = SheetProfiler.PlanSample(1, 8420);
            var second = SheetProfiler.PlanSample(1, 8420);
            Assert.Equal(first.Select(b => b.Item1), second.Select(b => b.Item1));
            Assert.Equal(first.Select(b => b.Item2), second.Select(b => b.Item2));
        }

        [Fact]
        public void SampleBlocksStayInsideTheSheet()
        {
            const int firstRow = 5;
            const int rowCount = 9000;
            var blocks = SheetProfiler.PlanSample(firstRow, rowCount);
            var lastAllowed = firstRow + rowCount - 1;

            foreach (var block in blocks)
            {
                Assert.True(block.Item1 >= firstRow, "block starts before the sheet");
                Assert.True(block.Item2 > 0, "empty block");
                Assert.True(block.Item1 + block.Item2 - 1 <= lastAllowed,
                    $"block {block.Item1}+{block.Item2} runs past row {lastAllowed}");
            }
        }

        [Fact]
        public void SampleBlocksDoNotOverlap()
        {
            var blocks = SheetProfiler.PlanSample(1, 8420)
                .OrderBy(b => b.Item1)
                .ToList();

            for (var i = 1; i < blocks.Count; i++)
            {
                var previousEnd = blocks[i - 1].Item1 + blocks[i - 1].Item2 - 1;
                Assert.True(blocks[i].Item1 > previousEnd,
                    $"block at {blocks[i].Item1} overlaps previous ending {previousEnd}");
            }
        }

        [Fact]
        public void EmptySheetPlansNothing()
        {
            Assert.Empty(SheetProfiler.PlanSample(1, 0));
            Assert.Empty(SheetProfiler.PlanSample(1, -5));
        }

        // ------------------------------------------------------------------
        // Number formats
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("yyyy-mm-dd", true)]
        [InlineData("yyyy/m/d", true)]
        [InlineData("m/d/yyyy", true)]
        [InlineData("[$-409]mmm-yy;@", true)]
        [InlineData("yyyy\"年\"m\"月\"d\"日\"", true)]
        [InlineData("General", false)]
        [InlineData("0.00", false)]
        [InlineData("#,##0.00", false)]
        [InlineData("0.00%", false)]
        [InlineData("@", false)]
        [InlineData("", false)]
        [InlineData(null, false)]
        public void DateFormatsAreDistinguishedFromNumberFormats(string format, bool expected)
        {
            Assert.Equal(expected, SheetProfiler.LooksLikeDateFormat(format));
        }

        [Fact]
        public void CurrencyFormatWithLiteralTextIsNotADate()
        {
            // The literal "d" inside quotes must not read as a day placeholder.
            Assert.False(SheetProfiler.LooksLikeDateFormat("\"USD\"#,##0.00"));
            Assert.False(SheetProfiler.LooksLikeDateFormat("[$¥-804]#,##0.00"));
        }

        [Fact]
        public void TimeOnlyFormatIsNotTreatedAsADate()
        {
            // Excel stores times as serials too, but a time column is not a date
            // column and converting it would produce 1899-12-30 everywhere.
            Assert.False(SheetProfiler.LooksLikeDateFormat("h:mm"));
            Assert.False(SheetProfiler.LooksLikeDateFormat("hh:mm:ss"));
        }

        // ------------------------------------------------------------------
        // OLE date conversion
        // ------------------------------------------------------------------

        [Fact]
        public void OleSerialsBecomeDatesInDateFormattedColumns()
        {
            var expected = new DateTime(2026, 1, 15);
            var rows = new List<object[]>
            {
                new object[] { "SO-1", expected.ToOADate(), 100.0 },
                new object[] { "SO-2", new DateTime(2026, 9, 10).ToOADate(), 200.0 }
            };

            SheetProfiler.ConvertOleDates(rows, new[] { "@", "yyyy-mm-dd", "#,##0.00" });

            Assert.IsType<DateTime>(rows[0][1]);
            Assert.Equal(expected, (DateTime)rows[0][1]);
            // The genuinely numeric column must be left alone.
            Assert.IsType<double>(rows[0][2]);
            Assert.Equal(100.0, rows[0][2]);
        }

        [Fact]
        public void ConvertedDatesThenProfileAsDates()
        {
            // The whole point of the conversion: without it this column would be
            // reported as "number 46037 ~ 46275", which looks plausible enough
            // that nothing would flag it.
            var rows = new List<object[]>
            {
                new object[] { new DateTime(2026, 1, 1).ToOADate() },
                new object[] { new DateTime(2026, 9, 10).ToOADate() }
            };

            var beforeConversion = TypeInference.ProfileColumns(rows.ToArray())[0];
            Assert.Equal(ColumnKind.Number, beforeConversion.Kind);

            SheetProfiler.ConvertOleDates(rows, new[] { "yyyy-mm-dd" });
            var afterConversion = TypeInference.ProfileColumns(rows.ToArray())[0];

            Assert.Equal(ColumnKind.Date, afterConversion.Kind);
            Assert.Equal("2026-01-01", afterConversion.MinDisplay);
            Assert.Equal("2026-09-10", afterConversion.MaxDisplay);
        }

        [Fact]
        public void ValuesOutsideExcelsDateRangeAreNotConverted()
        {
            var rows = new List<object[]>
            {
                new object[] { 0.0 },
                new object[] { -5.0 },
                new object[] { 9999999.0 }
            };

            SheetProfiler.ConvertOleDates(rows, new[] { "yyyy-mm-dd" });

            Assert.All(rows, row => Assert.IsType<double>(row[0]));
        }

        [Fact]
        public void ConversionToleratesRaggedRowsAndNulls()
        {
            var rows = new List<object[]>
            {
                new object[] { new DateTime(2026, 1, 1).ToOADate() },
                null,
                new object[] { },
                new object[] { "not a number" }
            };

            SheetProfiler.ConvertOleDates(rows, new[] { "yyyy-mm-dd", "General" });

            Assert.IsType<DateTime>(rows[0][0]);
            Assert.Equal("not a number", rows[3][0]);
        }

        [Fact]
        public void NullFormatsAreANoOp()
        {
            var rows = new List<object[]> { new object[] { 45000.0 } };
            SheetProfiler.ConvertOleDates(rows, null);
            Assert.IsType<double>(rows[0][0]);
        }
    }
}
