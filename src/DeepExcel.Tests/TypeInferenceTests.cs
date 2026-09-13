using System;
using System.Linq;
using DeepExcel.AddIn.Perception;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Column type inference.
    ///
    /// This is the highest-leverage change in the roadmap and also the most
    /// dangerous: a wrong type stated confidently is worse than no index at all,
    /// because the model acts on it. So the cases here are the awkward ones real
    /// workbooks contain -- currency text, totals rows, identifiers that look
    /// like dates -- not the clean ones.
    /// </summary>
    public class TypeInferenceTests
    {
        private static object[][] Rows(params object[][] rows) => rows;

        // ------------------------------------------------------------------
        // Value classification
        // ------------------------------------------------------------------

        [Theory]
        [InlineData(1200.0, ColumnKind.Number)]
        [InlineData(42, ColumnKind.Number)]
        [InlineData("1200", ColumnKind.Number)]
        [InlineData("1,200.50", ColumnKind.Number)]
        [InlineData("¥1,200.00", ColumnKind.Number)]
        [InlineData("$99.95", ColumnKind.Number)]
        [InlineData("(1,200)", ColumnKind.Number)]
        [InlineData("15%", ColumnKind.Number)]
        [InlineData("华东贸易", ColumnKind.Text)]
        [InlineData("SO-2026-0001", ColumnKind.Text)]
        [InlineData("2026-01-15", ColumnKind.Date)]
        [InlineData("2026/1/15", ColumnKind.Date)]
        [InlineData("TRUE", ColumnKind.Boolean)]
        [InlineData("是", ColumnKind.Boolean)]
        [InlineData("=A1*B1", ColumnKind.Formula)]
        [InlineData("", ColumnKind.Empty)]
        [InlineData("   ", ColumnKind.Empty)]
        [InlineData(null, ColumnKind.Empty)]
        public void ValuesAreClassified(object value, ColumnKind expected)
        {
            Assert.Equal(expected, TypeInference.ClassifyValue(value));
        }

        [Fact]
        public void CurrencyFormattedTextIsANumberNotText()
        {
            // Reporting a money column as text would send the model down a
            // string-parsing path for what is arithmetic.
            Assert.Equal(1200.0, TypeInference.ToNumber("¥1,200.00"));
            Assert.Equal(-1200.0, TypeInference.ToNumber("(1,200.00)"));
            Assert.Equal(0.15, TypeInference.ToNumber("15%"));
        }

        [Fact]
        public void EightDigitIdentifiersAreNotDates()
        {
            // "20260101" parses as yyyyMMdd, but an order number of the same
            // shape is far more common. Misreading it would make the model try
            // date arithmetic on an identifier.
            Assert.Null(TypeInference.ToDate("20260101"));
            Assert.Equal(ColumnKind.Number, TypeInference.ClassifyValue("20260101"));

            // A separator makes the intent unambiguous.
            Assert.NotNull(TypeInference.ToDate("2026-01-01"));
        }

        [Theory]
        [InlineData(1, "A")]
        [InlineData(26, "Z")]
        [InlineData(27, "AA")]
        [InlineData(52, "AZ")]
        [InlineData(53, "BA")]
        [InlineData(702, "ZZ")]
        [InlineData(703, "AAA")]
        public void ColumnLettersMatchExcel(int index, string expected)
        {
            Assert.Equal(expected, TypeInference.ColumnLetter(index));
        }

        // ------------------------------------------------------------------
        // Header detection
        // ------------------------------------------------------------------

        [Fact]
        public void TextRowAboveDataIsAHeader()
        {
            var rows = Rows(
                new object[] { "订单号", "日期", "金额" },
                new object[] { "SO-1", "2026-01-01", 100.0 },
                new object[] { "SO-2", "2026-01-02", 200.0 });
            Assert.True(TypeInference.LooksLikeHeaderRow(rows));
        }

        [Fact]
        public void AllTextTableDoesNotLoseItsFirstDataRow()
        {
            // Every row is text, so "first row is text" proves nothing. Treating
            // it as a header would silently drop a record.
            var rows = Rows(
                new object[] { "华东", "张三" },
                new object[] { "华南", "李四" },
                new object[] { "华北", "王五" });
            Assert.False(TypeInference.LooksLikeHeaderRow(rows));
        }

        [Fact]
        public void NumericFirstRowIsNotAHeader()
        {
            var rows = Rows(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 4.0, 5.0, 6.0 });
            Assert.False(TypeInference.LooksLikeHeaderRow(rows));
        }

        [Fact]
        public void SingleCellTitleRowIsNotTreatedAsAHeader()
        {
            // A merged report title above the real header.
            var rows = Rows(
                new object[] { "2026 年销售报表", null, null },
                new object[] { "订单号", "日期", "金额" },
                new object[] { "SO-1", "2026-01-01", 100.0 });
            Assert.False(TypeInference.LooksLikeHeaderRow(rows));
        }

        [Fact]
        public void TooFewRowsMeansNoHeaderGuess()
        {
            Assert.False(TypeInference.LooksLikeHeaderRow(Rows(new object[] { "a", "b" })));
            Assert.False(TypeInference.LooksLikeHeaderRow(null));
        }

        // ------------------------------------------------------------------
        // Column profiling
        // ------------------------------------------------------------------

        [Fact]
        public void ProfilesATypicalSalesTable()
        {
            var rows = Rows(
                new object[] { "SO-1", DateTime.Parse("2026-01-01"), "华东贸易", 10.0 },
                new object[] { "SO-2", DateTime.Parse("2026-03-15"), "华南商贸", 25.0 },
                new object[] { "SO-3", DateTime.Parse("2026-09-10"), "华东贸易", 5.0 });

            var profiles = TypeInference.ProfileColumns(
                rows, firstRowNumber: 2, headers: new[] { "订单号", "日期", "客户", "数量" });

            Assert.Equal(4, profiles.Count);

            Assert.Equal(ColumnKind.Text, profiles[0].Kind);
            Assert.True(profiles[0].IsUnique);

            Assert.Equal(ColumnKind.Date, profiles[1].Kind);
            Assert.Equal("2026-01-01", profiles[1].MinDisplay);
            Assert.Equal("2026-09-10", profiles[1].MaxDisplay);

            Assert.Equal(ColumnKind.Text, profiles[2].Kind);
            Assert.Equal(2, profiles[2].DistinctCount);
            Assert.False(profiles[2].IsUnique);

            Assert.Equal(ColumnKind.Number, profiles[3].Kind);
            Assert.Equal("5", profiles[3].MinDisplay);
            Assert.Equal("25", profiles[3].MaxDisplay);
        }

        [Fact]
        public void BlankCellsAreCountedAndLocated()
        {
            // "3 个空值(行 88,402,771)" is directly actionable; "has gaps" is not.
            var rows = Rows(
                new object[] { 1.0 },
                new object[] { null },
                new object[] { 3.0 },
                new object[] { "" });

            var profile = TypeInference.ProfileColumns(rows, firstRowNumber: 10)[0];

            Assert.Equal(2, profile.EmptyCount);
            Assert.Equal(new[] { 11, 13 }, profile.EmptyRows.ToArray());
            Assert.Equal(2, profile.NonEmptyCount);
        }

        [Fact]
        public void AStrayTotalsRowLowersConfidenceInsteadOfFlippingTheType()
        {
            // A numeric column with one "合计" cell is still a numeric column,
            // but the model should be warned rather than misled.
            var rows = Rows(
                new object[] { 10.0 }, new object[] { 20.0 }, new object[] { 30.0 },
                new object[] { 40.0 }, new object[] { 50.0 }, new object[] { 60.0 },
                new object[] { 70.0 }, new object[] { 80.0 }, new object[] { 90.0 },
                new object[] { "合计" });

            var profile = TypeInference.ProfileColumns(rows)[0];

            Assert.Equal(ColumnKind.Number, profile.Kind);
            Assert.Equal(0.9, profile.Confidence, 3);
            Assert.False(profile.IsLowConfidence); // exactly at the threshold
        }

        [Fact]
        public void GenuinelyMixedColumnsAreFlaggedForRealReading()
        {
            var rows = Rows(
                new object[] { 10.0 }, new object[] { "北京" }, new object[] { DateTime.Now },
                new object[] { true }, new object[] { "上海" });

            var profile = TypeInference.ProfileColumns(rows)[0];

            Assert.True(profile.IsLowConfidence);
            Assert.Contains("⚠", profile.Render());
            Assert.Contains("实读", profile.Render());
        }

        [Fact]
        public void NoTypeMajorityBecomesMixed()
        {
            var rows = Rows(new object[] { 1.0 }, new object[] { "a" }, new object[] { DateTime.Now });
            Assert.Equal(ColumnKind.Mixed, TypeInference.ProfileColumns(rows)[0].Kind);
        }

        [Fact]
        public void FormulaColumnsReportTheFormulaNotTheValues()
        {
            var rows = Rows(new object[] { 100.0 }, new object[] { 200.0 });
            var profile = TypeInference.ProfileColumns(
                rows, formulas: new[] { "=D2*E2" })[0];

            Assert.Equal(ColumnKind.Formula, profile.Kind);
            Assert.Equal("=D2*E2", profile.FormulaSample);
            // Knowing it is computed is what stops the model overwriting it.
            Assert.Contains("=D2*E2", profile.Render());
        }

        [Fact]
        public void EmptyColumnIsReportedAsEmpty()
        {
            var rows = Rows(new object[] { null }, new object[] { "" }, new object[] { "  " });
            var profile = TypeInference.ProfileColumns(rows)[0];

            Assert.Equal(ColumnKind.Empty, profile.Kind);
            Assert.Equal(0, profile.NonEmptyCount);
            Assert.False(profile.IsLowConfidence);
        }

        [Fact]
        public void RaggedRowsDoNotThrow()
        {
            // Real sheets have rows of different lengths.
            var rows = Rows(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 4.0 },
                null,
                new object[] { 5.0, 6.0 });

            var profiles = TypeInference.ProfileColumns(rows);
            Assert.Equal(3, profiles.Count);
            Assert.Equal(ColumnKind.Number, profiles[0].Kind);
        }

        [Fact]
        public void LongTextIsTruncatedSoOneCellCannotEatTheBudget()
        {
            var rows = Rows(new object[] { new string('长', 500) });
            var profile = TypeInference.ProfileColumns(rows)[0];
            Assert.All(profile.Samples, s => Assert.True(s.Length <= 25, "sample length " + s.Length));
        }

        [Fact]
        public void DistinctCountIsCappedSoAHugeColumnCannotStall()
        {
            var rows = Enumerable.Range(0, 2000)
                .Select(i => new object[] { "value-" + i })
                .ToArray();

            var profile = TypeInference.ProfileColumns(rows)[0];

            Assert.True(profile.DistinctCapped);
            Assert.Equal(ColumnProfile.DistinctCap, profile.DistinctCount);
            // Capped means the count is a floor, so uniqueness is unknown and
            // must not be asserted.
            Assert.False(profile.IsUnique);
        }

        // ------------------------------------------------------------------
        // Rendering / token cost
        // ------------------------------------------------------------------

        [Fact]
        public void RenderedIndexIsSmallEnoughToBeWorthIt()
        {
            var index = new WorkbookIndex
            {
                WorkbookName = "销售.xlsx",
                Sheets =
                {
                    new SheetIndex
                    {
                        Name = "销售明细",
                        FirstRow = 1,
                        LastRow = 8420,
                        HasHeader = true,
                        HeaderRow = 1,
                        Sampled = true,
                        SampledRows = 600,
                        Columns = TypeInference.ProfileColumns(
                            Rows(
                                new object[] { "SO-1", DateTime.Parse("2026-01-01"), "华东贸易", 10.0, 12.5 },
                                new object[] { "SO-2", DateTime.Parse("2026-09-10"), "华南商贸", 25.0, 9.0 }),
                            firstRowNumber: 2,
                            headers: new[] { "订单号", "日期", "客户", "数量", "单价" })
                    }
                }
            };

            var text = index.Render();

            Assert.Contains("销售明细", text);
            Assert.Contains("8420", text);
            Assert.Contains("采样 600", text);
            Assert.Contains("read_range", text); // tells the model how to get real values
            // The whole premise is that this is cheaper than reading rows.
            Assert.True(index.EstimateTokens() < 600, "index cost " + index.EstimateTokens() + " tokens");
        }

        [Fact]
        public void RenderRespectsItsCharacterBudget()
        {
            var index = new WorkbookIndex { WorkbookName = "big.xlsx" };
            for (var i = 0; i < 200; i++)
            {
                index.Sheets.Add(new SheetIndex
                {
                    Name = "Sheet" + i,
                    LastRow = 1000,
                    Columns = TypeInference.ProfileColumns(
                        Rows(new object[] { "a", 1.0 }), headers: new[] { "列一", "列二" })
                });
            }

            var text = index.Render(maxCharacters: 2000);

            // A 200-sheet workbook must not crowd out the conversation.
            Assert.True(text.Length < 4000, "rendered " + text.Length + " characters");
            Assert.Contains("已省略", text);
        }

        [Fact]
        public void IncompleteIndexTellsTheModelToReadInstead()
        {
            var index = new WorkbookIndex { WorkbookName = "x.xlsx", Incomplete = "超时" };
            var text = index.Render();

            // Silently shipping a partial picture is the failure mode worth
            // avoiding: the model would treat absence as emptiness.
            Assert.Contains("索引不完整", text);
            Assert.Contains("实读", text);
        }
    }
}
