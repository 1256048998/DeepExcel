using System;
using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Skills;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Turning one successful run into a reusable skill.
    ///
    /// The judgement being tested is which recorded values are incidental to
    /// that run and should become parameters. Both failure directions make the
    /// skill useless: parameterise too much and the user re-answers everything
    /// they already said; too little and it only works on the original data.
    /// </summary>
    public class SkillTests
    {
        private static RecordedStep Step(string tool, params object[] pairs)
        {
            var step = new RecordedStep { Tool = tool };
            for (var i = 0; i + 1 < pairs.Length; i += 2)
            {
                step.Arguments[(string)pairs[i]] = pairs[i + 1];
            }
            return step;
        }

        // ------------------------------------------------------------------
        // Classification
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("address", "A1:F200", ParameterKind.Range)]
        [InlineData("address", "A1", ParameterKind.Range)]
        [InlineData("range", "Sheet1!A1:B2", ParameterKind.Range)]
        [InlineData("range", "'My Sheet'!$A$1:$F$200", ParameterKind.Range)]
        [InlineData("address", "A:A", ParameterKind.Range)]
        [InlineData("sheet", "销售明细", ParameterKind.Sheet)]
        [InlineData("sheet_name", "Sheet2", ParameterKind.Sheet)]
        [InlineData("start_date", "2026-01-01", ParameterKind.Date)]
        [InlineData("month", "2026/09", ParameterKind.Date)]
        [InlineData("path", @"C:\Users\alice\data.xlsx", ParameterKind.File)]
        [InlineData("threshold", "1000", ParameterKind.Number)]
        [InlineData("top_n", "10", ParameterKind.Number)]
        public void IncidentalValuesBecomeParameters(string argument, string value, ParameterKind expected)
        {
            Assert.Equal(expected, SkillParameterizer.Classify(argument, value));
        }

        [Theory]
        [InlineData("ascending", "true")]
        [InlineData("has_header", "true")]
        [InlineData("chart_type", "column")]
        [InlineData("direction", "desc")]
        [InlineData("operation", "sum")]
        public void TaskDecisionsStayFixed(string argument, string value)
        {
            // These are what the user already decided. Asking again every run
            // would make the skill slower than just describing the task.
            Assert.Null(SkillParameterizer.Classify(argument, value));
        }

        [Fact]
        public void CodeIsNeverParameterized()
        {
            // A VBA body with a placeholder spliced in is a template that can
            // silently produce different code than the one that was approved.
            Assert.Null(SkillParameterizer.Classify("code", "Sub Foo()\nRange(\"A1\").Value = 1\nEnd Sub"));
            Assert.Null(SkillParameterizer.Classify("script", "import pandas"));
        }

        [Fact]
        public void RowCountsAndIndexesStayFixed()
        {
            // Part of the recorded shape, not something the user chooses.
            Assert.Null(SkillParameterizer.Classify("count", "3"));
            Assert.Null(SkillParameterizer.Classify("start_row", "5"));
            Assert.Null(SkillParameterizer.Classify("index", "2"));
        }

        [Fact]
        public void EightDigitIdentifiersAreNotDates()
        {
            // Turning an order number into a date parameter would prompt for the
            // wrong thing on every run.
            Assert.False(SkillParameterizer.LooksLikeDate("20260101"));
            Assert.True(SkillParameterizer.LooksLikeDate("2026-01-01"));
        }

        [Fact]
        public void PlainTextArgumentsStayFixed()
        {
            Assert.Null(SkillParameterizer.Classify("title", "月度销售汇总"));
            Assert.Null(SkillParameterizer.Classify("formula", "=SUM(A1:A10)"));
        }

        // ------------------------------------------------------------------
        // Building
        // ------------------------------------------------------------------

        [Fact]
        public void BuildsASkillFromARecordedRun()
        {
            var skill = SkillParameterizer.Build(
                "月度销售报表",
                "把 A1:F200 按销售额降序排列，并生成柱状图",
                new[]
                {
                    Step("read_range", "address", "A1:F200"),
                    Step("sort_data", "address", "A1:F200", "column", "E", "ascending", false),
                    Step("create_chart", "address", "A1:F200", "chart_type", "column")
                });

            Assert.Equal("月度销售报表", skill.Name);
            Assert.Equal(3, skill.Steps.Count);

            // The same range across three steps is one parameter, not three the
            // user has to keep consistent.
            var ranges = skill.Parameters.Where(p => p.Kind == ParameterKind.Range).ToList();
            Assert.Single(ranges);
            Assert.Equal("A1:F200", ranges[0].DefaultValue);

            Assert.All(skill.Steps, s => Assert.Equal("{{range}}", s.Arguments["address"]));
            // The sort direction the user chose stays fixed.
            Assert.Equal("false", skill.Steps[1].Arguments["ascending"]);
            Assert.Equal("column", skill.Steps[2].Arguments["chart_type"]);
        }

        [Fact]
        public void DistinctValuesGetDistinctParameters()
        {
            var skill = SkillParameterizer.Build("跨表汇总", null, new[]
            {
                Step("read_range", "address", "A1:C100", "sheet", "明细"),
                Step("write_range", "address", "F1:H10", "sheet", "汇总")
            });

            Assert.Equal(2, skill.Parameters.Count(p => p.Kind == ParameterKind.Range));
            Assert.Equal(2, skill.Parameters.Count(p => p.Kind == ParameterKind.Sheet));
            // Names must not collide.
            Assert.Equal(skill.Parameters.Count, skill.Parameters.Select(p => p.Name).Distinct().Count());
        }

        [Fact]
        public void OriginalRequestGetsTheSamePlaceholders()
        {
            var skill = SkillParameterizer.Build(
                "排序",
                "把 A1:F200 按销售额降序排列",
                new[] { Step("sort_data", "address", "A1:F200") });

            // Otherwise the prompt would contain the recorded range while the
            // steps contain a placeholder, and the agent would see two answers.
            Assert.Contains("{{range}}", skill.OriginalRequest);
            Assert.DoesNotContain("A1:F200", skill.OriginalRequest);
        }

        [Fact]
        public void LongerValuesAreSubstitutedFirst()
        {
            // Replacing "A1" before "A1:F200" would corrupt the longer one into
            // "{{range}}:F200".
            var parameters = new List<SkillParameter>
            {
                new SkillParameter { Name = "cell", DefaultValue = "A1", Kind = ParameterKind.Range },
                new SkillParameter { Name = "range", DefaultValue = "A1:F200", Kind = ParameterKind.Range }
            };
            var result = SkillParameterizer.ReplaceValuesWithPlaceholders("从 A1:F200 开始，标题在 A1", parameters);

            Assert.Contains("{{range}}", result);
            Assert.Contains("{{cell}}", result);
            Assert.DoesNotContain("{{range}}:F200", result);
        }

        [Fact]
        public void EmptyRecordingProducesAnEmptySkill()
        {
            var skill = SkillParameterizer.Build("空", null, new RecordedStep[0]);
            Assert.Empty(skill.Steps);
            Assert.Empty(skill.Parameters);
        }

        [Fact]
        public void MalformedStepsAreSkippedRatherThanThrowing()
        {
            var skill = SkillParameterizer.Build("部分", null, new[]
            {
                null,
                new RecordedStep { Tool = null },
                Step("read_range", "address", "A1")
            });
            Assert.Single(skill.Steps);
        }

        [Fact]
        public void UnnamedSkillGetsAPlaceholderName()
        {
            Assert.Equal("未命名技能", SkillParameterizer.Build("  ", null, new RecordedStep[0]).Name);
        }

        // ------------------------------------------------------------------
        // Replay
        // ------------------------------------------------------------------

        [Fact]
        public void ReplayUsesSuppliedArgumentsOverDefaults()
        {
            var skill = SkillParameterizer.Build(
                "月报", "汇总 A1:F200",
                new[] { Step("sort_data", "address", "A1:F200") });

            var prompt = skill.ToPrompt(new Dictionary<string, string> { ["range"] = "A1:F900" });

            Assert.Contains("A1:F900", prompt);
            Assert.DoesNotContain("A1:F200", prompt);
            Assert.DoesNotContain("{{range}}", prompt);
        }

        [Fact]
        public void ReplayFallsBackToRecordedDefaults()
        {
            var skill = SkillParameterizer.Build(
                "月报", "汇总 A1:F200",
                new[] { Step("sort_data", "address", "A1:F200") });

            foreach (var arguments in new[]
            {
                null,
                new Dictionary<string, string>(),
                new Dictionary<string, string> { ["range"] = "   " }
            })
            {
                var prompt = skill.ToPrompt(arguments);
                Assert.Contains("A1:F200", prompt);
                Assert.DoesNotContain("{{range}}", prompt);
            }
        }

        [Fact]
        public void ReplayPromptCarriesIntentNotJustSteps()
        {
            // Replaying the recorded calls verbatim breaks as soon as the data
            // differs from the recording, which is most of the time. The agent
            // needs to know what was meant, not only what was called.
            var skill = SkillParameterizer.Build(
                "月度销售报表",
                "把 A1:F200 按销售额降序排列，并生成柱状图",
                new[]
                {
                    Step("sort_data", "address", "A1:F200"),
                    Step("create_chart", "address", "A1:F200")
                });

            var prompt = skill.ToPrompt(null);

            Assert.Contains("月度销售报表", prompt);
            Assert.Contains("按销售额降序排列", prompt);
            Assert.Contains("sort_data", prompt);
            Assert.Contains("create_chart", prompt);
            Assert.Contains("以意图为准", prompt);
            Assert.Contains("不要机械重放", prompt);
        }

        [Fact]
        public void ParameterLabelsAreHumanReadable()
        {
            var skill = SkillParameterizer.Build("x", null, new[]
            {
                Step("read_range", "address", "A1:F200", "sheet", "明细")
            });

            Assert.Contains(skill.Parameters, p => p.Label == "数据区域");
            Assert.Contains(skill.Parameters, p => p.Label == "工作表");
        }
    }
}
