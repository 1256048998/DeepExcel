using System.Linq;
using DeepExcel.AddIn.Skills;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// What leaves the machine when a skill is synced or shared.
    ///
    /// A skill's defaults are values captured from a real run, so this is the
    /// difference between "backed up my workflow" and "published my file
    /// layout". It is scrubbed here and again on the server: here so the data
    /// never leaves, there so a modified client cannot skip it.
    /// </summary>
    public class SkillScrubberTests
    {
        private static Skill Build()
        {
            return SkillParameterizer.Build(
                "月度销售报表",
                @"读取 C:\Users\alice\2026财务预算.xlsx 并按 A1:F200 汇总",
                new[]
                {
                    new RecordedStep
                    {
                        Tool = "read_range",
                        Arguments =
                        {
                            ["address"] = "A1:F200",
                            ["path"] = @"C:\Users\alice\2026财务预算.xlsx"
                        }
                    },
                    new RecordedStep { Tool = "sort_data", Arguments = { ["address"] = "A1:F200" } }
                });
        }

        [Fact]
        public void WindowsAndUncPathsAreRemoved()
        {
            Assert.DoesNotContain("alice", SkillScrubber.Scrub(@"读取 C:\Users\alice\budget.xlsx"));
            Assert.DoesNotContain("alice", SkillScrubber.Scrub(@"\\server\share\alice\x.xlsx"));
        }

        [Fact]
        public void BareWorkbookNamesAreRemoved()
        {
            // A file name alone can identify a project, a client or a quarter.
            var cleaned = SkillScrubber.Scrub("汇总 2026年Q3并购尽调.xlsx");
            Assert.DoesNotContain("并购尽调", cleaned);
        }

        [Fact]
        public void StructureIsPreserved()
        {
            // Placeholders and ranges are structure; scrubbing them would
            // destroy the skill.
            Assert.Equal("把 {{range}} 按 A1:F200 排序",
                SkillScrubber.Scrub("把 {{range}} 按 A1:F200 排序"));
        }

        [Fact]
        public void UploadCopyCarriesNoLocalPaths()
        {
            var uploaded = SkillScrubber.ForUpload(Build());

            Assert.DoesNotContain("alice", uploaded.OriginalRequest);
            Assert.DoesNotContain("财务预算", uploaded.OriginalRequest);
            Assert.All(uploaded.Steps, step =>
                Assert.All(step.Arguments.Values, value => Assert.DoesNotContain("alice", value)));
            Assert.All(uploaded.Parameters, parameter =>
                Assert.DoesNotContain("alice", parameter.DefaultValue ?? ""));
        }

        [Fact]
        public void FileParameterDefaultsAreDroppedEntirely()
        {
            var skill = SkillParameterizer.Build("导入", null, new[]
            {
                new RecordedStep
                {
                    Tool = "read_attachment",
                    Arguments = { ["path"] = @"D:\HR\salary.xlsx" }
                }
            });

            var uploaded = SkillScrubber.ForUpload(skill);
            var fileParameters = uploaded.Parameters.Where(p => p.Kind == ParameterKind.File).ToList();

            // The recipient's file is somewhere else, so the default is useless
            // to them and harmful to keep.
            Assert.All(fileParameters, p => Assert.Equal("", p.DefaultValue));
        }

        [Fact]
        public void RangeDefaultsSurviveBecauseTheyAreUseful()
        {
            var uploaded = SkillScrubber.ForUpload(Build());
            var range = uploaded.Parameters.First(p => p.Kind == ParameterKind.Range);

            // A1:F200 identifies nothing and is genuinely a good default.
            Assert.Equal("A1:F200", range.DefaultValue);
        }

        [Fact]
        public void TheLocalSkillIsNotModified()
        {
            // Scrubbing on upload must not degrade the copy the user runs every
            // month with working defaults.
            var original = Build();
            var beforeRequest = original.OriginalRequest;
            var beforeSteps = original.Steps.Select(s => string.Join(",", s.Arguments.Values)).ToArray();

            SkillScrubber.ForUpload(original);

            Assert.Equal(beforeRequest, original.OriginalRequest);
            Assert.Equal(beforeSteps,
                original.Steps.Select(s => string.Join(",", s.Arguments.Values)).ToArray());
        }

        [Fact]
        public void RedactionIsDetectableSoTheUserCanBeTold()
        {
            Assert.True(SkillScrubber.WouldRedact(Build()));

            var clean = SkillParameterizer.Build("排序", "把 A1:F200 排序", new[]
            {
                new RecordedStep { Tool = "sort_data", Arguments = { ["address"] = "A1:F200" } }
            });
            Assert.False(SkillScrubber.WouldRedact(clean));
        }

        [Fact]
        public void UploadCopyIsStillRunnable()
        {
            // Scrubbing must not break the thing being shared.
            var uploaded = SkillScrubber.ForUpload(Build());
            var prompt = uploaded.ToPrompt(null);

            Assert.Contains("read_range", prompt);
            Assert.Contains("sort_data", prompt);
            Assert.DoesNotContain("{{", prompt); // all placeholders resolved
        }

        [Fact]
        public void NullIsHandled()
        {
            Assert.Null(SkillScrubber.ForUpload(null));
            Assert.False(SkillScrubber.WouldRedact(null));
            Assert.Null(SkillScrubber.Scrub(null));
            Assert.Equal("", SkillScrubber.Scrub(""));
        }
    }
}
