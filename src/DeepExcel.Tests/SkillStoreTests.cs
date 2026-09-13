using System;
using System.IO;
using System.Linq;
using DeepExcel.AddIn.Skills;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Skill persistence.
    ///
    /// The library is the first thing a user accumulates rather than rents, so
    /// losing one to a bad write, or silently listing one that cannot run, both
    /// undermine the reason it exists.
    /// </summary>
    public class SkillStoreTests : IDisposable
    {
        private readonly string _directory;
        private readonly SkillStore _store;

        public SkillStoreTests()
        {
            _directory = Path.Combine(Path.GetTempPath(), "DeepExcelSkillTests", Guid.NewGuid().ToString("N"));
            _store = new SkillStore(_directory);
        }

        public void Dispose()
        {
            try { Directory.Delete(_directory, true); } catch (Exception) { }
        }

        private static Skill Make(string name, int steps = 2)
        {
            var skill = new Skill { Name = name };
            for (var i = 0; i < steps; i++)
            {
                skill.Steps.Add(new SkillStep { Tool = "read_range" });
            }
            return skill;
        }

        [Fact]
        public void EmptyLibraryListsNothing()
        {
            Assert.Empty(_store.List());
            Assert.Null(_store.Get("missing"));
        }

        [Fact]
        public void SavedSkillRoundTrips()
        {
            var skill = SkillParameterizer.Build("月报", "汇总 A1:F200", new[]
            {
                new RecordedStep { Tool = "sort_data", Arguments = { ["address"] = "A1:F200" } },
                new RecordedStep { Tool = "create_chart", Arguments = { ["address"] = "A1:F200" } }
            });

            Assert.True(_store.Save(skill));

            var loaded = _store.Get(skill.Id);
            Assert.NotNull(loaded);
            Assert.Equal("月报", loaded.Name);
            Assert.Equal(2, loaded.Steps.Count);
            Assert.Single(loaded.Parameters);
            Assert.Equal("A1:F200", loaded.Parameters[0].DefaultValue);
            // The replay prompt must survive the round-trip, since that is the
            // only thing that actually runs.
            Assert.Contains("A1:F200", loaded.ToPrompt(null));
        }

        [Fact]
        public void DeletingRemovesIt()
        {
            var skill = Make("临时");
            _store.Save(skill);
            Assert.True(_store.Delete(skill.Id));
            Assert.Null(_store.Get(skill.Id));
            Assert.False(_store.Delete(skill.Id));
        }

        [Fact]
        public void MostUsedSkillsComeFirst()
        {
            // A library is only useful if the skill someone reaches for is the
            // one they see.
            var rare = Make("很少用");
            var common = Make("常用");
            _store.Save(rare);
            _store.Save(common);

            for (var i = 0; i < 5; i++)
            {
                _store.RecordRun(common.Id);
            }

            Assert.Equal("常用", _store.List().First().Name);
            Assert.Equal(5, _store.Get(common.Id).RunCount);
            Assert.NotNull(_store.Get(common.Id).LastRunAt);
        }

        [Fact]
        public void CorruptFilesAreSkippedNotFatal()
        {
            var good = Make("好的");
            _store.Save(good);
            Directory.CreateDirectory(_directory);
            File.WriteAllText(Path.Combine(_directory, "broken.json"), "{ not json at all");

            // One bad file must not take the library with it.
            var listed = _store.List();
            Assert.Single(listed);
            Assert.Equal("好的", listed[0].Name);
        }

        [Fact]
        public void SkillsWithNoStepsAreNotListed()
        {
            // Listing one would produce a confusing failure at replay time
            // rather than at save time.
            var empty = new Skill { Name = "空技能" };
            _store.Save(empty);
            Assert.Empty(_store.List());
        }

        [Fact]
        public void IdsCannotEscapeTheSkillDirectory()
        {
            // Ids are generated, but a hand-edited file must not be able to
            // reach outside the folder.
            var malicious = new Skill { Id = @"..\..\..\evil", Name = "x" };
            malicious.Steps.Add(new SkillStep { Tool = "read_range" });
            _store.Save(malicious);

            var written = Directory.Exists(_directory)
                ? Directory.GetFiles(_directory, "*.json", SearchOption.AllDirectories)
                : new string[0];
            Assert.All(written, path =>
                Assert.StartsWith(Path.GetFullPath(_directory), Path.GetFullPath(path)));
        }

        [Fact]
        public void SavingNullOrIdlessSkillIsRejected()
        {
            Assert.False(_store.Save(null));
            Assert.False(_store.Save(new Skill { Id = "" }));
        }

        [Fact]
        public void RecordRunOnAMissingSkillIsHarmless()
        {
            _store.RecordRun("nope");
            _store.RecordRun(null);
            Assert.Empty(_store.List());
        }

        [Fact]
        public void OverwritingAnExistingSkillKeepsItReadable()
        {
            var skill = Make("原名");
            _store.Save(skill);

            skill.Name = "改名";
            Assert.True(_store.Save(skill));

            Assert.Equal("改名", _store.Get(skill.Id).Name);
            Assert.Single(_store.List());
            // The write-then-replace temp file must not linger.
            Assert.Empty(Directory.GetFiles(_directory, "*.tmp"));
        }
    }
}
