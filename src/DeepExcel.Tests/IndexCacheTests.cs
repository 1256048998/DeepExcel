using System;
using DeepExcel.AddIn.Perception;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Cache behaviour for the workbook structure summary.
    ///
    /// Both directions cost something real. Rebuilding too often puts a
    /// multi-second sheet scan in front of every message; rebuilding too rarely
    /// hands the model a structure that no longer matches the sheet, and it will
    /// act on the stale one.
    /// </summary>
    public class IndexCacheTests
    {
        private static WorkbookIndex SampleIndex(string sheetName = "Sheet1")
        {
            return new WorkbookIndex
            {
                WorkbookName = "test.xlsx",
                Sheets = { new SheetIndex { Name = sheetName, LastRow = 100 } }
            };
        }

        [Fact]
        public void MissEntryReturnsNull()
        {
            var cache = new IndexCache();
            Assert.Null(cache.TryGet("wb-1", DateTime.UtcNow));
            Assert.Null(cache.TryGet(null, DateTime.UtcNow));
        }

        [Fact]
        public void StoredIndexIsReturnedWithoutRebuilding()
        {
            var cache = new IndexCache();
            var now = DateTime.UtcNow;
            cache.Store("wb-1", SampleIndex(), now);

            var rendered = cache.TryGet("wb-1", now.AddMinutes(1));
            Assert.NotNull(rendered);
            Assert.Contains("Sheet1", rendered);
        }

        [Fact]
        public void EditingTheWorkbookInvalidatesImmediately()
        {
            var cache = new IndexCache();
            var now = DateTime.UtcNow;
            cache.Store("wb-1", SampleIndex(), now);
            Assert.NotNull(cache.TryGet("wb-1", now));

            cache.Invalidate("wb-1");

            // A stale structure is worse than none: the model would operate on
            // columns that have moved.
            Assert.Null(cache.TryGet("wb-1", now));
        }

        [Fact]
        public void StaleEntriesExpireEvenWithoutAnEdit()
        {
            // Not every change raises SheetChange -- external links, calculation
            // and other add-ins can all move data underneath us.
            var cache = new IndexCache();
            var now = DateTime.UtcNow;
            cache.Store("wb-1", SampleIndex(), now);

            Assert.NotNull(cache.TryGet("wb-1", now + IndexCache.MaxAge - TimeSpan.FromSeconds(1)));
            Assert.Null(cache.TryGet("wb-1", now + IndexCache.MaxAge + TimeSpan.FromSeconds(1)));
        }

        [Fact]
        public void RebuildingClearsTheDirtyFlag()
        {
            var cache = new IndexCache();
            var now = DateTime.UtcNow;
            cache.Store("wb-1", SampleIndex(), now);
            cache.Invalidate("wb-1");
            Assert.Null(cache.TryGet("wb-1", now));

            cache.Store("wb-1", SampleIndex("Updated"), now);

            var rendered = cache.TryGet("wb-1", now);
            Assert.NotNull(rendered);
            Assert.Contains("Updated", rendered);
        }

        [Fact]
        public void WorkbooksAreCachedIndependently()
        {
            var cache = new IndexCache();
            var now = DateTime.UtcNow;
            cache.Store("wb-1", SampleIndex("First"), now);
            cache.Store("wb-2", SampleIndex("Second"), now);

            // Editing one workbook must not force a rescan of the others.
            cache.Invalidate("wb-1");

            Assert.Null(cache.TryGet("wb-1", now));
            Assert.Contains("Second", cache.TryGet("wb-2", now));
        }

        [Fact]
        public void ClosingAWorkbookReleasesItsEntry()
        {
            var cache = new IndexCache();
            cache.Store("wb-1", SampleIndex(), DateTime.UtcNow);
            Assert.Equal(1, cache.Count);

            cache.Remove("wb-1");

            Assert.Equal(0, cache.Count);
            Assert.Null(cache.TryGet("wb-1", DateTime.UtcNow));
        }

        [Fact]
        public void InvalidatingAnUnknownWorkbookIsHarmless()
        {
            // SheetChange fires for workbooks we have never indexed.
            var cache = new IndexCache();
            cache.Invalidate("never-seen");
            cache.Invalidate(null);
            cache.Remove("never-seen");
            Assert.Equal(0, cache.Count);
        }

        [Fact]
        public void StoringNullIsIgnoredRatherThanCachingAnEmptyAnswer()
        {
            var cache = new IndexCache();
            cache.Store("wb-1", null, DateTime.UtcNow);
            Assert.Equal(0, cache.Count);
            Assert.Null(cache.TryGet("wb-1", DateTime.UtcNow));
        }

        [Fact]
        public void RepeatedInvalidationIsCheapAndIdempotent()
        {
            // This runs once per changed range, including inside a loop writing
            // thousands of cells.
            var cache = new IndexCache();
            cache.Store("wb-1", SampleIndex(), DateTime.UtcNow);

            for (var i = 0; i < 10000; i++)
            {
                cache.Invalidate("wb-1");
            }

            Assert.Equal(1, cache.Count);
            Assert.Null(cache.TryGet("wb-1", DateTime.UtcNow));
        }
    }
}
