using System;
using System.Collections.Generic;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// Caches the rendered index per workbook and rebuilds it when the workbook
    /// changes.
    ///
    /// Rebuilding on every request would put a multi-second sheet scan in front
    /// of every message. Never rebuilding would feed the model a stale structure
    /// after the user edits, which is worse than no index -- the model would
    /// operate on columns that have moved.
    ///
    /// Invalidation is driven by Excel's SheetChange event rather than by
    /// hashing the data, because hashing means reading it, which is the cost the
    /// cache exists to avoid.
    /// </summary>
    public sealed class IndexCache
    {
        private sealed class Entry
        {
            public WorkbookIndex Index;
            public string Rendered;
            public DateTime BuiltUtc;
            public bool Dirty;
        }

        /// <summary>
        /// Rebuilt after this long even without an edit, to catch changes made
        /// through paths that do not raise SheetChange (external links,
        /// calculation, another add-in).
        /// </summary>
        public static readonly TimeSpan MaxAge = TimeSpan.FromMinutes(10);

        private readonly Dictionary<string, Entry> _entries =
            new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
        private readonly object _lock = new object();

        public int Count
        {
            get { lock (_lock) { return _entries.Count; } }
        }

        /// <summary>
        /// Returns the cached render, or null when a rebuild is needed.
        /// </summary>
        public string TryGet(string workbookKey, DateTime nowUtc)
        {
            if (string.IsNullOrEmpty(workbookKey))
            {
                return null;
            }
            lock (_lock)
            {
                if (!_entries.TryGetValue(workbookKey, out var entry))
                {
                    return null;
                }
                if (entry.Dirty || nowUtc - entry.BuiltUtc > MaxAge)
                {
                    return null;
                }
                return entry.Rendered;
            }
        }

        public void Store(string workbookKey, WorkbookIndex index, DateTime nowUtc)
        {
            if (string.IsNullOrEmpty(workbookKey) || index == null)
            {
                return;
            }
            lock (_lock)
            {
                _entries[workbookKey] = new Entry
                {
                    Index = index,
                    Rendered = index.Render(),
                    BuiltUtc = nowUtc,
                    Dirty = false
                };
            }
        }

        /// <summary>
        /// Marks a workbook stale. Cheap on purpose: this runs on every
        /// SheetChange, including inside a loop that writes thousands of cells,
        /// so it must not do any work beyond setting a flag.
        /// </summary>
        public void Invalidate(string workbookKey)
        {
            if (string.IsNullOrEmpty(workbookKey))
            {
                return;
            }
            lock (_lock)
            {
                if (_entries.TryGetValue(workbookKey, out var entry))
                {
                    entry.Dirty = true;
                }
            }
        }

        public void Remove(string workbookKey)
        {
            if (string.IsNullOrEmpty(workbookKey))
            {
                return;
            }
            lock (_lock)
            {
                _entries.Remove(workbookKey);
            }
        }

        public WorkbookIndex PeekIndex(string workbookKey)
        {
            lock (_lock)
            {
                return _entries.TryGetValue(workbookKey ?? "", out var entry) ? entry.Index : null;
            }
        }
    }
}
