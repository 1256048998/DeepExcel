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

        /// <summary>
        /// 不管是否过期都返回上一次的渲染结果和建立时间。只在重建失败时用：
        /// 过时的结构总比没有强，但必须配上 <see cref="MarkStale"/> 的说明一起给模型。
        /// </summary>
        public bool TryGetStale(string workbookKey, out string rendered, out DateTime builtUtc)
        {
            rendered = null;
            builtUtc = default(DateTime);
            if (string.IsNullOrEmpty(workbookKey))
            {
                return false;
            }
            lock (_lock)
            {
                if (!_entries.TryGetValue(workbookKey, out var entry) || string.IsNullOrEmpty(entry.Rendered))
                {
                    return false;
                }
                rendered = entry.Rendered;
                builtUtc = entry.BuiltUtc;
                return true;
            }
        }

        /// <summary>
        /// 在「## 工作簿结构」标题下插一句：这次读不到 Excel，下面是多久以前的缓存。
        /// </summary>
        public static string MarkStale(string rendered, TimeSpan age, string reason)
        {
            if (string.IsNullOrEmpty(rendered))
            {
                return rendered;
            }
            var minutes = Math.Max(1, (int)Math.Round(age.TotalMinutes));
            var why = string.IsNullOrWhiteSpace(reason) ? "" : "（" + reason.Trim() + "）";
            var note = "注意：这次读不到 Excel" + why + "，以下沿用 " + minutes +
                       " 分钟前的缓存，之后的修改没有反映；涉及具体数值或结构时请先 read_range 实读。";
            var newline = rendered.IndexOf('\n');
            return newline < 0
                ? rendered + "\n" + note
                : rendered.Substring(0, newline + 1) + note + "\n" + rendered.Substring(newline + 1);
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
