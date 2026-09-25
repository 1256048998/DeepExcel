using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>工作表上的一个矩形区域（行列从 1 开始，含两端）。</summary>
    public struct CellRect
    {
        public const int MaxRow = 1048576;
        public const int MaxColumn = 16384;

        public string Sheet;
        public int Row1, Col1, Row2, Col2;

        public CellRect(string sheet, int row1, int col1, int row2, int col2)
        {
            Sheet = sheet ?? "";
            Row1 = Math.Min(row1, row2); Row2 = Math.Max(row1, row2);
            Col1 = Math.Min(col1, col2); Col2 = Math.Max(col1, col2);
        }

        public bool Overlaps(CellRect other) =>
            string.Equals(Sheet, other.Sheet, StringComparison.OrdinalIgnoreCase) &&
            Row1 <= other.Row2 && other.Row1 <= Row2 &&
            Col1 <= other.Col2 && other.Col1 <= Col2;

        /// <summary>向下 / 向右扩展（write_range 的二维数组、fill_formula_down 的行数）。</summary>
        public CellRect Resize(int rows, int columns) =>
            new CellRect(Sheet, Row1, Col1,
                Math.Min(MaxRow, Row1 + Math.Max(1, rows) - 1),
                Math.Min(MaxColumn, Col1 + Math.Max(1, columns) - 1));

        public string ToA1()
        {
            var sheet = Sheet.IndexOfAny(new[] { ' ', '-', '!', '\'' }) >= 0 ? "'" + Sheet.Replace("'", "''") + "'" : Sheet;
            var tl = ColumnName(Col1) + Row1;
            var br = ColumnName(Col2) + Row2;
            return sheet + "!" + (tl == br ? tl : tl + ":" + br);
        }

        public static string ColumnName(int column)
        {
            var name = "";
            while (column > 0)
            {
                var rem = (column - 1) % 26;
                name = (char)('A' + rem) + name;
                column = (column - 1) / 26;
            }
            return name;
        }

        private static readonly Regex CellPattern = new Regex(@"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$");
        private static readonly Regex ColumnPattern = new Regex(@"^\$?([A-Za-z]{1,3})$");
        private static readonly Regex RowPattern = new Regex(@"^\$?(\d{1,7})$");

        /// <summary>
        /// 解析 A1 地址：Sheet1!A1:C4、'My Sheet'!B2、$A$1、A:C（整列）、2:5（整行）。
        /// 没写表名时用 defaultSheet。命名区域、R1C1、多区域（逗号）不解析，返回 false。
        /// </summary>
        public static bool TryParse(string address, string defaultSheet, out CellRect rect)
        {
            rect = default;
            if (string.IsNullOrWhiteSpace(address)) return false;
            var text = address.Trim();
            var sheet = defaultSheet ?? "";
            var bang = text.LastIndexOf('!');
            if (bang >= 0)
            {
                sheet = text.Substring(0, bang).Trim();
                if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
                {
                    sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
                }
                text = text.Substring(bang + 1).Trim();
            }
            if (text.Length == 0 || text.Contains(",")) return false;

            var parts = text.Split(':');
            if (parts.Length > 2) return false;
            var a = parts[0];
            var b = parts.Length == 2 ? parts[1] : parts[0];

            if (TryCell(a, out var r1, out var c1) && TryCell(b, out var r2, out var c2))
            {
                rect = new CellRect(sheet, r1, c1, r2, c2);
                return true;
            }
            var ca = ColumnPattern.Match(a); var cb = ColumnPattern.Match(b);
            if (parts.Length == 2 && ca.Success && cb.Success)
            {
                rect = new CellRect(sheet, 1, ColumnIndex(ca.Groups[1].Value), MaxRow, ColumnIndex(cb.Groups[1].Value));
                return true;
            }
            var ra = RowPattern.Match(a); var rb = RowPattern.Match(b);
            if (parts.Length == 2 && ra.Success && rb.Success)
            {
                rect = new CellRect(sheet, int.Parse(ra.Groups[1].Value), 1, int.Parse(rb.Groups[1].Value), MaxColumn);
                return true;
            }
            return false;
        }

        private static bool TryCell(string text, out int row, out int column)
        {
            row = column = 0;
            var m = CellPattern.Match(text);
            if (!m.Success) return false;
            column = ColumnIndex(m.Groups[1].Value);
            row = int.Parse(m.Groups[2].Value);
            return row >= 1 && row <= MaxRow && column >= 1 && column <= MaxColumn;
        }

        private static int ColumnIndex(string letters)
        {
            var index = 0;
            foreach (var ch in letters.ToUpperInvariant()) index = index * 26 + (ch - 'A' + 1);
            return index;
        }
    }

    public enum LedgerVerdict
    {
        Allowed,
        /// <summary>目标区域在模型上次读取之后被用户手动改过</summary>
        StaleRead,
        /// <summary>目标区域有内容，但模型从没读过（也没写过）它</summary>
        NotRead,
    }

    /// <summary>
    /// 先读后写 + 读后被改检测（Claude Code 的「编辑前必须先 Read」「读取后文件被改过」）。
    ///
    /// 模型看不到工作簿本身，只看得到它读过的部分。以前它可以不看一眼就覆盖一块有内容的
    /// 区域；用户在它工作时手动改了单元格，它也会照着几分钟前读到的旧值写回去，把用户刚改
    /// 的内容冲掉。这里记下模型读过 / 写过的区域和用户手动改过的区域，写入前据此把关。
    ///
    /// 纯逻辑，不碰 COM：区域有没有内容由调用方回答。
    /// </summary>
    public class ReadLedger
    {
        private const int MaxEntries = 300;

        private sealed class Entry
        {
            public CellRect Rect;
            public long Seq;
            public bool Reported;
        }

        private readonly object _lock = new object();
        private readonly List<Entry> _known = new List<Entry>();
        private readonly List<Entry> _userEdits = new List<Entry>();
        private readonly List<string> _notices = new List<string>();
        private long _seq;

        /// <summary>模型读到了这块区域的内容（read_range / read_selection）。</summary>
        public void RecordRead(CellRect rect) => Add(_known, rect);

        /// <summary>模型自己写了这块区域：它知道这里现在是什么，等同于读过。</summary>
        public void RecordOwnWrite(CellRect rect) => Add(_known, rect);

        /// <summary>用户（或别的程序）在 Excel 里改了这块区域，不是我们的工具改的。</summary>
        public void RecordUserEdit(CellRect rect) => Add(_userEdits, rect);

        /// <summary>
        /// 回滚之类整体换掉内容的操作之后：模型之前读到的都作废，要重新读。
        /// </summary>
        public void ForgetReads()
        {
            lock (_lock) { _known.Clear(); }
        }

        public LedgerVerdict Check(CellRect target, Func<bool> targetHasContent, out List<string> changedAddresses)
        {
            changedAddresses = new List<string>();
            long lastKnown;
            lock (_lock)
            {
                var overlappingKnown = _known.Where(e => e.Rect.Overlaps(target)).ToList();
                lastKnown = overlappingKnown.Count == 0 ? -1 : overlappingKnown.Max(e => e.Seq);
                if (lastKnown >= 0)
                {
                    // 模型看过这里；看过之后用户又改了其中的单元格 → 它手里的是旧值
                    changedAddresses = _userEdits
                        .Where(e => e.Seq > lastKnown && e.Rect.Overlaps(target))
                        .Select(e => e.Rect.ToA1())
                        .Distinct()
                        .Take(10)
                        .ToList();
                    return changedAddresses.Count > 0 ? LedgerVerdict.StaleRead : LedgerVerdict.Allowed;
                }
            }
            // 从没看过：空白区域可以直接写（新建的表、新的一列），有内容就得先读
            bool hasContent;
            try { hasContent = targetHasContent == null || targetHasContent(); }
            catch { hasContent = false; }
            return hasContent ? LedgerVerdict.NotRead : LedgerVerdict.Allowed;
        }

        /// <summary>
        /// 还没告诉模型的用户改动（每条只报一次）。附在下一个工具结果里，模型才知道表变了。
        /// </summary>
        public List<string> TakeUnreportedUserEdits()
        {
            lock (_lock)
            {
                var pending = _userEdits.Where(e => !e.Reported).ToList();
                foreach (var e in pending) e.Reported = true;
                return pending.Select(e => e.Rect.ToA1()).Distinct().Take(20).ToList();
            }
        }

        /// <summary>
        /// 不是某个区域、但模型必须知道的事（例如用户在面板上回退到了之前的检查点）。
        /// 和用户改动一样，在下一个工具结果或下一条用户消息里告诉模型一次。
        /// </summary>
        public void AddNotice(string notice)
        {
            if (string.IsNullOrWhiteSpace(notice)) return;
            lock (_lock)
            {
                _notices.Add(notice);
                if (_notices.Count > 5) _notices.RemoveAt(0);
            }
        }

        public List<string> TakeNotices()
        {
            lock (_lock)
            {
                var pending = _notices.ToList();
                _notices.Clear();
                return pending;
            }
        }

        private void Add(List<Entry> list, CellRect rect)
        {
            lock (_lock)
            {
                list.Add(new Entry { Rect = rect, Seq = ++_seq });
                if (list.Count > MaxEntries) list.RemoveRange(0, list.Count - MaxEntries);
            }
        }
    }
}
