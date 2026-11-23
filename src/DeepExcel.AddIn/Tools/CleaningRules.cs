using System;

namespace DeepExcel.AddIn.Tools
{
    /// <summary>
    /// What each cleaning operation does to one cell value, free of Excel.
    ///
    /// <see cref="DataCleaner"/> applies these to the workbook and the change
    /// preview applies them to a snapshot of it. Sharing one implementation is
    /// the point: a preview that drifts from the executor reports changes that
    /// will not happen, or misses ones that will.
    ///
    /// Each rule takes a raw Value2 (string, double, bool or null) and returns
    /// the new value, or null when the cell is left alone.
    /// </summary>
    public static class CleaningRules
    {
        /// <summary>trim_spaces: leading and trailing whitespace, text only.</summary>
        public static string TrimSpaces(object value)
        {
            if (!(value is string s)) return null;
            var trimmed = s.Trim();
            return trimmed != s ? trimmed : null;
        }

        /// <summary>
        /// collapse_spaces: runs of spaces become one, and the edges are
        /// trimmed as well. Text only.
        /// </summary>
        public static string CollapseSpaces(object value)
        {
            if (!(value is string s) || s.Length == 0) return null;
            var collapsed = s.Trim();
            while (collapsed.Contains("  "))
            {
                collapsed = collapsed.Replace("  ", " ");
            }
            return collapsed != s ? collapsed : null;
        }

        /// <summary>
        /// remove_special_chars: control characters other than tab and line
        /// breaks, and non-breaking spaces. Text only.
        /// </summary>
        public static string RemoveSpecialChars(object value)
        {
            if (!(value is string s) || s.Length == 0) return null;
            var kept = new System.Text.StringBuilder(s.Length);
            foreach (var c in s)
            {
                if (IsSpecialChar(c)) continue;
                kept.Append(c);
            }
            var cleaned = kept.ToString();
            return cleaned != s ? cleaned : null;
        }

        /// <summary>Number format clean_amount applies to what it converts.</summary>
        public const string AmountFormat = "#,##0.00";

        /// <summary>
        /// clean_amount: currency symbols, thousands separators and 元/圆
        /// stripped, then parsed as a number. Text only; text that does not
        /// parse is left alone.
        /// </summary>
        public static double? CleanAmount(object value)
        {
            if (!(value is string s) || s.Length == 0) return null;
            var cleaned = s
                .Replace("¥", "").Replace("￥", "")
                .Replace("$", "").Replace("€", "").Replace("£", "")
                .Replace(",", "").Replace("，", "")
                .Replace("元", "").Replace("圆", "")
                .Trim();
            return double.TryParse(cleaned, out var number) ? number : (double?)null;
        }

        public static bool IsSpecialChar(char c)
        {
            if (char.IsControl(c) && c != '\t' && c != '\n' && c != '\r') return true;
            return c == '\u00A0';
        }
    }
}
