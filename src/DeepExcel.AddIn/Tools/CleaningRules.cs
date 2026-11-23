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
    }
}
