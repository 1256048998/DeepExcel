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
    }
}
