using System;
using System.Collections.Generic;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>How much ceremony an operation needs before it runs.</summary>
    public enum ConfirmationLevel
    {
        /// <summary>Reversible and small. Runs immediately.</summary>
        None,

        /// <summary>Shows the exact change set; the user confirms it.</summary>
        Preview,

        /// <summary>
        /// Cannot be simulated. A snapshot is taken and the user is told plainly
        /// that the preview does not cover what will run.
        /// </summary>
        SnapshotAndWarn
    }

    /// <summary>
    /// Which operations need confirmation, and which kind.
    ///
    /// The distinction between <see cref="ConfirmationLevel.Preview"/> and
    /// <see cref="ConfirmationLevel.SnapshotAndWarn"/> is the point of this
    /// type. Presenting both the same way would train users to click through a
    /// dialog that sometimes means "here is exactly what happens" and sometimes
    /// means "we have no idea" -- after which the first kind stops being read.
    /// </summary>
    public static class PreviewPolicy
    {
        /// <summary>
        /// Arbitrary code. The tool boundary is the only thing we can see; what
        /// runs inside is opaque, so there is nothing to simulate.
        /// </summary>
        private static readonly Dictionary<string, string> Unpreviewable =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["execute_vba"] = "VBA 代码可以做任何事，无法预先推算它会改动什么",
                ["execute_python"] = "Python 代码可以做任何事，无法预先推算它会改动什么",
                ["send_keys"] = "按键模拟会直接作用于 Excel 界面，无法预先推算结果"
            };

        /// <summary>
        /// Destructive or wide-reaching, and simulable.
        ///
        /// Read-only tools are absent on purpose: prompting for them would be
        /// noise, and noise is what makes people stop reading prompts.
        /// </summary>
        private static readonly HashSet<string> NeedsPreview =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "write_range", "write_table", "write_value", "write_formula",
                "fill_formula_down", "replace_formula",
                "delete_rows", "delete_columns", "delete_blank_rows",
                "clear_range", "delete_sheet",
                "remove_duplicates", "clean_data", "merge_cells", "unmerge_cells",
                "sort_data", "split_text_to_columns", "fill_blank_cells",
                "remove_special_chars", "clean_amount", "merge_columns",
                "rename_columns", "collapse_spaces", "text_to_number",
                "unify_date", "trim_spaces", "rollback"
            };

        /// <summary>
        /// Below this, a preview is more friction than protection: the user can
        /// see the result immediately and undo it.
        /// </summary>
        public const int SmallChangeThreshold = 10;

        public static ConfirmationLevel LevelFor(string toolName)
        {
            if (string.IsNullOrEmpty(toolName))
            {
                return ConfirmationLevel.None;
            }
            if (Unpreviewable.ContainsKey(toolName))
            {
                return ConfirmationLevel.SnapshotAndWarn;
            }
            return NeedsPreview.Contains(toolName)
                ? ConfirmationLevel.Preview
                : ConfirmationLevel.None;
        }

        public static string ReasonFor(string toolName)
        {
            return toolName != null && Unpreviewable.TryGetValue(toolName, out var reason)
                ? reason
                : null;
        }

        /// <summary>
        /// Whether a computed preview is worth interrupting for.
        ///
        /// A preview that changes nothing, or touches a couple of cells, is not
        /// worth a dialog. Structural changes and formula overwrites always are,
        /// however few cells they touch: deleting one row or replacing one
        /// formula is exactly the kind of small edit that is hard to notice and
        /// expensive to discover later.
        /// </summary>
        public static bool ShouldInterrupt(ChangePreview preview)
        {
            if (preview == null)
            {
                return false;
            }
            if (!preview.Previewable)
            {
                return true;
            }
            if (preview.IsNoOp)
            {
                return false;
            }
            if (preview.IsStructural || preview.FormulasOverwritten > 0)
            {
                return true;
            }
            return preview.AffectedCells >= SmallChangeThreshold;
        }
    }
}
