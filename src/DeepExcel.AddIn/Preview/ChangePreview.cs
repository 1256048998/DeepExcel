using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DeepExcel.AddIn.Preview
{
    public enum ChangeKind
    {
        /// <summary>An empty cell gains a value.</summary>
        Add,
        /// <summary>An existing value is replaced.</summary>
        Overwrite,
        /// <summary>A value is removed.</summary>
        Clear,
        /// <summary>Only the number format or style changes.</summary>
        Format,
        /// <summary>The whole row or column goes.</summary>
        Delete
    }

    public sealed class CellChange
    {
        public string Address { get; set; }
        public string Before { get; set; }
        public string After { get; set; }
        public ChangeKind Kind { get; set; }

        /// <summary>
        /// True when the cell currently holds a formula.
        ///
        /// Called out separately because overwriting a formula with a literal is
        /// the destructive edit users least expect and most regret: the number
        /// looks right immediately and stops updating forever.
        /// </summary>
        public bool OverwritesFormula { get; set; }

        public override string ToString()
        {
            switch (Kind)
            {
                case ChangeKind.Delete:
                    return Address + "  → 删除";
                case ChangeKind.Format:
                    return Address + "  " + Before + " → " + After + "  (格式)";
                default:
                    return Address + "  " + (string.IsNullOrEmpty(Before) ? "(空)" : Before)
                           + " → " + (string.IsNullOrEmpty(After) ? "(空)" : After);
            }
        }
    }

    /// <summary>
    /// What an operation will do, computed before it runs.
    ///
    /// The product currently works the other way round: it executes, then offers
    /// a rollback. That is why people will not point it at a workbook that
    /// matters. Showing the change set first moves the decision to before the
    /// damage instead of after it.
    ///
    /// <see cref="Previewable"/> being false is a first-class outcome, not a
    /// failure. VBA and Python cannot be simulated, and pretending otherwise
    /// would be worse than admitting it -- the user would trust a preview that
    /// does not cover what actually runs.
    /// </summary>
    public sealed class ChangePreview
    {
        public string ToolName { get; set; }

        /// <summary>False when the operation cannot be simulated.</summary>
        public bool Previewable { get; set; } = true;

        /// <summary>Why it cannot be previewed, shown to the user verbatim.</summary>
        public string NotPreviewableReason { get; set; }

        public List<CellChange> Changes { get; set; } = new List<CellChange>();

        /// <summary>Total affected cells, which may exceed <see cref="Changes"/>.</summary>
        public int AffectedCells { get; set; }

        public List<int> DeletedRows { get; set; } = new List<int>();
        public List<string> DeletedColumns { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();

        /// <summary>Rows/columns/sheets whose structure changes, not just values.</summary>
        public bool IsStructural
        {
            get { return DeletedRows.Count > 0 || DeletedColumns.Count > 0; }
        }

        public int FormulasOverwritten
        {
            get { return Changes.Count(c => c.OverwritesFormula); }
        }

        /// <summary>Shown in the UI when the list is truncated.</summary>
        public const int MaxListedChanges = 50;

        public bool IsTruncated
        {
            get { return AffectedCells > Changes.Count; }
        }

        /// <summary>
        /// True when nothing actually changes.
        ///
        /// Worth surfacing: a no-op usually means the model misread the target,
        /// and confirming it teaches the user the tool did nothing rather than
        /// leaving them to wonder.
        /// </summary>
        public bool IsNoOp
        {
            get { return Previewable && AffectedCells == 0 && !IsStructural; }
        }

        public string Summary()
        {
            if (!Previewable)
            {
                return "无法预览此操作" +
                       (string.IsNullOrEmpty(NotPreviewableReason) ? "" : "：" + NotPreviewableReason);
            }
            if (IsNoOp)
            {
                return "此操作不会改变任何内容";
            }

            var parts = new List<string>();
            if (AffectedCells > 0)
            {
                parts.Add("将修改 " + AffectedCells + " 个单元格");
            }
            if (DeletedRows.Count > 0)
            {
                parts.Add("删除 " + DeletedRows.Count + " 行");
            }
            if (DeletedColumns.Count > 0)
            {
                parts.Add("删除 " + DeletedColumns.Count + " 列");
            }
            return string.Join("、", parts.ToArray());
        }

        public string Render()
        {
            var builder = new StringBuilder();
            builder.AppendLine(Summary());

            if (!Previewable)
            {
                builder.AppendLine("已自动创建快照，执行后可回滚。");
                return builder.ToString();
            }

            foreach (var change in Changes.Take(MaxListedChanges))
            {
                builder.Append("  ").AppendLine(change.ToString());
            }
            if (IsTruncated)
            {
                builder.Append("  …其余 ").Append(AffectedCells - Changes.Count).AppendLine(" 处未列出");
            }

            foreach (var warning in Warnings)
            {
                builder.Append("  ⚠ ").AppendLine(warning);
            }
            return builder.ToString();
        }

        /// <summary>An operation that cannot be simulated.</summary>
        public static ChangePreview NotPreviewable(string toolName, string reason)
        {
            return new ChangePreview
            {
                ToolName = toolName,
                Previewable = false,
                NotPreviewableReason = reason
            };
        }
    }
}
