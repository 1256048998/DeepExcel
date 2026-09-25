using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DeepExcel.AddIn.Perception
{
    public sealed class SheetIndex
    {
        public string Name { get; set; }
        public int FirstRow { get; set; } = 1;
        public int LastRow { get; set; }
        public int HeaderRow { get; set; }
        public bool HasHeader { get; set; }
        public List<ColumnProfile> Columns { get; set; } = new List<ColumnProfile>();

        /// <summary>True when the profile came from a sample, not every row.</summary>
        public bool Sampled { get; set; }
        public int SampledRows { get; set; }
    }

    public sealed class NamedRangeIndex
    {
        public string Name { get; set; }
        public string RefersTo { get; set; }
    }

    /// <summary>
    /// A cross-sheet link, inferred from a formula that references another sheet.
    ///
    /// Only stated when an actual formula reference was found. Guessing a
    /// relationship from matching column names would invent joins that are not
    /// there, and the model would act on them.
    /// </summary>
    public sealed class SheetRelation
    {
        public string FromSheet { get; set; }
        public string ToSheet { get; set; }
        public string ViaColumn { get; set; }
    }

    /// <summary>
    /// The workbook as the model should see it.
    ///
    /// This replaces the current approach, where the model calls read_range and
    /// infers structure from raw rows. That fails in three ways at once on a real
    /// workbook: it cannot read 8000 rows, the rows it does read blow the token
    /// budget, and the structure it infers is often wrong. The index costs a few
    /// hundred tokens and states the structure outright.
    /// </summary>
    public sealed class WorkbookIndex
    {
        public string WorkbookName { get; set; }
        public string ActiveSheet { get; set; }
        public List<SheetIndex> Sheets { get; set; } = new List<SheetIndex>();
        public List<NamedRangeIndex> NamedRanges { get; set; } = new List<NamedRangeIndex>();
        public List<SheetRelation> Relations { get; set; } = new List<SheetRelation>();

        public DateTime BuiltAtUtc { get; set; } = DateTime.UtcNow;

        /// <summary>Content hash the index was built from; drives cache invalidation.</summary>
        public string SourceHash { get; set; }

        /// <summary>
        /// Set when indexing was abandoned, e.g. it exceeded its time budget.
        /// The model is told so it falls back to reading rather than trusting a
        /// partial picture.
        /// </summary>
        public string Incomplete { get; set; }

        /// <summary>
        /// The prompt form.
        ///
        /// Every character here is paid for on every request, so the format is
        /// terse and positional rather than labelled. A budget is enforced: a
        /// 40-sheet workbook must not crowd out the conversation itself.
        /// </summary>
        public string Render(int maxCharacters = 6000)
        {
            var builder = new StringBuilder();
            builder.Append("## 工作簿结构：").Append(WorkbookName ?? "(未命名)").AppendLine();
            if (!string.IsNullOrEmpty(Incomplete))
            {
                builder.Append("注意：索引不完整（").Append(Incomplete)
                       .AppendLine("）。涉及未列出的区域时请先实读数据。");
            }

            foreach (var sheet in Sheets)
            {
                if (builder.Length > maxCharacters)
                {
                    builder.AppendLine("…（其余工作表已省略，需要时请实读）");
                    break;
                }
                RenderSheet(builder, sheet);
            }

            if (NamedRanges.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("命名区域：");
                foreach (var range in NamedRanges.Take(20))
                {
                    builder.Append("  ").Append(range.Name).Append(" = ").AppendLine(range.RefersTo);
                }
            }

            if (Relations.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("跨表引用：");
                foreach (var relation in Relations.Take(20))
                {
                    builder.Append("  ").Append(relation.FromSheet)
                           .Append(" → ").Append(relation.ToSheet);
                    if (!string.IsNullOrEmpty(relation.ViaColumn))
                    {
                        builder.Append("（经 ").Append(relation.ViaColumn).Append(" 列）");
                    }
                    builder.AppendLine();
                }
            }

            builder.AppendLine();
            builder.AppendLine(
                "以上为结构摘要，非完整数据。标注「⚠类型不一致」的列请先实读确认；" +
                "需要具体单元格值时请调用 read_range。");
            return builder.ToString();
        }

        private static void RenderSheet(StringBuilder builder, SheetIndex sheet)
        {
            builder.AppendLine();
            builder.Append("### ").Append(sheet.Name);
            if (sheet.LastRow > 0)
            {
                builder.Append(" (").Append(sheet.FirstRow).Append("..").Append(sheet.LastRow).Append(" 行");
                if (sheet.Sampled)
                {
                    builder.Append("，采样 ").Append(sheet.SampledRows).Append(" 行");
                }
                builder.Append(")");
            }
            builder.AppendLine();

            if (sheet.HasHeader)
            {
                builder.Append("表头行: ").Append(sheet.HeaderRow).AppendLine();
            }
            else
            {
                builder.AppendLine("表头行: 无");
            }

            foreach (var column in sheet.Columns)
            {
                if (column.Kind == ColumnKind.Empty && string.IsNullOrEmpty(column.Header))
                {
                    // An empty unnamed column tells the model nothing.
                    continue;
                }
                builder.Append("  ").AppendLine(column.Render(sheet.Sampled));
            }
        }

        /// <summary>Rough token cost, for logging and budget checks.</summary>
        public int EstimateTokens()
        {
            // Mixed CJK/ASCII lands near 2 characters per token in practice.
            return Render().Length / 2;
        }
    }
}
