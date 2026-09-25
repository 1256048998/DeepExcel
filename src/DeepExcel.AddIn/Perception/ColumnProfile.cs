using System;
using System.Collections.Generic;
using System.Globalization;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// What a column holds. Deliberately coarse: the model needs to know whether
    /// something is a date or a number, not which of eleven numeric formats it is.
    /// </summary>
    public enum ColumnKind
    {
        Empty,
        Text,
        Number,
        Date,
        Boolean,
        Formula,
        /// <summary>No type accounts for most values. A warning sign for the model.</summary>
        Mixed
    }

    /// <summary>
    /// One column's inferred shape.
    ///
    /// The point of this type is to replace "read 8000 rows and hope the model
    /// works it out" with about 40 tokens that say what the column actually is.
    /// </summary>
    public sealed class ColumnProfile
    {
        /// <summary>Column letter, e.g. "C".</summary>
        public string Letter { get; set; }

        /// <summary>Header text, or null when the sheet has no header row.</summary>
        public string Header { get; set; }

        public ColumnKind Kind { get; set; }

        /// <summary>
        /// Share of non-empty sampled cells matching <see cref="Kind"/>, 0..1.
        ///
        /// Anything below <see cref="LowConfidenceThreshold"/> is surfaced to the
        /// model so it knows to read the data rather than trust this summary.
        /// Silently presenting a wrong type would be worse than presenting
        /// nothing: the model would act on it.
        /// </summary>
        public double Confidence { get; set; }

        public int NonEmptyCount { get; set; }
        public int EmptyCount { get; set; }

        /// <summary>Row numbers of blank cells, capped; enough to point at the problem.</summary>
        public List<int> EmptyRows { get; set; } = new List<int>();

        /// <summary>Distinct value count, capped at <see cref="DistinctCap"/>.</summary>
        public int DistinctCount { get; set; }

        /// <summary>True when the cap was hit, so DistinctCount is a floor.</summary>
        public bool DistinctCapped { get; set; }

        public bool IsUnique { get; set; }

        /// <summary>Numeric or date extent, formatted for display.</summary>
        public string MinDisplay { get; set; }
        public string MaxDisplay { get; set; }

        /// <summary>A couple of real values, so the model can see the shape.</summary>
        public List<string> Samples { get; set; } = new List<string>();

        /// <summary>First formula seen, normalised to R1C1-free A1 form.</summary>
        public string FormulaSample { get; set; }

        public const double LowConfidenceThreshold = 0.9;
        public const int DistinctCap = 1000;

        public bool IsLowConfidence
        {
            get { return Kind != ColumnKind.Empty && Confidence < LowConfidenceThreshold; }
        }

        /// <summary>
        /// One line for the prompt. Kept terse on purpose: this runs once per
        /// column on every request, so wording costs tokens forever.
        /// </summary>
        /// <param name="sampled">这张表只采样了部分行：取值范围、唯一、不同值个数、空值都只是样本结论，
        /// 要明说，否则模型会据此断言「全表只有 3 个空值」</param>
        public string Render(bool sampled = false)
        {
            var parts = new List<string>();
            parts.Add(Letter);
            parts.Add(string.IsNullOrEmpty(Header) ? "(无表头)" : Header);
            parts.Add(KindLabel(Kind));

            var hasFormulaSample = Kind == ColumnKind.Formula && !string.IsNullOrEmpty(FormulaSample);
            if (hasFormulaSample)
            {
                parts.Add(FormulaSample);
            }
            // 公式样例不是统计量；统计量（范围、唯一、不同值、空值）从这里开始
            var statsStart = parts.Count;
            if (!hasFormulaSample && !string.IsNullOrEmpty(MinDisplay) && !string.IsNullOrEmpty(MaxDisplay))
            {
                parts.Add(MinDisplay + " ~ " + MaxDisplay);
            }

            if (IsUnique && NonEmptyCount > 1)
            {
                parts.Add("唯一");
            }
            else if (Kind == ColumnKind.Text && DistinctCount > 0)
            {
                parts.Add(DistinctCount + (DistinctCapped ? "+" : "") + " 个不同值");
            }

            if (EmptyCount > 0)
            {
                var where = EmptyRows.Count > 0
                    ? "(行 " + string.Join(",", EmptyRows.ToArray()) + (EmptyCount > EmptyRows.Count ? "…" : "") + ")"
                    : "";
                parts.Add(EmptyCount + " 个空值" + where);
            }

            if (sampled && parts.Count > statsStart)
            {
                parts[statsStart] = "样本: " + parts[statsStart];
            }

            if (Samples.Count > 0 && Kind == ColumnKind.Text)
            {
                parts.Add("样例: " + string.Join(" / ", Samples.ToArray()));
            }

            if (IsLowConfidence)
            {
                // The model should read the data instead of trusting this row.
                parts.Add("⚠类型不一致，建议实读");
            }

            return string.Join("  ", parts.ToArray());
        }

        internal static string KindLabel(ColumnKind kind)
        {
            switch (kind)
            {
                case ColumnKind.Text: return "text";
                case ColumnKind.Number: return "number";
                case ColumnKind.Date: return "date";
                case ColumnKind.Boolean: return "bool";
                case ColumnKind.Formula: return "formula";
                case ColumnKind.Mixed: return "mixed";
                default: return "empty";
            }
        }

        internal static string FormatValue(object value)
        {
            if (value == null) return "";
            if (value is DateTime dt)
            {
                return dt.TimeOfDay == TimeSpan.Zero
                    ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : dt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
            if (value is double d)
            {
                return d == Math.Floor(d) && Math.Abs(d) < 1e15
                    ? ((long)d).ToString(CultureInfo.InvariantCulture)
                    : d.ToString("0.####", CultureInfo.InvariantCulture);
            }
            if (value is bool b) return b ? "TRUE" : "FALSE";

            var text = value.ToString();
            // Long free text would blow the token budget for no benefit.
            return text.Length > 24 ? text.Substring(0, 24) + "…" : text;
        }
    }
}
