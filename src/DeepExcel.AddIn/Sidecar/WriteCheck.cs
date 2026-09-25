using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>一个公式错误单元格。</summary>
    public class ErrorCell
    {
        public string Sheet { get; set; }
        public string Address { get; set; }
        /// <summary>#DIV/0!、#REF! 等</summary>
        public string Text { get; set; }

        public string Key => (Sheet ?? "").ToUpperInvariant() + "!" + (Address ?? "").ToUpperInvariant();
        public override string ToString() => QualifiedAddress(Sheet, Address) + " " + Text;

        internal static string QualifiedAddress(string sheet, string address)
        {
            if (string.IsNullOrEmpty(sheet)) return address;
            var quoted = sheet.IndexOfAny(new[] { ' ', '-', '!', '\'' }) >= 0 ? "'" + sheet.Replace("'", "''") + "'" : sheet;
            return quoted + "!" + address;
        }
    }

    /// <summary>
    /// 一次工作簿体检：全部表上的公式错误单元格 + 外部链接。
    /// 由 IExcelActions.CaptureHealth 采集（SpecialCells 是 Excel 原生查找，整本扫一遍通常几毫秒）。
    /// </summary>
    public class HealthSnapshot
    {
        /// <summary>公式错误单元格总数（真实数量，可能多于 Errors 里收集到的）</summary>
        public int ErrorCount { get; set; }
        /// <summary>收集到的错误单元格（有上限，见 Truncated）</summary>
        public List<ErrorCell> Errors { get; set; } = new List<ErrorCell>();
        /// <summary>错误太多只收集了一部分：此时按数量差判断新增</summary>
        public bool Truncated { get; set; }
        /// <summary>Workbook.LinkSources：外部工作簿链接</summary>
        public List<string> ExternalLinks { get; set; } = new List<string>();
        /// <summary>计算模式为手动：回读的值可能不是最新的</summary>
        public bool CalculationManual { get; set; }
    }

    /// <summary>写入后回读的一个单元格。</summary>
    public class CellSample
    {
        [JsonPropertyName("address")]
        public string Address { get; set; }

        [JsonPropertyName("value")]
        public string Value { get; set; }

        [JsonPropertyName("formula")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string Formula { get; set; }
    }

    /// <summary>附在工具结果里的体检结论（模型看得到每个字段）。</summary>
    public class WriteVerification
    {
        /// <summary>一句话结论：模型据此决定是继续修还是可以汇报</summary>
        [JsonPropertyName("summary")]
        public string Summary { get; set; }

        [JsonPropertyName("ok")]
        public bool Ok { get; set; }

        [JsonPropertyName("formula_errors_before")]
        public int FormulaErrorsBefore { get; set; }

        [JsonPropertyName("formula_errors_after")]
        public int FormulaErrorsAfter { get; set; }

        [JsonPropertyName("new_errors")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<string> NewErrors { get; set; }

        [JsonPropertyName("new_external_links")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<string> NewExternalLinks { get; set; }

        [JsonPropertyName("samples")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<CellSample> Samples { get; set; }
    }

    /// <summary>
    /// 写后自动体检（Claude Code 改完代码跑测试 / lint，而不是「我觉得改好了」）。
    ///
    /// 每次会改工作簿的工具执行前后各拍一次 HealthSnapshot，比较出新增的公式错误和外部链接；
    /// 写公式的工具再回读几个计算结果。结论附在工具返回里，模型拿到的是「做完了，这是验证
    /// 结果」，而不是「调用成功」。公式引用了不存在的表时，Excel 会把它当成外部文件，所以
    /// 「新增外部链接」同时覆盖了「引用不存在的表」。
    ///
    /// 纯逻辑，不碰 COM。
    /// </summary>
    public static class WriteCheck
    {
        public const int MaxListed = 5;

        /// <summary>只改外观、不会让公式出错的工具：不做体检，省两次整本扫描。</summary>
        public static readonly HashSet<string> CosmeticTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "set_number_format", "set_column_width", "set_cell_style",
            "merge_cells", "unmerge_cells", "apply_conditional_format", "freeze_panes",
            "highlight_duplicates", "filter_data",
            "create_chart", "create_combo_chart", "add_data_labels", "set_chart_title", "set_chart_colors",
            "set_pivot_value_display", "set_pivot_totals", "add_pivot_slicer",
        };

        /// <summary>写公式的工具：体检时回读几个计算结果。</summary>
        public static readonly HashSet<string> FormulaTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "write_formula", "fill_formula_down", "replace_formula", "copy_range",
        };

        public static bool NeedsCheck(string toolName) =>
            ToolMutationPolicy.IsMutating(toolName) && !CosmeticTools.Contains(toolName ?? "");

        public static WriteVerification Compare(HealthSnapshot before, HealthSnapshot after, List<CellSample> samples)
        {
            if (before == null || after == null) return null;

            var known = new HashSet<string>(before.Errors.Select(e => e.Key));
            // 写入前的错误只收集了一部分时，「不在名单里」不代表是新的：不逐个列出
            var fresh = before.Truncated
                ? new List<ErrorCell>()
                : after.Errors.Where(e => !known.Contains(e.Key)).ToList();
            // 只收集了一部分时逐格比较不可靠，按数量差算
            var newCount = before.Truncated || after.Truncated
                ? Math.Max(0, after.ErrorCount - before.ErrorCount)
                : fresh.Count;

            var linksBefore = new HashSet<string>(before.ExternalLinks ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            var newLinks = (after.ExternalLinks ?? new List<string>()).Where(l => !linksBefore.Contains(l)).Distinct().ToList();

            var parts = new List<string>();
            if (newCount > 0)
            {
                var listed = string.Join("、", fresh.Take(MaxListed));
                parts.Add($"新增 {newCount} 个公式错误" + (listed.Length > 0 ? "：" + listed + (newCount > MaxListed ? " 等" : "") : "")
                    + "。先查明原因并修正，不要直接宣布完成");
            }
            if (newLinks.Count > 0)
            {
                parts.Add($"新增外部链接 {string.Join("、", newLinks.Take(MaxListed))}——公式可能引用了不存在的工作表或外部文件，请核对表名");
            }
            if (after.CalculationManual)
            {
                parts.Add("Excel 计算模式为手动，回读的值可能不是最新结果");
            }

            var ok = newCount == 0 && newLinks.Count == 0;
            string summary;
            if (ok)
            {
                var errorsPart = after.ErrorCount < before.ErrorCount
                    ? $"公式错误 {before.ErrorCount}→{after.ErrorCount}"
                    : after.ErrorCount == 0 ? "无公式错误" : $"没有新增公式错误（原有 {after.ErrorCount} 个）";
                summary = "体检通过：" + errorsPart + "，无新增外部链接";
                if (parts.Count > 0) summary += "。" + string.Join("。", parts);
            }
            else
            {
                summary = "体检发现问题：" + string.Join("。", parts);
            }

            return new WriteVerification
            {
                Summary = summary,
                Ok = ok,
                FormulaErrorsBefore = before.ErrorCount,
                FormulaErrorsAfter = after.ErrorCount,
                NewErrors = fresh.Count > 0 ? fresh.Take(MaxListed).Select(e => e.ToString()).ToList() : null,
                NewExternalLinks = newLinks.Count > 0 ? newLinks : null,
                Samples = samples != null && samples.Count > 0 ? samples : null,
            };
        }
    }
}
