using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>一个单元格的改动：原值 → 新值（公式格显示公式）。</summary>
    public class CellChange
    {
        [JsonPropertyName("address")]
        public string Address { get; set; }

        [JsonPropertyName("before")]
        public string Before { get; set; }

        [JsonPropertyName("after")]
        public string After { get; set; }
    }

    /// <summary>一次写入实际改了什么（面板的内联 diff；模型也看得到）。</summary>
    public class CellChanges
    {
        /// <summary>内容真的变了的单元格数</summary>
        [JsonPropertyName("changed")]
        public int Changed { get; set; }

        /// <summary>比较过的单元格数（目标区域大小）</summary>
        [JsonPropertyName("cells")]
        public int Cells { get; set; }

        [JsonPropertyName("sheet")]
        public string Sheet { get; set; }

        /// <summary>前几处改动；Changed 多于这里的条数时面板写「另有 N 格」</summary>
        [JsonPropertyName("samples")]
        public List<CellChange> Samples { get; set; } = new List<CellChange>();
    }

    /// <summary>
    /// 内联 diff：写入前后各读一次目标区域的 Formula（一次 COM 调用，公式格给公式、常量格给值），
    /// 逐格比较。纯逻辑，数组是 0 起始的 [行, 列]，和 target 左上角对齐。
    /// </summary>
    public static class CellDiff
    {
        /// <summary>超过这个格数不比较：两次整块读取在大区域上不值得</summary>
        public const int MaxCells = 20000;
        public const int MaxSamples = 8;
        private const int MaxTextLength = 80;

        public static CellChanges Compare(CellRect target, object[,] before, object[,] after, int maxSamples = MaxSamples)
        {
            if (before == null || after == null) return null;
            var rows = Math.Min(before.GetLength(0), after.GetLength(0));
            var cols = Math.Min(before.GetLength(1), after.GetLength(1));
            var result = new CellChanges { Cells = rows * cols, Sheet = target.Sheet };

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var was = Text(before[r, c]);
                    var now = Text(after[r, c]);
                    if (string.Equals(was, now, StringComparison.Ordinal)) continue;
                    result.Changed++;
                    if (result.Samples.Count < maxSamples)
                    {
                        result.Samples.Add(new CellChange
                        {
                            Address = CellRect.ColumnName(target.Col1 + c) + (target.Row1 + r),
                            Before = Clip(was),
                            After = Clip(now),
                        });
                    }
                }
            }
            return result;
        }

        private static string Text(object value)
        {
            if (value == null) return "";
            if (value is double d) return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
            return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "";
        }

        private static string Clip(string text) =>
            text.Length > MaxTextLength ? text.Substring(0, MaxTextLength) + "…" : text;
    }
}
