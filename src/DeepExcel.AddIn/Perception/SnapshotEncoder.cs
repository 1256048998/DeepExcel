using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// sheet_snapshot 的纯逻辑部分：窗口大小、单元格值编码、动态数组识别。
    /// 快照交给侧车的 perception 包分析（src/DeepExcel.Sidecar/perception/grid.py 描述了格式），
    /// 宿主只负责一次批量读出、不做任何判断。
    /// </summary>
    public static class SnapshotEncoder
    {
        public const int DefaultMaxCells = 60000;
        public const int HardMaxCells = 100000;
        public const int MaxColumns = 100;
        public const int MaxFormulas = 20000;
        public const int MaxMerges = 500;
        public const int MaxSpillChecks = 500;
        public const int MaxTextLength = 60;
        public const int MaxFormulaLength = 300;

        /// <summary>快照窗口：列最多 MaxColumns，行数按格数上限折算</summary>
        public static (int Rows, int Columns) PlanWindow(int totalRows, int totalColumns, int maxCells)
        {
            if (totalRows <= 0 || totalColumns <= 0) return (0, 0);
            if (maxCells <= 0) maxCells = DefaultMaxCells;
            maxCells = Math.Min(maxCells, HardMaxCells);
            var columns = Math.Min(totalColumns, MaxColumns);
            var rows = Math.Min(totalRows, Math.Max(1, maxCells / columns));
            return (rows, columns);
        }

        /// <summary>
        /// Range.Value（不是 Value2，才能分出日期）的一格编码成 JSON 友好的值：
        /// null / double / bool / 截断的文本 / {"d": 日期} / {"e": 错误文本}。
        /// COM 里数字总是 double，错误值是 Int32（CVErr），日期是 DateTime。
        /// </summary>
        public static object EncodeValue(object value)
        {
            switch (value)
            {
                case null:
                    return null;
                case DBNull _:
                    return null;
                case double d:
                    return double.IsNaN(d) || double.IsInfinity(d) ? null : (object)d;
                case bool b:
                    return b;
                case int code:
                    return new Dictionary<string, string> { ["e"] = ErrorText(code) };
                case DateTime dt:
                    var text = dt.TimeOfDay == TimeSpan.Zero
                        ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        : dt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
                    return new Dictionary<string, string> { ["d"] = text };
                case decimal m:
                    return (double)m;
                case string s:
                    if (s.Length == 0) return null;
                    return s.Length > MaxTextLength ? s.Substring(0, MaxTextLength) + "…" : s;
                default:
                    var other = Convert.ToString(value, CultureInfo.InvariantCulture);
                    return string.IsNullOrEmpty(other) ? null : other;
            }
        }

        /// <summary>CVErr 的 Int32 值 = -2146826288 + (xlErr 常量 - 2000)</summary>
        public static string ErrorText(int code)
        {
            switch (code + 2146828288)
            {
                case 2000: return "#NULL!";
                case 2007: return "#DIV/0!";
                case 2015: return "#VALUE!";
                case 2023: return "#REF!";
                case 2029: return "#NAME?";
                case 2036: return "#NUM!";
                case 2042: return "#N/A";
                case 2043: return "#GETTING_DATA";
                case 2045: return "#SPILL!";
                case 2046: return "#CONNECT!";
                case 2047: return "#BLOCKED!";
                case 2048: return "#UNKNOWN!";
                case 2049: return "#FIELD!";
                case 2050: return "#CALC!";
                default: return "#ERROR";
            }
        }

        public static string ClipFormula(string formula)
        {
            if (formula == null) return null;
            return formula.Length > MaxFormulaLength ? formula.Substring(0, MaxFormulaLength) + "…" : formula;
        }

        private static readonly Regex DynamicArrayFunctions = new Regex(
            @"\b(?:_xlfn\.)?(?:FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|RANDARRAY|XLOOKUP|TEXTSPLIT|VSTACK|HSTACK|TOCOL|TOROW|" +
            @"WRAPROWS|WRAPCOLS|TAKE|DROP|CHOOSEROWS|CHOOSECOLS|EXPAND|MAKEARRAY|MAP|BYROW|BYCOL|SCAN|TRANSPOSE|FREQUENCY|MMULT)\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>只对可能溢出的公式去问 HasSpill：逐格跨 COM 问太慢</summary>
        public static bool MaySpill(string formula)
        {
            return !string.IsNullOrEmpty(formula) && DynamicArrayFunctions.IsMatch(formula);
        }
    }
}
