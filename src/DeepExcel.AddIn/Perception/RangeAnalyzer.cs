using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// 区域分析器 - 读取单元格/区域详细信息
    /// </summary>
    public class RangeAnalyzer
    {
        public RangeInfo Analyze(Range range)
        {
            if (range == null) return null;

            return new RangeInfo
            {
                Address = SafeGet(() => range.Address),
                WorksheetName = SafeGet(() => range.Worksheet.Name),
                RowCount = range.Rows.Count,
                ColumnCount = range.Columns.Count,
                Values = GetValues(range),
                Formulas = GetFormulas(range),
                NumberFormats = GetNumberFormats(range),
                FormatConditions = AnalyzeFormatConditions(range),
                MergeCells = AnalyzeMergeCells(range)
            };
        }

        private object[,] GetValues(Range range)
        {
            try
            {
                if (range.Cells.Count == 1)
                {
                    var single = new object[1, 1];
                    single[0, 0] = range.Value2;
                    return single;
                }
                return range.Value2 as object[,];
            }
            catch
            {
                return new object[0, 0];
            }
        }

        private string[,] GetFormulas(Range range)
        {
            try
            {
                if (range.Cells.Count == 1)
                {
                    var single = new string[1, 1];
                    single[0, 0] = range.Formula.ToString();
                    return single;
                }

                var formulas = range.Formula as object[,];
                if (formulas == null) return new string[0, 0];

                var result = new string[formulas.GetLength(0), formulas.GetLength(1)];
                for (int i = 0; i < formulas.GetLength(0); i++)
                {
                    for (int j = 0; j < formulas.GetLength(1); j++)
                    {
                        result[i, j] = formulas[i + 1, j + 1]?.ToString() ?? "";
                    }
                }
                return result;
            }
            catch
            {
                return new string[0, 0];
            }
        }

        private string[,] GetNumberFormats(Range range)
        {
            try
            {
                if (range.Cells.Count == 1)
                {
                    var single = new string[1, 1];
                    single[0, 0] = range.NumberFormat.ToString();
                    return single;
                }
                return range.NumberFormat as string[,];
            }
            catch
            {
                return new string[0, 0];
            }
        }

        private FormatConditionInfo[] AnalyzeFormatConditions(Range range)
        {
            var result = new List<FormatConditionInfo>();

            try
            {
                var formats = range.FormatConditions;
                foreach (FormatCondition fc in formats)
                {
                    result.Add(new FormatConditionInfo
                    {
                        Type = fc.Type.ToString(),
                        Operator = fc.Operator.ToString(),
                        Formula1 = SafeGet(() => fc.Formula1?.ToString()),
                        Formula2 = SafeGet(() => fc.Formula2?.ToString())
                    });
                }
            }
            catch
            {
                // 某些情况下没有条件格式
            }

            return result.ToArray();
        }

        private MergeCellInfo[] AnalyzeMergeCells(Range range)
        {
            var result = new List<MergeCellInfo>();

            try
            {
                var mergeAreas = range.Areas;
                foreach (Range area in mergeAreas)
                {
                    result.Add(new MergeCellInfo
                    {
                        Address = area.Address,
                        RowCount = area.Rows.Count,
                        ColumnCount = area.Columns.Count
                    });
                }
            }
            catch
            {
                // 某些区域没有合并单元格
            }

            return result.ToArray();
        }

        private string SafeGet(Func<string> getter, string defaultValue = "")
        {
            try { return getter() ?? defaultValue; }
            catch { return defaultValue; }
        }
    }
}
