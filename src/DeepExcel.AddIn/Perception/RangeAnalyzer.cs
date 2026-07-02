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
                    single[0, 0] = range.NumberFormat?.ToString() ?? "";
                    return single;
                }

                // NumberFormat 返回的是 object[,]（和 Value2 一样），
                // 不是 string[,]，直接 as string[,] 会返回 null。
                // 需要手动转换。
                var formats = range.NumberFormat as object[,];
                if (formats == null)
                {
                    // 多区域（Areas.Count > 1）或其他情况，返回空数组
                    return new string[0, 0];
                }

                int rows = formats.GetLength(0);
                int cols = formats.GetLength(1);
                var result = new string[rows, cols];
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        // Excel COM 数组是 1-based
                        result[i, j] = formats[i + 1, j + 1]?.ToString() ?? "";
                    }
                }
                return result;
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
                object merged = range.MergeCells;

                // 快速路径：没有任何合并单元格
                if (merged is bool b && !b)
                {
                    return result.ToArray();
                }

                // 整个区域是一个合并单元格
                if (merged is bool b2 && b2)
                {
                    result.Add(new MergeCellInfo
                    {
                        Address = range.Address,
                        RowCount = range.Rows.Count,
                        ColumnCount = range.Columns.Count
                    });
                    return result.ToArray();
                }

                // 混合情况：智能扫描
                // 策略：
                // 1. 小区域（<= 500 单元格）：全量扫描
                // 2. 大区域（> 500 单元格）：扫描第一行 + 最后一行 + 第一列 + 最后一列 + 每隔 N 行抽一行
                //    这样能覆盖标题、表尾、左侧表头的合并单元格，避免漏报
                int totalCells = range.Rows.Count * range.Columns.Count;
                var seen = new HashSet<string>();

                if (totalCells <= 500)
                {
                    // 小区域：全量扫描
                    foreach (Range cell in range.Cells)
                    {
                        ScanCellForMerge(cell, seen, result);
                    }
                }
                else
                {
                    // 大区域：智能采样扫描
                    int rows = range.Rows.Count;
                    int cols = range.Columns.Count;

                    // 1. 扫描第一行（标题行通常有合并）
                    for (int c = 1; c <= cols; c++)
                    {
                        Range cell = null;
                        try { cell = (Range)range.Cells[1, c]; } catch { }
                        if (cell != null) ScanCellForMerge(cell, seen, result);
                    }

                    // 2. 扫描最后一行
                    if (rows > 1)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            Range cell = null;
                            try { cell = (Range)range.Cells[rows, c]; } catch { }
                            if (cell != null) ScanCellForMerge(cell, seen, result);
                        }
                    }

                    // 3. 扫描第一列
                    if (cols > 1)
                    {
                        for (int r = 2; r < rows; r++)
                        {
                            Range cell = null;
                            try { cell = (Range)range.Cells[r, 1]; } catch { }
                            if (cell != null) ScanCellForMerge(cell, seen, result);
                        }
                    }

                    // 4. 扫描最后一列
                    if (cols > 1 && rows > 2)
                    {
                        for (int r = 2; r < rows; r++)
                        {
                            Range cell = null;
                            try { cell = (Range)range.Cells[r, cols]; } catch { }
                            if (cell != null) ScanCellForMerge(cell, seen, result);
                        }
                    }

                    // 5. 每隔 20 行抽一行扫描（中间区域采样）
                    if (rows > 10)
                    {
                        int step = rows > 100 ? 20 : 10;
                        for (int r = 5; r < rows - 1; r += step)
                        {
                            for (int c = 1; c <= cols; c++)
                            {
                                Range cell = null;
                                try { cell = (Range)range.Cells[r, c]; } catch { }
                                if (cell != null) ScanCellForMerge(cell, seen, result);
                            }
                        }
                    }
                }
            }
            catch
            {
                // 某些区域无法访问 MergeCells 属性
            }

            return result.ToArray();
        }

        private void ScanCellForMerge(Range cell, HashSet<string> seen, List<MergeCellInfo> result)
        {
            try
            {
                if (cell == null) return;
                if (cell.MergeCells is bool mc && mc)
                {
                    string addr = cell.MergeArea.Address;
                    if (seen.Add(addr))
                    {
                        result.Add(new MergeCellInfo
                        {
                            Address = addr,
                            RowCount = cell.MergeArea.Rows.Count,
                            ColumnCount = cell.MergeArea.Columns.Count
                        });
                    }
                }
            }
            catch { }
        }

        private string SafeGet(Func<string> getter, string defaultValue = "")
        {
            try { return getter() ?? defaultValue; }
            catch { return defaultValue; }
        }
    }
}
