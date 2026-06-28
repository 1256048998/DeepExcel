using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Bridge;

namespace DeepExcel.AddIn.Tools
{
    /// <summary>
    /// 数据清洗工具集
    /// 支持：统一日期格式、删除重复行、标红缺失值、去除空格、类型转换等
    /// </summary>
    public class DataCleaner
    {
        private readonly Application _app;

        public DataCleaner(Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 统一日期格式
        /// </summary>
        public ToolResult UnifyDateFormats(string rangeAddress, string targetFormat = "yyyy-mm-dd")
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int convertedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    try
                    {
                        var val = cell.Value2;
                        if (val == null) continue;

                        // 尝试解析为日期
                        if (double.TryParse(val.ToString(), out double dateNum))
                        {
                            // 可能是Excel序列号日期
                            cell.NumberFormat = targetFormat;
                            convertedCount++;
                        }
                        else if (DateTime.TryParse(val.ToString(), out DateTime dt))
                        {
                            cell.Value = dt;
                            cell.NumberFormat = targetFormat;
                            convertedCount++;
                        }
                    }
                    catch
                    {
                        // 跳过无法转换的单元格
                    }
                }

                return new ToolResult
                {
                    Name = "unify_date_formats",
                    Success = true,
                    Data = new { convertedCount = convertedCount, targetFormat = targetFormat }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "unify_date_formats",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 删除重复行
        /// </summary>
        public ToolResult RemoveDuplicates(string rangeAddress, int[] columns = null)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int removedCount = 0;

                // 从最后一行往上遍历比较
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;

                var seenRows = new HashSet<string>();

                for (int i = rows; i >= 1; i--)
                {
                    var rowKey = GetRowKey(range, i, cols, columns);
                    if (seenRows.Contains(rowKey))
                    {
                        ((Range)range.Rows[i]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        removedCount++;
                    }
                    else
                    {
                        seenRows.Add(rowKey);
                    }
                }

                return new ToolResult
                {
                    Name = "remove_duplicates",
                    Success = true,
                    Data = new { removedCount = removedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "remove_duplicates",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 标红缺失值（空单元格）
        /// </summary>
        public ToolResult HighlightMissingValues(string rangeAddress, string colorHex = "#FFFF00")
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int highlightedCount = 0;
                var color = ColorTranslator.FromHtml(colorHex);
                var oleColor = ColorTranslator.ToOle(color);

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val == null || string.IsNullOrWhiteSpace(val.ToString()))
                    {
                        cell.Interior.Color = oleColor;
                        highlightedCount++;
                    }
                }

                return new ToolResult
                {
                    Name = "highlight_missing",
                    Success = true,
                    Data = new { highlightedCount = highlightedCount, color = colorHex }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "highlight_missing",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 去除文本首尾空格
        /// </summary>
        public ToolResult TrimSpaces(string rangeAddress)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int trimmedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val != null && val is string s && s != s.Trim())
                    {
                        cell.Value = s.Trim();
                        trimmedCount++;
                    }
                }

                return new ToolResult
                {
                    Name = "trim_spaces",
                    Success = true,
                    Data = new { trimmedCount = trimmedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "trim_spaces",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 统一大小写
        /// </summary>
        public ToolResult ChangeCase(string rangeAddress, string caseType = "proper")
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int changedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val != null && val is string s && !string.IsNullOrEmpty(s))
                    {
                        string newVal = caseType switch
                        {
                            "upper" => s.ToUpper(),
                            "lower" => s.ToLower(),
                            "proper" => ToProperCase(s),
                            _ => s
                        };

                        if (newVal != s)
                        {
                            cell.Value = newVal;
                            changedCount++;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "change_case",
                    Success = true,
                    Data = new { changedCount = changedCount, caseType = caseType }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "change_case",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 文本转数字
        /// </summary>
        public ToolResult TextToNumber(string rangeAddress)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int convertedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val != null && val is string s && !string.IsNullOrEmpty(s))
                    {
                        if (double.TryParse(s, out double num))
                        {
                            cell.Value = num;
                            cell.NumberFormat = "General";
                            convertedCount++;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "text_to_number",
                    Success = true,
                    Data = new { convertedCount = convertedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "text_to_number",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 一键综合清洗
        /// </summary>
        public CleanResult FullClean(string rangeAddress)
        {
            var result = new CleanResult { Operations = new List<string>() };

            var r1 = TrimSpaces(rangeAddress);
            if (r1.Success) result.Operations.Add($"去除空格: {(r1.Data as dynamic)?.trimmedCount}个");

            var r2 = UnifyDateFormats(rangeAddress);
            if (r2.Success) result.Operations.Add($"统一日期格式: {(r2.Data as dynamic)?.convertedCount}个");

            var r3 = HighlightMissingValues(rangeAddress);
            if (r3.Success) result.Operations.Add($"标红缺失值: {(r3.Data as dynamic)?.highlightedCount}个");

            result.Success = true;
            return result;
        }

        private string GetRowKey(Range range, int rowIndex, int colCount, int[] columns)
        {
            var key = "";
            int startCol = range.Column;
            int startRow = range.Row;
            int row = startRow + rowIndex - 1;

            if (columns != null)
            {
                foreach (var col in columns)
                {
                    var cell = (Range)_app.Cells[row, startCol + col - 1];
                    key += cell.Value2?.ToString() + "|";
                }
            }
            else
            {
                for (int c = 0; c < colCount; c++)
                {
                    var cell = (Range)_app.Cells[row, startCol + c];
                    key += cell.Value2?.ToString() + "|";
                }
            }

            return key;
        }

        private string ToProperCase(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var chars = s.ToCharArray();
            if (chars.Length > 0) chars[0] = char.ToUpper(chars[0]);
            for (int i = 1; i < chars.Length; i++)
            {
                chars[i] = char.ToLower(chars[i]);
            }
            return new string(chars);
        }
    }

    public class CleanResult
    {
        public bool Success { get; set; }
        public List<string> Operations { get; set; }
        public string Error { get; set; }
    }
}
