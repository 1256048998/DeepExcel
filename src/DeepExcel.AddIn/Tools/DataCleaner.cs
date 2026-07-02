using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
                        // ★ 必须用 EntireRow.Delete() 删除整行，所有列才会一起上移。
                        // 错误做法：range.Rows[i].Delete(xlShiftUp) 只删除 range 覆盖列的单元格，
                        // 范围外列不会上移，导致列错位。
                        ((Range)range.Rows[i]).EntireRow.Delete();
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
        /// 删除空行（整行为空的行）
        /// </summary>
        public ToolResult DeleteBlankRows(string rangeAddress)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;
                int deletedCount = 0;

                for (int i = rows; i >= 1; i--)
                {
                    bool isBlank = true;
                    for (int j = 1; j <= cols; j++)
                    {
                        var cell = (Range)range.Cells[i, j];
                        var val = cell.Value2;
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                        {
                            isBlank = false;
                            break;
                        }
                    }
                    if (isBlank)
                    {
                        // ★ 必须用 EntireRow.Delete() 删除整行，所有列才会一起上移。
                        // 错误做法：range.Rows[i].Delete(xlShiftUp) 只删除 range 覆盖列的单元格，
                        // 范围外列不会上移，导致列错位。
                        ((Range)range.Rows[i]).EntireRow.Delete();
                        deletedCount++;
                    }
                }

                return new ToolResult
                {
                    Name = "delete_blank_rows",
                    Success = true,
                    Data = new { deletedCount = deletedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "delete_blank_rows",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 按分隔符拆分文本到多列（从指定列开始向右扩展）
        /// </summary>
        public ToolResult SplitTextToColumns(string rangeAddress, string delimiter = ",", bool hasHeader = false)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int startCol = range.Column;
                int startRow = range.Row;
                int rows = range.Rows.Count;
                int splitCount = 0;
                int maxParts = 0;

                var splitData = new List<List<string>>();
                for (int i = 1; i <= rows; i++)
                {
                    var cell = (Range)range.Cells[i, 1];
                    var val = cell.Value2;
                    string text = val?.ToString() ?? "";
                    var parts = text.Split(new[] { delimiter }, StringSplitOptions.None);
                    splitData.Add(parts.ToList());
                    if (parts.Length > maxParts) maxParts = parts.Length;
                    splitCount++;
                }

                if (maxParts <= 1)
                {
                    return new ToolResult
                    {
                        Name = "split_text_to_columns",
                        Success = true,
                        Data = new { splitCount = 0, newColumnCount = 0, reason = "no_split_needed" }
                    };
                }

                for (int i = 0; i < splitCount; i++)
                {
                    for (int j = 0; j < splitData[i].Count; j++)
                    {
                        var cell = (Range)_app.Cells[startRow + i, startCol + j];
                        cell.Value = splitData[i][j];
                    }
                    for (int j = splitData[i].Count; j < maxParts; j++)
                    {
                        var cell = (Range)_app.Cells[startRow + i, startCol + j];
                        cell.Value = "";
                    }
                }

                return new ToolResult
                {
                    Name = "split_text_to_columns",
                    Success = true,
                    Data = new { splitCount = splitCount, newColumnCount = maxParts }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "split_text_to_columns",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 向下填充空白单元格（用上方最近的非空值填充）
        /// </summary>
        public ToolResult FillBlankCells(string rangeAddress, string direction = "down")
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;
                int filledCount = 0;

                for (int c = 1; c <= cols; c++)
                {
                    object lastValue = null;
                    for (int r = 1; r <= rows; r++)
                    {
                        var cell = (Range)range.Cells[r, c];
                        var val = cell.Value2;
                        if (val == null || string.IsNullOrWhiteSpace(val.ToString()))
                        {
                            if (lastValue != null)
                            {
                                cell.Value = lastValue;
                                filledCount++;
                            }
                        }
                        else
                        {
                            lastValue = val;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "fill_blank_cells",
                    Success = true,
                    Data = new { filledCount = filledCount, direction = direction }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "fill_blank_cells",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 标记/高亮重复值
        /// </summary>
        public ToolResult HighlightDuplicates(string rangeAddress, string colorHex = "#FFC7CE")
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int highlightedCount = 0;
                var color = ColorTranslator.FromHtml(colorHex);
                var oleColor = ColorTranslator.ToOle(color);
                var seen = new Dictionary<string, List<Range>>();

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    string key = val?.ToString() ?? "";
                    if (string.IsNullOrEmpty(key)) continue;

                    if (!seen.ContainsKey(key))
                    {
                        seen[key] = new List<Range>();
                    }
                    seen[key].Add(cell);
                }

                foreach (var kvp in seen)
                {
                    if (kvp.Value.Count > 1)
                    {
                        foreach (var cell in kvp.Value)
                        {
                            cell.Interior.Color = oleColor;
                            highlightedCount++;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "highlight_duplicates",
                    Success = true,
                    Data = new { highlightedCount = highlightedCount, color = colorHex }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "highlight_duplicates",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 去除特殊字符（非打印字符、不可见字符）
        /// </summary>
        public ToolResult RemoveSpecialChars(string rangeAddress)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int cleanedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val != null && val is string s && !string.IsNullOrEmpty(s))
                    {
                        char[] arr = s.ToCharArray();
                        List<char> result = new List<char>();
                        foreach (char c in arr)
                        {
                            if (char.IsControl(c) && c != '\t' && c != '\n' && c != '\r')
                                continue;
                            if (c == '\u00A0')
                                continue;
                            result.Add(c);
                        }
                        string newVal = new string(result.ToArray());
                        if (newVal != s)
                        {
                            cell.Value = newVal;
                            cleanedCount++;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "remove_special_chars",
                    Success = true,
                    Data = new { cleanedCount = cleanedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "remove_special_chars",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 金额清洗：去除货币符号、千分位，转为纯数字
        /// </summary>
        public ToolResult CleanAmount(string rangeAddress)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int cleanedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val != null && val is string s && !string.IsNullOrEmpty(s))
                    {
                        string cleaned = s
                            .Replace("¥", "").Replace("￥", "")
                            .Replace("$", "").Replace("€", "").Replace("£", "")
                            .Replace(",", "").Replace("，", "")
                            .Replace("元", "").Replace("圆", "")
                            .Trim();

                        if (double.TryParse(cleaned, out double num))
                        {
                            cell.Value = num;
                            cell.NumberFormat = "#,##0.00";
                            cleanedCount++;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "clean_amount",
                    Success = true,
                    Data = new { cleanedCount = cleanedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "clean_amount",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 合并多列为一列（用指定分隔符连接）
        /// </summary>
        public ToolResult MergeColumns(string rangeAddress, string delimiter = " ", string targetColumn = "")
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;
                int mergedCount = 0;

                int targetCol;
                if (string.IsNullOrEmpty(targetColumn))
                {
                    targetCol = range.Column;
                }
                else
                {
                    targetCol = _app.Range[targetColumn + "1"].Column;
                }

                for (int r = 1; r <= rows; r++)
                {
                    var parts = new List<string>();
                    for (int c = 1; c <= cols; c++)
                    {
                        var cell = (Range)range.Cells[r, c];
                        var val = cell.Value2;
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                        {
                            parts.Add(val.ToString());
                        }
                    }
                    var targetCell = (Range)_app.Cells[range.Row + r - 1, targetCol];
                    targetCell.Value = string.Join(delimiter, parts);
                    mergedCount++;
                }

                return new ToolResult
                {
                    Name = "merge_columns",
                    Success = true,
                    Data = new { mergedCount = mergedCount, delimiter = delimiter }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "merge_columns",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 批量重命名列（修改第一行的列标题）
        /// </summary>
        public ToolResult RenameColumns(string rangeAddress, string[] newNames)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int renamedCount = 0;

                for (int i = 0; i < newNames.Length && i < range.Columns.Count; i++)
                {
                    var cell = (Range)range.Cells[1, i + 1];
                    cell.Value = newNames[i];
                    renamedCount++;
                }

                return new ToolResult
                {
                    Name = "rename_columns",
                    Success = true,
                    Data = new { renamedCount = renamedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "rename_columns",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 压缩内部多余空格（多个连续空格变一个）
        /// </summary>
        public ToolResult CollapseSpaces(string rangeAddress)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                int cleanedCount = 0;

                foreach (Range cell in range.Cells)
                {
                    var val = cell.Value2;
                    if (val != null && val is string s && !string.IsNullOrEmpty(s))
                    {
                        string trimmed = s.Trim();
                        while (trimmed.Contains("  "))
                        {
                            trimmed = trimmed.Replace("  ", " ");
                        }
                        if (trimmed != s)
                        {
                            cell.Value = trimmed;
                            cleanedCount++;
                        }
                    }
                }

                return new ToolResult
                {
                    Name = "collapse_spaces",
                    Success = true,
                    Data = new { cleanedCount = cleanedCount }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "collapse_spaces",
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
            if (r1.Success) result.Operations.Add($"去除首尾空格: {(r1.Data as dynamic)?.trimmedCount}个");

            var r2 = CollapseSpaces(rangeAddress);
            if (r2.Success) result.Operations.Add($"压缩内部空格: {(r2.Data as dynamic)?.cleanedCount}个");

            var r3 = UnifyDateFormats(rangeAddress);
            if (r3.Success) result.Operations.Add($"统一日期格式: {(r3.Data as dynamic)?.convertedCount}个");

            var r4 = HighlightMissingValues(rangeAddress);
            if (r4.Success) result.Operations.Add($"标红缺失值: {(r4.Data as dynamic)?.highlightedCount}个");

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
