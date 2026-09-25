using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Perception;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 首次使用的宿主一侧：get_starter 交出每张表的前几十行（推荐在面板里算，见 UI 的 utils/starter.ts），
    /// insert_sample 把面板给的示例数据写进一张新表。两者都不经过侧车、不花模型调用。
    /// </summary>
    public static class StarterOutline
    {
        public const int MaxSheets = 12;
        public const int MaxRows = 51;   // 表头 + 50 行
        public const int MaxColumns = 30;

        public const int MaxSampleRows = 500;
        public const int MaxSampleColumns = 30;
        public const int MaxSampleText = 200;

        /// <summary>Value2 的一格换成 JSON 友好的值：数字、文本、布尔、错误文本或 null</summary>
        public static object ToCell(object value)
        {
            switch (value)
            {
                case null:
                case DBNull _:
                    return null;
                case double d:
                    return double.IsNaN(d) || double.IsInfinity(d) ? null : (object)d;
                case bool b:
                    return b;
                case int code:
                    return SnapshotEncoder.ErrorText(code);
                case DateTime dt:
                    return dt.ToOADate();
                case string s:
                    if (s.Length == 0) return null;
                    return s.Length > MaxSampleText ? s.Substring(0, MaxSampleText) : s;
                default:
                    return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>和已有表不重名的表名：示例-销售明细、示例-销售明细(2)……（Excel 表名最长 31 字）</summary>
        public static string UniqueSheetName(string wanted, IEnumerable<string> existing)
        {
            var taken = new HashSet<string>(existing ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var baseName = Sanitize(wanted);
            if (!taken.Contains(baseName)) return baseName;
            for (int i = 2; ; i++)
            {
                var suffix = "(" + i + ")";
                var candidate = (baseName.Length + suffix.Length > 31 ? baseName.Substring(0, 31 - suffix.Length) : baseName) + suffix;
                if (!taken.Contains(candidate)) return candidate;
            }
        }

        private static string Sanitize(string name)
        {
            var cleaned = new string((name ?? "").Where(ch => "\\/?*[]:".IndexOf(ch) < 0).ToArray()).Trim('\'', ' ');
            if (cleaned.Length == 0) cleaned = "示例";
            return cleaned.Length > 31 ? cleaned.Substring(0, 31) : cleaned;
        }

        /// <summary>面板给的示例行 → 写入用的二维数组。超出上限或类型不对返回 null 并给原因。</summary>
        public static object[,] ParseRows(JsonElement rows, out string error)
        {
            error = null;
            if (rows.ValueKind != JsonValueKind.Array || rows.GetArrayLength() == 0)
            {
                error = "示例数据为空";
                return null;
            }
            int height = rows.GetArrayLength();
            int width = rows.EnumerateArray().Select(r => r.ValueKind == JsonValueKind.Array ? r.GetArrayLength() : 0).Max();
            if (height > MaxSampleRows || width == 0 || width > MaxSampleColumns)
            {
                error = "示例数据的行列数超出范围";
                return null;
            }
            var data = new object[height, width];
            int r = 0;
            foreach (var row in rows.EnumerateArray())
            {
                int c = 0;
                if (row.ValueKind == JsonValueKind.Array)
                {
                    foreach (var cell in row.EnumerateArray())
                    {
                        switch (cell.ValueKind)
                        {
                            case JsonValueKind.Number: data[r, c] = cell.GetDouble(); break;
                            case JsonValueKind.True: data[r, c] = true; break;
                            case JsonValueKind.False: data[r, c] = false; break;
                            case JsonValueKind.String:
                                var text = cell.GetString() ?? "";
                                // 以 = 开头的会被当成公式写入：示例里不需要，也不接受
                                if (text.StartsWith("=", StringComparison.Ordinal)) text = "'" + text;
                                data[r, c] = text.Length > MaxSampleText ? text.Substring(0, MaxSampleText) : text;
                                break;
                            default: data[r, c] = null; break;
                        }
                        c++;
                    }
                }
                r++;
            }
            return data;
        }

        /// <summary>get_starter 的内容：每张可见表的总行列数、前 51 行 × 30 列、哪些列是日期</summary>
        public static object Describe(Excel.Workbook wb)
        {
            if (wb == null) return new { workbook_name = "", sheets = new object[0] };
            var sheets = new List<object>();
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (sheets.Count >= MaxSheets) break;
                if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible) continue;
                sheets.Add(OutlineOf(ws));
            }
            return new { workbook_name = wb.Name, sheets };
        }

        private static object OutlineOf(Excel.Worksheet ws)
        {
            var used = ws.UsedRange;
            int rows = used.Rows.Count, cols = used.Columns.Count;
            int r = Math.Min(rows, MaxRows), c = Math.Min(cols, MaxColumns);
            var grid = new List<object[]>();
            var raw = ws.Range[used.Cells[1, 1], used.Cells[r, c]].Value2;
            if (raw is object[,] values)
            {
                int r0 = values.GetLowerBound(0), c0 = values.GetLowerBound(1);
                for (int i = 0; i < r; i++)
                {
                    var row = new object[c];
                    for (int j = 0; j < c; j++) row[j] = ToCell(values[r0 + i, c0 + j]);
                    grid.Add(row);
                }
            }
            else
            {
                grid.Add(new[] { ToCell(raw) });
            }

            // Value2 里日期只是序列号：看第二行的数字格式才分得出来
            var dateColumns = new bool[c];
            if (r >= 2)
            {
                for (int j = 0; j < c; j++)
                {
                    try
                    {
                        var format = Convert.ToString(((Excel.Range)used.Cells[2, j + 1]).NumberFormat, CultureInfo.InvariantCulture);
                        dateColumns[j] = SheetProfiler.LooksLikeDateFormat(format);
                    }
                    catch (Exception) { }
                }
            }
            return new { name = ws.Name, rows, columns = cols, grid, date_columns = dateColumns };
        }

        /// <summary>
        /// 示例写进最后一张表之后的新表，返回实际表名。不碰用户已有的表；
        /// 文本型数字靠前导撇号写成文本（示例里故意埋的问题）。
        /// </summary>
        public static string InsertSample(Excel.Workbook wb, string wanted, object[,] data, IDictionary<int, string> numberFormats)
        {
            if (wb.ProtectStructure) throw new InvalidOperationException("工作簿结构受保护，无法新建工作表");
            var existing = wb.Worksheets.Cast<Excel.Worksheet>().Select(s => s.Name).ToList();
            var name = UniqueSheetName(wanted, existing);
            var last = wb.Worksheets[wb.Worksheets.Count];
            var ws = (Excel.Worksheet)wb.Worksheets.Add(After: last);
            ws.Name = name;

            int height = data.GetLength(0), width = data.GetLength(1);
            ws.Range[ws.Cells[1, 1], ws.Cells[height, width]].Value2 = data;
            if (numberFormats != null && height > 1)
            {
                foreach (var f in numberFormats)
                {
                    if (f.Key < 0 || f.Key >= width || string.IsNullOrEmpty(f.Value)) continue;
                    ws.Range[ws.Cells[2, f.Key + 1], ws.Cells[height, f.Key + 1]].NumberFormat = f.Value;
                }
            }
            ws.Range[ws.Cells[1, 1], ws.Cells[1, width]].Font.Bold = true;
            ws.Range[ws.Cells[1, 1], ws.Cells[height, width]].Columns.AutoFit();
            ws.Activate();
            return name;
        }

        public static Dictionary<int, string> ParseNumberFormats(JsonElement formats)
        {
            var result = new Dictionary<int, string>();
            if (formats.ValueKind != JsonValueKind.Object) return result;
            foreach (var f in formats.EnumerateObject())
            {
                if (int.TryParse(f.Name, out var col) && f.Value.ValueKind == JsonValueKind.String) result[col] = f.Value.GetString();
            }
            return result;
        }
    }

    public partial class MessageBridge
    {
        private string HandleGetStarter()
        {
            try
            {
                return MakeResponse("starter", StarterOutline.Describe(_excelApp.ActiveWorkbook));
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "get_starter failed: " + ex.Message);
                return MakeError("读取工作簿结构失败：" + ex.Message);
            }
        }

        private string HandleInsertSample(Message msg)
        {
            try
            {
                if (!msg.Payload.HasValue || !msg.Payload.Value.TryGetProperty("rows", out var rowsEl)) return MakeError("缺少示例数据");
                var payload = msg.Payload.Value;
                var wanted = payload.TryGetProperty("sheet_name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String
                    ? nameEl.GetString() : "示例";
                var data = StarterOutline.ParseRows(rowsEl, out var parseError);
                if (data == null) return MakeError(parseError);
                var formats = payload.TryGetProperty("number_formats", out var formatsEl)
                    ? StarterOutline.ParseNumberFormats(formatsEl) : null;

                var wb = _excelApp.ActiveWorkbook;
                if (wb == null) return MakeError("没有打开的工作簿");
                var name = StarterOutline.InsertSample(wb, wanted, data, formats);
                Logger.Instance.Info("MessageBridge", $"Sample sheet inserted: {data.GetLength(0)}x{data.GetLength(1)}");
                return MakeResponse("sample_inserted", new { sheet = name });
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "insert_sample failed: " + ex.Message);
                return MakeError(ex is InvalidOperationException ? ex.Message : "插入示例失败：" + ex.Message);
            }
        }
    }
}
