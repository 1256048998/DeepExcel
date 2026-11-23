using System;
using System.Collections.Generic;
using DeepExcel.AddIn.Sidecar;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>
    /// Where an operation lands, and reading what is there now.
    /// </summary>
    public sealed partial class PreviewBuilder
    {
        /// <summary>A target range in the form the reader understands.</summary>
        internal sealed class Target
        {
            /// <summary>Null means the active sheet.</summary>
            public string Sheet;

            /// <summary>Local address, without a sheet prefix.</summary>
            public string Address;

            /// <summary>Parsed geometry; absent for named ranges and the like.</summary>
            public CellRect? Rect;

            public long ExpectedCells
            {
                get
                {
                    if (Rect == null) return -1;
                    var r = Rect.Value;
                    return (long)(r.Row2 - r.Row1 + 1) * (r.Col2 - r.Col1 + 1);
                }
            }
        }

        /// <summary>
        /// A rectangular block of current cells, addressed from its top-left.
        /// </summary>
        internal sealed class Grid
        {
            private readonly IReadOnlyList<CellSnapshot> _cells;

            public Grid(CellRect rect, IReadOnlyList<CellSnapshot> cells)
            {
                Rect = rect;
                _cells = cells;
            }

            public CellRect Rect { get; }
            public int Rows => Rect.Row2 - Rect.Row1 + 1;
            public int Cols => Rect.Col2 - Rect.Col1 + 1;
            public IEnumerable<CellSnapshot> Cells => _cells;

            /// <summary>Zero-based row and column within the block.</summary>
            public CellSnapshot this[int row, int col] => _cells[row * Cols + col];
        }

        /// <summary>
        /// Picks the address argument and splits off its sheet.
        ///
        /// Tools take addresses like "Data!A1:C20", but the reader resolves an
        /// address against one worksheet; handing it the prefixed form fails
        /// and yields nothing, which used to read as "nothing will change".
        /// Whole columns are bounded by the used range so "A:A" means the data,
        /// not a million empty cells.
        /// </summary>
        private Target ResolveTarget(IDictionary<string, object> args, params string[] keys)
        {
            string raw = null;
            foreach (var key in keys)
            {
                raw = GetString(args, key);
                if (!string.IsNullOrEmpty(raw)) break;
            }
            if (string.IsNullOrEmpty(raw))
            {
                return null;
            }

            var target = new Target
            {
                Sheet = GetString(args, "sheet") ?? GetString(args, "sheet_name"),
                Address = raw
            };
            if (!CellRect.TryParse(raw, null, out var rect))
            {
                return target;
            }

            if (!string.IsNullOrEmpty(rect.Sheet))
            {
                target.Sheet = rect.Sheet;
            }
            if (rect.Row1 == 1 && rect.Row2 == CellRect.MaxRow)
            {
                var used = _reader.UsedRowCount(target.Sheet);
                rect = new CellRect(rect.Sheet, 1, rect.Col1, Math.Max(1, used), rect.Col2);
            }
            target.Rect = rect;
            target.Address = LocalAddress(rect);
            return target;
        }

        /// <summary>
        /// Reads the whole target as a block, or explains why that is not
        /// possible. Returns null on success.
        ///
        /// Both failures matter for the same reason: an empty or partial read
        /// undercounts the changes, and a small enough count is auto-allowed
        /// without asking the user.
        /// </summary>
        private ChangePreview ReadGrid(string toolName, Target target, out Grid grid)
        {
            grid = null;
            if (target.Rect == null)
            {
                return ChangePreview.NotPreviewable(toolName, "无法解析区域地址「" + target.Address + "」，已自动创建快照");
            }

            var cells = _reader.Read(target.Sheet, target.Address);
            if (cells.Count == 0)
            {
                return Unreadable(toolName);
            }
            if (cells.Count < target.ExpectedCells)
            {
                return ChangePreview.NotPreviewable(toolName,
                    $"区域超过 {ExcelRangeReader.MaxCellsRead} 个单元格，无法逐格预览，已自动创建快照");
            }

            grid = new Grid(target.Rect.Value, cells);
            return null;
        }

        private static ChangePreview Unreadable(string toolName)
        {
            return ChangePreview.NotPreviewable(toolName, "读取目标区域失败，无法预览，已自动创建快照");
        }

        internal static string LocalAddress(CellRect rect)
        {
            var topLeft = CellRect.ColumnName(rect.Col1) + rect.Row1;
            var bottomRight = CellRect.ColumnName(rect.Col2) + rect.Row2;
            return topLeft == bottomRight ? topLeft : topLeft + ":" + bottomRight;
        }

        /// <summary>
        /// A list-of-text argument (new_names, operations), read the way
        /// ToolDispatcher.GetArgArray reads it.
        ///
        /// The sidecar's permission request carries arrays as raw JSON text;
        /// tests pass real lists. The executor calls GetString() on each JSON
        /// element, which throws on anything but a string or null, and then
        /// drops the whole list -- so a list with a number in it yields an
        /// empty list here too.
        /// </summary>
        internal static List<string> GetStringArray(IDictionary<string, object> args, string key)
        {
            var list = new List<string>();
            if (args == null || !args.TryGetValue(key, out var raw) || raw == null)
            {
                return list;
            }
            if (raw is string text)
            {
                try
                {
                    using (var doc = System.Text.Json.JsonDocument.Parse(text))
                    {
                        if (doc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Array) return list;
                        foreach (var item in doc.RootElement.EnumerateArray())
                        {
                            if (item.ValueKind == System.Text.Json.JsonValueKind.Null) list.Add(null);
                            else if (item.ValueKind == System.Text.Json.JsonValueKind.String) list.Add(item.GetString());
                            else return new List<string>();
                        }
                    }
                }
                catch (System.Text.Json.JsonException)
                {
                    return new List<string>();
                }
                return list;
            }
            if (raw is System.Collections.IEnumerable sequence)
            {
                foreach (var item in sequence)
                {
                    list.Add(item?.ToString());
                }
            }
            return list;
        }

        /// <summary>
        /// The value the executor will see. Readers that only fill in the
        /// display text (older fakes) are treated as holding that text.
        /// </summary>
        internal static object RawValue(CellSnapshot cell)
        {
            if (cell == null) return null;
            if (cell.Value != null) return cell.Value;
            return string.IsNullOrEmpty(cell.Display) ? null : cell.Display;
        }

        /// <summary>A date as the preview shows it: yyyy-MM-dd, plus the time if any.</summary>
        internal static string FormatDate(DateTime date)
        {
            return date.TimeOfDay == TimeSpan.Zero
                ? date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                : date.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// What a number shows once given a yyyy-mm-dd format, on Excel's
        /// calendar: serial 1 is 1900-01-01, and serial 60 is the 29 February
        /// 1900 Excel counts although it never existed. Out-of-range values
        /// show as #### in Excel.
        /// </summary>
        internal static string SerialToDateText(double serial)
        {
            if (serial < 0 || serial >= 2958466) return "####";
            if (serial < 1) return "1900-01-00";
            var day = Math.Floor(serial);
            if (day == 60) return "1900-02-29";
            var date = DateTime.FromOADate(day < 60 ? day + 1 : day);
            return date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        }

        private static bool IsFormulaText(string text)
        {
            return !string.IsNullOrEmpty(text) && text.StartsWith("=", StringComparison.Ordinal);
        }
    }
}
