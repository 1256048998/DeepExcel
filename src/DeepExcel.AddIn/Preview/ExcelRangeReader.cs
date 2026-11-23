using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>
    /// Reads current cell state from the live workbook.
    ///
    /// Values and formulas are fetched as whole blocks: a preview runs while the
    /// user waits, and per-cell COM calls over a few hundred cells would make
    /// the confirmation dialog feel broken.
    /// </summary>
    public sealed class ExcelRangeReader : IRangeReader
    {
        /// <summary>
        /// Cap on cells read for a preview.
        ///
        /// Beyond this the preview is truncated anyway, so reading more only
        /// costs time. The count shown to the user comes from the range size,
        /// not from this list, so truncation does not understate the impact.
        /// </summary>
        public const int MaxCellsRead = 5000;

        private readonly Application _app;

        public ExcelRangeReader(Application app)
        {
            _app = app;
        }

        public IReadOnlyList<CellSnapshot> Read(string sheet, string address)
        {
            var cells = new List<CellSnapshot>();
            if (_app == null || string.IsNullOrEmpty(address))
            {
                return cells;
            }

            try
            {
                var worksheet = ResolveSheet(sheet);
                if (worksheet == null)
                {
                    return cells;
                }

                var range = worksheet.Range[address];
                if (range == null)
                {
                    return cells;
                }

                var rowCount = range.Rows.Count;
                var columnCount = range.Columns.Count;
                if ((long)rowCount * columnCount > MaxCellsRead)
                {
                    // Read the leading block only; the caller reports the true
                    // size separately.
                    rowCount = Math.Max(1, MaxCellsRead / Math.Max(1, columnCount));
                    range = range.Resize[rowCount, columnCount];
                }

                var values = range.Value2;
                var formulas = range.Formula;
                var firstRow = range.Row;
                var firstColumn = range.Column;

                // Value (unlike Value2) comes back as DateTime for cells
                // formatted as dates. One more block read is what it costs to
                // show a date as a date instead of its serial number, and to
                // know which numbers are already dates.
                object typed = null;
                try { typed = range.Value; } catch (Exception) { }

                if (!(values is object[,] valueGrid))
                {
                    cells.Add(Snapshot(CellAddress(firstRow, firstColumn), values, formulas as string, typed));
                    return cells;
                }

                var typedGrid = typed as object[,];
                var formulaGrid = formulas as object[,];
                var lowerRow = valueGrid.GetLowerBound(0);
                var upperRow = valueGrid.GetUpperBound(0);
                var lowerColumn = valueGrid.GetLowerBound(1);
                var upperColumn = valueGrid.GetUpperBound(1);

                for (var r = lowerRow; r <= upperRow; r++)
                {
                    for (var c = lowerColumn; c <= upperColumn; c++)
                    {
                        var formula = formulaGrid != null ? formulaGrid[r, c] as string : null;
                        cells.Add(Snapshot(
                            CellAddress(firstRow + (r - lowerRow), firstColumn + (c - lowerColumn)),
                            valueGrid[r, c], formula, typedGrid != null ? typedGrid[r, c] : null));
                    }
                }
            }
            catch (Exception)
            {
                // An unreadable range yields no snapshots. The caller turns that
                // into "not previewable" rather than "nothing will change".
                cells.Clear();
            }
            return cells;
        }

        public int UsedRowCount(string sheet)
        {
            try
            {
                var worksheet = ResolveSheet(sheet);
                if (worksheet == null)
                {
                    return 0;
                }
                var used = worksheet.UsedRange;
                return used == null ? 0 : used.Row + used.Rows.Count - 1;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public int UsedColumnCount(string sheet)
        {
            try
            {
                var worksheet = ResolveSheet(sheet);
                if (worksheet == null)
                {
                    return 0;
                }
                var used = worksheet.UsedRange;
                return used == null ? 0 : used.Column + used.Columns.Count - 1;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public bool SheetExists(string sheet)
        {
            return ResolveSheet(sheet) != null;
        }

        private Worksheet ResolveSheet(string sheet)
        {
            try
            {
                if (string.IsNullOrEmpty(sheet))
                {
                    return DeepExcel.AddIn.Executor.ExcelTarget.ActiveSheet(_app);
                }
                return DeepExcel.AddIn.Executor.ExcelTarget.Workbook(_app)?.Worksheets[sheet] as Worksheet;
            }
            catch (Exception)
            {
                // A named sheet that does not exist is a real failure, not a
                // reason to silently preview the active sheet instead.
                return null;
            }
        }

        private static CellSnapshot Snapshot(string address, object value2, string formula, object typed)
        {
            var isDate = typed is DateTime;
            return new CellSnapshot
            {
                Address = address,
                Value = value2,
                Formula = formula,
                IsDate = isDate,
                Display = isDate ? PreviewBuilder.FormatDate((DateTime)typed) : Format(value2)
            };
        }

        private static string CellAddress(int row, int column)
        {
            return Perception.TypeInference.ColumnLetter(column) + row;
        }

        private static string Format(object value)
        {
            if (value == null) return "";
            if (value is double d)
            {
                return d == Math.Floor(d) && Math.Abs(d) < 1e15
                    ? ((long)d).ToString(CultureInfo.InvariantCulture)
                    : d.ToString("0.####", CultureInfo.InvariantCulture);
            }
            if (value is bool b) return b ? "TRUE" : "FALSE";
            var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            return text.Length > 40 ? text.Substring(0, 40) + "…" : text;
        }
    }
}
