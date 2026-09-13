using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// Builds a <see cref="WorkbookIndex"/> from a live workbook.
    ///
    /// Two constraints shape everything here:
    ///
    ///   * COM round-trips dominate. Values are read as whole blocks
    ///     (Range.Value2 returns a 2-D array in one call); reading cell by cell
    ///     on an 8000-row sheet would take minutes and freeze Excel.
    ///   * It runs on the UI thread during a user action, so it has a time
    ///     budget. Exceeding it produces a partial index marked Incomplete
    ///     rather than a hang -- the model is told to read instead of trusting
    ///     an index that quietly stops early.
    /// </summary>
    public sealed class SheetProfiler
    {
        /// <summary>Head rows always read: real tables put their structure at the top.</summary>
        public const int HeadRows = 100;

        /// <summary>Additional rows sampled from the body, evenly spaced.</summary>
        public const int BodySampleRows = 500;

        /// <summary>Columns beyond this are almost always a stray used-range.</summary>
        public const int MaxColumns = 64;

        private readonly TimeSpan _budget;

        public SheetProfiler(TimeSpan? budget = null)
        {
            _budget = budget ?? TimeSpan.FromSeconds(5);
        }

        public WorkbookIndex Build(Workbook workbook, Application app)
        {
            var index = new WorkbookIndex();
            if (workbook == null)
            {
                index.Incomplete = "没有打开的工作簿";
                return index;
            }

            var clock = Stopwatch.StartNew();
            index.WorkbookName = Safe(() => workbook.Name, "(未命名)");
            index.ActiveSheet = Safe(() => (app?.ActiveSheet as Worksheet)?.Name, null);

            try
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (clock.Elapsed > _budget)
                    {
                        index.Incomplete = "超过 " + _budget.TotalSeconds + " 秒时间预算";
                        break;
                    }
                    if (Safe(() => sheet.Visible, XlSheetVisibility.xlSheetVisible)
                        != XlSheetVisibility.xlSheetVisible)
                    {
                        // Hidden sheets are usually scratch space; including them
                        // costs tokens and invites the model to edit them.
                        continue;
                    }

                    var profile = ProfileSheet(sheet, index);
                    if (profile != null)
                    {
                        index.Sheets.Add(profile);
                    }
                }
            }
            catch (Exception ex)
            {
                index.Incomplete = "读取工作表失败：" + ex.GetType().Name;
            }

            try { CollectNamedRanges(workbook, index); }
            catch (Exception) { }

            return index;
        }

        private SheetIndex ProfileSheet(Worksheet sheet, WorkbookIndex index)
        {
            var name = Safe(() => sheet.Name, null);
            if (name == null)
            {
                return null;
            }

            Range used = null;
            try { used = sheet.UsedRange; }
            catch (Exception) { }
            if (used == null)
            {
                return new SheetIndex { Name = name, LastRow = 0 };
            }

            var firstRow = Safe(() => used.Row, 1);
            var firstColumn = Safe(() => used.Column, 1);
            var rowCount = Safe(() => used.Rows.Count, 0);
            var columnCount = Math.Min(Safe(() => used.Columns.Count, 0), MaxColumns);

            var result = new SheetIndex
            {
                Name = name,
                FirstRow = firstRow,
                LastRow = firstRow + rowCount - 1
            };

            if (rowCount <= 0 || columnCount <= 0)
            {
                return result;
            }

            var blocks = PlanSample(firstRow, rowCount);
            result.Sampled = blocks.Sum(b => b.Item2) < rowCount;
            result.SampledRows = blocks.Sum(b => b.Item2);

            var rows = new List<object[]>();
            var rowNumbers = new List<int>();
            foreach (var block in blocks)
            {
                var read = ReadBlock(sheet, block.Item1, block.Item2, firstColumn, columnCount);
                if (read == null)
                {
                    continue;
                }
                for (var i = 0; i < read.Length; i++)
                {
                    rows.Add(read[i]);
                    rowNumbers.Add(block.Item1 + i);
                }
            }

            if (rows.Count == 0)
            {
                return result;
            }

            var sample = rows.ToArray();
            string[] headers = null;
            var dataStart = 0;
            if (TypeInference.LooksLikeHeaderRow(sample))
            {
                result.HasHeader = true;
                result.HeaderRow = rowNumbers[0];
                headers = sample[0].Select(v => v?.ToString()?.Trim()).ToArray();
                dataStart = 1;
            }

            var bodyFirstRow = dataStart < rowNumbers.Count ? rowNumbers[dataStart] : result.FirstRow;
            var metadata = ReadRowMetadata(sheet, bodyFirstRow, firstColumn, columnCount);

            // Must happen before profiling: otherwise every date column is
            // classified from its OLE serial and reported as a number.
            ConvertOleDates(rows, metadata.NumberFormats);

            var body = rows.Skip(dataStart).ToArray();
            result.Columns = TypeInference.ProfileColumns(
                body, bodyFirstRow, headers, metadata.Formulas);
            CollectRelations(name, metadata.Formulas, index);
            return result;
        }

        /// <summary>
        /// Which row blocks to read: the head, then an evenly spaced sample of
        /// the body.
        ///
        /// Evenly spaced rather than random so the result is reproducible --
        /// an index that changes between two runs on an unchanged sheet would
        /// make every bug report unreproducible.
        /// </summary>
        internal static List<Tuple<int, int>> PlanSample(int firstRow, int rowCount)
        {
            var blocks = new List<Tuple<int, int>>();
            if (rowCount <= 0)
            {
                return blocks;
            }

            var head = Math.Min(HeadRows, rowCount);
            blocks.Add(Tuple.Create(firstRow, head));

            var remaining = rowCount - head;
            if (remaining <= 0)
            {
                return blocks;
            }
            if (remaining <= BodySampleRows)
            {
                blocks.Add(Tuple.Create(firstRow + head, remaining));
                return blocks;
            }

            // Spread BodySampleRows over the remainder in evenly spaced chunks.
            const int chunk = 25;
            var chunks = BodySampleRows / chunk;
            var stride = remaining / chunks;
            for (var i = 0; i < chunks; i++)
            {
                var start = firstRow + head + i * stride;
                var size = Math.Min(chunk, firstRow + rowCount - start);
                if (size > 0)
                {
                    blocks.Add(Tuple.Create(start, size));
                }
            }
            return blocks;
        }

        /// <summary>
        /// One COM call per block. The returned array is 1-based in both
        /// dimensions, and a single-cell range comes back as a scalar rather
        /// than an array -- both have caused index bugs here before.
        /// </summary>
        private static object[][] ReadBlock(
            Worksheet sheet, int startRow, int rowCount, int startColumn, int columnCount)
        {
            try
            {
                var topLeft = (Range)sheet.Cells[startRow, startColumn];
                var bottomRight = (Range)sheet.Cells[startRow + rowCount - 1, startColumn + columnCount - 1];
                var block = sheet.Range[topLeft, bottomRight];
                var raw = block.Value2;

                if (raw == null)
                {
                    return null;
                }
                if (!(raw is object[,] grid))
                {
                    // Single cell.
                    return new[] { new[] { raw } };
                }

                var lowerRow = grid.GetLowerBound(0);
                var upperRow = grid.GetUpperBound(0);
                var lowerColumn = grid.GetLowerBound(1);
                var upperColumn = grid.GetUpperBound(1);

                var rows = new List<object[]>();
                for (var r = lowerRow; r <= upperRow; r++)
                {
                    var row = new object[upperColumn - lowerColumn + 1];
                    for (var c = lowerColumn; c <= upperColumn; c++)
                    {
                        row[c - lowerColumn] = NormalizeCell(grid[r, c]);
                    }
                    rows.Add(row);
                }
                return rows.ToArray();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object NormalizeCell(object value)
        {
            return value;
        }

        /// <summary>Formula text and number format for one row, read once per column.</summary>
        internal sealed class RowMetadata
        {
            public string[] Formulas;
            public string[] NumberFormats;
        }

        private static RowMetadata ReadRowMetadata(
            Worksheet sheet, int row, int startColumn, int columnCount)
        {
            var metadata = new RowMetadata
            {
                Formulas = new string[columnCount],
                NumberFormats = new string[columnCount]
            };
            try
            {
                for (var c = 0; c < columnCount; c++)
                {
                    var cell = (Range)sheet.Cells[row, startColumn + c];
                    if (Safe<bool>(() => (bool)cell.HasFormula, false))
                    {
                        metadata.Formulas[c] = Safe(() => cell.Formula as string, null);
                    }
                    metadata.NumberFormats[c] = Safe(() => cell.NumberFormat as string, null);
                }
            }
            catch (Exception)
            {
            }
            return metadata;
        }

        /// <summary>
        /// Range.Value2 is used because it is the fast, locale-independent read,
        /// but it returns dates as OLE serial numbers. Without this conversion
        /// every date column profiles as a number in the 45000s -- which looks
        /// plausible, so nothing would flag it, and the model would be told a
        /// date column is numeric.
        ///
        /// The cell's number format is what distinguishes a date from a genuine
        /// number, since the value itself cannot.
        /// </summary>
        internal static void ConvertOleDates(List<object[]> rows, string[] numberFormats)
        {
            if (numberFormats == null)
            {
                return;
            }
            for (var c = 0; c < numberFormats.Length; c++)
            {
                if (!LooksLikeDateFormat(numberFormats[c]))
                {
                    continue;
                }
                foreach (var row in rows)
                {
                    if (row == null || c >= row.Length || !(row[c] is double serial))
                    {
                        continue;
                    }
                    // Excel's usable date range; outside it the value is not a date.
                    if (serial < 1 || serial > 2958465)
                    {
                        continue;
                    }
                    try { row[c] = DateTime.FromOADate(serial); }
                    catch (ArgumentException) { }
                }
            }
        }

        internal static bool LooksLikeDateFormat(string numberFormat)
        {
            if (string.IsNullOrEmpty(numberFormat) || numberFormat == "General")
            {
                return false;
            }
            // Strip literals so a currency symbol such as "d" inside quotes does
            // not read as a day placeholder.
            var cleaned = Regex.Replace(numberFormat, "\"[^\"]*\"", "");
            cleaned = Regex.Replace(cleaned, @"\[[^\]]*\]", "");
            var lower = cleaned.ToLowerInvariant();

            // A date format must contain a year, month or day placeholder.
            // "h:mm" alone is a time, which Excel also stores as a serial.
            return lower.Contains("yy") || lower.Contains("mmm")
                || (lower.Contains("d") && lower.Contains("m"))
                || lower.Contains("年") || lower.Contains("月");
        }

        // Sheet references look like 'Other Sheet'!A1 or Sheet2!A1.
        private static readonly Regex SheetReference =
            new Regex(@"(?:'([^']+)'|([A-Za-z0-9_一-龥]+))!\$?[A-Z]{1,3}\$?\d+",
                RegexOptions.Compiled);

        private static void CollectRelations(string fromSheet, string[] formulas, WorkbookIndex index)
        {
            if (formulas == null)
            {
                return;
            }
            for (var c = 0; c < formulas.Length; c++)
            {
                var formula = formulas[c];
                if (string.IsNullOrEmpty(formula))
                {
                    continue;
                }
                foreach (Match match in SheetReference.Matches(formula))
                {
                    var target = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                    if (string.IsNullOrEmpty(target) ||
                        target.Equals(fromSheet, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    var letter = TypeInference.ColumnLetter(c + 1);
                    if (!index.Relations.Any(r =>
                            r.FromSheet == fromSheet && r.ToSheet == target && r.ViaColumn == letter))
                    {
                        index.Relations.Add(new SheetRelation
                        {
                            FromSheet = fromSheet,
                            ToSheet = target,
                            ViaColumn = letter
                        });
                    }
                }
            }
        }

        private static void CollectNamedRanges(Workbook workbook, WorkbookIndex index)
        {
            foreach (Name name in workbook.Names)
            {
                var text = Safe(() => name.Name, null);
                var refersTo = Safe(() => name.RefersTo as string, null);
                if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(refersTo))
                {
                    continue;
                }
                // Excel's own bookkeeping names are noise to the model.
                if (text.StartsWith("_xl", StringComparison.OrdinalIgnoreCase) ||
                    text.Contains("Print_Area") || text.Contains("Print_Titles"))
                {
                    continue;
                }
                index.NamedRanges.Add(new NamedRangeIndex { Name = text, RefersTo = refersTo });
            }
        }

        private static T Safe<T>(Func<T> getter, T fallback = default)
        {
            try { return getter(); }
            catch (Exception) { return fallback; }
        }
    }
}
