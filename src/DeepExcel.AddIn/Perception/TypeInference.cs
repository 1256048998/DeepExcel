using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// Infers column shape from sampled values.
    ///
    /// Pure: it takes arrays and returns profiles, with no Excel dependency. That
    /// is deliberate -- this is the logic most likely to be wrong in a way that
    /// silently misleads the model, so it has to be testable against awkward real
    /// data (mixed types, stray totals rows, dates stored as text) rather than
    /// only against whatever workbook happens to be open.
    /// </summary>
    public static class TypeInference
    {
        private const int MaxSamples = 3;
        private const int MaxEmptyRowsReported = 5;

        /// <summary>
        /// Detects the header row.
        ///
        /// Only the first row is considered, and only when it is entirely text
        /// while the rows below are not. Guessing more aggressively is worse than
        /// not guessing: treating a data row as a header shifts every row number
        /// the model reasons about, and it has no way to notice.
        /// </summary>
        public static bool LooksLikeHeaderRow(object[][] rows)
        {
            if (rows == null || rows.Length < 2)
            {
                return false;
            }

            var first = rows[0];
            if (first == null || first.Length == 0)
            {
                return false;
            }

            var textCells = 0;
            var nonEmpty = 0;
            foreach (var cell in first)
            {
                if (IsEmpty(cell)) continue;
                nonEmpty++;
                if (cell is string s && !LooksNumeric(s)) textCells++;
            }
            // A header with one populated cell is more likely a stray title.
            if (nonEmpty < 2 || textCells != nonEmpty)
            {
                return false;
            }

            // The body must not look like headers too, or a text-only table
            // would lose its first data row.
            var bodyTextRows = 0;
            var examined = 0;
            for (var i = 1; i < rows.Length && examined < 5; i++)
            {
                if (rows[i] == null) continue;
                examined++;
                var allText = true;
                var any = false;
                foreach (var cell in rows[i])
                {
                    if (IsEmpty(cell)) continue;
                    any = true;
                    if (!(cell is string s) || LooksNumeric(s)) { allText = false; break; }
                }
                if (any && allText) bodyTextRows++;
            }
            return examined == 0 || bodyTextRows < examined;
        }

        /// <summary>
        /// Profiles every column.
        /// </summary>
        /// <param name="rows">Sampled rows, each the same width. Nulls allowed.</param>
        /// <param name="firstRowNumber">Sheet row number of rows[0], for reporting blanks.</param>
        /// <param name="headers">Header text per column, or null.</param>
        /// <param name="formulas">Formula text per column where present, or null.</param>
        public static List<ColumnProfile> ProfileColumns(
            object[][] rows,
            int firstRowNumber = 1,
            string[] headers = null,
            string[] formulas = null)
        {
            var profiles = new List<ColumnProfile>();
            if (rows == null || rows.Length == 0)
            {
                return profiles;
            }

            var width = rows.Where(r => r != null).Select(r => r.Length).DefaultIfEmpty(0).Max();
            for (var column = 0; column < width; column++)
            {
                profiles.Add(ProfileColumn(rows, column, firstRowNumber, headers, formulas));
            }
            return profiles;
        }

        private static ColumnProfile ProfileColumn(
            object[][] rows, int column, int firstRowNumber, string[] headers, string[] formulas)
        {
            var profile = new ColumnProfile
            {
                Letter = ColumnLetter(column + 1),
                Header = headers != null && column < headers.Length ? headers[column] : null
            };

            var values = new List<object>();
            for (var i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                var cell = row != null && column < row.Length ? row[column] : null;
                if (IsEmpty(cell))
                {
                    profile.EmptyCount++;
                    if (profile.EmptyRows.Count < MaxEmptyRowsReported)
                    {
                        profile.EmptyRows.Add(firstRowNumber + i);
                    }
                    continue;
                }
                values.Add(cell);
            }

            profile.NonEmptyCount = values.Count;

            var formula = formulas != null && column < formulas.Length ? formulas[column] : null;
            if (!string.IsNullOrEmpty(formula))
            {
                profile.Kind = ColumnKind.Formula;
                profile.Confidence = 1.0;
                profile.FormulaSample = formula.Length > 40 ? formula.Substring(0, 40) + "…" : formula;
                return profile;
            }

            if (values.Count == 0)
            {
                profile.Kind = ColumnKind.Empty;
                profile.Confidence = 1.0;
                return profile;
            }

            var counts = new Dictionary<ColumnKind, int>();
            foreach (var value in values)
            {
                var kind = ClassifyValue(value);
                counts[kind] = counts.TryGetValue(kind, out var existing) ? existing + 1 : 1;
            }

            var dominant = counts.OrderByDescending(kv => kv.Value).First();
            profile.Confidence = (double)dominant.Value / values.Count;
            profile.Kind = profile.Confidence >= 0.5 ? dominant.Key : ColumnKind.Mixed;

            FillStatistics(profile, values);
            return profile;
        }

        private static void FillStatistics(ColumnProfile profile, List<object> values)
        {
            var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in values)
            {
                if (distinct.Count >= ColumnProfile.DistinctCap)
                {
                    profile.DistinctCapped = true;
                    break;
                }
                distinct.Add(ColumnProfile.FormatValue(value));
            }
            profile.DistinctCount = distinct.Count;
            profile.IsUnique = !profile.DistinctCapped && distinct.Count == values.Count;

            if (profile.Kind == ColumnKind.Number)
            {
                var numbers = values.Select(ToNumber).Where(n => n.HasValue).Select(n => n.Value).ToList();
                if (numbers.Count > 0)
                {
                    profile.MinDisplay = ColumnProfile.FormatValue(numbers.Min());
                    profile.MaxDisplay = ColumnProfile.FormatValue(numbers.Max());
                }
            }
            else if (profile.Kind == ColumnKind.Date)
            {
                var dates = values.Select(ToDate).Where(d => d.HasValue).Select(d => d.Value).ToList();
                if (dates.Count > 0)
                {
                    profile.MinDisplay = ColumnProfile.FormatValue(dates.Min());
                    profile.MaxDisplay = ColumnProfile.FormatValue(dates.Max());
                }
            }

            // Samples come from distinct values so a column of one repeated value
            // does not print it three times.
            foreach (var sample in distinct)
            {
                if (profile.Samples.Count >= MaxSamples) break;
                if (!string.IsNullOrEmpty(sample)) profile.Samples.Add(sample);
            }
        }

        public static ColumnKind ClassifyValue(object value)
        {
            if (IsEmpty(value)) return ColumnKind.Empty;
            if (value is bool) return ColumnKind.Boolean;
            if (value is DateTime) return ColumnKind.Date;
            if (value is double || value is int || value is long || value is decimal || value is float)
            {
                return ColumnKind.Number;
            }

            if (value is string text)
            {
                var trimmed = text.Trim();
                if (trimmed.Length == 0) return ColumnKind.Empty;
                if (trimmed.StartsWith("=")) return ColumnKind.Formula;
                if (LooksBoolean(trimmed)) return ColumnKind.Boolean;
                // Dates first: "2026-01-01" is also not numeric, but checking
                // numbers first would misread "20260101" as a number, which is
                // the correct reading anyway.
                if (ToDate(trimmed).HasValue) return ColumnKind.Date;
                if (LooksNumeric(trimmed)) return ColumnKind.Number;
                return ColumnKind.Text;
            }

            return ColumnKind.Text;
        }

        public static bool IsEmpty(object value)
        {
            if (value == null || value == DBNull.Value) return true;
            return value is string s && s.Trim().Length == 0;
        }

        private static bool LooksBoolean(string text)
        {
            return text.Equals("true", StringComparison.OrdinalIgnoreCase)
                || text.Equals("false", StringComparison.OrdinalIgnoreCase)
                || text == "是" || text == "否";
        }

        /// <summary>
        /// Accepts the shapes a spreadsheet actually produces: thousands
        /// separators, currency symbols, trailing percent, parenthesised
        /// negatives. A column of "¥1,200.00" is a number column, and reporting
        /// it as text would send the model down a string-parsing path.
        /// </summary>
        public static bool LooksNumeric(string text)
        {
            return ToNumber(text).HasValue;
        }

        public static double? ToNumber(object value)
        {
            if (value is double d) return d;
            if (value is int i) return i;
            if (value is long l) return l;
            if (value is decimal m) return (double)m;
            if (value is float f) return f;

            if (!(value is string raw)) return null;
            var text = raw.Trim();
            if (text.Length == 0) return null;

            var negative = false;
            if (text.StartsWith("(") && text.EndsWith(")"))
            {
                negative = true;
                text = text.Substring(1, text.Length - 2).Trim();
            }

            var percent = text.EndsWith("%");
            if (percent) text = text.Substring(0, text.Length - 1).Trim();

            foreach (var symbol in new[] { "¥", "$", "€", "£", "￥", "元", "人民币" })
            {
                if (text.StartsWith(symbol)) text = text.Substring(symbol.Length).Trim();
                if (text.EndsWith(symbol)) text = text.Substring(0, text.Length - symbol.Length).Trim();
            }
            text = text.Replace(",", "").Replace("，", "");
            if (text.Length == 0) return null;

            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return null;
            }
            if (percent) parsed /= 100.0;
            return negative ? -parsed : parsed;
        }

        private static readonly string[] DateFormats =
        {
            "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyy年M月d日",
            "yyyy-M-d", "yyyy/M/d",
            "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm", "yyyy/MM/dd HH:mm:ss",
            "MM/dd/yyyy", "dd/MM/yyyy", "yyyyMMdd"
        };

        public static DateTime? ToDate(object value)
        {
            if (value is DateTime dt) return dt;
            if (!(value is string raw)) return null;
            var text = raw.Trim();
            // "20260101" parses as yyyyMMdd but so would an 8-digit order number.
            // Requiring a separator keeps identifiers out of the date bucket.
            if (text.Length < 6 || (text.Length == 8 && text.All(char.IsDigit)))
            {
                return null;
            }
            if (DateTime.TryParseExact(text, DateFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var exact))
            {
                return exact;
            }
            return null;
        }

        /// <summary>1 -> A, 27 -> AA.</summary>
        public static string ColumnLetter(int oneBasedIndex)
        {
            if (oneBasedIndex < 1) return "";
            var letters = "";
            var n = oneBasedIndex;
            while (n > 0)
            {
                var remainder = (n - 1) % 26;
                letters = (char)('A' + remainder) + letters;
                n = (n - 1) / 26;
            }
            return letters;
        }
    }
}
