using System;
using System.Globalization;
using System.Text;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>
    /// What a formula becomes when AutoFill copies it some rows down: relative
    /// row references move with it, $-anchored ones stay.
    ///
    /// A small scanner rather than a regular expression, because the parts
    /// that must not move are positional -- text inside "string literals",
    /// quoted sheet names, structured references like [@金额], and names or
    /// function calls that merely look like references (LOG10(, Sheet2).
    /// Column-only ranges (A:A) do not move; row-only ranges (2:2) do.
    /// </summary>
    internal static class FormulaShift
    {
        private const int MaxRow = 1048576;

        public static string ShiftRows(string formula, int rows)
        {
            if (string.IsNullOrEmpty(formula) || rows == 0) return formula;

            var result = new StringBuilder(formula.Length + 8);
            var i = 0;
            while (i < formula.Length)
            {
                var ch = formula[i];
                if (ch == '"' || ch == '\'')
                {
                    var end = QuotedEnd(formula, i, ch);
                    result.Append(formula, i, end - i);
                    i = end;
                    continue;
                }
                if (ch == '[')
                {
                    var end = BracketEnd(formula, i);
                    result.Append(formula, i, end - i);
                    i = end;
                    continue;
                }

                var tokenStart = i == 0 || !(IsIdent(formula[i - 1]) || formula[i - 1] == '$');
                if (tokenStart)
                {
                    if (TryRowRange(formula, i, rows, out var length, out var shifted) ||
                        TryCell(formula, i, rows, out length, out shifted))
                    {
                        result.Append(shifted);
                        i += length;
                        continue;
                    }
                }

                if (IsIdent(ch))
                {
                    // Copy the whole identifier so nothing inside a name or a
                    // number is mistaken for a reference.
                    var end = i;
                    while (end < formula.Length && IsIdent(formula[end])) end++;
                    result.Append(formula, i, end - i);
                    i = end;
                    continue;
                }

                result.Append(ch);
                i++;
            }
            return result.ToString();
        }

        private static bool IsIdent(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\';
        }

        /// <summary>End (exclusive) of a quoted run; a doubled quote is an escape.</summary>
        private static int QuotedEnd(string text, int start, char quote)
        {
            var i = start + 1;
            while (i < text.Length)
            {
                if (text[i] == quote)
                {
                    if (i + 1 < text.Length && text[i + 1] == quote)
                    {
                        i += 2;
                        continue;
                    }
                    return i + 1;
                }
                i++;
            }
            return text.Length;
        }

        private static int BracketEnd(string text, int start)
        {
            var depth = 0;
            var i = start;
            while (i < text.Length)
            {
                if (text[i] == '[') depth++;
                else if (text[i] == ']') depth--;
                i++;
                if (depth == 0) break;
            }
            return i;
        }

        /// <summary>$A$1, A1, $A1, A$1 -- not LOG10( and not part of a longer name.</summary>
        private static bool TryCell(string text, int start, int rows, out int length, out string shifted)
        {
            length = 0;
            shifted = null;
            var i = start;
            var colAbs = i < text.Length && text[i] == '$';
            if (colAbs) i++;
            var lettersStart = i;
            while (i < text.Length && IsAsciiLetter(text[i])) i++;
            var letters = i - lettersStart;
            if (letters < 1 || letters > 3) return false;
            var rowAbs = i < text.Length && text[i] == '$';
            if (rowAbs) i++;
            var digitsStart = i;
            while (i < text.Length && char.IsDigit(text[i])) i++;
            if (i == digitsStart || i - digitsStart > 7) return false;
            if (i < text.Length && (IsIdent(text[i]) || text[i] == '(' || text[i] == '!')) return false;

            var row = int.Parse(text.Substring(digitsStart, i - digitsStart), CultureInfo.InvariantCulture);
            var newRow = rowAbs ? row : row + rows;
            length = i - start;
            shifted = newRow < 1 || newRow > MaxRow
                ? "#REF!"
                : text.Substring(start, digitsStart - start) + newRow.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        /// <summary>Whole rows such as 2:5 or $2:$2.</summary>
        private static bool TryRowRange(string text, int start, int rows, out int length, out string shifted)
        {
            length = 0;
            shifted = null;
            var i = start;
            if (!TryRowPart(text, ref i, out var abs1, out var row1)) return false;
            if (i >= text.Length || text[i] != ':') return false;
            i++;
            if (!TryRowPart(text, ref i, out var abs2, out var row2)) return false;
            if (i < text.Length && (IsIdent(text[i]) || text[i] == '(')) return false;

            var new1 = abs1 ? row1 : row1 + rows;
            var new2 = abs2 ? row2 : row2 + rows;
            length = i - start;
            shifted = new1 < 1 || new2 < 1 || new1 > MaxRow || new2 > MaxRow
                ? "#REF!"
                : (abs1 ? "$" : "") + new1.ToString(CultureInfo.InvariantCulture) + ":" +
                  (abs2 ? "$" : "") + new2.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static bool TryRowPart(string text, ref int i, out bool absolute, out int row)
        {
            row = 0;
            absolute = i < text.Length && text[i] == '$';
            var start = absolute ? i + 1 : i;
            var end = start;
            while (end < text.Length && char.IsDigit(text[end])) end++;
            if (end == start || end - start > 7) return false;
            row = int.Parse(text.Substring(start, end - start), CultureInfo.InvariantCulture);
            i = end;
            return true;
        }

        private static bool IsAsciiLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }
    }
}
