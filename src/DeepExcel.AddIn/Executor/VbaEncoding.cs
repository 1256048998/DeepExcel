using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Executor
{
    public sealed class VbaEncodeResult
    {
        public string Code { get; set; }
        /// <summary>用到了 DeepExcelU 辅助函数，需要把 <see cref="VbaEncoding.HelperFunction"/> 一起注入</summary>
        public bool UsesHelper { get; set; }
        /// <summary>改写后有语句的续行超过 VBA 上限（24），这段代码编译不了</summary>
        public bool TooLong { get; set; }
        public int StringsRewritten { get; set; }
    }

    /// <summary>
    /// VBA 字符串的代码页安全。AddFromString 按系统 ANSI 代码页把代码转成字节：英文 Windows（1252）
    /// 上中文字面量全变成 "?"。表示不了的字符改写成运行时拼出来的表达式：
    ///
    /// - 只改写当前代码页表示不了的字符（中文 Windows 上中文原样保留，不白白拉长代码）
    /// - 一两个字用 ChrW(n)；更长的一段用 DeepExcelU("9500...")，每个字 4 位十六进制，
    ///   比 ChrW(38144) &amp; 短三倍——VBA 单行上限 1023 字符，80 个汉字就能撑爆旧写法
    /// - 行太长时在拼接处续行
    /// - Const 里不能调用函数：过程内的 Const 改成 Dim + 赋值，模块级的改成同名 Property Get
    /// - 注释不动（表示不了的字变成 ? 也不影响运行）
    /// </summary>
    public static class VbaEncoding
    {
        public const string HelperName = "DeepExcelU";
        public const int MaxLineLength = 900;
        public const int MaxContinuations = 24;
        private const int HelperChunk = 150;  // 每段 DeepExcelU 最多这么多字（600 个十六进制字符，行首再长也放得下）
        private const char Separator = '\u0001';

        public static string HelperFunction =>
            "Private Function " + HelperName + "(ByVal h As String) As String\r\n" +
            "    Dim i As Long, s As String\r\n" +
            "    For i = 1 To Len(h) Step 4\r\n" +
            // 字符串 "&H9500" 按 Integer 解析成负数；ChrW 接受 -32768..65535，负数照样映射回原字符
            "        s = s & ChrW(CLng(\"&H\" & Mid$(h, i, 4)))\r\n" +
            "    Next i\r\n" +
            "    " + HelperName + " = s\r\n" +
            "End Function";

        private static readonly Regex ProcStart = new Regex(
            @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+",
            RegexOptions.IgnoreCase);
        private static readonly Regex ProcEnd = new Regex(@"^\s*End\s+(?:Sub|Function|Property)\b", RegexOptions.IgnoreCase);
        private static readonly Regex ConstLine = new Regex(
            @"^(?<indent>\s*)(?:(?<scope>Public|Private|Global)\s+)?Const\s+(?<name>[A-Za-z_]\w*)(?:\s+As\s+\w+)?\s*=\s*(?<value>.*)$",
            RegexOptions.IgnoreCase);

        public static VbaEncodeResult Encode(string code, Encoding ansi = null)
        {
            var result = new VbaEncodeResult { Code = code };
            if (string.IsNullOrEmpty(code)) return result;
            var strict = Encoding.GetEncoding((ansi ?? Encoding.Default).CodePage,
                EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

            var output = new StringBuilder(code.Length + 64);
            var lines = code.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var depth = 0;
            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var codeOnly = CodeOutsideStrings(line);
                if (ProcStart.IsMatch(codeOnly)) depth++;
                else if (ProcEnd.IsMatch(codeOnly)) depth = Math.Max(0, depth - 1);

                var rewritten = RewriteLine(line, strict, result);
                if (rewritten != line)
                {
                    rewritten = RewriteConst(rewritten, depth > 0);
                    rewritten = Wrap(rewritten, result);
                }
                output.Append(rewritten.Replace(Separator.ToString(), " & "));
                if (i < lines.Length - 1) output.Append("\r\n");
            }
            result.Code = output.ToString();
            return result;
        }

        /// <summary>一行里需要改写的字符串字面量换成表达式（片段之间用 Separator 占位，稍后决定是否续行）</summary>
        private static string RewriteLine(string line, Encoding strict, VbaEncodeResult result)
        {
            if (line.IndexOf('"') < 0) return line;
            var output = new StringBuilder(line.Length);
            var i = 0;
            while (i < line.Length)
            {
                var ch = line[i];
                if (ch == '\'')
                {
                    output.Append(line, i, line.Length - i);  // 注释：原样
                    break;
                }
                if (ch != '"')
                {
                    output.Append(ch);
                    i++;
                    continue;
                }
                // 字符串字面量
                var content = new StringBuilder();
                var j = i + 1;
                var closed = false;
                while (j < line.Length)
                {
                    if (line[j] == '"')
                    {
                        if (j + 1 < line.Length && line[j + 1] == '"') { content.Append('"'); j += 2; continue; }
                        closed = true;
                        break;
                    }
                    content.Append(line[j]);
                    j++;
                }
                var original = line.Substring(i, (closed ? j + 1 : j) - i);
                var text = content.ToString();
                if (!closed || Representable(text, strict))
                {
                    output.Append(original);
                }
                else
                {
                    output.Append(Expression(text, strict, result));
                    result.StringsRewritten++;
                }
                i = closed ? j + 1 : j;
            }
            return output.ToString();
        }

        private static bool Representable(string text, Encoding strict)
        {
            if (text.Length == 0) return true;
            try
            {
                strict.GetBytes(text);
                return true;
            }
            catch (EncoderFallbackException)
            {
                return false;
            }
        }

        private static string Expression(string text, Encoding strict, VbaEncodeResult result)
        {
            var parts = new List<string>();
            var plain = new StringBuilder();
            var foreign = new StringBuilder();

            void FlushPlain()
            {
                if (plain.Length == 0) return;
                parts.Add("\"" + plain.ToString().Replace("\"", "\"\"") + "\"");
                plain.Clear();
            }

            void FlushForeign()
            {
                if (foreign.Length == 0) return;
                var run = foreign.ToString();
                if (run.Length <= 2)
                {
                    foreach (var c in run) parts.Add("ChrW(" + ((int)c).ToString(CultureInfo.InvariantCulture) + ")");
                }
                else
                {
                    for (var start = 0; start < run.Length; start += HelperChunk)
                    {
                        var chunk = run.Substring(start, Math.Min(HelperChunk, run.Length - start));
                        var hex = new StringBuilder(chunk.Length * 4);
                        foreach (var c in chunk) hex.Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
                        parts.Add(HelperName + "(\"" + hex + "\")");
                    }
                    result.UsesHelper = true;
                }
                foreign.Clear();
            }

            foreach (var c in text)
            {
                if (Representable(c.ToString(), strict))
                {
                    FlushForeign();
                    plain.Append(c);
                }
                else
                {
                    FlushPlain();
                    foreign.Append(c);
                }
            }
            FlushPlain();
            FlushForeign();
            // 用括号包起来，改写后的表达式在任何位置都保持原来的优先级
            return parts.Count == 1 ? parts[0] : "(" + string.Join(Separator.ToString(), parts) + ")";
        }

        /// <summary>Const 的值不能是函数调用：过程内改成 Dim + 赋值，模块级改成同名 Property Get</summary>
        private static string RewriteConst(string line, bool insideProcedure)
        {
            var codeOnly = CodeOutsideStrings(line.Replace(Separator, ' '));
            var match = ConstLine.Match(line);
            if (!match.Success || !ConstLine.IsMatch(codeOnly)) return line;
            var name = match.Groups["name"].Value;
            var value = match.Groups["value"].Value;
            var indent = match.Groups["indent"].Value;
            if (insideProcedure)
            {
                return $"{indent}Dim {name} As String: {name} = {value}";
            }
            var scope = match.Groups["scope"].Value;
            var visibility = scope.Equals("Public", StringComparison.OrdinalIgnoreCase) ||
                             scope.Equals("Global", StringComparison.OrdinalIgnoreCase) ? "Public" : "Private";
            return $"{indent}{visibility} Property Get {name}() As String\r\n{indent}    {name} = {value}\r\n{indent}End Property";
        }

        /// <summary>超过 MaxLineLength 的物理行，在改写产生的拼接处续行</summary>
        private static string Wrap(string line, VbaEncodeResult result)
        {
            if (line.IndexOf(Separator) < 0) return line;
            var output = new StringBuilder(line.Length + 32);
            var segments = line.Split(Separator);
            var physical = 0;
            var continuations = 0;
            for (var k = 0; k < segments.Length; k++)
            {
                var segment = segments[k];
                var firstBreak = segment.IndexOf('\n');
                var head = firstBreak < 0 ? segment.Length : firstBreak;
                // 看下一段放不放得下：放不下就先续行（不是等已经超了才断）
                if (k > 0 && physical + 3 + head > MaxLineLength - 2)
                {
                    output.Append(" & _\r\n        ");
                    physical = 8;
                    continuations++;
                }
                else if (k > 0)
                {
                    output.Append(" & ");
                    physical += 3;
                }
                output.Append(segment);
                var lastBreak = segment.LastIndexOf('\n');
                physical = lastBreak < 0 ? physical + segment.Length : segment.Length - lastBreak - 1;
            }
            if (continuations > MaxContinuations) result.TooLong = true;
            return output.ToString();
        }

        private static string CodeOutsideStrings(string line)
        {
            var output = new StringBuilder(line.Length);
            var inString = false;
            foreach (var ch in line)
            {
                if (ch == '"') { inString = !inString; output.Append('"'); continue; }
                if (!inString && ch == '\'') break;
                output.Append(inString ? ' ' : ch);
            }
            return output.ToString();
        }
    }
}
