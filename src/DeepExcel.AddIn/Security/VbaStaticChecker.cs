using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Security
{
    public enum VbaIssueLevel { Warning, Error }

    public sealed class VbaIssue
    {
        public VbaIssueLevel Level { get; set; }
        /// <summary>1 起的行号（相对模型给的代码）；0 表示整段</summary>
        public int Line { get; set; }
        public string Message { get; set; }

        public override string ToString() => Line > 0 ? $"第 {Line} 行：{Message}" : Message;
    }

    /// <summary>
    /// VBA 编译前静态检查。VBA 的编译错误和死循环在 Excel 里的表现是弹出对话框或整个 Excel
    /// 卡死；在注入之前把模型常犯的错误挑出来，error 级直接退回给模型改。
    ///
    /// 只做行级词法分析（字符串、注释），不做完整语法分析：宁可漏报，不能误报
    /// ——误报会让模型改一段本来能跑的代码。
    /// </summary>
    public static class VbaStaticChecker
    {
        public const int MaxContinuations = 24;

        private static readonly Regex ProcStart = new Regex(
            @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+",
            RegexOptions.IgnoreCase);
        private static readonly Regex ProcEnd = new Regex(@"^\s*End\s+(?:Sub|Function|Property)\b", RegexOptions.IgnoreCase);
        private static readonly Regex BlockStart = new Regex(@"^\s*(?:(?:Public|Private)\s+)?(?:Type|Enum)\s+\w+", RegexOptions.IgnoreCase);
        private static readonly Regex BlockEnd = new Regex(@"^\s*End\s+(?:Type|Enum)\b", RegexOptions.IgnoreCase);
        private static readonly Regex ModuleLevelOk = new Regex(
            @"^\s*(?:$|Option\b|Dim\b|Private\b|Public\b|Global\b|Const\b|Declare\b|Attribute\b|#|Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var)\b|Implements\b|Event\b)",
            RegexOptions.IgnoreCase);
        private static readonly Regex OptionLine = new Regex(@"^\s*Option\s+(?:Explicit|Base|Compare|Private)\b", RegexOptions.IgnoreCase);
        private static readonly Regex FormulaAssign = new Regex(@"\.(?:Formula\w*|Value2?)\s*=|\bFormula\w*\s*:=", RegexOptions.IgnoreCase);
        private static readonly Regex ShowCall = new Regex(@"\.\s*Show\b", RegexOptions.IgnoreCase);        private static readonly Regex SelfModify = new Regex(@"\b(?:VBProject|VBComponents|CodeModule|VBE)\b", RegexOptions.IgnoreCase);
        private static readonly Regex DoStart = new Regex(@"^\s*Do\b(?<cond>.*)$", RegexOptions.IgnoreCase);
        private static readonly Regex LoopEnd = new Regex(@"^\s*Loop\b(?<cond>.*)$", RegexOptions.IgnoreCase);
        private static readonly Regex WhileStart = new Regex(@"^\s*While\s+(?<cond>.+)$", RegexOptions.IgnoreCase);
        private static readonly Regex WendEnd = new Regex(@"^\s*Wend\b", RegexOptions.IgnoreCase);
        // 跳出所有循环：Exit Sub / Function / Property、GoTo（不含 On Error GoTo）、独立的 End
        private static readonly Regex LeaveAll = new Regex(
            @"\bExit\s+(?:Sub|Function|Property)\b|(?<!On\s+Error\s+)\bGoTo\b|^\s*End\s*$", RegexOptions.IgnoreCase);
        private static readonly Regex ExitDo = new Regex(@"\bExit\s+Do\b", RegexOptions.IgnoreCase);
        private static readonly Regex AlwaysTrue = new Regex(@"^\s*(?:While\s+(?:True|1|-1)|Until\s+(?:False|0))\s*$", RegexOptions.IgnoreCase);
        private static readonly Regex ResumeNext = new Regex(@"\bOn\s+Error\s+Resume\s+Next\b", RegexOptions.IgnoreCase);
        private static readonly Regex ErrorGoto0 = new Regex(@"\bOn\s+Error\s+GoTo\s+0\b", RegexOptions.IgnoreCase);

        /// <summary>一行拆成：代码部分（字符串内容换成空格）、是否在字符串里结束（引号不配对）</summary>
        internal static (string Code, bool Unterminated, bool HasComment) Lex(string line)
        {
            var code = new StringBuilder(line.Length);
            var inString = false;
            for (var i = 0; i < line.Length; i++)
            {
                var ch = line[i];
                if (ch == '"')
                {
                    if (inString && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        code.Append("  ");
                        i++;
                        continue;
                    }
                    inString = !inString;
                    code.Append('"');
                    continue;
                }
                if (!inString && ch == '\'')
                {
                    return (code.ToString(), false, true);
                }
                code.Append(inString ? ' ' : ch);
            }
            var text = code.ToString();
            // Rem 注释：只认语句开头的 Rem
            if (!inString && Regex.IsMatch(text, @"^\s*Rem(?:\s|$)", RegexOptions.IgnoreCase))
            {
                return ("", false, true);
            }
            return (text, inString, false);
        }

        public static List<VbaIssue> Check(string code)
        {
            var issues = new List<VbaIssue>();
            if (string.IsNullOrWhiteSpace(code)) return issues;
            var lines = code.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var hasProcedures = lines.Any(l => ProcStart.IsMatch(Lex(l).Code));

            var depth = 0;          // 在 Sub/Function 里
            var blockDepth = 0;     // 在 Type/Enum 里
            var seenProcedure = false;
            var continuation = 0;
            var continuationStart = 0;
            var resumeNextLine = 0;
            var loops = new List<OpenLoop>();  // 末尾是最内层

            for (var index = 0; index < lines.Length; index++)
            {
                var lineNo = index + 1;
                var raw = lines[index];
                var lexed = Lex(raw);
                var codePart = lexed.Code;
                var trimmed = codePart.Trim();

                if (lexed.Unterminated)
                {
                    issues.Add(Error(lineNo, "引号不配对：VBA 的字符串不能跨行；字符串里的双引号要写成两个 \"\""));
                }

                if (raw.IndexOf("\\\"", StringComparison.Ordinal) >= 0 && FormulaAssign.IsMatch(codePart))
                {
                    issues.Add(Error(lineNo,
                        "公式字符串里用了 \\\"：VBA 不认反斜杠转义，字符串里的双引号要写成两个 \"\"（例如 \"=IF(A1=\"\"是\"\",1,0)\"）"));
                }

                // 续行：上一行以 _ 结尾时，这一行是同一条语句的后半截
                var continuedFromAbove = continuation > 0;
                var continues = Regex.IsMatch(codePart, @"\s_\s*$");
                if (continues)
                {
                    if (continuation == 0) continuationStart = lineNo;
                    continuation++;
                    if (continuation == MaxContinuations + 1)
                    {
                        issues.Add(Error(continuationStart, $"一条语句的续行（行尾 _）超过 {MaxContinuations} 个，VBA 无法编译；请拆成多条语句或用数组"));
                    }
                }
                else
                {
                    continuation = 0;
                }

                if (trimmed.Length == 0) continue;

                if (OptionLine.IsMatch(codePart) && seenProcedure)
                {
                    issues.Add(Error(lineNo, "Option 语句必须放在模块最前面（所有 Sub / Function 之前）"));
                }

                if (SelfModify.IsMatch(codePart))
                {
                    issues.Add(Error(lineNo, "不能读写 VBA 工程本身（VBProject / VBComponents / CodeModule）"));
                }

                if (ShowCall.IsMatch(codePart))
                {
                    issues.Add(Error(lineNo, "不能用 .Show 打开窗体或对话框：它会等用户操作，Excel 会卡住；请直接操作单元格"));
                }
                // MsgBox / InputBox 在 VBAExecutor.TryPrepareCode 里已经拦下

                if (BlockStart.IsMatch(codePart) && depth == 0) { blockDepth++; continue; }
                if (BlockEnd.IsMatch(codePart)) { blockDepth = Math.Max(0, blockDepth - 1); continue; }

                if (ProcStart.IsMatch(codePart))
                {
                    depth++;
                    seenProcedure = true;
                    continue;
                }
                if (ProcEnd.IsMatch(codePart))
                {
                    depth = Math.Max(0, depth - 1);
                    if (depth == 0)
                    {
                        ReportEndlessLoops(loops, issues);
                        loops.Clear();
                        if (resumeNextLine > 0)
                        {
                            issues.Add(Warning(resumeNextLine, "On Error Resume Next 之后没有 On Error GoTo 0：出错会被静默跳过，结果可能不完整"));
                            resumeNextLine = 0;
                        }
                    }
                    continue;
                }

                if (hasProcedures && depth == 0 && blockDepth == 0 && !continuedFromAbove && !ModuleLevelOk.IsMatch(codePart))
                {
                    issues.Add(Error(lineNo, "这一行在 Sub / Function 外面：可执行语句必须写在 Sub … End Sub 里"));
                    continue;
                }

                if (ResumeNext.IsMatch(codePart)) resumeNextLine = lineNo;
                if (ErrorGoto0.IsMatch(codePart)) resumeNextLine = 0;

                // 循环出口
                var doMatch = DoStart.Match(codePart);
                if (doMatch.Success)
                {
                    var cond = doMatch.Groups["cond"].Value;
                    var needsExit = string.IsNullOrWhiteSpace(cond) || AlwaysTrue.IsMatch(cond);
                    loops.Add(new OpenLoop { Line = lineNo, NeedsExit = needsExit, IsDo = true });
                    continue;
                }
                var whileMatch = WhileStart.Match(codePart);
                if (whileMatch.Success)
                {
                    var cond = whileMatch.Groups["cond"].Value.Trim();
                    var needsExit = Regex.IsMatch(cond, @"^(?:True|1|-1)$", RegexOptions.IgnoreCase);
                    loops.Add(new OpenLoop { Line = lineNo, NeedsExit = needsExit, IsDo = false });
                    continue;
                }
                var loopMatch = LoopEnd.Match(codePart);
                if ((loopMatch.Success || WendEnd.IsMatch(codePart)) && loops.Count > 0)
                {
                    var open = loops[loops.Count - 1];
                    loops.RemoveAt(loops.Count - 1);
                    var closingCondition = loopMatch.Success ? loopMatch.Groups["cond"].Value : "";
                    var conditionAtEnd = !string.IsNullOrWhiteSpace(closingCondition) && !AlwaysTrue.IsMatch(closingCondition);
                    if (open.NeedsExit && !open.HasExit && !conditionAtEnd)
                    {
                        issues.Add(Error(open.Line, "循环没有出口（没有条件、也没有 Exit Do），会让 Excel 卡死"));
                    }
                    continue;
                }
                if (loops.Count > 0)
                {
                    if (LeaveAll.IsMatch(codePart))
                    {
                        foreach (var open in loops) open.HasExit = true;
                    }
                    else if (ExitDo.IsMatch(codePart))
                    {
                        var innermostDo = loops.LastOrDefault(l => l.IsDo);
                        if (innermostDo != null) innermostDo.HasExit = true;
                    }
                }
            }

            ReportEndlessLoops(loops, issues);
            return issues;
        }

        private sealed class OpenLoop
        {
            public int Line;
            public bool NeedsExit;  // 没有条件（Do … Loop）或条件恒真（Do While True）
            public bool HasExit;
            public bool IsDo;
        }

        private static void ReportEndlessLoops(IEnumerable<OpenLoop> loops, List<VbaIssue> issues)
        {
            foreach (var open in loops.Where(l => l.NeedsExit && !l.HasExit))
            {
                issues.Add(Error(open.Line, "循环没有出口（没有条件、也没有 Exit Do），会让 Excel 卡死"));
            }
        }

        /// <summary>给模型的一段话：error 全列出，最多 8 条</summary>
        public static string Describe(IEnumerable<VbaIssue> issues)
        {
            var list = issues.Take(8).Select(i => "- " + i).ToList();
            return list.Count == 0 ? null : string.Join("\n", list);
        }

        private static VbaIssue Error(int line, string message) => new VbaIssue { Level = VbaIssueLevel.Error, Line = line, Message = message };
        private static VbaIssue Warning(int line, string message) => new VbaIssue { Level = VbaIssueLevel.Warning, Line = line, Message = message };
    }
}
