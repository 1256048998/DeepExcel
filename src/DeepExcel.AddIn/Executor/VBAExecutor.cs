using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// VBA执行引擎 - 生成并执行VBA代码
    /// 安全网：执行前自动快照，失败自动回滚
    /// </summary>
    public class VBAExecutor
    {
        private readonly Microsoft.Office.Interop.Excel.Application _app;
        private const string DefaultMacroName = "DeepExcel_TempMacro";
        private const int ExcelBusyHResult = unchecked((int)0x800AC472);
        private const string RunOkMarker = "__DEEPEXCEL_VBA_OK__";
        private const string RunErrorMarker = "__DEEPEXCEL_VBA_ERROR__";
        private const char RunResultSeparator = (char)30;
        private static readonly Regex ProcedureRegex = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Static)\s+)*Sub\s+([^\s(]+)\s*\(([^)]*)\)",
            RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);

        public VBAExecutor(Microsoft.Office.Interop.Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 执行VBA代码
        /// </summary>
        public ToolResult Execute(string vbaCode, string macroName = null)
        {
            if (string.IsNullOrWhiteSpace(vbaCode))
            {
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = "VBA代码为空"
                };
            }

            if (!TryPrepareCode(vbaCode, macroName, out var procCode, out var entryPoint, out var preparationError))
            {
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = preparationError,
                    Suggestion = "VBA 入口必须是不带参数的 Sub；过程名和变量名请使用英文、数字和下划线。"
                };
            }

            // ★ P0-3 沙箱校验：阻止 LLM 执行 Shell / WScript.Shell / 文件系统 / 网络请求
            var sandboxError = DeepExcel.AddIn.Security.CodeSandbox.ValidateVba(procCode);
            if (sandboxError != null)
            {
                Logger.Instance.Warning("VBAExecutor", "Code blocked by sandbox: " + sandboxError);
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = sandboxError
                };
            }

            string moduleName = null;
            string wrapperName = null;
            string stage = "preflight";
            VBProject vbProject = null;
            VBComponent module = null;
            Exception executionError = null;
            bool screenUpdatingChanged = false;
            bool previousScreenUpdating = true;

            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null)
                {
                    return new ToolResult
                    {
                        Name = "execute_vba",
                        Success = false,
                        Error = "没有活动工作簿"
                    };
                }

                stage = "VBA project access";
                vbProject = wb.VBProject as VBProject;
                if (vbProject == null)
                {
                    return new ToolResult
                    {
                        Name = "execute_vba",
                        Success = false,
                        Error = "VBA项目不可访问"
                    };
                }

                // Every execution gets an isolated module. Reusing the old
                // DeepExcelModule left stale procedures behind after the second
                // run, producing intermittent "macro unavailable" errors.
                stage = "temporary module creation";
                moduleName = "DeepExcelTmp" + Guid.NewGuid().ToString("N").Substring(0, 8);
                module = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                module.Name = moduleName;

                // ★ VBA Unicode 编码转换：系统 ANSI 代码页不支持中文时，AddFromString 会把中文变 "?"
                // 将 VBA 字符串字面量中的非 ASCII 字符自动转换为 ChrW() 调用
                stage = "code injection";
                string encodedCode = EncodeVbaUnicode(procCode);
                wrapperName = "DeepExcelInvoke" + Guid.NewGuid().ToString("N").Substring(0, 8);
                encodedCode += "\r\n\r\n" + BuildInvocationWrapper(entryPoint, wrapperName);
                module.CodeModule.AddFromString(encodedCode);

                // ★ 诊断：记录注入后的实际代码内容 + 模块状态，便于排查"宏被禁用"
                try
                {
                    var lineCount = module.CodeModule.CountOfLines;
                    var firstLines = lineCount > 0 ? module.CodeModule.get_Lines(1, Math.Min(lineCount, 8)) : "";
                    Logger.Instance.Info("VBAExecutor",
                        $"After AddFromString: module={module.Name}, entry={entryPoint}, lineCount={lineCount}, firstLines=\n{firstLines}");
                }
                catch (Exception logEx) { Logger.Instance.Warning("VBAExecutor", "Log code failed: " + logEx.Message); }

                // ★ 第一性原理：_app.Run 在 UI 线程同步执行 VBA，期间 WebView2 渲染进程
                // 无法与 UI 线程通信（消息泵被阻塞）。如果 VBA 执行超过几秒，
                // WebView2 进程会因心跳超时崩溃，导致对话面板白屏/消失。
                // 解决：执行前处理一次消息泵，让 WebView2 有机会完成 pending 的渲染。
                var workbookName = wb.Name.Replace("'", "''");
                var qualifiedMacroName = $"'{workbookName}'!{moduleName}.{wrapperName}";
                Logger.Instance.Info("VBAExecutor", $"VBA Run START: macro={qualifiedMacroName}, entry={entryPoint}, codeLen={procCode.Length}");
                System.Windows.Forms.Application.DoEvents();

                // 4. 执行宏
                // ★ 临时关闭 Application.ScreenUpdating 提升性能 + 防闪烁
                stage = "macro execution";
                previousScreenUpdating = _app.ScreenUpdating;
                _app.ScreenUpdating = false;
                screenUpdatingChanged = true;
                var runSw = System.Diagnostics.Stopwatch.StartNew();
                var runResult = RunMacroWithBusyRetry(qualifiedMacroName);
                ThrowIfVbaReportedError(runResult);
                runSw.Stop();

                // 执行后立即处理消息泵，恢复 WebView2 心跳
                System.Windows.Forms.Application.DoEvents();
                Logger.Instance.Info("VBAExecutor", $"VBA Run END: elapsed={runSw.ElapsedMilliseconds}ms");

                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = true,
                    Data = new { macro = entryPoint }
                };
            }
            catch (Exception ex)
            {
                executionError = ex;
            }
            finally
            {
                if (screenUpdatingChanged)
                {
                    try { _app.ScreenUpdating = previousScreenUpdating; }
                    catch (Exception ex) { Logger.Instance.Warning("VBAExecutor", "Restore ScreenUpdating failed: " + ex.Message); }
                }

                // Remove only the unique module created by this invocation.
                // Never remove Module1/Module2: those may belong to the user.
                try
                {
                    if (module != null && vbProject != null)
                    {
                        vbProject.VBComponents.Remove(module);
                    }
                }
                catch (Exception cleanupError)
                {
                    Logger.Instance.Warning("VBAExecutor", "Temporary module cleanup failed: " + cleanupError.Message);
                }
            }

            if (executionError != null)
            {
                var hresult = executionError.HResult;
                var friendlyError = BuildFriendlyError(executionError, stage);
                Logger.Instance.Error("VBAExecutor",
                    $"Execution failed: stage={stage}, entry={entryPoint}, hresult=0x{hresult:X8}", executionError);

                // No automatic rollback: the dispatcher backed the workbook up
                // before this call (fail-closed) and reports that snapshot id
                // as backup_snapshot_id, so the user or the model can restore
                // it explicitly.
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = friendlyError,
                    Suggestion = BuildFailureSuggestion(executionError, stage),
                    Data = new
                    {
                        rolledBack = false,
                        stage,
                        hresult = $"0x{hresult:X8}"
                    }
                };
            }

            return new ToolResult
            {
                Name = "execute_vba",
                Success = false,
                Error = "VBA 执行未返回结果"
            };
        }

        private object RunMacroWithBusyRetry(string qualifiedMacroName)
        {
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    return _app.Run(qualifiedMacroName);
                }
                catch (COMException ex) when (ex.HResult == ExcelBusyHResult && attempt < maxAttempts)
                {
                    Logger.Instance.Warning("VBAExecutor", $"Excel busy during Application.Run; retry {attempt}/{maxAttempts}");
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(150 * attempt);
                }
            }
            return null;
        }

        internal static string BuildInvocationWrapper(string entryPoint, string wrapperName)
        {
            return
                $"Public Function {wrapperName}() As String\r\n" +
                "    On Error GoTo DeepExcel_Error\r\n" +
                $"    Call {entryPoint}\r\n" +
                $"    {wrapperName} = \"{RunOkMarker}\"\r\n" +
                "    Exit Function\r\n" +
                "DeepExcel_Error:\r\n" +
                $"    {wrapperName} = \"{RunErrorMarker}\" & Chr(30) & CStr(Err.Number) & Chr(30) & Err.Source & Chr(30) & Err.Description & Chr(30) & CStr(Erl)\r\n" +
                "    Err.Clear\r\n" +
                "End Function";
        }

        private static void ThrowIfVbaReportedError(object runResult)
        {
            var result = runResult?.ToString() ?? "";
            if (result == RunOkMarker) return;
            if (!result.StartsWith(RunErrorMarker + RunResultSeparator, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("VBA 包装器未返回有效执行状态。");
            }

            var parts = result.Split(RunResultSeparator);
            int number = 0;
            int line = 0;
            if (parts.Length > 1) int.TryParse(parts[1], out number);
            var source = parts.Length > 2 ? parts[2] : "";
            var description = parts.Length > 3 ? parts[3] : "未知 VBA 运行时错误";
            if (parts.Length > 4) int.TryParse(parts[4], out line);
            throw new VbaRuntimeException(number, source, description, line);
        }

        /// <summary>
        /// Normalize model output and select the actual parameterless entry Sub.
        /// </summary>
        internal static bool TryPrepareCode(string code, string requestedMacroName,
            out string preparedCode, out string entryPoint, out string error)
        {
            preparedCode = null;
            entryPoint = null;
            error = null;

            if (string.IsNullOrWhiteSpace(code))
            {
                error = "VBA 代码为空";
                return false;
            }

            code = StripCodeFence(code).Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
            var matches = ProcedureRegex.Matches(code);
            if (matches.Count == 0)
            {
                entryPoint = DefaultMacroName;
                preparedCode = $"Public Sub {entryPoint}()\r\n{code}\r\nEnd Sub";
            }
            else
            {
                Match selected = null;
                if (!string.IsNullOrWhiteSpace(requestedMacroName))
                {
                    foreach (Match match in matches)
                    {
                        if (string.Equals(match.Groups[1].Value, requestedMacroName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            selected = match;
                            break;
                        }
                    }
                    if (selected == null)
                    {
                        error = $"找不到指定的 VBA 入口过程：{requestedMacroName}";
                        return false;
                    }
                }
                else
                {
                    foreach (Match match in matches)
                    {
                        if (string.Equals(match.Groups[1].Value, DefaultMacroName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            selected = match;
                            break;
                        }
                    }
                    if (selected == null)
                    {
                        foreach (Match match in matches)
                        {
                            if (string.IsNullOrWhiteSpace(match.Groups[2].Value))
                            {
                                selected = match;
                                break;
                            }
                        }
                    }
                }

                if (selected == null || !string.IsNullOrWhiteSpace(selected.Groups[2].Value))
                {
                    error = "VBA 入口过程必须是一个不带参数的 Sub";
                    return false;
                }

                entryPoint = selected.Groups[1].Value;
                if (!Regex.IsMatch(entryPoint, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                {
                    var nameGroup = selected.Groups[1];
                    preparedCode = code.Remove(nameGroup.Index, nameGroup.Length)
                        .Insert(nameGroup.Index, DefaultMacroName);
                    entryPoint = DefaultMacroName;
                }
                else
                {
                    preparedCode = code;
                }
            }

            var invalidToken = FindFirstNonAsciiCodeToken(preparedCode);
            if (invalidToken != null)
            {
                error = $"VBA 标识符包含非英文字符：{invalidToken}。中文可用于字符串和注释，但过程名、变量名必须使用英文。";
                return false;
            }

            var executableText = GetCodeOutsideStringsAndComments(preparedCode);
            if (Regex.IsMatch(executableText,
                @"\b(?:MsgBox|InputBox|Application\s*\.\s*InputBox)\b",
                RegexOptions.IgnoreCase))
            {
                error = "VBA 代码包含交互式弹窗（MsgBox/InputBox），会阻塞 Excel。请改为直接操作单元格并通过工具结果返回状态。";
                return false;
            }
            if (Regex.IsMatch(executableText, @"^\s*End\s*$",
                RegexOptions.IgnoreCase | RegexOptions.Multiline))
            {
                error = "VBA 代码包含独立的 End 语句，可能强制终止 Excel 的 VBA 运行环境，已拒绝执行。";
                return false;
            }

            return true;
        }

        private static string StripCodeFence(string code)
        {
            var match = Regex.Match(code,
                @"^\s*```(?:vba|vb|visual\s*basic)?\s*\r?\n([\s\S]*?)\r?\n```\s*$",
                RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : code;
        }

        private static string FindFirstNonAsciiCodeToken(string code)
        {
            var visibleCode = GetCodeOutsideStringsAndComments(code);
            var match = Regex.Match(visibleCode, @"[^\x00-\x7F]+");
            return match.Success ? match.Value : null;
        }

        private static string GetCodeOutsideStringsAndComments(string code)
        {
            var result = new StringBuilder(code.Length);
            foreach (var rawLine in code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                bool inString = false;
                for (int i = 0; i < rawLine.Length; i++)
                {
                    char ch = rawLine[i];
                    if (ch == '"')
                    {
                        if (inString && i + 1 < rawLine.Length && rawLine[i + 1] == '"')
                        {
                            i++;
                            result.Append("  ");
                            continue;
                        }
                        inString = !inString;
                        result.Append(' ');
                        continue;
                    }
                    if (!inString && ch == '\'') break;
                    result.Append(inString ? ' ' : ch);
                }
                result.AppendLine();
            }
            return result.ToString();
        }

        private sealed class VbaRuntimeException : Exception
        {
            public int VbaNumber { get; }
            public string VbaSource { get; }
            public int VbaLine { get; }

            public VbaRuntimeException(int number, string source, string description, int line)
                : base(description)
            {
                VbaNumber = number;
                VbaSource = source;
                VbaLine = line;
            }
        }

        private static string BuildFriendlyError(Exception ex, string stage)
        {
            if (ex is VbaRuntimeException vba)
            {
                var lineText = vba.VbaLine > 0 ? $"（第 {vba.VbaLine} 行）" : "";
                return $"VBA 运行时错误 {vba.VbaNumber}{lineText}：{vba.Message}";
            }
            var message = ex.Message ?? "未知错误";
            if (message.IndexOf("Visual Basic Project", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.Contains("不信任"))
            {
                return "Excel 未授权 DeepExcel 访问 VBA 工程对象模型。";
            }
            if (ex.HResult == ExcelBusyHResult)
            {
                return "Excel 当前正忙，暂时无法执行 VBA。";
            }
            if (message.Contains("无法运行") && message.Contains("宏"))
            {
                return "VBA 入口已注入，但 Excel 无法运行它；代码可能存在编译错误，或当前工作簿策略阻止宏执行。";
            }
            return $"VBA 在“{stage}”阶段失败：{message}";
        }

        private static string BuildFailureSuggestion(Exception ex, string stage)
        {
            if (ex is VbaRuntimeException)
            {
                return "已捕获 VBA 原始错误且保留执行前快照；请让 AI 根据错误号和描述修正代码。如工作簿被部分修改，可从“历史版本”手动回滚。";
            }
            var message = ex.Message ?? "";
            if (message.IndexOf("Visual Basic Project", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.Contains("不信任"))
            {
                return "在 Excel 的“文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置”中，仅勾选“信任对 VBA 工程对象模型的访问”，然后重启 Excel；不建议启用所有宏。";
            }
            if (ex.HResult == ExcelBusyHResult)
            {
                return "先按 Esc 结束单元格编辑并关闭 Excel 弹窗，再重试。";
            }
            if (stage == "macro execution")
            {
                return "已保留执行前快照；如工作簿被部分修改，可从“历史版本”手动回滚。请让 AI 根据原始 VBA 错误修正代码后重试。";
            }
            return "VBA 尚未开始执行，工作簿内容未被修改。";
        }

        /// <summary>
        /// ★ VBA Unicode 编码转换：将 VBA 代码字符串字面量中的非 ASCII 字符转为 ChrW() 调用。
        /// 解决系统 ANSI 代码页不支持中文（如英文 Windows cp1252）导致 AddFromString 中文变 "?" 的问题。
        /// 转换后 VBA 代码全为 ASCII，VBA 解析器不会出错；ChrW() 在运行时返回正确 Unicode 字符。
        /// 例如："销售数据" → ChrW(38144) & ChrW(21806) & ChrW(25968) & ChrW(25454)
        /// </summary>
        internal static string EncodeVbaUnicode(string code)
        {
            if (string.IsNullOrEmpty(code)) return code;

            var result = new StringBuilder(code.Length * 2);
            int i = 0;
            bool inString = false;
            var stringBuf = new StringBuilder();

            while (i < code.Length)
            {
                char c = code[i];

                if (!inString)
                {
                    if (c == '"')
                    {
                        inString = true;
                        stringBuf.Clear();
                        i++;
                    }
                    else
                    {
                        result.Append(c);
                        i++;
                    }
                }
                else
                {
                    // 在字符串内
                    if (c == '"')
                    {
                        // 检查是否是转义的 ""
                        if (i + 1 < code.Length && code[i + 1] == '"')
                        {
                            stringBuf.Append('"');
                            i += 2;
                        }
                        else
                        {
                            // 字符串结束，处理收集到的内容
                            inString = false;
                            i++;
                            result.Append(EncodeStringToChrW(stringBuf.ToString()));
                        }
                    }
                    else
                    {
                        stringBuf.Append(c);
                        i++;
                    }
                }
            }

            // 异常情况：代码以未闭合的字符串结尾
            if (inString && stringBuf.Length > 0)
            {
                result.Append(EncodeStringToChrW(stringBuf.ToString()));
            }

            return result.ToString();
        }

        /// <summary>
        /// 将 VBA 字符串内容转换为 ChrW() 调用表达式。
        /// 全 ASCII 的字符串保持原样（"hello"）。
        /// 包含非 ASCII 的字符串拆分为 ChrW() & "ascii" 形式。
        /// </summary>
        private static string EncodeStringToChrW(string content)
        {
            // 检查是否有非 ASCII 字符
            bool hasNonAscii = false;
            foreach (char ch in content)
            {
                if (ch > 127) { hasNonAscii = true; break; }
            }

            if (!hasNonAscii)
            {
                // 全 ASCII，保持原样
                return "\"" + content.Replace("\"", "\"\"") + "\"";
            }

            // 包含非 ASCII，转换为 ChrW() 调用
            var parts = new List<string>();
            var asciiBuf = new StringBuilder();

            foreach (char ch in content)
            {
                if (ch <= 127)
                {
                    asciiBuf.Append(ch);
                }
                else
                {
                    if (asciiBuf.Length > 0)
                    {
                        parts.Add("\"" + asciiBuf.ToString().Replace("\"", "\"\"") + "\"");
                        asciiBuf.Clear();
                    }
                    parts.Add("ChrW(" + (int)ch + ")");
                }
            }
            if (asciiBuf.Length > 0)
            {
                parts.Add("\"" + asciiBuf.ToString().Replace("\"", "\"\"") + "\"");
            }

            if (parts.Count == 0)
                return "\"\"";

            return string.Join(" & ", parts);
        }
    }
}
