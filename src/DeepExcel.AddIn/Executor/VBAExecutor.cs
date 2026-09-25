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

            // 编译前静态检查：编译错误和死循环在 Excel 里是弹窗或整个卡死，注入前就退回给模型
            var issues = DeepExcel.AddIn.Security.VbaStaticChecker.Check(StripCodeFence(vbaCode));
            var blocking = issues.FindAll(i => i.Level == DeepExcel.AddIn.Security.VbaIssueLevel.Error);
            if (blocking.Count > 0)
            {
                Logger.Instance.Info("VBAExecutor", $"Static check rejected the code: {blocking.Count} error(s)");
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = $"VBA 代码有 {blocking.Count} 处问题，没有执行：\n" +
                            DeepExcel.AddIn.Security.VbaStaticChecker.Describe(blocking),
                    Suggestion = "按上面逐条改好后重新调用 execute_vba（行号是你给的代码里的行号）",
                };
            }
            var warnings = DeepExcel.AddIn.Security.VbaStaticChecker.Describe(
                issues.FindAll(i => i.Level == DeepExcel.AddIn.Security.VbaIssueLevel.Warning));

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
            AppStateGuard appState = null;
            List<string> restoredState = null;
            DialogGuard dialogGuard = null;
            bool vbeWasVisible = false;

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

                // ★ 代码页安全：AddFromString 按系统 ANSI 代码页转码，表示不了的字（英文 Windows 上的中文）
                // 会变成 "?"。这些字符串改写成运行时拼出来的表达式（见 VbaEncoding）
                stage = "code injection";
                var encoded = VbaEncoding.Encode(procCode);
                if (encoded.TooLong)
                {
                    return new ToolResult
                    {
                        Name = "execute_vba",
                        Success = false,
                        Error = "代码里的中文字符串太长，改写后超过了 VBA 单条语句的续行上限",
                        Suggestion = "把长字符串拆成几个变量分别赋值，或直接用 write_value / write_range 写进单元格",
                    };
                }
                string encodedCode = encoded.Code;
                if (encoded.UsesHelper)
                {
                    encodedCode += "\r\n\r\n" + VbaEncoding.HelperFunction;
                }
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
                // ★ 记下全局开关（执行后把代码改了没改回来的恢复），临时关闭 ScreenUpdating 提升性能 + 防闪烁
                stage = "macro execution";
                appState = AppStateGuard.Capture(_app);
                vbeWasVisible = IsVbeVisible();
                _app.ScreenUpdating = false;
                // 执行期间弹出的对话框没人能点：后台线程按「绝不点是」的规则代点
                dialogGuard = DialogGuard.Start(DialogGuard.ProcessOf(SafeHwnd()));

                // ★ 先显式编译：让 Application.Run 去触发编译的话，编译错误弹窗点掉之后 VBE 会进入
                // 中断模式，Run 一直不返回（Excel 卡死）。「调试 → 编译」出错只弹窗、不进中断模式
                stage = "compile";
                if (!CompileProject(module, dialogGuard, out var compileError))
                {
                    return new ToolResult
                    {
                        Name = "execute_vba",
                        Success = false,
                        Error = compileError,
                        Suggestion = "代码没有执行、工作簿没有被修改。按编译错误改好后重新调用 execute_vba",
                    };
                }

                stage = "macro execution";
                var runSw = System.Diagnostics.Stopwatch.StartNew();
                object runResult;
                try
                {
                    runResult = RunMacroWithBusyRetry(qualifiedMacroName);
                }
                finally
                {
                    dialogGuard.Dispose();
                }
                ThrowIfVbaReportedError(runResult);
                runSw.Stop();
                restoredState = appState.Restore();

                // 执行后立即处理消息泵，恢复 WebView2 心跳
                System.Windows.Forms.Application.DoEvents();
                Logger.Instance.Info("VBAExecutor", $"VBA Run END: elapsed={runSw.ElapsedMilliseconds}ms");

                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = true,
                    Data = new { macro = entryPoint },
                    Warning = JoinNotes(warnings, AppStateGuard.Describe(restoredState),
                        DialogGuard.Describe(dialogGuard.Handled)),
                };
            }
            catch (Exception ex)
            {
                executionError = ex;
            }
            finally
            {
                if (appState != null && restoredState == null)
                {
                    restoredState = appState.Restore();
                }
                // 编译错误会把 VBE 窗口拉出来；弹窗点掉之后把它收回去
                if (dialogGuard != null && dialogGuard.Handled.Count > 0 && !vbeWasVisible)
                {
                    HideVbe();
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
                // 编译错误的具体原因只写在 VBE 弹窗上：弹窗内容比 COM 异常有用得多
                var dialogs = DialogGuard.Describe(dialogGuard?.Handled);
                if (dialogs != null) friendlyError += "\n" + dialogs;
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
                    Warning = AppStateGuard.Describe(restoredState),
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

        private const int CompileProjectControlId = 578;  // VBE「调试 → 编译 VBAProject」
        private const int CompileTimeoutMs = 10000;

        private static bool IsEnabled(Microsoft.Office.Core.CommandBarControl control)
        {
            try { return control.Enabled; }
            catch { return false; }
        }

        /// <summary>
        /// 编译整个工程。编译错误会弹对话框（DialogGuard 点掉）并把光标停在出错行：这时返回 false，
        /// 错误信息 = 弹窗内容 + 出错那一行的代码。编译命令拿不到（极少见）时当作通过，交给 Run。
        /// </summary>
        private bool CompileProject(VBComponent module, DialogGuard guard, out string error)
        {
            error = null;
            Microsoft.Office.Core.CommandBarControl control = null;
            try
            {
                control = _app.VBE.CommandBars.FindControl(Id: CompileProjectControlId);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("VBAExecutor", "Compile command unavailable: " + ex.Message);
            }
            // 已经编译过的工程该命令是灰的；新加了模块一定可用
            if (control == null || !control.Enabled) return true;

            var before = guard.Handled.Count;
            // Execute 只是把命令投递出去：编译在 Excel 处理消息时才发生。编译通过后该命令变灰；
            // 编译出错则弹窗（DialogGuard 记下并点掉），命令保持可用
            control.Execute();
            var dialogs = guard.Handled;
            var waited = System.Diagnostics.Stopwatch.StartNew();
            while (dialogs.Count == before && waited.ElapsedMilliseconds < CompileTimeoutMs)
            {
                System.Windows.Forms.Application.DoEvents();
                if (!IsEnabled(control)) return true;
                Thread.Sleep(20);
                dialogs = guard.Handled;
            }
            if (dialogs.Count == before)
            {
                Logger.Instance.Warning("VBAExecutor", "Compile did not finish in time; running anyway");
                return true;
            }

            var message = dialogs[dialogs.Count - 1].Text;
            var lineText = "";
            try
            {
                module.CodeModule.CodePane.GetSelection(out var startLine, out _, out _, out _);
                if (startLine > 0)
                {
                    lineText = "；出错的是这一行：" + module.CodeModule.get_Lines(startLine, 1).Trim();
                }
            }
            catch { }
            error = "VBA 编译错误，没有执行：" + (string.IsNullOrWhiteSpace(message) ? "（弹窗没有文字）" : message) + lineText;
            return false;
        }

        private static string JoinNotes(params string[] notes)
        {
            var present = Array.FindAll(notes, n => !string.IsNullOrWhiteSpace(n));
            return present.Length == 0 ? null : string.Join("\n", present);
        }

        private int SafeHwnd()
        {
            try { return _app.Hwnd; }
            catch { return 0; }
        }

        private bool IsVbeVisible()
        {
            try { return _app.VBE.MainWindow.Visible; }
            catch { return false; }
        }

        private void HideVbe()
        {
            try { _app.VBE.MainWindow.Visible = false; }
            catch (Exception ex) { Logger.Instance.Warning("VBAExecutor", "Hide VBE failed: " + ex.Message); }
        }

        private object RunMacroWithBusyRetry(string qualifiedMacroName)
        {
            // 瞬时错误表示调用被拒绝、宏没开始跑，重试不会重复执行
            return ComErrors.Retry(() => _app.Run(qualifiedMacroName), "Application.Run",
                () => System.Windows.Forms.Application.DoEvents());
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
            if (ComErrors.IsTransient(ex))
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
            if (ComErrors.IsTransient(ex))
            {
                return "先按 Esc 结束单元格编辑并关闭 Excel 弹窗，再重试。";
            }
            if (stage == "macro execution")
            {
                return "已保留执行前快照；如工作簿被部分修改，可从“历史版本”手动回滚。请让 AI 根据原始 VBA 错误修正代码后重试。";
            }
            return "VBA 尚未开始执行，工作簿内容未被修改。";
        }
    }
}
