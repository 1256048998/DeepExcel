using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// Python执行引擎：只做纯计算。
    ///
    /// CodeSandbox 禁掉了 openpyxl/pandas/win32com/xlwings 等，脚本碰不到工作簿，
    /// 所以这里不建快照、失败也不回滚。以前脚本一出语法错误就触发回滚，
    /// 而旧回滚会关掉工作簿、覆盖磁盘原文件——对一个根本没改过工作簿的操作纯属破坏。
    /// </summary>
    public class PythonExecutor
    {
        private readonly Application _excelApp;
        private string _pythonPath;

        public bool PythonAvailable => !string.IsNullOrEmpty(_pythonPath);
        public string PythonPath => _pythonPath;

        public PythonExecutor(Application excelApp)
        {
            _excelApp = excelApp;
            _pythonPath = FindPythonPath();
        }

        /// <summary>
        /// 执行Python脚本
        /// </summary>
        public ToolResult Execute(string pythonCode, Dictionary<string, object> context = null)
        {
            if (!PythonAvailable)
            {
                return new ToolResult
                {
                    Name = "execute_python",
                    Success = false,
                    Error = "未找到Python环境。请先安装Python并将其添加到PATH。"
                };
            }

            if (string.IsNullOrWhiteSpace(pythonCode))
            {
                return new ToolResult
                {
                    Name = "execute_python",
                    Success = false,
                    Error = "Python代码为空"
                };
            }

            // ★ P0-3 沙箱校验：阻止 LLM 执行任意系统命令/外泄数据
            var sandboxError = DeepExcel.AddIn.Security.CodeSandbox.ValidatePython(pythonCode);
            if (sandboxError != null)
            {
                Logger.Instance.Warning("PythonExecutor", "Code blocked by sandbox: " + sandboxError);
                return new ToolResult
                {
                    Name = "execute_python",
                    Success = false,
                    Error = sandboxError
                };
            }

            var tempScript = Path.GetTempFileName() + ".py";
            var tempInput = Path.GetTempFileName() + ".json";
            var tempOutput = Path.GetTempFileName() + ".json";
            var tempTrace = Path.GetTempFileName() + ".trace";

            try
            {
                // 生成上下文脚本
                var script = BuildScriptWithContext(pythonCode, tempInput, tempOutput, context, tempTrace);
                File.WriteAllText(tempScript, script.Text);

                // 写入输入上下文
                if (context != null)
                {
                    var inputJson = System.Text.Json.JsonSerializer.Serialize(context);
                    File.WriteAllText(tempInput, inputJson);
                }

                // 执行脚本（带超时，防止子进程死循环阻塞 Excel 主线程）
                var psi = new ProcessStartInfo
                {
                    FileName = _pythonPath,
                    Arguments = $"\"{tempScript}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                // ★ 强制 UTF-8 模式：系统 ANSI 代码页 cp1252 不支持中文，
                // python print() 中文会 UnicodeEncodeError 崩溃（project_memory 已记录同类问题）
                psi.EnvironmentVariables["PYTHONUTF8"] = "1";
                psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

                // ★ Job Object 管整棵进程树：超时或插件退出时孙进程一起结束
                using var job = ProcessJob.Create(MemoryLimitBytes);
                using var process = Process.Start(psi);
                job.Assign(process);
                // ★ 异步读取 stdout/stderr，避免管道满导致子进程阻塞；各自只保留前 MaxOutputChars 个字符
                var outputTask = BoundedOutput.ReadAsync(process.StandardOutput, MaxOutputChars);
                var errorTask = BoundedOutput.ReadAsync(process.StandardError, MaxOutputChars);

                // ★ 超时：子进程死循环时结束整棵树，防止 Excel 主线程被冻结
                if (!process.WaitForExit(TimeoutMs))
                {
                    Logger.Instance.Error("PythonExecutor",
                        $"Python script timed out after {TimeoutMs / 1000}s, terminating job (pid={process.Id})");
                    try
                    {
                        job.Terminate();
                        if (!job.Active) process.Kill();
                        process.WaitForExit(2000);
                    }
                    catch (Exception killEx)
                    {
                        Logger.Instance.Error("PythonExecutor", "Kill failed: " + killEx.Message);
                    }
                    var stuck = StuckLine(SafeReadAllText(tempTrace), tempScript, script);
                    return new ToolResult
                    {
                        Name = "execute_python",
                        Success = false,
                        Error = $"Python 脚本执行超时（{TimeoutMs / 1000} 秒），已结束。" +
                                (stuck ?? "没能定位到卡在哪一行。"),
                        Suggestion = "检查那一行所在的循环能否结束；数据量大时先在少量数据上试，或改用 Excel 公式 / 工具直接处理",
                    };
                }
                process.WaitForExit();  // 等异步读取把管道读完

                var output = outputTask.Result;
                var error = errorTask.Result;

                if (process.ExitCode == 0)
                {
                    var resultData = ReadOutput(tempOutput);
                    return new ToolResult
                    {
                        Name = "execute_python",
                        Success = true,
                        Data = new { output = output.Describe(), result = resultData }
                    };
                }
                else
                {
                    var message = MapScriptLines(error.Describe(), tempScript, script);
                    if (message.Contains("MemoryError"))
                    {
                        message += $"\n（脚本占用内存超过了 {MemoryLimitBytes / (1024 * 1024 * 1024)} GB 上限）";
                    }
                    return new ToolResult
                    {
                        Name = "execute_python",
                        Success = false,
                        Error = message,
                        Data = new { output = output.Describe() }
                    };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "execute_python",
                    Success = false,
                    Error = ex.Message
                };
            }
            finally
            {
                // 清理临时文件
                SafeDelete(tempScript);
                SafeDelete(tempInput);
                SafeDelete(tempOutput);
                SafeDelete(tempTrace);
            }
        }

        /// <summary>超时（测试里调短）</summary>
        internal int TimeoutMs { get; set; } = 30000;
        internal const int MaxOutputChars = 64 * 1024;
        internal const long MemoryLimitBytes = 2L * 1024 * 1024 * 1024;

        /// <summary>生成的脚本 + 用户代码在其中的位置（报错行号换算回用户代码的行号）</summary>
        internal sealed class GeneratedScript
        {
            public string Text;
            /// <summary>用户代码第 1 行在脚本里是第 FirstUserLine 行</summary>
            public int FirstUserLine;
            public int UserLineCount;

            public int? ToUserLine(int scriptLine)
            {
                var line = scriptLine - FirstUserLine + 1;
                return line >= 1 && line <= UserLineCount ? line : (int?)null;
            }
        }

        private static readonly System.Text.RegularExpressions.Regex FrameLine = new System.Text.RegularExpressions.Regex(
            @"File ""(?<file>[^""]+)"", line (?<line>\d+)");

        /// <summary>
        /// 超时前 faulthandler 写下的调用栈（最内层在前）里，找脚本本身最内层的那一帧，
        /// 换算成用户代码的行号并带上那一行的内容。
        /// </summary>
        internal static string StuckLine(string trace, string scriptPath, GeneratedScript script)
        {
            if (string.IsNullOrEmpty(trace)) return null;
            foreach (System.Text.RegularExpressions.Match match in FrameLine.Matches(trace))
            {
                if (!SamePath(match.Groups["file"].Value, scriptPath)) continue;
                var userLine = script.ToUserLine(int.Parse(match.Groups["line"].Value));
                if (userLine == null) continue;
                var code = UserCodeLine(script, userLine.Value);
                return $"超时时正在执行你代码的第 {userLine} 行：{code}";
            }
            return null;
        }

        /// <summary>Traceback 里「File "临时脚本", line N」换成用户代码的行号，模型对着自己的代码就能改</summary>
        internal static string MapScriptLines(string text, string scriptPath, GeneratedScript script)
        {
            if (string.IsNullOrEmpty(text)) return text;
            return FrameLine.Replace(text, match =>
            {
                if (!SamePath(match.Groups["file"].Value, scriptPath)) return match.Value;
                var userLine = script.ToUserLine(int.Parse(match.Groups["line"].Value));
                return userLine == null ? "File \"<DeepExcel 引导代码>\"" : $"你的代码第 {userLine} 行";
            });
        }

        private static string UserCodeLine(GeneratedScript script, int userLine)
        {
            var lines = script.Text.Replace("\r\n", "\n").Split('\n');
            var index = script.FirstUserLine - 1 + userLine - 1;
            return index >= 0 && index < lines.Length ? lines[index].Trim() : "";
        }

        private static bool SamePath(string a, string b)
        {
            try { return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase); }
            catch { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        }

        private static string SafeReadAllText(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                // faulthandler 可能还开着这个文件：共享读
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch { return null; }
        }

        /// <summary>
        /// 检查Python是否可用，附带检查必要库
        /// </summary>
        public string CheckPythonEnvironment()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Python路径: {_pythonPath ?? "未找到"}");

            if (PythonAvailable)
            {
                var version = RunPythonCommand("--version");
                sb.AppendLine($"版本: {version.Trim()}");

                var libs = new[] { "openpyxl", "pandas", "numpy" };
                foreach (var lib in libs)
                {
                    var exists = CheckLibrary(lib);
                    sb.AppendLine($"{lib}: {(exists ? "✓" : "✗")}");
                }
            }

            return sb.ToString();
        }

        private string FindPythonPath()
        {
            // 按优先级查找Python
            var candidates = new List<string>
            {
                "python",
                "python3",
                @"C:\Python311\python.exe",
                @"C:\Python310\python.exe",
                @"C:\Python39\python.exe",
                @"C:\Python38\python.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python311", "python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python310", "python.exe"),
            };

            // 查找Anaconda/Miniconda
            var condaPaths = new[]
            {
                @"C:\Users\" + Environment.UserName + @"\anaconda3\python.exe",
                @"C:\Users\" + Environment.UserName + @"\miniconda3\python.exe",
                @"C:\ProgramData\Anaconda3\python.exe",
                @"C:\ProgramData\Miniconda3\python.exe"
            };
            candidates.AddRange(condaPaths);

            // 从PATH查找
            var pathVar = Environment.GetEnvironmentVariable("PATH");
            if (pathVar != null)
            {
                foreach (var dir in pathVar.Split(';'))
                {
                    var pyExe = Path.Combine(dir, "python.exe");
                    if (File.Exists(pyExe) && !candidates.Contains(pyExe))
                    {
                        candidates.Add(pyExe);
                    }
                }
            }

            foreach (var candidate in candidates)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = candidate,
                        Arguments = "--version",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    using var process = Process.Start(psi);
                    if (process != null)
                    {
                        process.WaitForExit(2000);
                        if (process.ExitCode == 0)
                        {
                            return candidate;
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return null;
        }

        internal GeneratedScript BuildScriptWithContext(
            string userCode,
            string inputPath,
            string outputPath,
            Dictionary<string, object> context,
            string tracePath = null)
        {
            var sb = new StringBuilder();

            sb.AppendLine("import json");
            sb.AppendLine("import sys");
            if (tracePath != null)
            {
                // 超时前一点点把所有线程的调用栈写进文件：被结束之后还能说出卡在哪一行
                var dumpAfter = Math.Max(0.5, (TimeoutMs - 1500) / 1000.0).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
                sb.AppendLine("import faulthandler as _deepexcel_fh");
                sb.AppendLine($"_deepexcel_trace = open(r'{tracePath}', 'w', encoding='utf-8')");
                sb.AppendLine($"_deepexcel_fh.dump_traceback_later({dumpAfter}, exit=False, file=_deepexcel_trace)");
            }
            sb.AppendLine();
            sb.AppendLine("# 上下文输入");
            sb.AppendLine("ctx = {}");
            sb.AppendLine($"try:");
            sb.AppendLine($"    with open(r'{inputPath.Replace("\\", "\\\\")}', 'r', encoding='utf-8') as f:");
            sb.AppendLine($"        ctx = json.load(f)");
            sb.AppendLine("except: pass");
            sb.AppendLine();

            // ★ 把 context 注入为顶层 Python 变量，AI 代码可直接用 workbook_path / active_sheet 等
            // 之前只放进 ctx dict，AI 代码用 workbook_path 会 KeyError
            if (context != null)
            {
                sb.AppendLine("# 上下文变量（可直接使用）");
                foreach (var kvp in context)
                {
                    var pyVal = ToPythonLiteral(kvp.Value);
                    sb.AppendLine($"{kvp.Key} = {pyVal}");
                }
                sb.AppendLine();
            }

            sb.AppendLine("result = {}");
            sb.AppendLine();

            // 辅助函数：写入Excel区域（通过csv）
            sb.AppendLine("def write_range(address, values):");
            sb.AppendLine("    \"\"\"写入Excel区域: address如Sheet1!A1, values为二维数组\"\"\"");
            sb.AppendLine("    result['__write_range'] = result.get('__write_range', [])");
            sb.AppendLine("    result['__write_range'].append({'address': address, 'values': values})");
            sb.AppendLine();
            sb.AppendLine("def set_cell(address, value):");
            sb.AppendLine("    \"\"\"设置单个单元格\"\"\"");
            sb.AppendLine("    result['__write_range'] = result.get('__write_range', [])");
            sb.AppendLine("    result['__write_range'].append({'address': address, 'values': [[value]]})");
            sb.AppendLine();

            sb.AppendLine("# ========== 用户代码开始 ==========");
            var firstUserLine = CountLines(sb) + 1;
            sb.AppendLine(userCode);
            sb.AppendLine("# ========== 用户代码结束 ==========");
            sb.AppendLine();

            sb.AppendLine("# 输出结果");
            sb.AppendLine("with open(r'" + outputPath.Replace("\\", "\\\\") + "', 'w', encoding='utf-8') as f:");
            sb.AppendLine("    json.dump(result, f, ensure_ascii=False, default=str)");

            return new GeneratedScript
            {
                Text = sb.ToString(),
                FirstUserLine = firstUserLine,
                UserLineCount = userCode.Replace("\r\n", "\n").Split('\n').Length,
            };
        }

        private static int CountLines(StringBuilder sb)
        {
            var count = 0;
            for (var i = 0; i < sb.Length; i++) if (sb[i] == '\n') count++;
            return count;
        }

        private object ReadOutput(string outputPath)
        {
            try
            {
                if (!File.Exists(outputPath)) return null;
                var json = File.ReadAllText(outputPath);
                return System.Text.Json.JsonSerializer.Deserialize<object>(json);
            }
            catch
            {
                return null;
            }
        }

        private string RunPythonCommand(string args)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = _pythonPath,
                    Arguments = args,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(psi);
                return process.StandardOutput.ReadToEnd();
            }
            catch
            {
                return "";
            }
        }

        private bool CheckLibrary(string libName)
        {
            var output = RunPythonCommand($"-c \"import {libName}; print('ok')\"");
            return output.Contains("ok");
        }

        private void SafeDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        /// <summary>
        /// 把 C# 值转为 Python 字面量字符串（用于代码注入）
        /// </summary>
        private string ToPythonLiteral(object value)
        {
            if (value == null) return "None";
            if (value is bool b) return b ? "True" : "False";
            if (value is int || value is long || value is double || value is float)
                return value.ToString().Replace(",", "."); // 防本地化
            // 字符串：用双引号，转义反斜杠和引号
            var s = value.ToString() ?? "";
            s = s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "").Replace("\n", "\\n");
            return $"\"{s}\"";
        }
    }
}
