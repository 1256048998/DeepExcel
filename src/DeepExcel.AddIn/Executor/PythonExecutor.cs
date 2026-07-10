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
    /// Python执行引擎
    /// 支持执行Python脚本，与Excel交互（通过临时csv文件传递数据）
    /// 要求用户系统已安装Python和openpyxl/pandas
    /// </summary>
    public class PythonExecutor
    {
        private readonly Application _excelApp;
        private readonly SnapshotManager _snapshots;
        private string _pythonPath;

        public bool PythonAvailable => !string.IsNullOrEmpty(_pythonPath);
        public string PythonPath => _pythonPath;

        public PythonExecutor(Application excelApp, SnapshotManager snapshots)
        {
            _excelApp = excelApp;
            _snapshots = snapshots;
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

            // 执行前快照
            var snapshotId = _snapshots.CreateSnapshot();
            var tempScript = Path.GetTempFileName() + ".py";
            var tempInput = Path.GetTempFileName() + ".json";
            var tempOutput = Path.GetTempFileName() + ".json";

            try
            {
                // 生成上下文脚本
                var fullScript = BuildScriptWithContext(pythonCode, tempInput, tempOutput, context);
                File.WriteAllText(tempScript, fullScript);

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

                using var process = Process.Start(psi);
                // ★ 异步读取 stdout/stderr，避免管道满导致子进程阻塞
                var outputTask = process.StandardOutput.ReadToEndAsync();
                var errorTask = process.StandardError.ReadToEndAsync();

                // ★ 30 秒超时：子进程死循环时强制 kill，防止 Excel 主线程被冻结
                const int timeoutMs = 30000;
                if (!process.WaitForExit(timeoutMs))
                {
                    try
                    {
                        Logger.Instance.Error("PythonExecutor",
                            $"Python script timed out after {timeoutMs / 1000}s, killing process (pid={process.Id})");
                        process.Kill();
                        process.WaitForExit(2000);
                    }
                    catch (Exception killEx)
                    {
                        Logger.Instance.Error("PythonExecutor", "Kill failed: " + killEx.Message);
                    }
                    _snapshots.Rollback(snapshotId);
                    return new ToolResult
                    {
                        Name = "execute_python",
                        Success = false,
                        Error = $"Python 脚本执行超时（{timeoutMs / 1000} 秒）。请简化代码或减少数据处理量。",
                        Data = new { snapshotId }
                    };
                }

                var output = outputTask.Result;
                var error = errorTask.Result;

                if (process.ExitCode == 0)
                {
                    var resultData = ReadOutput(tempOutput);
                    return new ToolResult
                    {
                        Name = "execute_python",
                        Success = true,
                        Data = new { output, result = resultData, snapshotId }
                    };
                }
                else
                {
                    // 失败回滚
                    _snapshots.Rollback(snapshotId);
                    return new ToolResult
                    {
                        Name = "execute_python",
                        Success = false,
                        Error = error,
                        Data = new { output, snapshotId }
                    };
                }
            }
            catch (Exception ex)
            {
                _snapshots.Rollback(snapshotId);
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
            }
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

        private string BuildScriptWithContext(
            string userCode,
            string inputPath,
            string outputPath,
            Dictionary<string, object> context)
        {
            var sb = new StringBuilder();

            sb.AppendLine("import json");
            sb.AppendLine("import sys");
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
            sb.AppendLine(userCode);
            sb.AppendLine("# ========== 用户代码结束 ==========");
            sb.AppendLine();

            sb.AppendLine("# 输出结果");
            sb.AppendLine("with open(r'" + outputPath.Replace("\\", "\\\\") + "', 'w', encoding='utf-8') as f:");
            sb.AppendLine("    json.dump(result, f, ensure_ascii=False, default=str)");

            return sb.ToString();
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
