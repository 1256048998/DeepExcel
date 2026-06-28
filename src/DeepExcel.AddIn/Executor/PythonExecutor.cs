using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Bridge;

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

                // 执行脚本
                var psi = new ProcessStartInfo
                {
                    FileName = _pythonPath,
                    Arguments = $"\"{tempScript}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(psi);
                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                process.WaitForExit();

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

            sb.AppendLine("==========");
            sb.AppendLine("# 用户代码开始");
            sb.AppendLine(userCode);
            sb.AppendLine("# 用户代码结束");
            sb.AppendLine("==========");
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
    }
}
