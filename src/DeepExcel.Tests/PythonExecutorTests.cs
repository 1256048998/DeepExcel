using System;
using System.Diagnostics;
using System.Text.Json;
using System.Threading;
using DeepExcel.AddIn.Executor;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>本机有 python 时才跑（在 PATH 里找）</summary>
    public sealed class PythonFactAttribute : FactAttribute
    {
        private static readonly Lazy<bool> Available = new Lazy<bool>(() => new PythonExecutor(null).PythonAvailable);

        public PythonFactAttribute()
        {
            if (!Available.Value) Skip = "本机没有 Python";
        }
    }

    /// <summary>Python 子进程：整棵进程树、输出上限、超时报行号、报错行号换算回用户代码。</summary>
    public class PythonExecutorTests
    {
        private static string Json(object value) => JsonSerializer.Serialize(value);

        [PythonFact]
        public void A_timeout_names_the_line_the_script_was_stuck_on()
        {
            var python = new PythonExecutor(null) { TimeoutMs = 4000 };
            var result = python.Execute("total = 0\nwhile True:\n    total += 1\n");
            Assert.False(result.Success);
            Assert.Contains("超时", result.Error);
            Assert.Matches(@"你代码的第 [23] 行", result.Error);
        }

        [PythonFact]
        public void Tracebacks_point_at_the_users_own_line_numbers()
        {
            var result = new PythonExecutor(null).Execute("a = 1\nb = a / 0\n");
            Assert.False(result.Success);
            Assert.Contains("ZeroDivisionError", result.Error);
            Assert.Contains("你的代码第 2 行", result.Error);
            Assert.DoesNotContain(".py\", line", result.Error);  // 临时脚本路径不再出现
        }

        [PythonFact]
        public void Runaway_output_is_capped()
        {
            var result = new PythonExecutor(null).Execute("for i in range(20000):\n    print('x' * 20)\n");
            Assert.True(result.Success, result.Error);
            var output = (string)result.Data.GetType().GetProperty("output").GetValue(result.Data);
            Assert.Contains("输出太长", output);
            Assert.True(output.Length < PythonExecutor.MaxOutputChars + 2000);
        }

        [PythonFact]
        public void Normal_scripts_still_return_their_result()
        {
            var result = new PythonExecutor(null).Execute("result['sum'] = sum(range(10))\nprint('done')\n");
            Assert.True(result.Success, result.Error);
            var data = Json(result.Data);
            Assert.Contains("\"sum\":45", data);
            Assert.Contains("done", data);
        }

        [PythonFact]
        public void Terminating_the_job_takes_grandchildren_with_it()
        {
            var pythonPath = new PythonExecutor(null).PythonPath;
            var psi = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = "-c \"import subprocess, sys, time; p = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(60)']); print(p.pid, flush=True); time.sleep(60)\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            using (var job = ProcessJob.Create())
            using (var parent = Process.Start(psi))
            {
                Assert.True(job.Assign(parent));
                var grandchildPid = int.Parse(parent.StandardOutput.ReadLine().Trim());
                var grandchild = Process.GetProcessById(grandchildPid);
                job.Terminate();
                Assert.True(parent.WaitForExit(5000));
                Assert.True(grandchild.WaitForExit(5000), "孙进程没有跟着结束");
            }
        }

        [PythonFact]
        public void Closing_the_job_also_ends_the_tree()
        {
            var pythonPath = new PythonExecutor(null).PythonPath;
            var psi = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = "-c \"import time; time.sleep(60)\"",
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            var job = ProcessJob.Create();
            using (var process = Process.Start(psi))
            {
                Assert.True(job.Assign(process));
                job.Dispose();
                Assert.True(process.WaitForExit(5000));
            }
        }
    }

    public class ScriptLineMappingTests
    {
        [Fact]
        public void Only_lines_inside_the_user_code_are_mapped()
        {
            var script = new PythonExecutor.GeneratedScript
            {
                Text = "import json\n# start\nx = 1\ny = 2\n# end\n",
                FirstUserLine = 3,
                UserLineCount = 2,
            };
            Assert.Equal(1, script.ToUserLine(3));
            Assert.Equal(2, script.ToUserLine(4));
            Assert.Null(script.ToUserLine(2));
            Assert.Null(script.ToUserLine(5));

            var mapped = PythonExecutor.MapScriptLines(
                "File \"C:\\t\\a.py\", line 4, in <module>\nFile \"C:\\t\\a.py\", line 1\nFile \"other.py\", line 4",
                "C:\\t\\a.py", script);
            Assert.Contains("你的代码第 2 行", mapped);
            Assert.Contains("<DeepExcel 引导代码>", mapped);
            Assert.Contains("File \"other.py\", line 4", mapped);

            var stuck = PythonExecutor.StuckLine("Thread 0x1 (most recent call first):\n  File \"C:\\t\\a.py\", line 4 in <module>\n",
                "C:\\t\\a.py", script);
            Assert.Equal("超时时正在执行你代码的第 2 行：y = 2", stuck);
        }
    }
}
