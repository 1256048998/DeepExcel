using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace DeepExcel.Repair
{
    /// <summary>
    /// DeepExcel.Repair.exe
    ///
    /// Two audiences, one binary:
    ///   * the installer calls "--repair --quiet" instead of shelling out to
    ///     powershell.exe with register-user.ps1;
    ///   * a user whose ribbon tab is missing double-clicks it and gets a window
    ///     with one button.
    ///
    /// Exit codes: 0 healthy, 1 blocking problems remain, 2 usage error.
    /// </summary>
    public static class Program
    {
        private const int ExitHealthy = 0;
        private const int ExitBlocked = 1;
        private const int ExitUsage = 2;

        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);
        private const int AttachParentProcess = -1;

        [STAThread]
        public static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                RepairWindow.Run();
                return ExitHealthy;
            }

            // Built as winexe so double-clicking never flashes a console. When
            // invoked from a shell, borrow the parent's console for output.
            AttachConsole(AttachParentProcess);
            try
            {
                // Diagnostic text is Chinese; the console's default OEM code page
                // renders it as question marks, which makes a support report
                // useless. Failure here is not worth aborting over.
                Console.OutputEncoding = Encoding.UTF8;
            }
            catch (Exception)
            {
            }

            var quiet = false;
            var mode = (string)null;
            foreach (var arg in args)
            {
                switch (arg.ToLowerInvariant())
                {
                    case "--quiet":
                    case "-q":
                        quiet = true;
                        break;
                    case "--verify":
                    case "--repair":
                    case "--activation-only":
                    case "--bundle":
                        if (mode != null)
                        {
                            Console.Error.WriteLine("一次只能指定一种模式。");
                            return ExitUsage;
                        }
                        mode = arg.ToLowerInvariant();
                        break;
                    case "--help":
                    case "-h":
                    case "/?":
                        PrintUsage();
                        return ExitHealthy;
                    default:
                        Console.Error.WriteLine("未知参数：" + arg);
                        PrintUsage();
                        return ExitUsage;
                }
            }

            if (mode == null)
            {
                PrintUsage();
                return ExitUsage;
            }

            var output = new StringBuilder();
            Action<string> log = line =>
            {
                output.AppendLine(line);
                if (!quiet)
                {
                    Console.WriteLine(line);
                }
            };

            switch (mode)
            {
                case "--activation-only":
                {
                    var failure = ComProbe.TryActivate(AddInIdentity.MainProgId);
                    if (failure != null)
                    {
                        Console.Error.WriteLine(failure);
                        return ExitBlocked;
                    }
                    log(string.Format("{0} 位进程 COM 激活成功。", ComProbe.BitnessOfCurrentProcess));
                    return ExitHealthy;
                }

                case "--verify":
                {
                    var report = Engine.Diagnose(log);
                    return report.HasBlocking ? ExitBlocked : ExitHealthy;
                }

                case "--repair":
                {
                    var report = Engine.Repair(log);
                    return report.HasBlocking ? ExitBlocked : ExitHealthy;
                }

                case "--bundle":
                {
                    var report = Engine.Diagnose(log);
                    var path = DiagnosticBundle.Write(AddInIdentity.InstallDirectory, report);
                    Console.WriteLine("诊断包已生成：" + path);
                    return report.HasBlocking ? ExitBlocked : ExitHealthy;
                }
            }

            return ExitUsage;
        }

        private static void PrintUsage()
        {
            Console.WriteLine(@"DeepExcel 诊断与修复工具

  DeepExcel.Repair.exe                无参数运行会打开图形界面
  DeepExcel.Repair.exe --verify       只检查，不修改任何内容
  DeepExcel.Repair.exe --repair       检查并自动修复
  DeepExcel.Repair.exe --bundle       检查并导出脱敏诊断包到桌面
  DeepExcel.Repair.exe --activation-only
                                      仅测试当前位数的 COM 激活

  --quiet / -q                        不向控制台输出（退出码仍然有效）

退出码：0 正常，1 仍存在阻断性问题，2 参数错误");
        }
    }

    /// <summary>Diagnose and repair, shared by the CLI and the window.</summary>
    public static class Engine
    {
        public static Report Diagnose(Action<string> log)
        {
            var installDirectory = AddInIdentity.InstallDirectory;
            var report = new Report();

            log("安装目录：" + installDirectory);
            EnvironmentChecks.Verify(report, installDirectory);

            // Registry verification needs the DLL to read its assembly version.
            // If it is missing the environment check already reported it, and
            // every registry finding would be noise.
            var dllPath = AddInIdentity.AddInDllPath;
            if (System.IO.File.Exists(dllPath))
            {
                RegistryRepair.Verify(report, dllPath);

                var failure = ComProbe.TryActivate(AddInIdentity.MainProgId);
                if (failure != null)
                {
                    report.Add(new Finding(Codes.Activation64, Severity.Blocking, failure, true));
                }

                ProbeThirtyTwoBit(report);
            }

            log(report.Render());
            return report;
        }

        /// <summary>
        /// A process can only activate COM in its own bitness, so 32-bit Office
        /// support is verified by the dedicated x86 probe. Its absence is a
        /// packaging error, not a user-machine problem, so it is a warning.
        /// </summary>
        private static void ProbeThirtyTwoBit(Report report)
        {
            var probe = System.IO.Path.Combine(AddInIdentity.InstallDirectory, "DeepExcel.Probe32.exe");
            if (!System.IO.File.Exists(probe))
            {
                report.Add(new Finding(
                    Codes.Activation32, Severity.Warning,
                    "未找到 DeepExcel.Probe32.exe，无法验证 32 位 Excel 兼容性",
                    false));
                return;
            }

            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo(probe, "--activation-only")
                {
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    // The probe writes UTF-8; without this the pipe is decoded
                    // with the OEM code page and the failure reason -- the whole
                    // reason we ran it -- comes back as question marks.
                    StandardErrorEncoding = Encoding.UTF8,
                    StandardOutputEncoding = Encoding.UTF8
                };
                using (var process = System.Diagnostics.Process.Start(startInfo))
                {
                    var stderr = process.StandardError.ReadToEnd();
                    process.StandardOutput.ReadToEnd();
                    if (!process.WaitForExit(30000))
                    {
                        try { process.Kill(); } catch (Exception) { }
                        report.Add(new Finding(
                            Codes.Activation32, Severity.Blocking, "32 位 COM 激活探测超时", true));
                        return;
                    }
                    if (process.ExitCode != 0)
                    {
                        report.Add(new Finding(
                            Codes.Activation32, Severity.Blocking,
                            string.IsNullOrEmpty(stderr) ? "32 位 COM 激活失败" : stderr.Trim(),
                            true));
                    }
                }
            }
            catch (Exception ex)
            {
                report.Add(new Finding(
                    Codes.Activation32, Severity.Warning,
                    "无法启动 32 位探测进程：" + ex.Message, false));
            }
        }

        public static Report Repair(Action<string> log)
        {
            var installDirectory = AddInIdentity.InstallDirectory;
            var before = Diagnose(log);

            if (before.HasUnrepairableBlocking)
            {
                log("");
                log("存在无法自动修复的问题，已停止。请按上面的诊断码联系支持。");
                return before;
            }

            if (!before.HasBlocking && before.Findings.Count == 0)
            {
                return before;
            }

            log("");
            log("开始修复…");

            var dllPath = AddInIdentity.AddInDllPath;
            if (System.IO.File.Exists(dllPath))
            {
                RegistryRepair.RegisterAll(dllPath);
                log("  已重建 32/64 位 COM 注册与 Excel 加载项项");
            }

            var cleared = RegistryRepair.ClearResiliencyBlocks();
            if (cleared > 0)
            {
                log(string.Format("  已清理 {0} 条 Excel 禁用/崩溃记录", cleared));
            }

            var unblocked = EnvironmentChecks.RemoveMarkOfTheWeb(installDirectory);
            if (unblocked > 0)
            {
                log(string.Format("  已解除 {0} 个文件的下载来源标记", unblocked));
            }

            var killed = EnvironmentChecks.KillOrphanSidecars(installDirectory);
            if (killed > 0)
            {
                log(string.Format("  已结束 {0} 个残留 Sidecar 进程", killed));
            }

            // Otherwise a breadcrumb from an already-fixed failure would be
            // re-reported on every run. Excel writes a new one if it recurs.
            EnvironmentChecks.ClearLastLoadFailure();

            log("");
            log("重新检查…");
            var after = Diagnose(log);
            if (!after.HasBlocking)
            {
                log("");
                log("修复完成。请完全关闭并重新打开 Excel。");
            }
            return after;
        }
    }
}
