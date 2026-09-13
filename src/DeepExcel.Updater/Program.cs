using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DeepExcel.AddIn.Updates;

namespace DeepExcel.Updater
{
    /// <summary>
    /// DeepExcel.Updater.exe — installs a staged, signature-verified update.
    ///
    /// Launched by the add-in from inside the staging directory, never from the
    /// install directory: Inno Setup cannot overwrite a running executable, so
    /// an updater running in place would block the install it just started.
    /// That is also why this binary depends on nothing but GAC assemblies — it
    /// has to run as a lone copied file.
    ///
    /// It re-verifies everything the add-in already verified. The package sat on
    /// disk in a user-writable directory in between, and a check performed at
    /// download time is not a check performed at execution time.
    ///
    /// Two things it deliberately will not do: terminate Excel (the user may
    /// have unsaved work), and install while any host application is still
    /// running (the install would half-fail in a way nobody can diagnose).
    /// Both simply defer the update to the next restart.
    ///
    /// Exit codes: 0 installed, 2 usage, 3 stage rejected, 4 digest mismatch,
    /// 5 host still running, 6 installer failed, 7 unexpected error.
    /// </summary>
    public static class Program
    {
        private const int ExitInstalled = 0;
        private const int ExitUsage = 2;
        private const int ExitStageRejected = 3;
        private const int ExitDigestMismatch = 4;
        private const int ExitHostRunning = 5;
        private const int ExitInstallerFailed = 6;
        private const int ExitUnexpected = 7;

        private const int DefaultWaitSeconds = 180;
        private const int InstallerTimeoutSeconds = 900;

        /// <summary>Host applications that hold the add-in files open.</summary>
        private static readonly string[] HostProcessNames = { "EXCEL", "et", "wps" };

        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);
        private const int AttachParentProcess = -1;

        private static string _logPath;

        /// <summary>
        /// Whether a failure may open a dialog. Off unless --notify is passed.
        ///
        /// Default-off because this process can run at any moment: in the
        /// background, from a test, from a script. A modal box appearing on
        /// somebody's desktop with no preceding user action is an interruption
        /// they did not ask for and cannot place. UpdateService passes --notify
        /// only on the path where the user just clicked "restart and install"
        /// and is therefore expecting to hear how it went.
        /// </summary>
        private static bool _notify;

        [STAThread]
        public static int Main(string[] args)
        {
            // winexe, so an update that succeeds is completely invisible — which
            // is the entire user-facing promise. Borrow the parent console when
            // run from a shell for diagnosis.
            AttachConsole(AttachParentProcess);
            try { Console.OutputEncoding = Encoding.UTF8; } catch (Exception) { }

            try
            {
                return Run(args);
            }
            catch (Exception ex)
            {
                Log("未预期的错误：" + ex);
                ReportFailure("更新失败：" + ex.Message);
                return ExitUnexpected;
            }
        }

        private static int Run(string[] args)
        {
            Dictionary<string, string> options = ParseArguments(args);
            if (options == null ||
                !options.TryGetValue("stage", out string stage) || string.IsNullOrWhiteSpace(stage) ||
                !options.TryGetValue("installed-version", out string installedVersion))
            {
                Console.Error.WriteLine(
                    "用法: DeepExcel.Updater.exe --stage <目录> --installed-version <版本>\n" +
                    "                            [--channel stable] [--wait-pid <pid>]\n" +
                    "                            [--wait-seconds 180] [--relaunch <exe>]\n" +
                    "                            [--notify] [--dry-run]\n" +
                    "  --notify   失败时弹窗告知用户；仅供面板发起的更新使用");
                return ExitUsage;
            }

            _notify = options.ContainsKey("notify");

            stage = Path.GetFullPath(stage);
            _logPath = Path.Combine(stage, UpdateStage.LogFileName);
            Log("=== 开始 ===");
            Log("stage=" + stage + " installed=" + installedVersion);

            string channel = options.TryGetValue("channel", out string c) && !string.IsNullOrWhiteSpace(c)
                ? c
                : UpdateManifest.DefaultChannel;
            bool dryRun = options.ContainsKey("dry-run");

            // 1. Wait for the process that asked for this to go away.
            if (options.TryGetValue("wait-pid", out string pidText) &&
                int.TryParse(pidText, NumberStyles.None, CultureInfo.InvariantCulture, out int pid))
            {
                int waitSeconds = DefaultWaitSeconds;
                if (options.TryGetValue("wait-seconds", out string waitText))
                {
                    int.TryParse(waitText, NumberStyles.None, CultureInfo.InvariantCulture, out waitSeconds);
                }
                // A recycled pid is possible in principle; the only consequence
                // is waiting for an unrelated process and deferring the update,
                // and the host-process sweep below is the check that matters.
                if (!WaitForExit(pid, waitSeconds))
                {
                    Log("调用方进程 " + pid + " 在 " + waitSeconds + " 秒内未退出，本次放弃");
                    return ExitHostRunning;
                }
            }

            // 2. Re-authenticate the staged manifest from scratch.
            UpdateCheckResult check = UpdateStage.Load(
                stage, installedVersion, channel, EmbeddedUpdateKey.Verifier);
            if (!check.Accepted)
            {
                Log("更新清单被拒绝（" + check.Rejection + "）：" + check.Detail);
                if (check.Rejection != UpdateRejection.NotNewer)
                {
                    ReportFailure("更新校验失败：" + check.Detail);
                }
                return ExitStageRejected;
            }
            Log("清单已验签：v" + check.Release.Version + " 通道 " + check.Release.Channel);

            // 3. Re-hash the package immediately before executing it.
            string package = Path.Combine(stage, UpdateStage.PackageFileName);
            if (!UpdateStage.VerifyPackage(package, check.Release, out string digestProblem))
            {
                Log("更新包校验失败：" + digestProblem);
                ReportFailure("更新包已损坏或被替换，已中止安装。");
                return ExitDigestMismatch;
            }
            Log("更新包校验通过：" + check.Release.Sha256);

            if (dryRun)
            {
                Log("dry-run：跳过安装");
                return ExitInstalled;
            }

            // 4. Refuse while any host application still holds the files. This
            //    also covers the second Excel window the user forgot about, and
            //    it sits here rather than earlier so a bad package is reported
            //    as a bad package instead of as "please close Excel".
            string running = FindRunningHost();
            if (running != null)
            {
                Log("检测到仍在运行的宿主进程：" + running);
                ReportFailure("请先关闭 " + running + " 再安装 DeepExcel 更新。更新已保留，下次重启时会继续。");
                return ExitHostRunning;
            }

            // 5. Install.
            int installerExit = RunInstaller(package, stage);
            if (installerExit != 0)
            {
                Log("安装程序返回 " + installerExit);
                ReportFailure(string.Format(
                    CultureInfo.InvariantCulture,
                    "安装 DeepExcel v{0} 失败（代码 {1}）。详见 {2}",
                    check.Release.Version, installerExit, Path.Combine(stage, "install.log")));
                return ExitInstallerFailed;
            }
            Log("安装成功：v" + check.Release.Version);

            // 6. Put the user back where they were. Never fatal.
            if (options.TryGetValue("relaunch", out string host) && !string.IsNullOrWhiteSpace(host))
            {
                Relaunch(host);
            }

            // The staging directory is not deleted here: this executable is
            // running inside it. The next update check sees the new version as
            // current and prunes the whole thing.
            Log("=== 完成 ===");
            return ExitInstalled;
        }

        // ------------------------------------------------------------------

        private static bool WaitForExit(int pid, int seconds)
        {
            try
            {
                using (Process process = Process.GetProcessById(pid))
                {
                    Log("等待进程 " + pid + " 退出…");
                    return process.WaitForExit(Math.Max(1, seconds) * 1000);
                }
            }
            catch (ArgumentException)
            {
                // Already gone before we looked.
                return true;
            }
            catch (InvalidOperationException)
            {
                return true;
            }
        }

        private static string FindRunningHost()
        {
            foreach (string name in HostProcessNames)
            {
                try
                {
                    Process[] found = Process.GetProcessesByName(name);
                    try
                    {
                        if (found.Length > 0)
                        {
                            return name;
                        }
                    }
                    finally
                    {
                        foreach (Process process in found)
                        {
                            process.Dispose();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("枚举进程 " + name + " 失败：" + ex.Message);
                }
            }
            return null;
        }

        private static int RunInstaller(string package, string stage)
        {
            // /SUPPRESSMSGBOXES is needed as well as /VERYSILENT: WizardSilent
            // returns False under /VERYSILENT in this installer, so the
            // completion box would otherwise appear and wait for a click that
            // nobody is there to give.
            string arguments = string.Format(
                CultureInfo.InvariantCulture,
                "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=\"{0}\"",
                Path.Combine(stage, "install.log"));
            Log("运行安装程序：" + package + " " + arguments);

            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = package,
                    Arguments = arguments,
                    WorkingDirectory = stage,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                process.Start();
                if (!process.WaitForExit(InstallerTimeoutSeconds * 1000))
                {
                    Log("安装程序超时；不强制结束，交由用户处理");
                    return -1;
                }
                return process.ExitCode;
            }
        }

        private static void Relaunch(string executable)
        {
            try
            {
                if (!File.Exists(executable))
                {
                    Log("宿主程序不存在，跳过重启：" + executable);
                    return;
                }
                // No document arguments: reopening the user's workbooks is their
                // call, and getting it wrong means reopening the wrong file.
                Process.Start(new ProcessStartInfo
                {
                    FileName = executable,
                    UseShellExecute = true,
                })?.Dispose();
                Log("已重新启动 " + executable);
            }
            catch (Exception ex)
            {
                Log("重新启动失败：" + ex.Message);
            }
        }

        private static Dictionary<string, string> ParseArguments(string[] args)
        {
            var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (args == null)
            {
                return options;
            }
            for (int i = 0; i < args.Length; i++)
            {
                string argument = args[i];
                if (!argument.StartsWith("--", StringComparison.Ordinal))
                {
                    return null;
                }
                string name = argument.Substring(2);
                if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                {
                    options[name] = args[++i];
                }
                else
                {
                    options[name] = "";
                }
            }
            return options;
        }

        /// <summary>
        /// A failed update is the only thing the user needs to see.
        ///
        /// Success stays silent on purpose: the whole point of the updater is
        /// that new versions arrive without anybody having to do anything.
        /// </summary>
        private static void ReportFailure(string message)
        {
            Console.Error.WriteLine(message);
            Log(message);
            if (!_notify)
            {
                return;
            }
            try
            {
                MessageBox.Show(
                    message + "\n\nDeepExcel 仍可继续使用当前版本。",
                    "DeepExcel 更新", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception)
            {
            }
        }

        private static void Log(string message)
        {
            string line = string.Format(
                CultureInfo.InvariantCulture, "[{0:yyyy-MM-dd HH:mm:ss}] {1}", DateTime.Now, message);
            Console.WriteLine(line);
            if (_logPath == null)
            {
                return;
            }
            try
            {
                File.AppendAllText(_logPath, line + Environment.NewLine, new UTF8Encoding(false));
            }
            catch (Exception)
            {
            }
        }
    }
}
