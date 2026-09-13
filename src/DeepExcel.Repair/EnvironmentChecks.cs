using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace DeepExcel.Repair
{
    /// <summary>
    /// Runtime prerequisites and the file-level state that silently blocks
    /// loading (Mark of the Web, orphaned sidecar processes).
    /// </summary>
    public static class EnvironmentChecks
    {
        private const string WebView2Client = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";
        private const int DotNet48Release = 528040;

        public static void Verify(Report report, string installDirectory)
        {
            var dllPath = Path.Combine(installDirectory, AddInIdentity.AddInDllName);
            if (!File.Exists(dllPath))
            {
                report.Add(new Finding(
                    Codes.AddInDll, Severity.Blocking,
                    "加载项主文件缺失：" + dllPath + "（安装不完整，请重新运行安装器）",
                    false));
            }

            if (!IsDotNet48Installed())
            {
                report.Add(new Finding(
                    Codes.DotNet48, Severity.Blocking,
                    ".NET Framework 4.8 未安装",
                    false));
            }

            if (!IsWebView2Installed())
            {
                report.Add(new Finding(
                    Codes.WebView2, Severity.Blocking,
                    "WebView2 运行时未安装（面板将无法显示）",
                    false));
            }

            var python = Path.Combine(installDirectory, "python", "python.exe");
            if (!File.Exists(python))
            {
                report.Add(new Finding(
                    Codes.EmbeddedPython, Severity.Blocking,
                    "内置 Python 缺失：" + python,
                    false));
            }

            var sidecar = Path.Combine(installDirectory, "sidecar", "sidecar.py");
            if (!File.Exists(sidecar))
            {
                report.Add(new Finding(
                    Codes.SidecarScript, Severity.Blocking,
                    "AI Sidecar 缺失：" + sidecar,
                    false));
            }

            var blocked = CountMarkOfTheWeb(installDirectory);
            if (blocked > 0)
            {
                report.Add(new Finding(
                    Codes.MarkOfTheWeb, Severity.Blocking,
                    string.Format("{0} 个文件带有下载来源标记，Excel 信任中心会静默拦截", blocked),
                    true));
            }

            var orphans = FindOrphanSidecars(installDirectory).Count;
            if (orphans > 0)
            {
                report.Add(new Finding(
                    Codes.OrphanSidecar, Severity.Warning,
                    string.Format("{0} 个残留 Sidecar 进程仍在运行", orphans),
                    true));
            }

            VerifyLastLoad(report);
        }

        /// <summary>
        /// Registration can be perfect while the managed constructor still
        /// throws -- Excel then hides the add-in with no visible cause. The
        /// add-in drops a breadcrumb on that path (ThisAddIn.OnConnection) so
        /// the reason survives to the next time the user runs this tool.
        /// </summary>
        private static void VerifyLastLoad(Report report)
        {
            var path = Path.Combine(AddInIdentity.LogDirectory, "last-load-failure.txt");
            if (!File.Exists(path))
            {
                return;
            }

            string detail;
            try
            {
                detail = File.ReadAllText(path).Trim().Replace("\t", " | ");
            }
            catch (Exception ex)
            {
                detail = "(无法读取: " + ex.Message + ")";
            }

            report.Add(new Finding(
                Codes.LastLoadFailed, Severity.Warning,
                "上次 Excel 启动时加载项初始化失败：" + detail,
                false));
        }

        /// <summary>
        /// Cleared after a repair so a stale breadcrumb from a fixed problem
        /// does not keep alarming the user. Excel rewrites it if it recurs.
        /// </summary>
        public static void ClearLastLoadFailure()
        {
            try
            {
                var path = Path.Combine(AddInIdentity.LogDirectory, "last-load-failure.txt");
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        public static bool IsDotNet48Installed()
        {
            using (var key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
            {
                if (key == null)
                {
                    return false;
                }
                var install = key.GetValue("Install");
                var release = key.GetValue("Release");
                return install != null && Convert.ToInt32(install) == 1
                    && release != null && Convert.ToInt32(release) >= DotNet48Release;
            }
        }

        public static bool IsWebView2Installed()
        {
            // Evergreen runtime records itself under EdgeUpdate in either hive
            // and either view depending on per-user vs per-machine install.
            var paths = new[]
            {
                @"SOFTWARE\Microsoft\EdgeUpdate\Clients\" + WebView2Client,
                @"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\" + WebView2Client
            };
            foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            {
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
                    {
                        foreach (var path in paths)
                        {
                            using (var key = baseKey.OpenSubKey(path))
                            {
                                var pv = key == null ? null : key.GetValue("pv") as string;
                                if (!string.IsNullOrEmpty(pv) && pv != "0.0.0.0")
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        // ------------------------------------------------------------------
        // Mark of the Web
        // ------------------------------------------------------------------

        private const uint GenericRead = 0x80000000;
        private const uint FileShareRead = 0x00000001;
        private const uint OpenExisting = 3;
        private static readonly IntPtr InvalidHandle = new IntPtr(-1);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateFileW(
            string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes,
            uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFileW(string lpFileName);

        /// <summary>
        /// .NET's File APIs reject the "path:stream" syntax, so the alternate
        /// data stream has to be reached through Win32 directly.
        /// </summary>
        private static bool HasZoneIdentifier(string path)
        {
            var handle = CreateFileW(path + ":Zone.Identifier", GenericRead, FileShareRead,
                IntPtr.Zero, OpenExisting, 0, IntPtr.Zero);
            if (handle == InvalidHandle)
            {
                return false;
            }
            CloseHandle(handle);
            return true;
        }

        private static IEnumerable<string> BlockableFiles(string installDirectory)
        {
            var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".dll", ".exe", ".ps1", ".config", ".py", ".html", ".js", ".css", ".json"
            };
            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(installDirectory, "*", SearchOption.AllDirectories);
            }
            catch (Exception)
            {
                yield break;
            }
            foreach (var file in files)
            {
                if (extensions.Contains(Path.GetExtension(file)))
                {
                    yield return file;
                }
            }
        }

        public static int CountMarkOfTheWeb(string installDirectory)
        {
            var count = 0;
            foreach (var file in BlockableFiles(installDirectory))
            {
                if (HasZoneIdentifier(file))
                {
                    count++;
                }
            }
            return count;
        }

        public static int RemoveMarkOfTheWeb(string installDirectory)
        {
            var removed = 0;
            foreach (var file in BlockableFiles(installDirectory))
            {
                if (HasZoneIdentifier(file) && DeleteFileW(file + ":Zone.Identifier"))
                {
                    removed++;
                }
            }
            return removed;
        }

        // ------------------------------------------------------------------
        // Orphaned sidecar processes
        // ------------------------------------------------------------------

        /// <summary>
        /// Only processes launched from our own install directory are touched.
        /// Matching on the name "python" alone would kill unrelated work.
        /// </summary>
        public static List<Process> FindOrphanSidecars(string installDirectory)
        {
            var expected = Path.Combine(installDirectory, "python", "python.exe");
            var found = new List<Process>();
            Process[] candidates;
            try
            {
                candidates = Process.GetProcessesByName("python");
            }
            catch (Exception)
            {
                return found;
            }

            foreach (var process in candidates)
            {
                try
                {
                    if (string.Equals(process.MainModule.FileName, expected, StringComparison.OrdinalIgnoreCase))
                    {
                        found.Add(process);
                        continue;
                    }
                }
                catch (Exception)
                {
                    // Access denied on another user's process: not ours, skip.
                }
                process.Dispose();
            }
            return found;
        }

        public static int KillOrphanSidecars(string installDirectory)
        {
            var killed = 0;
            foreach (var process in FindOrphanSidecars(installDirectory))
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(5000);
                    killed++;
                }
                catch (Exception)
                {
                    // Already gone or protected; nothing useful to do.
                }
                finally
                {
                    process.Dispose();
                }
            }
            return killed;
        }
    }
}
