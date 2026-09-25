using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Config
{
    /// <summary>
    /// WebView2 用户数据目录的选择与清理。
    ///
    /// 以前每个 Excel 进程新建一个 <c>WebView2_&lt;pid&gt;</c>，从不清理：开发机
    /// 2026-09-25 已积累 162 个、约 3.1 GB，每台用户机器都在无上限地吃磁盘。
    /// 而且面板的 localStorage（常用提示词、「自动加载历史」开关）跟着 pid 目录走，
    /// 每次重开 Excel 都是一个空目录——用户存的提示词重启就没了。
    ///
    /// WebView2 允许多个进程共享同一个数据目录，前提是环境参数一致；所以固定用
    /// <c>WebView2</c> 共享目录，创建失败才回退到按 pid 的目录，并在启动时清理
    /// 进程已不存在的旧 pid 目录。
    /// </summary>
    public static class WebViewDataFolder
    {
        public const string SharedFolderName = "WebView2";
        private const string PerProcessPrefix = "WebView2_";
        private static readonly Regex PerProcessPattern = new Regex(@"^WebView2_(\d+)$", RegexOptions.CultureInvariant);

        public static string Root(string localAppData) => Path.Combine(localAppData, "DeepExcel");

        public static string SharedPath(string localAppData) => Path.Combine(Root(localAppData), SharedFolderName);

        public static string PerProcessPath(string localAppData, int pid) =>
            Path.Combine(Root(localAppData), PerProcessPrefix + pid);

        /// <summary>
        /// 目录名是 WebView2_&lt;pid&gt; 且该 pid 已不在运行的旧目录。
        /// 只认这个命名模式：绝不碰共享目录或别的东西。pid 被复用、仍在运行时宁可留着。
        /// </summary>
        public static IList<string> StalePerProcessFolders(IEnumerable<string> directoryNames, int currentPid, Func<int, bool> isProcessAlive)
        {
            var stale = new List<string>();
            foreach (var name in directoryNames)
            {
                var match = PerProcessPattern.Match(name ?? "");
                if (!match.Success) continue;
                if (!int.TryParse(match.Groups[1].Value, out var pid)) continue;
                if (pid == currentPid) continue;
                bool alive;
                try { alive = isProcessAlive(pid); }
                catch { alive = true; }
                if (!alive) stale.Add(name);
            }
            return stale;
        }

        /// <summary>删除进程已退出的旧 pid 目录。失败静默（可能仍被占用），下次启动再试。</summary>
        public static int CleanupStale(string localAppData, int currentPid, Func<int, bool> isProcessAlive, Action<string> log = null)
        {
            var root = Root(localAppData);
            if (!Directory.Exists(root)) return 0;
            var names = new List<string>();
            foreach (var dir in Directory.GetDirectories(root, PerProcessPrefix + "*"))
                names.Add(Path.GetFileName(dir));

            int removed = 0;
            foreach (var name in StalePerProcessFolders(names, currentPid, isProcessAlive))
            {
                try
                {
                    Directory.Delete(Path.Combine(root, name), recursive: true);
                    removed++;
                }
                catch (Exception ex)
                {
                    log?.Invoke("WebView2 stale folder not removed: " + name + " (" + ex.GetType().Name + ")");
                }
            }
            return removed;
        }

        public static bool IsProcessAlive(int pid)
        {
            try
            {
                using (var p = System.Diagnostics.Process.GetProcessById(pid))
                    return !p.HasExited;
            }
            catch (ArgumentException)
            {
                return false; // 没有这个进程
            }
        }
    }
}
