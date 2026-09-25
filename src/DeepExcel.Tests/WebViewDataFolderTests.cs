using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DeepExcel.AddIn.Config;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// WebView2 数据目录：固定共享目录 + 清理旧的按 pid 目录。
    ///
    /// 以前每个 Excel 进程一个 WebView2_&lt;pid&gt;，从不清理（开发机 162 个、约 3.1 GB），
    /// 面板 localStorage 也跟着每次重启清零。
    /// </summary>
    public class WebViewDataFolderTests : IDisposable
    {
        private readonly string _localAppData;

        public WebViewDataFolderTests()
        {
            _localAppData = Path.Combine(Path.GetTempPath(), "de-webview-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_localAppData);
        }

        public void Dispose()
        {
            try { Directory.Delete(_localAppData, recursive: true); } catch { }
        }

        [Fact]
        public void Shared_path_is_stable_across_processes()
        {
            Assert.Equal(
                Path.Combine(_localAppData, "DeepExcel", "WebView2"),
                WebViewDataFolder.SharedPath(_localAppData));
        }

        [Fact]
        public void Only_dead_per_process_folders_are_stale()
        {
            var alive = new HashSet<int> { 200, 300 };
            var stale = WebViewDataFolder.StalePerProcessFolders(
                new[] { "WebView2_100", "WebView2_200", "WebView2_300", "WebView2_400", "WebView2", "logs", "WebView2_abc", "WebView2_12x", null },
                currentPid: 300,
                isProcessAlive: pid => alive.Contains(pid));

            Assert.Equal(new[] { "WebView2_100", "WebView2_400" }, stale.OrderBy(s => s).ToArray());
        }

        [Fact]
        public void A_liveness_probe_that_throws_keeps_the_folder()
        {
            // 查不清就当它还活着：宁可多留一个目录，也不能删掉正在用的
            var stale = WebViewDataFolder.StalePerProcessFolders(
                new[] { "WebView2_100" }, currentPid: 1, isProcessAlive: _ => throw new InvalidOperationException());
            Assert.Empty(stale);
        }

        [Fact]
        public void Cleanup_removes_dead_folders_and_never_touches_the_shared_one()
        {
            var root = WebViewDataFolder.Root(_localAppData);
            foreach (var name in new[] { "WebView2", "WebView2_100", "WebView2_200", "logs" })
            {
                Directory.CreateDirectory(Path.Combine(root, name, "EBWebView"));
                File.WriteAllText(Path.Combine(root, name, "EBWebView", "x.dat"), "x");
            }

            int removed = WebViewDataFolder.CleanupStale(_localAppData, currentPid: 999, isProcessAlive: pid => pid == 200);

            Assert.Equal(1, removed);
            Assert.False(Directory.Exists(Path.Combine(root, "WebView2_100")));
            Assert.True(Directory.Exists(Path.Combine(root, "WebView2_200")));
            Assert.True(Directory.Exists(Path.Combine(root, "WebView2")));
            Assert.True(Directory.Exists(Path.Combine(root, "logs")));
        }

        [Fact]
        public void Cleanup_without_a_root_is_a_no_op()
        {
            Assert.Equal(0, WebViewDataFolder.CleanupStale(Path.Combine(_localAppData, "missing"), 1, _ => false));
        }

        [Fact]
        public void The_current_process_is_reported_alive()
        {
            Assert.True(WebViewDataFolder.IsProcessAlive(System.Diagnostics.Process.GetCurrentProcess().Id));
            Assert.False(WebViewDataFolder.IsProcessAlive(int.MaxValue));
        }
    }
}
