using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 副本 Excel 进程登记簿。
    ///
    /// 插件崩溃、Excel 被强关时，拉起的隐藏 Excel 没人收：它不可见，用户在任务栏看不到，
    /// 只会一直占着内存和副本文件。登记 PID + 创建时间（PID 会被系统复用，只比 PID 可能误杀
    /// 用户后来开的 Excel），启动时和每次试跑前回收主人已经不在的那些。
    /// </summary>
    public static class LabProcessRegistry
    {
        private static readonly object Gate = new object();

        /// <summary>副本目录保留多久：超过这个时间的一定是残留（一次试跑最多 20 秒）</summary>
        private static readonly TimeSpan StaleCopyAge = TimeSpan.FromHours(1);

        public static string Root { get; set; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DeepExcel", "lab");

        private static string FilePath => Path.Combine(Root, "processes.json");

        public sealed class Entry
        {
            public int Pid { get; set; }
            public long StartedUtcTicks { get; set; }
            public int OwnerPid { get; set; }
            public long OwnerStartedUtcTicks { get; set; }
        }

        public static void Register(int pid)
        {
            if (pid <= 0) return;
            lock (Gate)
            {
                var entries = Load();
                entries.RemoveAll(e => e.Pid == pid);
                var owner = Process.GetCurrentProcess();
                entries.Add(new Entry
                {
                    Pid = pid,
                    StartedUtcTicks = StartTicks(pid),
                    OwnerPid = owner.Id,
                    OwnerStartedUtcTicks = StartTicks(owner.Id),
                });
                Save(entries);
            }
        }

        public static void Unregister(int pid)
        {
            lock (Gate)
            {
                var entries = Load();
                if (entries.RemoveAll(e => e.Pid == pid) > 0) Save(entries);
            }
        }

        /// <summary>结束主人已退出的副本进程，删掉过期的副本目录。返回结束了几个进程。</summary>
        public static int Reap()
        {
            int killed = 0;
            lock (Gate)
            {
                var entries = Load();
                var keep = new List<Entry>();
                foreach (var entry in entries)
                {
                    bool ownerAlive = IsSameProcess(entry.OwnerPid, entry.OwnerStartedUtcTicks);
                    bool labAlive = IsSameProcess(entry.Pid, entry.StartedUtcTicks);
                    if (!labAlive) continue;
                    if (ownerAlive) { keep.Add(entry); continue; }
                    try
                    {
                        Process.GetProcessById(entry.Pid).Kill();
                        killed++;
                        Logger.Instance.Warning("LabProcessRegistry", $"Reaped orphan lab Excel pid={entry.Pid}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Instance.Warning("LabProcessRegistry", $"Reap pid={entry.Pid} failed: {ex.Message}");
                        keep.Add(entry);
                    }
                }
                if (keep.Count != entries.Count) Save(keep);
            }
            DeleteStaleCopies();
            return killed;
        }

        /// <summary>给一次试跑建一个独立目录：副本用原文件名，VBA 里读 ThisWorkbook.Name 的结果才一致</summary>
        public static string NewCopyDirectory()
        {
            var dir = Path.Combine(Root, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static void DeleteCopyDirectory(string dir)
        {
            if (string.IsNullOrEmpty(dir)) return;
            // 只删登记目录下的子目录，路径算错也不会删到别处
            var full = Path.GetFullPath(dir).TrimEnd('\\');
            var root = Path.GetFullPath(Root).TrimEnd('\\') + "\\";
            if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return;
            try { Directory.Delete(full, true); }
            catch (Exception ex) { Logger.Instance.Warning("LabProcessRegistry", "Delete copy failed: " + ex.Message); }
        }

        private static void DeleteStaleCopies()
        {
            try
            {
                if (!Directory.Exists(Root)) return;
                foreach (var dir in Directory.GetDirectories(Root))
                {
                    if (DateTime.UtcNow - Directory.GetCreationTimeUtc(dir) > StaleCopyAge)
                    {
                        DeleteCopyDirectory(dir);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("LabProcessRegistry", "Stale copy sweep failed: " + ex.Message);
            }
        }

        internal static bool IsSameProcess(int pid, long startedUtcTicks)
        {
            if (pid <= 0) return false;
            var actual = StartTicks(pid);
            if (actual == 0) return false;
            // 创建时间取自同一个内核值，1 秒容差只为吸收序列化精度
            return Math.Abs(actual - startedUtcTicks) < TimeSpan.TicksPerSecond;
        }

        private static long StartTicks(int pid)
        {
            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    return p.HasExited ? 0 : p.StartTime.ToUniversalTime().Ticks;
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static List<Entry> Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new List<Entry>();
                return JsonSerializer.Deserialize<List<Entry>>(File.ReadAllText(FilePath)) ?? new List<Entry>();
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("LabProcessRegistry", "Registry unreadable, starting over: " + ex.Message);
                return new List<Entry>();
            }
        }

        private static void Save(List<Entry> entries)
        {
            try
            {
                Directory.CreateDirectory(Root);
                var tmp = FilePath + ".tmp";
                File.WriteAllText(tmp, JsonSerializer.Serialize(entries.ToList()));
                if (File.Exists(FilePath)) File.Delete(FilePath);
                File.Move(tmp, FilePath);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("LabProcessRegistry", "Registry write failed: " + ex.Message);
            }
        }
    }
}
