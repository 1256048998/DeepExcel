using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace DeepExcel.AddIn.Collaboration
{
    /// <summary>
    /// 工作簿记忆的面板入口（查看 / 修改 / 清除）。记忆本身由侧车维护（workbook_memory.py）：
    /// 这里只读写同一组文件，目录口径与侧车、WPS 端一致——
    /// %LOCALAPPDATA%\DeepExcel\workbooks\&lt;SHA256(workbookKey) 前 16 字节 hex&gt;\NOTES.md / history.jsonl / meta.json。
    /// 用户在这里保存的是用户自己的修改：可以解除禁区（模型不能）。侧车下一轮看到文件变了会重新注入。
    /// </summary>
    public static class WorkbookMemoryStore
    {
        public const int MaxNotesChars = 6000;

        /// <summary>测试用：DEEPEXCEL_MEMORY_DIR 覆盖根目录（与侧车同一个变量）</summary>
        public static string Root
        {
            get
            {
                var overridden = Environment.GetEnvironmentVariable("DEEPEXCEL_MEMORY_DIR");
                if (!string.IsNullOrEmpty(overridden)) return overridden;
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "DeepExcel", "workbooks");
            }
        }

        public static string KeyHash(string workbookKey)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(workbookKey ?? ""));
                return BitConverter.ToString(bytes, 0, 16).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>没保存过的工作簿（key 不是路径）没有记忆</summary>
        public static bool IsPersistentKey(string workbookKey) =>
            !string.IsNullOrEmpty(workbookKey) && (workbookKey.Contains("\\") || workbookKey.Contains("/"));

        public static string DirectoryFor(string workbookKey) => Path.Combine(Root, KeyHash(workbookKey));

        /// <summary>面板显示用：{ available, workbookName, notes, historyCount }</summary>
        public static object Describe(string workbookKey, string workbookName, string error = null)
        {
            if (!IsPersistentKey(workbookKey))
            {
                return new { available = false, workbookName, notes = "", historyCount = 0, error };
            }
            var dir = DirectoryFor(workbookKey);
            return new
            {
                available = true,
                workbookName,
                notes = ReadText(Path.Combine(dir, "NOTES.md")),
                historyCount = CountLines(Path.Combine(dir, "history.jsonl")),
                error,
            };
        }

        /// <summary>保存用户改过的记忆；不允许时返回原因</summary>
        public static string Save(string workbookKey, string workbookName, string notes)
        {
            if (!IsPersistentKey(workbookKey)) return "工作簿还没保存过，保存后才有记忆";
            var text = (notes ?? "").Replace("\r\n", "\n").Trim();
            if (text.Length > MaxNotesChars) return $"记忆太长（{text.Length} 字），上限 {MaxNotesChars} 字";
            var dir = DirectoryFor(workbookKey);
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "NOTES.md");
            if (text.Length == 0)
            {
                TryDelete(path);
            }
            else
            {
                var temp = path + ".tmp";
                File.WriteAllText(temp, text + "\n", new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temp, path, null);
                else File.Move(temp, path);
            }
            WriteMeta(dir, workbookKey, workbookName);
            return null;
        }

        /// <summary>清除这个工作簿的全部记忆（笔记和操作记录）</summary>
        public static void Clear(string workbookKey)
        {
            if (!IsPersistentKey(workbookKey)) return;
            var dir = DirectoryFor(workbookKey);
            foreach (var name in new[] { "NOTES.md", "history.jsonl", "meta.json" })
            {
                TryDelete(Path.Combine(dir, name));
            }
            try
            {
                if (Directory.Exists(dir) && !Directory.EnumerateFileSystemEntries(dir).Any()) Directory.Delete(dir);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static void WriteMeta(string dir, string workbookKey, string workbookName)
        {
            try
            {
                var json = System.Text.Json.JsonSerializer.Serialize(new
                {
                    workbook_key = workbookKey,
                    workbook_name = workbookName,
                    updated_at = DateTimeOffset.UtcNow.ToUnixTimeSeconds(),
                });
                File.WriteAllText(Path.Combine(dir, "meta.json"), json, new UTF8Encoding(false));
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static string ReadText(string path)
        {
            try { return File.Exists(path) ? File.ReadAllText(path, Encoding.UTF8) : ""; }
            catch (IOException) { return ""; }
            catch (UnauthorizedAccessException) { return ""; }
        }

        private static int CountLines(string path)
        {
            try { return File.Exists(path) ? File.ReadLines(path).Count(l => l.Trim().Length > 0) : 0; }
            catch (IOException) { return 0; }
            catch (UnauthorizedAccessException) { return 0; }
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
