using System;
using System.IO;
using System.Text.Json;
using DeepExcel.AddIn.Collaboration;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 工作簿记忆的面板入口。目录口径必须与侧车（workbook_memory.py）和 WPS 端一致，
    /// 否则同一个文件在两个宿主里各记各的；期望值由侧车 key_hash 算出。
    /// </summary>
    [Collection("MemoryDirEnv")]
    public class WorkbookMemoryStoreTests : IDisposable
    {
        private const string Key = @"C:\财务\2026 预算.xlsx";
        private readonly string _root;
        private readonly string _previous;

        public WorkbookMemoryStoreTests()
        {
            _root = Path.Combine(Path.GetTempPath(), "DeepExcelMemoryTests", Guid.NewGuid().ToString("N"));
            _previous = Environment.GetEnvironmentVariable("DEEPEXCEL_MEMORY_DIR");
            Environment.SetEnvironmentVariable("DEEPEXCEL_MEMORY_DIR", _root);
        }

        public void Dispose()
        {
            Environment.SetEnvironmentVariable("DEEPEXCEL_MEMORY_DIR", _previous);
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        private static JsonElement Describe(string key, string error = null) =>
            JsonSerializer.SerializeToElement(WorkbookMemoryStore.Describe(key, "2026 预算.xlsx", error));

        [Fact]
        public void The_directory_matches_the_sidecar_and_wps()
        {
            Assert.Equal("745f66b0b53be193d2368fd7504c8f6a", WorkbookMemoryStore.KeyHash(Key));
            Assert.Equal(Path.Combine(_root, "745f66b0b53be193d2368fd7504c8f6a"), WorkbookMemoryStore.DirectoryFor(Key));
        }

        [Fact]
        public void Unsaved_workbooks_have_no_memory()
        {
            Assert.False(Describe("工作簿1").GetProperty("available").GetBoolean());
            Assert.NotNull(WorkbookMemoryStore.Save("工作簿1", "工作簿1", "x"));
        }

        [Fact]
        public void Save_describe_and_clear()
        {
            Assert.Null(WorkbookMemoryStore.Save(Key, "2026 预算.xlsx", "## 禁区\r\n- 汇总\r\n"));
            var state = Describe(Key);
            Assert.True(state.GetProperty("available").GetBoolean());
            Assert.Equal("## 禁区\n- 汇总\n", state.GetProperty("notes").GetString());
            Assert.Contains("上限", WorkbookMemoryStore.Save(Key, "x", new string('y', WorkbookMemoryStore.MaxNotesChars + 1)));

            File.WriteAllText(Path.Combine(WorkbookMemoryStore.DirectoryFor(Key), "history.jsonl"), "{\"tool\":\"a\"}\n{\"tool\":\"b\"}\n");
            Assert.Equal(2, Describe(Key).GetProperty("historyCount").GetInt32());

            WorkbookMemoryStore.Clear(Key);
            Assert.Equal("", Describe(Key).GetProperty("notes").GetString());
            Assert.False(Directory.Exists(WorkbookMemoryStore.DirectoryFor(Key)));
        }
    }
}
