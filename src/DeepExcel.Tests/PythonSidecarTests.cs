using System.Windows.Forms;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    public class PythonSidecarTests
    {
        [Fact]
        public void PythonSidecar_CanBeConstructedWithDeps()
        {
            var excelActions = new FakeExcelActions();
            var uiControl = new UserControl();  // 真实 Control 用于 Invoke
            var sidecar = new PythonSidecar(excelActions, null, uiControl);
            Assert.NotNull(sidecar);
            sidecar.Dispose();
        }

        [Fact]
        public void PythonSidecar_GetSidecarPath_ReturnsValidPath()
        {
            // 静态方法测试，不依赖进程
            var path = PythonSidecar.GetSidecarPath();
            Assert.Contains("sidecar.py", path);
        }

        [Fact]
        public void PythonSidecar_GetPythonPath_ReturnsNonEmpty()
        {
            var py = PythonSidecar.GetPythonPath();
            Assert.False(string.IsNullOrEmpty(py));
        }

        [Fact]
        public void Tool_result_carries_the_warning_and_the_backup_id_to_the_model()
        {
            // sidecar 把整条 tool_result 原样交给模型；以前 warning 根本不在消息里
            var json = PythonSidecar.BuildToolResultJson("c1", true, new { n = 1 }, null, null, null,
                "backup-1", "已自动按 has_header=true 排序");
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            var root = doc.RootElement;
            Assert.Equal("tool_result", root.GetProperty("type").GetString());
            Assert.Equal("已自动按 has_header=true 排序", root.GetProperty("warning").GetString());
            Assert.Equal("backup-1", root.GetProperty("backup_snapshot_id").GetString());
        }

        [Fact]
        public void Unserializable_data_still_keeps_the_warning_and_backup_id()
        {
            var json = PythonSidecar.BuildToolResultJson("c1", true, new object[,] { { new System.IO.MemoryStream() } },
                null, null, null, "backup-1", "w");
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            Assert.Equal("backup-1", doc.RootElement.GetProperty("backup_snapshot_id").GetString());
            Assert.Equal("w", doc.RootElement.GetProperty("warning").GetString());
        }

        [Fact]
        public void Restore_history_carries_the_engine_session_only_when_there_is_one()
        {
            var messages = new System.Collections.Generic.List<DeepExcel.AddIn.Collaboration.HistoryMessage>
            {
                new DeepExcel.AddIn.Collaboration.HistoryMessage { role = "user", content = "汇总一下" },
            };
            using (var doc = System.Text.Json.JsonDocument.Parse(
                PythonSidecar.BuildRestoreHistoryJson(messages, "sess-1", "fp-1")))
            {
                var root = doc.RootElement;
                Assert.Equal("restore_history", root.GetProperty("type").GetString());
                Assert.Equal("sess-1", root.GetProperty("resume_session_id").GetString());
                Assert.Equal("fp-1", root.GetProperty("resume_fingerprint").GetString());
                Assert.Equal(1, root.GetProperty("messages").GetArrayLength());
            }
            using (var doc = System.Text.Json.JsonDocument.Parse(
                PythonSidecar.BuildRestoreHistoryJson(messages, null, null)))
            {
                Assert.False(doc.RootElement.TryGetProperty("resume_session_id", out _));
            }
        }

        // ★ IExcelActions mock 已抽到共享的 FakeExcelActions.cs，
        // 避免接口增加方法时每个测试文件都要各自补一遍 stub
    }
}
