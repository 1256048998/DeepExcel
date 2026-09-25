using System;
using System.Collections.Generic;
using DeepExcel.AddIn.Account;
using DeepExcel.AddIn.Bridge;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Telemetry must never be able to hurt the product it measures, and must
    /// never carry workbook content off the machine.
    /// </summary>
    public class TelemetryReporterTests
    {
        private static TelemetryReporter NewReporter()
        {
            // A signed-out session: nothing can be sent, which is the state that
            // exercises buffering. Each test gets its own outbox so nothing lands
            // in the real %APPDATA% one.
            var root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DeepExcelTelemetryTests",
                Guid.NewGuid().ToString("N"));
            return new TelemetryReporter(new SessionManager(new TokenVault(root)),
                outboxPath: System.IO.Path.Combine(root, "outbox", "pending.jsonl"));
        }

        [Fact]
        public void RecordingNeverThrows()
        {
            using (var reporter = NewReporter())
            {
                // Called from user-facing paths; an exception here would surface
                // as a failed spreadsheet operation.
                reporter.Record("task_complete", null);
                reporter.Record("task_complete", new Dictionary<string, object> { ["outcome"] = "success" });
                reporter.Record(null);
                reporter.Record("");
                Assert.True(reporter.BufferedCount >= 2);
            }
        }

        [Fact]
        public void BufferIsBoundedSoAnOfflineWeekCannotGrowMemory()
        {
            using (var reporter = NewReporter())
            {
                for (int i = 0; i < 1000; i++)
                {
                    reporter.Record("session_start", new Dictionary<string, object> { ["host"] = "excel" });
                }
                // Oldest events are dropped rather than accumulating.
                Assert.True(reporter.BufferedCount <= 200,
                    "buffer grew to " + reporter.BufferedCount);
            }
        }

        [Fact]
        public void DisabledReporterRecordsNothing()
        {
            using (var reporter = NewReporter())
            {
                reporter.Enabled = false;
                reporter.Record("task_complete", new Dictionary<string, object> { ["outcome"] = "success" });
                Assert.Equal(0, reporter.BufferedCount);
            }
        }

        [Fact]
        public void InstallIdIsOpaqueAndStable()
        {
            using (var first = NewReporter())
            using (var second = NewReporter())
            {
                // Random, not a machine fingerprint: it answers "how many
                // installs", not "whose machine is this".
                Assert.Equal(32, first.InstallId.Length);
                Assert.DoesNotContain(Environment.MachineName, first.InstallId,
                    StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain(Environment.UserName, first.InstallId,
                    StringComparison.OrdinalIgnoreCase);
                // Persisted, so the same install is not counted twice.
                Assert.Equal(first.InstallId, second.InstallId);
            }
        }

        [Fact]
        public void FlushWithoutASessionIsANoOp()
        {
            using (var reporter = NewReporter())
            {
                reporter.Record("session_start", new Dictionary<string, object>());
                reporter.FlushAsync().GetAwaiter().GetResult();
                // Nowhere to send it yet, so it stays buffered rather than
                // being dropped.
                Assert.Equal(1, reporter.BufferedCount);
            }
        }

        // ------------------------------------------------------------------
        // Error classification
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("Operation timed out after 30s", "timeout")]
        [InlineData("REGDB_E_CLASSNOTREG: Class not registered", "com_not_registered")]
        [InlineData("Index was out of range", "range_out_of_bounds")]
        [InlineData("401 Unauthorized: invalid api key", "auth_failed")]
        [InlineData("429 rate limit exceeded", "rate_limited")]
        [InlineData("VBA project access denied", "permission_denied")]
        [InlineData("Something entirely novel", "other")]
        [InlineData("", "unknown")]
        [InlineData(null, "unknown")]
        public void ErrorsAreClassifiedIntoAFixedVocabulary(string message, string expected)
        {
            Assert.Equal(expected, MessageBridge.ClassifyError(message));
        }

        [Theory]
        [InlineData(null, "success")]
        [InlineData("success", "success")]
        [InlineData("interrupted", "cancelled")]
        [InlineData("max_turns", "error")]
        [InlineData("error", "error")]
        public void RunSummaryOutcomeDecidesTheTraceOutcome(string runOutcome, string expected)
        {
            // stream_end 以前一律记成 success：达到最大轮次、API 报错的任务也被当成可用结果，
            // 还会被推荐保存为技能。
            Assert.Equal(expected, MessageBridge.TraceOutcomeFromRunSummary(runOutcome));
        }

        [Theory]
        [InlineData("写入后出现 #NAME? 错误 3 处", "formula_name")]
        [InlineData("B2 返回 #VALUE!", "formula_value")]
        [InlineData("新增 #REF! 1 处", "formula_ref")]
        [InlineData("「汇总」在工作簿记忆的禁区里，已拒绝写入", "protected_zone")]
        [InlineData("目标区域 A1:C3 已有内容，但你还没有读过它，本次未写入。", "unread_target")]
        [InlineData("Input validation error: 'file' is a required property", "bad_arguments")]
        [InlineData("运行时错误 13：类型不匹配", "type_mismatch")]
        [InlineData("编译错误：缺少 End Sub（第 12 行）", "vba_compile")]
        public void DomainErrorsGetTheirOwnCategory(string message, string expected)
        {
            // 知识技能的「常见报错」按这些类别聚合，泛泛的 other 用不上
            Assert.Equal(expected, MessageBridge.ClassifyError(message));
        }

        private static System.Text.Json.JsonElement Event(string json) =>
            System.Text.Json.JsonDocument.Parse(json).RootElement;

        [Fact]
        public void FailedToolEndBecomesAToolError()
        {
            string name = MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_end\",\"id\":\"t1\",\"name\":\"mcp__excel__write_formula\",\"ok\":false," +
                "\"error\":{\"code\":\"tool_failed\",\"message\":\"写入后 C2 出现 #NAME?\"}}"), out string code);

            Assert.Equal("write_formula", name);
            Assert.Equal("formula_name", code);
        }

        [Fact]
        public void SpecificSidecarCodesAreKeptAndStopsAreNotReported()
        {
            Assert.Equal("find", MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_end\",\"name\":\"find\",\"ok\":false,\"error\":{\"code\":\"denied\",\"message\":\"x\"}}"),
                out string denied));
            Assert.Equal("denied", denied);

            Assert.Null(MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_end\",\"name\":\"find\",\"ok\":false,\"error\":{\"code\":\"interrupted\"}}"), out _));
            Assert.Null(MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_end\",\"name\":\"find\",\"ok\":true}"), out _));
            Assert.Null(MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_start\",\"name\":\"find\"}"), out _));
            // 工具名不合规（可能夹带内容）就不报
            Assert.Null(MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_end\",\"name\":\"C:\\\\Users\\\\alice\",\"ok\":false}"), out _));
        }

        [Fact]
        public void ToolErrorNeverCarriesTheMessage()
        {
            MessageBridge.ToolErrorFromUiEvent(Event(
                "{\"kind\":\"tool_end\",\"name\":\"write_value\",\"ok\":false," +
                "\"error\":{\"code\":\"Cannot write 128000 to alice\",\"message\":\"Cannot write 128000 to alice\"}}"),
                out string code);
            Assert.DoesNotContain("alice", code);
            Assert.DoesNotContain(" ", code);
        }

        [Fact]
        public void ClassificationNeverEchoesTheOriginalMessage()
        {
            // The whole point: raw error text quotes cell values and file paths,
            // so the classifier must emit a label, never the input.
            const string sensitive = @"Cannot write 128000 to 'C:\Users\alice\2026 财务预算.xlsx'!Sheet1!C4";
            var code = MessageBridge.ClassifyError(sensitive);

            Assert.DoesNotContain("alice", code);
            Assert.DoesNotContain("财务预算", code);
            Assert.DoesNotContain("128000", code);
            Assert.DoesNotContain("C4", code);
            // A short identifier, not a sentence.
            Assert.True(code.Length <= 40);
            Assert.DoesNotContain(" ", code);
        }
    }
}
