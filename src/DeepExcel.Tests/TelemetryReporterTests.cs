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
            // exercises buffering.
            return new TelemetryReporter(new SessionManager(new TokenVault(
                System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DeepExcelTelemetryTests",
                    Guid.NewGuid().ToString("N")))));
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
