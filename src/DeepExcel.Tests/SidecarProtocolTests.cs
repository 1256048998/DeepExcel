using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    public class SidecarProtocolTests
    {
        [Fact]
        public void ToolResult_HasSuggestionAndContextFields()
        {
            var tr = new ToolResult
            {
                Name = "write_formula",
                Success = false,
                Error = "type mismatch",
                Suggestion = "目标区域包含文本，无法求和。建议改用 COUNTA。",
                Context = new { active_sheet = "Sheet1" },
            };
            Assert.NotNull(tr.Suggestion);
            Assert.NotNull(tr.Context);
        }

        [Fact]
        public void SidecarProtocol_TypeConstants_MatchSpec()
        {
            // C# → Python
            Assert.Equal("user_message", SidecarProtocol.TypeUserMessage);
            Assert.Equal("cancel", SidecarProtocol.TypeCancel);
            Assert.Equal("tool_result", SidecarProtocol.TypeToolResult);
            Assert.Equal("config", SidecarProtocol.TypeConfig);
            Assert.Equal("clarify_answer", SidecarProtocol.TypeClarifyAnswer);

            // Python → C#
            Assert.Equal("stream_delta", SidecarProtocol.TypeStreamDelta);
            Assert.Equal("tool_call", SidecarProtocol.TypeToolCall);
            Assert.Equal("clarify", SidecarProtocol.TypeClarify);
            Assert.Equal("stream_end", SidecarProtocol.TypeStreamEnd);
        }
    }
}
