using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>权限模式：宿主只认三个值、只转发不保存</summary>
    public class PermissionModeTests
    {
        private static Message Msg(string payloadJson) => new Message
        {
            Type = "user_message",
            Payload = payloadJson == null ? (JsonElement?)null : JsonDocument.Parse(payloadJson).RootElement,
        };

        [Theory]
        [InlineData("{\"content\":\"x\",\"permission_mode\":\"plan\"}", "plan")]
        [InlineData("{\"content\":\"x\",\"permission_mode\":\"accept_writes\"}", "accept_writes")]
        [InlineData("{\"mode\":\"default\"}", "default")]
        [InlineData("{\"content\":\"x\"}", null)]
        [InlineData("{\"permission_mode\":\"bypass\"}", null)]
        [InlineData("{\"permission_mode\":3}", null)]
        [InlineData(null, null)]
        public void Only_the_three_modes_are_forwarded(string payload, string expected)
        {
            Assert.Equal(expected, MessageBridge.ReadPermissionMode(Msg(payload)));
        }

        [Fact]
        public void Protocol_knows_exactly_three_modes()
        {
            Assert.True(SidecarProtocol.IsPermissionMode("default"));
            Assert.True(SidecarProtocol.IsPermissionMode("accept_writes"));
            Assert.True(SidecarProtocol.IsPermissionMode("plan"));
            Assert.False(SidecarProtocol.IsPermissionMode("Plan"));
            Assert.False(SidecarProtocol.IsPermissionMode(""));
            Assert.False(SidecarProtocol.IsPermissionMode(null));
        }

        [Fact]
        public void Mode_is_not_part_of_the_saved_config()
        {
            // 「本次会话自动应用写入」不能持久化：配置对象上不该有任何权限模式字段
            foreach (var property in typeof(DeepExcel.AddIn.Config.AppConfig).GetProperties())
            {
                Assert.DoesNotContain("PermissionMode", property.Name);
            }
        }
    }
}
