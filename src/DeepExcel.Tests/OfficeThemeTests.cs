using DeepExcel.AddIn.Bridge;
using Xunit;

namespace DeepExcel.Tests
{
    public class OfficeThemeTests
    {
        [Theory]
        [InlineData(0, "light")]   // 彩色
        [InlineData(5, "light")]   // 白色
        [InlineData(3, "dark")]    // 深灰
        [InlineData(4, "dark")]    // 黑色
        [InlineData(6, "system")]  // 使用系统设置
        [InlineData(1, "light")]   // 2013 浅灰
        [InlineData(2, "dark")]    // 2013 深灰
        [InlineData(99, "system")]
        public void MapsOfficeUiThemeToPanelTheme(int value, string expected)
        {
            Assert.Equal(expected, OfficeTheme.FromRegistryValue(value));
        }

        [Fact]
        public void MissingValueFollowsSystem()
        {
            Assert.Equal("system", OfficeTheme.FromRegistryValue(null));
        }

        [Fact]
        public void ReadNeverThrows()
        {
            Assert.Contains(OfficeTheme.Read(), new[] { "light", "dark", "system" });
        }
    }
}
