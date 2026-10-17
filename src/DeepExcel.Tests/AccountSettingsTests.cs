using System.Text.Json;
using DeepExcel.AddIn.Config;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// The panel no longer asks for a server address, so the add-in must always
    /// know one: the built-in address unless config.json names another.
    /// </summary>
    public class AccountSettingsTests
    {
        [Fact]
        public void EmptyServerUrlFallsBackToTheBuiltInAddress()
        {
            Assert.Equal(AccountSettings.DefaultServerUrl, new AccountSettings().ResolveServerUrl());
            Assert.Equal(AccountSettings.DefaultServerUrl,
                new AccountSettings { ServerUrl = "   " }.ResolveServerUrl());
        }

        [Fact]
        public void ConfiguredServerUrlWinsForPrivateDeployments()
        {
            var settings = new AccountSettings { ServerUrl = " https://de.internal.example " };
            Assert.Equal("https://de.internal.example", settings.ResolveServerUrl());
        }

        [Fact]
        public void BuiltInAddressIsHttps()
        {
            Assert.StartsWith("https://", AccountSettings.DefaultServerUrl);
        }

        [Fact]
        public void OldConfigWithoutAccountSectionGetsTheDefault()
        {
            var config = JsonSerializer.Deserialize<AppConfig>("{\"Providers\":{}}");
            Assert.NotNull(config.Account);
            Assert.Equal(AccountSettings.DefaultServerUrl, config.Account.ResolveServerUrl());
        }
    }
}
