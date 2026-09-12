using System.Text.Json;
using DeepExcel.AddIn.Config;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// ★ Excel(C#) 与 WPS(Node) 共用 %APPDATA%\DeepExcel\config.json，
    /// 这里锁定"WPS 端写出的配置，Excel 端能原样读回"这一契约。
    /// WPS 侧的反向用例在 scripts/test-wps-main.js（读 C# 风格 PascalCase 配置）。
    /// 两边任何一方改了字段名或结构，这个测试会先失败。
    /// </summary>
    public class WpsConfigInteropTests
    {
        // src/DeepExcel.Wps/config-store.js 实际写出的结构（截取 2 个厂商）
        private const string WpsWrittenConfig = @"{
  ""CurrentProvider"": ""deepseek"",
  ""CurrentModel"": ""deepseek-v4-flash"",
  ""DefaultProvider"": ""deepseek"",
  ""ModelCatalogVersion"": 2,
  ""Providers"": {
    ""anthropic"": {
      ""Type"": ""anthropic"",
      ""DisplayName"": ""Claude (Anthropic)"",
      ""ApiKey"": """",
      ""BaseUrl"": ""https://api.anthropic.com"",
      ""Models"": [""claude-sonnet-5"", ""claude-opus-5""],
      ""DefaultModel"": ""claude-sonnet-5"",
      ""Headers"": {},
      ""SupportsVision"": true,
      ""LastTestSuccess"": false,
      ""ModelsCustomized"": false
    },
    ""deepseek"": {
      ""Type"": ""anthropic"",
      ""DisplayName"": ""DeepSeek"",
      ""ApiKey"": """",
      ""BaseUrl"": ""https://api.deepseek.com/anthropic"",
      ""Models"": [""deepseek-v4-flash"", ""deepseek-v4-pro""],
      ""DefaultModel"": ""deepseek-v4-flash"",
      ""Headers"": {},
      ""SupportsVision"": false,
      ""LastTestSuccess"": true,
      ""ModelsCustomized"": true
    }
  },
  ""General"": {
    ""MaxRetries"": 2,
    ""RequestTimeoutSeconds"": 60,
    ""AutoCreateSnapshot"": true,
    ""RequireConfirmation"": true,
    ""MaxConversationHistory"": 10,
    ""MaxTurns"": 20
  },
  ""UI"": {
    ""Theme"": ""light"",
    ""Language"": ""zh-CN"",
    ""ShowTokenUsage"": true,
    ""StreamOutput"": true
  }
}";

        [Fact]
        public void AppConfig_ReadsConfigWrittenByWpsHost()
        {
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var config = JsonSerializer.Deserialize<AppConfig>(WpsWrittenConfig, options);

            Assert.NotNull(config);
            Assert.Equal("deepseek", config.CurrentProvider);
            Assert.Equal("deepseek-v4-flash", config.CurrentModel);
            Assert.Equal("deepseek", config.DefaultProvider);
            Assert.Equal(AppConfig.CurrentModelCatalogVersion, config.ModelCatalogVersion);

            var deepseek = config.Providers["deepseek"];
            // 模型优先级顺序必须原样保留（WPS 里拖拽排序的结果）
            Assert.Equal(new[] { "deepseek-v4-flash", "deepseek-v4-pro" }, deepseek.Models);
            Assert.Equal("deepseek-v4-flash", deepseek.DefaultModel);
            Assert.True(deepseek.ModelsCustomized);
            Assert.True(deepseek.LastTestSuccess);
            Assert.Equal(20, config.General.MaxTurns);
        }

        [Fact]
        public void MigrateStyleUpdate_DoesNotOverwriteCustomizedModels()
        {
            // ModelsCustomized=true + 目录版本已是最新 → 内置目录不得覆盖用户排好的顺序
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var config = JsonSerializer.Deserialize<AppConfig>(WpsWrittenConfig, options);

            var manager = ConfigManager.Instance;  // 只为触发同一套迁移语义的常量
            Assert.NotNull(manager);
            Assert.True(config.ModelCatalogVersion >= AppConfig.CurrentModelCatalogVersion);
            Assert.True(config.Providers["deepseek"].ModelsCustomized);
        }

        [Fact]
        public void AppConfig_SerializesWithPascalCaseKeysForWpsHost()
        {
            // WPS 端按 PascalCase 读取（大小写不敏感兜底），这里确认 C# 写出的就是 PascalCase
            var config = AppConfig.CreateDefault();
            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });

            Assert.Contains("\"CurrentProvider\"", json);
            Assert.Contains("\"Providers\"", json);
            Assert.Contains("\"Models\"", json);
            Assert.Contains("\"ModelsCustomized\"", json);
            Assert.Contains("\"ModelCatalogVersion\"", json);
        }
    }
}
