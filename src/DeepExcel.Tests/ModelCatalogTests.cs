using DeepExcel.AddIn.Config;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// ★ 内置模型目录的边界：只允许"补充"，绝不允许"减少"。
    /// 厂商模型迭代速度远快于我们更新内置目录，用内置目录去覆盖/裁剪用户列表，
    /// 会让用户原本能用的模型凭空消失——这个测试就是钉死这条规则。
    /// </summary>
    public class ModelCatalogTests
    {
        [Fact]
        public void MergeModels_KeepsEveryExistingModelAndOrder()
        {
            var existing = new[] { "my-gateway-model", "claude-sonnet-5", "claude-opus-4.8" };
            var catalog = new[] { "claude-sonnet-5", "claude-opus-5" };

            var merged = AppConfig.MergeModels(existing, catalog);

            // 用户已有的一个都不能少，顺序不变（拖拽排好的优先级不能被打乱）
            Assert.Equal("my-gateway-model", merged[0]);
            Assert.Equal("claude-sonnet-5", merged[1]);
            Assert.Equal("claude-opus-4.8", merged[2]);
            // 目录里的新模型追加到末尾
            Assert.Equal("claude-opus-5", merged[3]);
            Assert.Equal(4, merged.Length);
        }

        [Fact]
        public void MergeModels_IsCaseInsensitiveAndHandlesNulls()
        {
            Assert.Equal(new[] { "GPT-5.5" }, AppConfig.MergeModels(new[] { "GPT-5.5" }, new[] { "gpt-5.5" }));
            Assert.Equal(new[] { "gpt-5" }, AppConfig.MergeModels(null, new[] { "gpt-5" }));
            Assert.Equal(new[] { "gpt-5" }, AppConfig.MergeModels(new[] { "gpt-5" }, null));
            Assert.Empty(AppConfig.MergeModels(null, null));
        }

        [Fact]
        public void DefaultCatalog_CoversEveryShippedProvider()
        {
            // 防止"随手精简"厂商：默认配置里这些厂商必须都在
            var config = AppConfig.CreateDefault();
            foreach (var key in new[]
            {
                "anthropic", "deepseek", "stepfun", "openai",
                "kimi", "qwen", "zhipu", "minimax", "doubao", "custom",
            })
            {
                Assert.True(config.Providers.ContainsKey(key), $"provider missing: {key}");
                Assert.NotEmpty(config.Providers[key].Models);
                Assert.False(string.IsNullOrEmpty(config.Providers[key].DefaultModel), $"no default model: {key}");
            }
        }

        [Fact]
        public void DefaultCatalog_KeepsBothClaudeGenerations()
        {
            // 4.8 / haiku-5 这些老名字仍要保留，升级不能把用户能用的模型删掉
            var models = AppConfig.CreateDefault().Providers["anthropic"].Models;
            Assert.Contains("claude-sonnet-5", models);
            Assert.Contains("claude-opus-5", models);
            Assert.Contains("claude-opus-4.8", models);
            Assert.Contains("claude-haiku-5", models);
        }
    }
}
