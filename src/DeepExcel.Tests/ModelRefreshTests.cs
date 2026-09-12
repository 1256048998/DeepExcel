using System;
using System.IO;
using System.Reflection;
using DeepExcel.AddIn.Bridge;
using Xunit;

namespace DeepExcel.Tests
{
    public class ModelRefreshTests
    {
        public ModelRefreshTests()
        {
            AppDomain.CurrentDomain.AssemblyResolve -= ResolveTestDependency;
            AppDomain.CurrentDomain.AssemblyResolve += ResolveTestDependency;
        }

        private static Assembly ResolveTestDependency(object sender, ResolveEventArgs args)
        {
            if (!args.Name.StartsWith("System.Runtime.CompilerServices.Unsafe", StringComparison.Ordinal))
                return null;
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "System.Runtime.CompilerServices.Unsafe.dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        }

        [Fact]
        public void ParseModelIds_ReadsOpenAiAndAnthropicDataShape()
        {
            var models = MessageBridge.ParseModelIds(
                "{\"data\":[{\"id\":\"model-b\"},{\"id\":\"model-a\"}]}");

            // ★ 保持厂商接口返回顺序（新模型通常在最前），不再按字母排序
            Assert.Equal(new[] { "model-b", "model-a" }, models);
        }

        [Fact]
        public void ParseModelIds_ReadsStringArrayAndRemovesDuplicates()
        {
            var models = MessageBridge.ParseModelIds(
                "{\"models\":[\"alpha\",\"ALPHA\",\"beta\"]}");

            Assert.Equal(new[] { "alpha", "beta" }, models);
        }

        [Fact]
        public void ParseModelIds_UnknownShapeReturnsEmptyList()
        {
            var models = MessageBridge.ParseModelIds("{\"result\":true}");

            Assert.Empty(models);
        }

        [Fact]
        public void BuildNextPageUrl_FollowsAnthropicCursor()
        {
            var next = MessageBridge.BuildNextPageUrl(
                "https://api.anthropic.com/v1/models",
                "{\"data\":[{\"id\":\"claude-opus-5\"}],\"has_more\":true,\"last_id\":\"claude-opus-5\"}");

            Assert.Equal("https://api.anthropic.com/v1/models?limit=100&after_id=claude-opus-5", next);
        }

        [Fact]
        public void BuildNextPageUrl_ReturnsNullWhenNoMorePages()
        {
            Assert.Null(MessageBridge.BuildNextPageUrl(
                "https://api.deepseek.com/models",
                "{\"data\":[{\"id\":\"deepseek-v4-pro\"}],\"has_more\":false}"));
            Assert.Null(MessageBridge.BuildNextPageUrl(
                "https://api.openai.com/v1/models",
                "{\"data\":[{\"id\":\"gpt-5.5\"}]}"));
        }

        [Theory]
        // 智谱：对话走 /api/anthropic，模型列表在 /api/paas/v4/models —— 必须靠厂商已知端点兜底
        [InlineData("zhipu", "https://api.z.ai/api/anthropic", "https://api.z.ai/api/paas/v4/models")]
        // 豆包：对话走 /api/compatible，模型列表在 /api/v3/models
        [InlineData("doubao", "https://ark.cn-beijing.volces.com/api/compatible", "https://ark.cn-beijing.volces.com/api/v3/models")]
        // 阶跃星辰：对话走 /step_plan，模型列表在 /v1/models
        [InlineData("stepfun", "https://api.stepfun.com/step_plan", "https://api.stepfun.com/v1/models")]
        // 通义千问：剥离 /anthropic 后的 /compatible-mode/v1/models 就是正确地址
        [InlineData("qwen", "https://dashscope.aliyuncs.com/compatible-mode/anthropic", "https://dashscope.aliyuncs.com/compatible-mode/v1/models")]
        public void BuildModelEndpointCandidates_CoversProviderRealEndpoint(
            string provider, string baseUrl, string expected)
        {
            var candidates = MessageBridge.BuildModelEndpointCandidates(provider, baseUrl);

            Assert.Contains(expected, candidates);
        }

        [Fact]
        public void BuildModelEndpointCandidates_DoesNotWalkPastHost()
        {
            var candidates = MessageBridge.BuildModelEndpointCandidates("custom", "https://gateway.internal/v1");

            Assert.Contains("https://gateway.internal/v1/models", candidates);
            Assert.DoesNotContain("https://gateway.internal//models", candidates);
            Assert.DoesNotContain("https:/models", candidates);
        }
    }
}
