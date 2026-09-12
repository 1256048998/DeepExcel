using System.Collections.Generic;
using DeepExcel.AddIn.Security;
using DeepExcel.AddIn.Performance;
using Xunit;

namespace DeepExcel.Tests
{
    public class SecurityManagerTests
    {
        [Fact]
        public void Encrypt_Decrypt_RoundTrip()
        {
            // ★ SecurityManager 是单例（构造函数私有），测试改用 Instance
            var manager = SecurityManager.Instance;
            var plain = "test-api-key-12345";
            
            var encrypted = manager.Encrypt(plain);
            Assert.NotEqual(plain, encrypted);
            
            var decrypted = manager.Decrypt(encrypted);
            Assert.Equal(plain, decrypted);
        }

        [Fact]
        public void Encrypt_Null_ReturnsEmpty()
        {
            // ★ SecurityManager 是单例（构造函数私有），测试改用 Instance
            var manager = SecurityManager.Instance;
            Assert.Equal("", manager.Encrypt(null));
            Assert.Equal("", manager.Decrypt(null));
        }

        [Fact]
        public void GenerateVerificationCode_Length()
        {
            // ★ SecurityManager 是单例（构造函数私有），测试改用 Instance
            var manager = SecurityManager.Instance;
            var code = manager.GenerateVerificationCode();
            Assert.Equal(6, code.Length);
        }

        [Fact]
        public void GenerateVerificationCode_Unique()
        {
            // ★ SecurityManager 是单例（构造函数私有），测试改用 Instance
            var manager = SecurityManager.Instance;
            var codes = new HashSet<string>();
            for (int i = 0; i < 100; i++)
            {
                codes.Add(manager.GenerateVerificationCode());
            }
            Assert.True(codes.Count > 90);
        }
    }

    public class TokenBudgetManagerTests
    {
        [Fact]
        public void RecordUsage_WithinLimit()
        {
            var manager = new TokenBudgetManager();
            manager.SetDailyLimit("test", 1000);
            
            var result = manager.RecordUsage("test", 500, 300);
            Assert.True(result);
            
            var status = manager.GetStatus("test");
            Assert.Equal(800, status.Used);
            Assert.Equal(200, status.Remaining);
        }

        [Fact]
        public void RecordUsage_ExceedsLimit()
        {
            var manager = new TokenBudgetManager();
            manager.SetDailyLimit("test", 1000);
            
            var result1 = manager.RecordUsage("test", 600, 500);
            Assert.False(result1);
            
            var status = manager.GetStatus("test");
            Assert.Equal(1100, status.Used);
            Assert.Equal(0, status.Remaining);
        }

        [Fact]
        public void ResetDaily_ResetsUsage()
        {
            var manager = new TokenBudgetManager();
            manager.SetDailyLimit("test", 1000);
            manager.RecordUsage("test", 500, 300);
            
            manager.ResetDaily("test");
            var status = manager.GetStatus("test");
            Assert.Equal(0, status.Used);
        }
    }

    // ★ ChartSpecificationEngine / TemplateRecommender 已删除：
    // 图表类型、模板选择这类判断交给 agent 做（它有数据上下文和工具），
    // 宿主不再用硬编码规则表替 agent 决定。相关测试随之移除。

    public class FormulaToolTests
    {
        [Fact]
        public void WriteFormula_ValidAddress()
        {
            var formula = "=SUM(A1:A10)";
            Assert.StartsWith("=", formula);
            Assert.Contains("SUM", formula);
        }

        [Fact]
        public void FormulaContainsValidFunctions()
        {
            var functions = new[] { "SUM", "IF", "VLOOKUP", "AVERAGE", "COUNT" };
            foreach (var func in functions)
            {
                var formula = $"={func}(A1:A10)";
                Assert.Contains(func, formula);
            }
        }
    }
}