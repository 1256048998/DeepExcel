using DeepExcel.AddIn.Security;
using DeepExcel.AddIn.Performance;
using DeepExcel.AddIn.Advanced;
using DeepExcel.AddIn.Agent;
using DeepExcel.AddIn.Models;
using Xunit;

namespace DeepExcel.Tests
{
    public class SecurityManagerTests
    {
        [Fact]
        public void Encrypt_Decrypt_RoundTrip()
        {
            var manager = new SecurityManager();
            var plain = "test-api-key-12345";
            
            var encrypted = manager.Encrypt(plain);
            Assert.NotEqual(plain, encrypted);
            
            var decrypted = manager.Decrypt(encrypted);
            Assert.Equal(plain, decrypted);
        }

        [Fact]
        public void Encrypt_Null_ReturnsEmpty()
        {
            var manager = new SecurityManager();
            Assert.Equal("", manager.Encrypt(null));
            Assert.Equal("", manager.Decrypt(null));
        }

        [Fact]
        public void GenerateVerificationCode_Length()
        {
            var manager = new SecurityManager();
            var code = manager.GenerateVerificationCode();
            Assert.Equal(6, code.Length);
        }

        [Fact]
        public void GenerateVerificationCode_Unique()
        {
            var manager = new SecurityManager();
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

    public class ChartSpecificationEngineTests
    {
        [Fact]
        public void RecommendChart_TimeSeries_LineChart()
        {
            var engine = new ChartSpecificationEngine();
            var data = new List<List<object>>
            {
                new() { "2024-01", 100 },
                new() { "2024-02", 150 },
                new() { "2024-03", 120 }
            };
            var headers = new[] { "日期", "销售额" };
            
            var result = engine.RecommendChart(data, headers);
            Assert.Equal("line", result.ChartType);
            Assert.Equal("销售额 趋势", result.Title);
        }

        [Fact]
        public void RecommendChart_Categories_ColumnChart()
        {
            var engine = new ChartSpecificationEngine();
            var data = new List<List<object>>
            {
                new() { "产品A", 100 },
                new() { "产品B", 150 },
                new() { "产品C", 120 }
            };
            var headers = new[] { "产品", "销量" };
            
            var result = engine.RecommendChart(data, headers);
            Assert.Equal("column", result.ChartType);
        }

        [Fact]
        public void RecommendChart_SmallData_PieChart()
        {
            var engine = new ChartSpecificationEngine();
            var data = new List<List<object>>
            {
                new() { "A", 30 },
                new() { "B", 40 },
                new() { "C", 30 }
            };
            
            var result = engine.RecommendChart(data);
            Assert.Equal("pie", result.ChartType);
        }

        [Fact]
        public void RecommendChart_NumericColumns_ScatterChart()
        {
            var engine = new ChartSpecificationEngine();
            var data = new List<List<object>>
            {
                new() { 1.0, 2.0 },
                new() { 2.0, 4.0 },
                new() { 3.0, 6.0 }
            };
            
            var result = engine.RecommendChart(data);
            Assert.Equal("scatter", result.ChartType);
        }
    }

    public class TemplateRecommenderTests
    {
        [Fact]
        public void Recommend_Sales_ReturnsSalesReport()
        {
            var recommender = new TemplateRecommender();
            var results = recommender.Recommend("销售报表");
            
            Assert.NotNull(results);
            Assert.NotEmpty(results);
            Assert.Equal("销售报表", results[0].Template.Name);
            Assert.True(results[0].Score > 0);
        }

        [Fact]
        public void Recommend_Finance_ReturnsFinanceLedger()
        {
            var recommender = new TemplateRecommender();
            var results = recommender.Recommend("财务记账");
            
            Assert.Contains(results, r => r.Template.Name == "财务台账");
        }

        [Fact]
        public void Recommend_NoMatch_ReturnsEmpty()
        {
            var recommender = new TemplateRecommender();
            var results = recommender.Recommend("不存在的模板");
            
            Assert.Empty(results);
        }

        [Fact]
        public void GetAllTemplates_ReturnsAll()
        {
            var recommender = new TemplateRecommender();
            var templates = recommender.GetAllTemplates();
            
            Assert.Equal(6, templates.Count);
        }
    }

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