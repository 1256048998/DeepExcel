using System;
using System.Collections.Generic;
using System.Linq;

namespace DeepExcel.AddIn.Advanced
{
    /// <summary>
    /// 图表规范引擎 - 根据数据特征自动选择最佳图表类型和样式
    /// </summary>
    public class ChartSpecificationEngine
    {
        public ChartRecommendation RecommendChart(List<List<object>> data, string[] headers = null)
        {
            var analysis = AnalyzeData(data, headers);
            var chartType = DetermineChartType(analysis);
            var style = DetermineChartStyle(analysis);

            return new ChartRecommendation
            {
                ChartType = chartType,
                Style = style,
                Analysis = analysis,
                Title = GenerateTitle(analysis, headers),
                XLabel = headers?.FirstOrDefault(),
                YLabel = headers?.Skip(1).FirstOrDefault()
            };
        }

        private DataAnalysisResult AnalyzeData(List<List<object>> data, string[] headers)
        {
            if (data == null || data.Count == 0)
                return new DataAnalysisResult();

            var result = new DataAnalysisResult
            {
                RowCount = data.Count,
                ColumnCount = data.FirstOrDefault()?.Count ?? 0,
                HasHeaders = headers != null && headers.Length > 0
            };

            // 检测数据类型分布
            var numericCols = new List<int>();
            var stringCols = new List<int>();

            for (int col = 0; col < result.ColumnCount; col++)
            {
                int numericCount = 0;
                int stringCount = 0;

                foreach (var row in data)
                {
                    if (col < row.Count)
                    {
                        var val = row[col];
                        if (val is double || val is int || (val is string s && double.TryParse(s, out _)))
                            numericCount++;
                        else if (val is string)
                            stringCount++;
                    }
                }

                if (numericCount > data.Count * 0.5)
                    numericCols.Add(col);
                else if (stringCount > data.Count * 0.5)
                    stringCols.Add(col);
            }

            result.NumericColumnCount = numericCols.Count;
            result.StringColumnCount = stringCols.Count;
            result.IsTimeSeries = DetectTimeSeries(data, stringCols);
            result.HasCategories = stringCols.Count > 0;

            return result;
        }

        private bool DetectTimeSeries(List<List<object>> data, List<int> stringCols)
        {
            foreach (var col in stringCols)
            {
                int dateCount = 0;
                foreach (var row in data)
                {
                    if (col < row.Count)
                    {
                        var val = row[col];
                        if (DateTime.TryParse(val?.ToString(), out _))
                            dateCount++;
                    }
                }
                if (dateCount > data.Count * 0.5)
                    return true;
            }
            return false;
        }

        private string DetermineChartType(DataAnalysisResult analysis)
        {
            // 根据数据特征选择最佳图表类型
            if (analysis.RowCount == 0 || analysis.ColumnCount == 0)
                return "column";

            // 时间序列数据 → 折线图
            if (analysis.IsTimeSeries && analysis.NumericColumnCount >= 1)
                return "line";

            // 单个数值列 + 分类 → 柱状图
            if (analysis.NumericColumnCount == 1 && analysis.HasCategories)
                return "column";

            // 多个数值列对比 → 柱状图或雷达图
            if (analysis.NumericColumnCount >= 2)
                return "column";

            // 部分占比 → 饼图
            if (analysis.NumericColumnCount == 1 && analysis.RowCount < 10)
                return "pie";

            // 散点图数据
            if (analysis.NumericColumnCount >= 2 && !analysis.HasCategories)
                return "scatter";

            return "column";
        }

        private ChartStyle DetermineChartStyle(DataAnalysisResult analysis)
        {
            // 根据数据规模和类型选择样式
            if (analysis.RowCount > 50)
                return ChartStyle.Minimal;

            if (analysis.ColumnCount >= 4)
                return ChartStyle.Professional;

            return ChartStyle.Standard;
        }

        private string GenerateTitle(DataAnalysisResult analysis, string[] headers)
        {
            if (headers != null && headers.Length >= 2)
            {
                return $"{headers[1]} 趋势";
            }
            return "数据分析图表";
        }
    }

    public class DataAnalysisResult
    {
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public bool HasHeaders { get; set; }
        public int NumericColumnCount { get; set; }
        public int StringColumnCount { get; set; }
        public bool IsTimeSeries { get; set; }
        public bool HasCategories { get; set; }
    }

    public class ChartRecommendation
    {
        public string ChartType { get; set; }
        public ChartStyle Style { get; set; }
        public DataAnalysisResult Analysis { get; set; }
        public string Title { get; set; }
        public string XLabel { get; set; }
        public string YLabel { get; set; }
    }

    public enum ChartStyle
    {
        Standard,
        Professional,
        Minimal,
        Modern
    }

    /// <summary>
    /// 模板推荐系统 - 根据用户场景和数据类型推荐Excel模板
    /// </summary>
    public class TemplateRecommender
    {
        private static readonly List<ExcelTemplate> Templates = new()
        {
            new ExcelTemplate
            {
                Id = "sales_report",
                Name = "销售报表",
                Description = "月度销售数据分析报表，包含趋势图和同比分析",
                Categories = new[] { "销售", "报表", "数据分析" },
                RequiredFeatures = new[] { "chart", "pivot_table", "formula" },
                Icon = "📊"
            },
            new ExcelTemplate
            {
                Id = "finance_ledger",
                Name = "财务台账",
                Description = "收支明细、余额追踪、月度汇总",
                Categories = new[] { "财务", "记账", "预算" },
                RequiredFeatures = new[] { "formula", "chart" },
                Icon = "💰"
            },
            new ExcelTemplate
            {
                Id = "project_tracker",
                Name = "项目进度跟踪",
                Description = "任务分解、进度追踪、甘特图",
                Categories = new[] { "项目", "任务", "进度" },
                RequiredFeatures = new[] { "chart", "formula" },
                Icon = "📋"
            },
            new ExcelTemplate
            {
                Id = "inventory",
                Name = "库存管理",
                Description = "库存盘点、预警提醒、出入库记录",
                Categories = new[] { "库存", "仓储", "管理" },
                RequiredFeatures = new[] { "formula", "conditional_formatting" },
                Icon = "📦"
            },
            new ExcelTemplate
            {
                Id = "survey_analysis",
                Name = "问卷分析",
                Description = "问卷数据录入、统计分析、可视化",
                Categories = new[] { "问卷", "调查", "统计" },
                RequiredFeatures = new[] { "chart", "pivot_table" },
                Icon = "📝"
            },
            new ExcelTemplate
            {
                Id = "dashboard",
                Name = "数据仪表盘",
                Description = "多维度数据展示、实时监控",
                Categories = new[] { "仪表盘", "监控", "可视化" },
                RequiredFeatures = new[] { "chart", "pivot_table", "formula" },
                Icon = "🎯"
            }
        };

        /// <summary>
        /// 根据用户查询推荐模板
        /// </summary>
        public List<TemplateRecommendation> Recommend(string query, DataAnalysisResult dataAnalysis = null)
        {
            var recommendations = new List<TemplateRecommendation>();

            foreach (var template in Templates)
            {
                var score = CalculateMatchScore(template, query, dataAnalysis);
                if (score > 0)
                {
                    recommendations.Add(new TemplateRecommendation
                    {
                        Template = template,
                        Score = score,
                        MatchReason = GetMatchReason(template, query)
                    });
                }
            }

            return recommendations.OrderByDescending(r => r.Score).Take(5).ToList();
        }

        /// <summary>
        /// 获取所有模板
        /// </summary>
        public List<ExcelTemplate> GetAllTemplates()
        {
            return Templates.ToList();
        }

        private double CalculateMatchScore(ExcelTemplate template, string query, DataAnalysisResult dataAnalysis)
        {
            double score = 0;
            var queryLower = query.ToLower();

            // 名称匹配
            if (template.Name.ToLower().Contains(queryLower))
                score += 3;

            // 描述匹配
            if (template.Description.ToLower().Contains(queryLower))
                score += 2;

            // 分类匹配
            foreach (var cat in template.Categories)
            {
                if (cat.ToLower().Contains(queryLower))
                    score += 1;
            }

            // 数据特征匹配
            if (dataAnalysis != null)
            {
                if (template.RequiredFeatures.Contains("chart") && dataAnalysis.NumericColumnCount >= 1)
                    score += 1;
                if (template.RequiredFeatures.Contains("pivot_table") && dataAnalysis.RowCount >= 20)
                    score += 1;
            }

            return score;
        }

        private string GetMatchReason(ExcelTemplate template, string query)
        {
            var queryLower = query.ToLower();
            if (template.Name.ToLower().Contains(queryLower))
                return $"名称包含 '{query}'";
            if (template.Description.ToLower().Contains(queryLower))
                return $"描述包含 '{query}'";
            return "匹配相关分类";
        }
    }

    public class ExcelTemplate
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string[] Categories { get; set; }
        public string[] RequiredFeatures { get; set; }
        public string Icon { get; set; }
    }

    public class TemplateRecommendation
    {
        public ExcelTemplate Template { get; set; }
        public double Score { get; set; }
        public string MatchReason { get; set; }
    }
}
