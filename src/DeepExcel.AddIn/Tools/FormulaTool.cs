using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Perception;

namespace DeepExcel.AddIn.Tools
{
    /// <summary>
    /// 公式工具 - 生成并插入Excel公式
    /// 支持常见公式模式，以及AI辅助公式生成
    /// </summary>
    public class FormulaTool
    {
        private readonly IExcelActions _excelActions;
        private readonly RangeAnalyzer _rangeAnalyzer;
        private readonly Application _app;

        public FormulaTool(IExcelActions excelActions, RangeAnalyzer rangeAnalyzer, Application app)
        {
            _excelActions = excelActions;
            _rangeAnalyzer = rangeAnalyzer;
            _app = app;
        }

        /// <summary>
        /// 直接写入公式
        /// </summary>
        public ToolResult WriteFormula(string address, string formula)
        {
            return _excelActions.WriteFormula(address, formula);
        }

        /// <summary>
        /// 智能填充公式到下方单元格
        /// </summary>
        public ToolResult FillDown(string fromAddress, int rowCount)
        {
            try
            {
                var app = _app;
                var fromRange = app.Range[fromAddress];
                var toRange = fromRange.Resize[rowCount, fromRange.Columns.Count];
                fromRange.AutoFill(toRange, XlAutoFillType.xlFillDefault);
                return new ToolResult { Name = "fill_formula", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "fill_formula", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 查找并替换公式中的引用
        /// </summary>
        public ToolResult ReplaceInFormulas(string rangeAddress, string find, string replace)
        {
            try
            {
                var app = _app;
                var range = app.Range[rangeAddress];
                int count = 0;

                foreach (Range cell in range.Cells)
                {
                    if ((bool)cell.HasFormula)
                    {
                        var formula = cell.Formula.ToString();
                        if (formula.Contains(find))
                        {
                            cell.Formula = formula.Replace(find, replace);
                            count++;
                        }
                    }
                }

                return new ToolResult { Name = "replace_formula", Success = true, Data = new { replacedCount = count } };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "replace_formula", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 分析指定区域的公式，返回结构化信息
        /// </summary>
        public FormulaAnalysis AnalyzeFormulas(string rangeAddress)
        {
            try
            {
                var app = _app;
                var range = app.Range[rangeAddress];
                var analysis = new FormulaAnalysis();
                var formulaCells = new List<FormulaCellInfo>();

                foreach (Range cell in range.Cells)
                {
                    if ((bool)cell.HasFormula)
                    {
                        formulaCells.Add(new FormulaCellInfo
                        {
                            Address = cell.Address,
                            Formula = cell.Formula.ToString(),
                            Value = cell.Value2?.ToString()
                        });
                    }
                }

                analysis.FormulaCount = formulaCells.Count;
                analysis.Formulas = formulaCells.ToArray();

                return analysis;
            }
            catch (Exception ex)
            {
                return new FormulaAnalysis { Error = ex.Message };
            }
        }
    }

    public class FormulaAnalysis
    {
        public int FormulaCount { get; set; }
        public FormulaCellInfo[] Formulas { get; set; }
        public string Error { get; set; }
    }

    public class FormulaCellInfo
    {
        public string Address { get; set; }
        public string Formula { get; set; }
        public string Value { get; set; }
    }
}
