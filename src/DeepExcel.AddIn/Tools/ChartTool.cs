using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Tools
{
    /// <summary>
    /// 图表与透视表创建工具
    /// </summary>
    public class ChartTool
    {
        private readonly Application _app;

        public ChartTool(Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 创建图表
        /// </summary>
        public ToolResult CreateChart(
            string dataRange,
            string chartType = "Column",
            string title = "",
            string xLabel = "",
            string yLabel = "",
            string chartSheet = "")
        {
            Logger.Instance.Info("ChartTool", $"CreateChart: dataRange={dataRange}, chartType={chartType}, title={title}");
            try
            {
                var dataRng = _app.Range[dataRange];
                Worksheet ws = dataRng.Worksheet;
                Logger.Instance.Info("ChartTool", $"  worksheet={ws.Name}, range address={dataRng.Address}");

                // 确定图表位置
                Chart chart;
                if (!string.IsNullOrEmpty(chartSheet))
                {
                    chart = (Chart)_app.Charts.Add();
                    chart.Name = chartSheet;
                }
                else
                {
                    // 嵌入到当前工作表
                    // ★ AddChart2 在某些 Excel 版本/区域会抛 HRESULT 异常，用 AddChart 兼容老版本
                    var shapes = ws.Shapes;
                    Shape shape;
                    try
                    {
                        // 先尝试 AddChart2（Excel 2013+）
                        var chartTypeEnum = ParseChartType(chartType);
                        shape = shapes.AddChart2(-1, chartTypeEnum);
                    }
                    catch (Exception ex2)
                    {
                        Logger.Instance.Warning("ChartTool", "AddChart2 failed, falling back to AddChart: " + ex2.Message);
                        // 回退到 AddChart（Excel 2007+）
                        shape = shapes.AddChart(ParseChartType(chartType));
                    }
                    chart = shape.Chart;
                }

                // 设置数据源
                chart.SetSourceData(dataRng);

                // 设置图表类型
                chart.ChartType = ParseChartType(chartType);

                // 设置标题
                if (!string.IsNullOrEmpty(title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                }

                // 设置坐标轴标题
                if (!string.IsNullOrEmpty(xLabel))
                {
                    var xAxis = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                    xAxis.HasTitle = true;
                    xAxis.AxisTitle.Text = xLabel;
                }
                if (!string.IsNullOrEmpty(yLabel))
                {
                    var yAxis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                    yAxis.HasTitle = true;
                    yAxis.AxisTitle.Text = yLabel;
                }

                Logger.Instance.Info("ChartTool", "CreateChart success");
                return new ToolResult
                {
                    Name = "create_chart",
                    Success = true,
                    Data = new { chartType = chartType, title = title }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "CreateChart failed: dataRange=" + dataRange, ex);
                return new ToolResult
                {
                    Name = "create_chart",
                    Success = false,
                    Error = ex.Message + " (dataRange=" + dataRange + ")"
                };
            }
        }

        /// <summary>
        /// 创建数据透视表
        /// </summary>
        public ToolResult CreatePivotTable(
            string sourceRange,
            string destinationSheet,
            string pivotTableName = "PivotTable1",
            string[] rowFields = null,
            string[] columnFields = null,
            string[] valueFields = null,
            string valueFunction = "Sum")
        {
            try
            {
                var wsSource = _app.Range[sourceRange].Worksheet;
                var wb = wsSource.Parent as Workbook;

                // 创建目标工作表
                Worksheet wsDest;
                try
                {
                    wsDest = (Worksheet)wb.Sheets[destinationSheet];
                }
                catch
                {
                    wsDest = (Worksheet)wb.Sheets.Add();
                    wsDest.Name = destinationSheet;
                }

                var pivotCache = wb.PivotCaches().Create(
                    XlPivotTableSourceType.xlDatabase,
                    sourceRange);

                var pivotTable = pivotCache.CreatePivotTable(
                    wsDest.Range["A1"],
                    pivotTableName);

                // 添加行字段
                if (rowFields != null)
                {
                    foreach (var field in rowFields)
                    {
                        ((PivotField)pivotTable.PivotFields(field)).Orientation = XlPivotFieldOrientation.xlRowField;
                    }
                }

                // 添加列字段
                if (columnFields != null)
                {
                    foreach (var field in columnFields)
                    {
                        ((PivotField)pivotTable.PivotFields(field)).Orientation = XlPivotFieldOrientation.xlColumnField;
                    }
                }

                // 添加值字段
                if (valueFields != null)
                {
                    var func = ParseFunction(valueFunction);
                    foreach (var field in valueFields)
                    {
                        pivotTable.AddDataField(pivotTable.PivotFields(field), field + " " + valueFunction, func);
                    }
                }

                return new ToolResult
                {
                    Name = "create_pivot_table",
                    Success = true,
                    Data = new { pivotTableName = pivotTableName, sheet = destinationSheet }
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Name = "create_pivot_table",
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        private XlChartType ParseChartType(string type)
        {
            return type.ToLower() switch
            {
                "column" or "bar" => XlChartType.xlColumnClustered,
                "line" => XlChartType.xlLine,
                "pie" => XlChartType.xlPie,
                "scatter" => XlChartType.xlXYScatter,
                "area" => XlChartType.xlArea,
                "doughnut" => XlChartType.xlDoughnut,
                "bar_horizontal" => XlChartType.xlBarClustered,
                _ => XlChartType.xlColumnClustered
            };
        }

        private XlConsolidationFunction ParseFunction(string func)
        {
            return func.ToLower() switch
            {
                "sum" => XlConsolidationFunction.xlSum,
                "count" => XlConsolidationFunction.xlCount,
                "average" or "avg" => XlConsolidationFunction.xlAverage,
                "max" => XlConsolidationFunction.xlMax,
                "min" => XlConsolidationFunction.xlMin,
                "product" => XlConsolidationFunction.xlProduct,
                _ => XlConsolidationFunction.xlSum
            };
        }
    }
}
