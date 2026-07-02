using System;
using System.Collections.Generic;
using System.Drawing;
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

        /// <summary>
        /// 为图表添加数据标签
        /// </summary>
        public ToolResult AddDataLabels(string chartName = "", string position = "outside_end", bool showValue = true, bool showCategoryName = false, bool showPercentage = false)
        {
            try
            {
                Chart chart = FindChart(chartName);
                if (chart == null)
                    return new ToolResult { Name = "add_data_labels", Success = false, Error = "Chart not found" };

                var seriesCollection = (dynamic)chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;
                int labeledCount = 0;

                var labelPos = ParseLabelPosition(position);

                for (int i = 1; i <= seriesCount; i++)
                {
                    dynamic series = seriesCollection.Item(i);
                    series.HasDataLabels = true;
                    dynamic dataLabels = series.DataLabels();
                    dataLabels.Position = labelPos;
                    dataLabels.ShowValue = showValue;
                    dataLabels.ShowCategoryName = showCategoryName;
                    dataLabels.ShowPercentage = showPercentage;
                    labeledCount++;
                }

                return new ToolResult
                {
                    Name = "add_data_labels",
                    Success = true,
                    Data = new { seriesCount = labeledCount, position = position }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "AddDataLabels failed", ex);
                return new ToolResult { Name = "add_data_labels", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 设置图表标题
        /// </summary>
        public ToolResult SetChartTitle(string chartName = "", string title = "")
        {
            try
            {
                Chart chart = FindChart(chartName);
                if (chart == null)
                    return new ToolResult { Name = "set_chart_title", Success = false, Error = "Chart not found" };

                chart.HasTitle = !string.IsNullOrEmpty(title);
                if (chart.HasTitle)
                {
                    chart.ChartTitle.Text = title;
                }

                return new ToolResult
                {
                    Name = "set_chart_title",
                    Success = true,
                    Data = new { title = title }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "SetChartTitle failed", ex);
                return new ToolResult { Name = "set_chart_title", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 设置图表系列颜色
        /// </summary>
        public ToolResult SetChartColors(string chartName = "", string[] colorHexes = null)
        {
            try
            {
                Chart chart = FindChart(chartName);
                if (chart == null)
                    return new ToolResult { Name = "set_chart_colors", Success = false, Error = "Chart not found" };

                if (colorHexes == null || colorHexes.Length == 0)
                    return new ToolResult { Name = "set_chart_colors", Success = false, Error = "No colors provided" };

                var seriesCollection = (dynamic)chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;
                int coloredCount = 0;

                for (int i = 1; i <= seriesCount && i <= colorHexes.Length; i++)
                {
                    try
                    {
                        dynamic series = seriesCollection.Item(i);
                        var color = ColorTranslator.FromHtml(colorHexes[i - 1]);
                        var oleColor = ColorTranslator.ToOle(color);
                        series.Format.Fill.ForeColor.RGB = oleColor;
                        series.Format.Line.ForeColor.RGB = oleColor;
                        coloredCount++;
                    }
                    catch { }
                }

                return new ToolResult
                {
                    Name = "set_chart_colors",
                    Success = true,
                    Data = new { coloredCount = coloredCount }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "SetChartColors failed", ex);
                return new ToolResult { Name = "set_chart_colors", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 创建组合图（柱状+折线双轴）
        /// </summary>
        public ToolResult CreateComboChart(
            string dataRange,
            string title = "",
            string xLabel = "",
            string yLabel = "",
            string secondaryYLabel = "",
            int lineSeriesIndex = 2)
        {
            try
            {
                var dataRng = _app.Range[dataRange];
                Worksheet ws = dataRng.Worksheet;
                Logger.Instance.Info("ChartTool", $"CreateComboChart: dataRange={dataRange}");

                Shape shape;
                try
                {
                    shape = ws.Shapes.AddChart2(-1, XlChartType.xlColumnClustered);
                }
                catch
                {
                    shape = ws.Shapes.AddChart(XlChartType.xlColumnClustered);
                }
                Chart chart = shape.Chart;

                chart.SetSourceData(dataRng);

                if (!string.IsNullOrEmpty(title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                }

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

                var seriesCollection = (dynamic)chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;
                if (lineSeriesIndex > 0 && lineSeriesIndex <= seriesCount)
                {
                    dynamic lineSeries = seriesCollection.Item(lineSeriesIndex);
                    lineSeries.ChartType = XlChartType.xlLine;
                    lineSeries.AxisGroup = XlAxisGroup.xlSecondary;

                    if (!string.IsNullOrEmpty(secondaryYLabel))
                    {
                        var secAxis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary);
                        secAxis.HasTitle = true;
                        secAxis.AxisTitle.Text = secondaryYLabel;
                    }
                }

                Logger.Instance.Info("ChartTool", "CreateComboChart success");
                return new ToolResult
                {
                    Name = "create_combo_chart",
                    Success = true,
                    Data = new { title = title, lineSeriesIndex = lineSeriesIndex }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "CreateComboChart failed", ex);
                return new ToolResult { Name = "create_combo_chart", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 导出图表为图片
        /// </summary>
        public ToolResult ExportChartAsImage(string chartName = "", string outputPath = "", string format = "png")
        {
            try
            {
                Chart chart = FindChart(chartName);
                if (chart == null)
                    return new ToolResult { Name = "export_chart", Success = false, Error = "Chart not found" };

                if (string.IsNullOrEmpty(outputPath))
                {
                    string tempDir = System.IO.Path.GetTempPath();
                    string fileName = "chart_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "." + format;
                    outputPath = System.IO.Path.Combine(tempDir, fileName);
                }

                var filterName = format.ToLower() switch
                {
                    "png" => "PNG",
                    "jpg" or "jpeg" => "JPEG",
                    "gif" => "GIF",
                    "bmp" => "BMP",
                    _ => "PNG"
                };

                chart.Export(outputPath, filterName);

                return new ToolResult
                {
                    Name = "export_chart",
                    Success = true,
                    Data = new { outputPath = outputPath, format = format }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "ExportChartAsImage failed", ex);
                return new ToolResult { Name = "export_chart", Success = false, Error = ex.Message };
            }
        }

        private Chart FindChart(string chartName)
        {
            if (!string.IsNullOrEmpty(chartName))
            {
                try
                {
                    Chart chartObj = (Chart)_app.Charts[chartName];
                    if (chartObj != null) return chartObj;
                }
                catch { }
            }

            Worksheet ws = _app.ActiveSheet as Worksheet;
            if (ws == null) return null;

            Shape firstChartShape = null;
            foreach (Shape shape in ws.Shapes)
            {
                if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    if (string.IsNullOrEmpty(chartName))
                    {
                        return shape.Chart;
                    }
                    if (firstChartShape == null) firstChartShape = shape;
                    if (shape.Name == chartName || shape.Chart.ChartTitle?.Text == chartName)
                    {
                        return shape.Chart;
                    }
                }
            }

            return firstChartShape?.Chart;
        }

        private XlDataLabelPosition ParseLabelPosition(string pos)
        {
            return pos.ToLower() switch
            {
                "outside_end" => XlDataLabelPosition.xlLabelPositionOutsideEnd,
                "inside_end" => XlDataLabelPosition.xlLabelPositionInsideEnd,
                "inside_base" => XlDataLabelPosition.xlLabelPositionInsideBase,
                "center" => XlDataLabelPosition.xlLabelPositionCenter,
                "above" => XlDataLabelPosition.xlLabelPositionAbove,
                "below" => XlDataLabelPosition.xlLabelPositionBelow,
                "left" => XlDataLabelPosition.xlLabelPositionLeft,
                "right" => XlDataLabelPosition.xlLabelPositionRight,
                _ => XlDataLabelPosition.xlLabelPositionOutsideEnd
            };
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

        /// <summary>
        /// 刷新数据透视表
        /// </summary>
        public ToolResult RefreshPivotTable(string pivotTableName = "", string sheetName = "")
        {
            try
            {
                var pivotTable = FindPivotTable(pivotTableName, sheetName);
                if (pivotTable == null)
                    return new ToolResult { Name = "refresh_pivot", Success = false, Error = "Pivot table not found" };

                pivotTable.RefreshTable();

                return new ToolResult
                {
                    Name = "refresh_pivot",
                    Success = true,
                    Data = new { pivotTableName = pivotTable.Name }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "RefreshPivotTable failed", ex);
                return new ToolResult { Name = "refresh_pivot", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 透视表日期分组（按年/月/日）
        /// </summary>
        public ToolResult GroupPivotDateField(string pivotTableName, string fieldName, string groupBy = "month", string sheetName = "")
        {
            try
            {
                var pivotTable = FindPivotTable(pivotTableName, sheetName);
                if (pivotTable == null)
                    return new ToolResult { Name = "group_pivot_date", Success = false, Error = "Pivot table not found" };

                var pivotField = (PivotField)pivotTable.PivotFields(fieldName);
                if (pivotField == null)
                    return new ToolResult { Name = "group_pivot_date", Success = false, Error = "Field not found: " + fieldName };

                var range = pivotField.DataRange;
                if (range == null)
                    return new ToolResult { Name = "group_pivot_date", Success = false, Error = "No data in field" };

                bool byYears = false, byMonths = false, byDays = false, byQuarters = false;
                foreach (var part in groupBy.ToLower().Split(new[] { ',', ';', '+' }))
                {
                    switch (part.Trim())
                    {
                        case "year": byYears = true; break;
                        case "month": byMonths = true; break;
                        case "day": byDays = true; break;
                        case "quarter": byQuarters = true; break;
                    }
                }
                if (!byYears && !byMonths && !byDays && !byQuarters) byMonths = true;

                dynamic cell = range.Cells[1, 1];
                cell.Group(
                    Type.Missing, Type.Missing,
                    Type.Missing,
                    new object[] { byYears, byQuarters, byMonths, false, byDays, false, false });

                return new ToolResult
                {
                    Name = "group_pivot_date",
                    Success = true,
                    Data = new { fieldName = fieldName, groupBy = groupBy }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "GroupPivotDateField failed", ex);
                return new ToolResult { Name = "group_pivot_date", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 设置透视表值显示方式
        /// </summary>
        public ToolResult SetPivotValueDisplay(string pivotTableName, string valueField, string displayType = "normal", string baseField = "", string sheetName = "")
        {
            try
            {
                var pivotTable = FindPivotTable(pivotTableName, sheetName);
                if (pivotTable == null)
                    return new ToolResult { Name = "set_pivot_value_display", Success = false, Error = "Pivot table not found" };

                dynamic pivotDyn = pivotTable;
                var dataField = (PivotField)pivotDyn.DataFields(valueField);
                if (dataField == null)
                    return new ToolResult { Name = "set_pivot_value_display", Success = false, Error = "Value field not found: " + valueField };

                switch (displayType.ToLower())
                {
                    case "percent_of_column":
                        dataField.Calculation = XlPivotFieldCalculation.xlPercentOfColumn;
                        break;
                    case "percent_of_row":
                        dataField.Calculation = XlPivotFieldCalculation.xlPercentOfRow;
                        break;
                    case "percent_of_total":
                        dataField.Calculation = XlPivotFieldCalculation.xlPercentOfTotal;
                        break;
                    case "percent_of_parent":
                        dataField.Calculation = XlPivotFieldCalculation.xlPercentOfParent;
                        break;
                    case "running_total":
                        dataField.Calculation = XlPivotFieldCalculation.xlRunningTotal;
                        if (!string.IsNullOrEmpty(baseField))
                            dataField.BaseField = baseField;
                        break;
                    case "rank":
                        dataField.Calculation = (XlPivotFieldCalculation)61;
                        break;
                    default:
                        dataField.Calculation = XlPivotFieldCalculation.xlNoAdditionalCalculation;
                        break;
                }

                return new ToolResult
                {
                    Name = "set_pivot_value_display",
                    Success = true,
                    Data = new { valueField = valueField, displayType = displayType }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "SetPivotValueDisplay failed", ex);
                return new ToolResult { Name = "set_pivot_value_display", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 控制透视表总计显示
        /// </summary>
        public ToolResult SetPivotTotals(string pivotTableName, bool showRowTotals = true, bool showColumnTotals = true, string sheetName = "")
        {
            try
            {
                var pivotTable = FindPivotTable(pivotTableName, sheetName);
                if (pivotTable == null)
                    return new ToolResult { Name = "set_pivot_totals", Success = false, Error = "Pivot table not found" };

                pivotTable.RowGrand = showRowTotals;
                pivotTable.ColumnGrand = showColumnTotals;

                return new ToolResult
                {
                    Name = "set_pivot_totals",
                    Success = true,
                    Data = new { rowTotals = showRowTotals, columnTotals = showColumnTotals }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "SetPivotTotals failed", ex);
                return new ToolResult { Name = "set_pivot_totals", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 为透视表添加切片器
        /// </summary>
        public ToolResult AddPivotSlicer(string pivotTableName, string fieldName, string sheetName = "")
        {
            try
            {
                var pivotTable = FindPivotTable(pivotTableName, sheetName);
                if (pivotTable == null)
                    return new ToolResult { Name = "add_pivot_slicer", Success = false, Error = "Pivot table not found" };

                Worksheet ws = pivotTable.Parent as Worksheet;
                var wb = ws.Parent as Workbook;

                var slicerCaches = wb.SlicerCaches;
                var slicerCache = slicerCaches.Add(pivotTable, fieldName);

                var slicers = slicerCache.Slicers;
                var left = (double)pivotTable.TableRange2.Left + (double)pivotTable.TableRange2.Width + 20;
                var top = (double)pivotTable.TableRange2.Top;

                slicers.Add(ws, Type.Missing, Type.Missing, left, top, 150, 200);

                return new ToolResult
                {
                    Name = "add_pivot_slicer",
                    Success = true,
                    Data = new { fieldName = fieldName }
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ChartTool", "AddPivotSlicer failed", ex);
                return new ToolResult { Name = "add_pivot_slicer", Success = false, Error = ex.Message };
            }
        }

        private PivotTable FindPivotTable(string pivotTableName, string sheetName)
        {
            Worksheet ws = null;
            if (!string.IsNullOrEmpty(sheetName))
            {
                try { ws = (Worksheet)_app.Sheets[sheetName]; }
                catch { }
            }
            if (ws == null) ws = _app.ActiveSheet as Worksheet;
            if (ws == null) return null;

            dynamic pivotTables = ws.PivotTables();
            int count = pivotTables.Count;
            if (count == 0) return null;

            if (string.IsNullOrEmpty(pivotTableName))
                return (PivotTable)pivotTables.Item(1);

            try { return (PivotTable)pivotTables.Item(pivotTableName); }
            catch { return (PivotTable)pivotTables.Item(1); }
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
