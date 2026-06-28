using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Tools;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// 工具执行调度器：接收 sidecar 的 tool_call，在 STA 线程执行 Excel 操作
    /// 从 Orchestrator.ExecuteToolAsync 移植，保留 GetArg/GetArgArray 的 JsonElement 处理
    /// </summary>
    public class ToolDispatcher
    {
        private readonly IExcelActions _excel;
        private readonly Application _excelApp;

        /// <summary>
        /// read_range 返回的最大行数限制。
        /// Claude Agent SDK 内部消息缓冲区限制 1MB（1048576 bytes），
        /// 超过会导致 "JSON message exceeded maximum buffer size" 崩溃。
        /// 200 行 × 20 列 × 平均 30 字节 ≈ 120KB，留足余量。
        /// </summary>
        private const int MaxReadRangeRows = 200;

        /// <summary>
        /// 共享的 JSON 序列化选项：注册 2D 数组转换器，避免序列化 RangeInfo 时崩溃。
        /// </summary>
        private static readonly JsonSerializerOptions _jsonOptions = BuildJsonOptions();

        private static JsonSerializerOptions BuildJsonOptions()
        {
            var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            opts.Converters.Add(new Object2DArrayConverter());
            opts.Converters.Add(new String2DArrayConverter());
            return opts;
        }

        public ToolDispatcher(IExcelActions excel, Application excelApp = null)
        {
            _excel = excel;
            _excelApp = excelApp;
        }

        /// <summary>
        /// 同步执行工具（必须在 STA 主线程调用）
        /// </summary>
        public ToolResult Execute(string toolName, Dictionary<string, object> args)
        {
            Logger.Instance.Info("ToolDispatcher", "Execute: " + toolName + ", args keys=" + (args == null ? "null" : string.Join(",", args.Keys)));

            try
            {
                switch (toolName)
                {
                    // 注：每个 case 的执行结果（含 success/error）在 switch 结束后统一记录
                    case "echo":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = new { echo = GetArg<string>(args, "text") },
                        };

                    case "read_range":
                        var address = GetArg<string>(args, "address");
                        var rangeData = _excel.ReadRange(address);
                        // ★ Claude Agent SDK 内部消息缓冲区限制 1MB（1048576 bytes），
                        // read_range 返回的 Values/Formulas 二维数组序列化后可能超过此限制，
                        // 导致 SDK 抛 "JSON message exceeded maximum buffer size" 崩溃。
                        // 解决：对超过 MaxRows 的数据截断，并附加 truncated 标志告知模型。
                        var truncatedData = TruncateRangeData(rangeData, MaxReadRangeRows);
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = truncatedData,
                            Suggestion = GenerateRangeSuggestion(rangeData),
                            Context = BuildExcelSnapshot(),
                        };

                    case "write_formula":
                        var cellAddress = GetArg<string>(args, "address");
                        var formula = GetArg<string>(args, "formula");
                        Logger.Instance.Info("ToolDispatcher", "write_formula: address=[" + (cellAddress ?? "null") + "], formula=[" + (formula ?? "null") + "]");
                        var result = _excel.WriteFormula(cellAddress, formula);
                        if (!result.Success)
                        {
                            result.Suggestion = GenerateFormulaSuggestion(formula, result.Error);
                        }
                        return result;

                    case "read_workbook":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = _excel.ReadWorkbook(),
                            Context = BuildExcelSnapshot(),
                        };

                    case "read_selection":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = _excel.GetSelection(),
                            Context = BuildExcelSnapshot(),
                        };

                    case "fill_formula_down":
                        var fromAddr = GetArg<string>(args, "from_address");
                        var rowCount = GetArg<int>(args, "row_count");
                        return ExecuteFillDown(fromAddr, rowCount);

                    case "replace_formula":
                        var rangeAddr = GetArg<string>(args, "range_address");
                        var find = GetArg<string>(args, "find");
                        var replace = GetArg<string>(args, "replace");
                        return ExecuteReplaceFormula(rangeAddr, find, replace);

                    case "clean_data":
                        var cleanRange = GetArg<string>(args, "range_address");
                        var ops = GetArgArray(args, "operations");
                        return ExecuteCleanData(cleanRange, ops);

                    case "create_chart":
                        var dataRange = GetArg<string>(args, "data_range");
                        var chartType = GetArg<string>(args, "chart_type") ?? "column";
                        var chartTitle = GetArg<string>(args, "title") ?? "";
                        var xLabel = GetArg<string>(args, "x_label") ?? "";
                        var yLabel = GetArg<string>(args, "y_label") ?? "";
                        return ExecuteCreateChart(dataRange, chartType, chartTitle, xLabel, yLabel);

                    case "create_pivot_table":
                        var sourceRange = GetArg<string>(args, "source_range");
                        var destSheet = GetArg<string>(args, "destination_sheet");
                        var pivotName = GetArg<string>(args, "pivot_table_name") ?? "PivotTable1";
                        var rowFields = GetArgArray(args, "row_fields");
                        var colFields = GetArgArray(args, "column_fields");
                        var valFields = GetArgArray(args, "value_fields");
                        var valFunc = GetArg<string>(args, "value_function") ?? "Sum";
                        return ExecuteCreatePivot(sourceRange, destSheet, pivotName, rowFields, colFields, valFields, valFunc);

                    case "execute_vba":
                        var code = GetArg<string>(args, "code");
                        return _excel.ExecuteVBA(code);

                    case "execute_python":
                        var pyCode = GetArg<string>(args, "code");
                        return _excel.ExecutePython(pyCode);

                    case "add_sheet":
                        return _excel.AddSheet(GetArg<string>(args, "name"));

                    case "delete_sheet":
                        return _excel.DeleteSheet(GetArg<string>(args, "name"));

                    case "rename_sheet":
                        return _excel.RenameSheet(
                            GetArg<string>(args, "old_name"),
                            GetArg<string>(args, "new_name"));

                    case "set_number_format":
                        return _excel.SetNumberFormat(
                            GetArg<string>(args, "address"),
                            GetArg<string>(args, "format"));

                    case "set_column_width":
                        return _excel.SetColumnWidth(
                            GetArg<string>(args, "address"),
                            GetArg<double>(args, "width"),
                            GetArg<bool>(args, "auto_fit"));

                    case "sort_data":
                        return _excel.SortData(
                            GetArg<string>(args, "range_address"),
                            GetArg<string>(args, "sort_column"),
                            GetArg<bool>(args, "descending"));

                    case "filter_data":
                        return _excel.FilterData(
                            GetArg<string>(args, "range_address"),
                            GetArg<int>(args, "column_index"),
                            GetArg<string>(args, "criteria"));

                    case "create_snapshot":
                        var snapshotId = _excel.CreateSnapshot();
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = !string.IsNullOrEmpty(snapshotId),
                            Data = new { snapshot_id = snapshotId },
                        };

                    case "rollback":
                        var sid = GetArg<string>(args, "snapshot_id");
                        var success = _excel.Rollback(sid);
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = success,
                        };

                    default:
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = false,
                            Error = $"未知工具: {toolName}",
                        };
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "Execute failed: " + toolName, ex);
                return new ToolResult
                {
                    Name = toolName,
                    Success = false,
                    Error = ex.Message,
                };
            }
        }

        /// <summary>
        /// 执行后日志记录（在 PythonSidecar.HandleToolCall 调用 Execute 后调用）
        /// </summary>
        public void LogResult(string toolName, ToolResult result)
        {
            if (result == null) return;
            if (result.Success)
            {
                Logger.Instance.Info("ToolDispatcher", $"Execute result: tool={toolName}, success=True");
            }
            else
            {
                Logger.Instance.Warning("ToolDispatcher", $"Execute result: tool={toolName}, success=False, error={result.Error}");
            }
        }

        /// <summary>
        /// 构建当前 Excel 上下文快照（附在 tool_result.context）
        /// 注意：只放元信息，避免二维数组（Object[,]）导致 System.Text.Json 序列化失败崩溃。
        /// 实际数据由 sidecar 通过 read_selection / read_range 工具按需读取。
        /// </summary>
        public object BuildExcelSnapshot()
        {
            try
            {
                object selectionInfo = null;
                try
                {
                    if (_excelApp != null)
                    {
                        var sel = _excelApp.Selection;
                        if (sel is Range range)
                        {
                            selectionInfo = new
                            {
                                address = range.Address,
                                worksheet = SafeGetWorksheetName(range),
                                rowCount = range.Rows.Count,
                                columnCount = range.Columns.Count,
                            };
                        }
                    }
                }
                catch { }

                return new
                {
                    workbook = _excel.ReadWorkbook(),
                    selection = selectionInfo ?? new { address = "" },
                    timestamp = DateTime.Now.ToString("o"),
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "BuildExcelSnapshot failed", ex);
                return new { error = ex.Message };
            }
        }

        private static string SafeGetWorksheetName(Range range)
        {
            try { return range.Worksheet.Name; }
            catch { return ""; }
        }

        /// <summary>
        /// 截断 RangeInfo 的二维数组到 MaxRows 行，避免 SDK 1MB 消息缓冲区限制。
        /// 返回一个动态对象（包含截断标志和原始行数），供序列化发送给 sidecar。
        /// </summary>
        private object TruncateRangeData(object rangeData, int maxRows)
        {
            if (rangeData == null) return null;

            try
            {
                var type = rangeData.GetType();

                // 通过反射读取字段（避免直接依赖 RangeInfo 类型）
                var addressProp = type.GetProperty("Address");
                var wsProp = type.GetProperty("WorksheetName");
                var rowProp = type.GetProperty("RowCount");
                var colProp = type.GetProperty("ColumnCount");
                var valuesProp = type.GetProperty("Values");
                var formulasProp = type.GetProperty("Formulas");
                var numFmtProp = type.GetProperty("NumberFormats");

                int originalRows = (int)(rowProp?.GetValue(rangeData) ?? 0);
                int cols = (int)(colProp?.GetValue(rangeData) ?? 0);

                // 不需要截断
                if (originalRows <= maxRows)
                {
                    return rangeData;
                }

                Logger.Instance.Info("ToolDispatcher",
                    $"TruncateRangeData: originalRows={originalRows}, maxRows={maxRows}, truncating");

                // 截断二维数组
                object[,] origValues = valuesProp?.GetValue(rangeData) as object[,];
                string[,] origFormulas = formulasProp?.GetValue(rangeData) as string[,];
                string[,] origNumFmt = numFmtProp?.GetValue(rangeData) as string[,];

                int rowStart = origValues?.GetLowerBound(0) ?? 0;
                int colStart = origValues?.GetLowerBound(1) ?? 0;

                var newValues = new object[maxRows, cols];
                var newFormulas = origFormulas != null ? new string[maxRows, cols] : null;
                var newNumFmt = origNumFmt != null ? new string[maxRows, cols] : null;

                for (int i = 0; i < maxRows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        if (origValues != null)
                            newValues[i, j] = origValues[rowStart + i, colStart + j];
                        if (newFormulas != null)
                            newFormulas[i, j] = origFormulas[rowStart + i, colStart + j];
                        if (newNumFmt != null)
                            newNumFmt[i, j] = origNumFmt[rowStart + i, colStart + j];
                    }
                }

                // 返回截断后的匿名对象（保持与 RangeInfo 相同的字段名）
                return new
                {
                    Address = addressProp?.GetValue(rangeData),
                    WorksheetName = wsProp?.GetValue(rangeData),
                    RowCount = maxRows,
                    ColumnCount = cols,
                    Values = newValues,
                    Formulas = newFormulas,
                    NumberFormats = newNumFmt,
                    // 截断标志：告知模型数据被截断，原始行数
                    truncated = true,
                    original_row_count = originalRows,
                    truncation_hint = $"数据超过 {maxRows} 行被截断。如需查看后续数据，请用更精确的地址（如 A201:B400）再次调用 read_range。",
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "TruncateRangeData failed, returning original", ex);
                return rangeData;
            }
        }

        /// <summary>
        /// read_range 返回数据的类型检测：检测到 text/mixed 时返回澄清提示。
        /// ★ 直接从 RangeInfo 对象读取，不序列化整个对象（避免 2D 数组序列化崩溃）。
        /// </summary>
        private string GenerateRangeSuggestion(object rangeData)
        {
            try
            {
                // 尝试读取 RangeInfo.Values 进行类型检测
                if (rangeData == null) return null;

                // 通过反射读取 Values 属性（避免直接依赖 RangeInfo 类型）
                var valuesProp = rangeData.GetType().GetProperty("Values");
                if (valuesProp == null) return null;
                var values = valuesProp.GetValue(rangeData) as object[,];
                if (values == null) return null;

                // 1-based 数组安全访问
                int rowStart = values.GetLowerBound(0);
                int rowEnd = values.GetUpperBound(0);
                int colStart = values.GetLowerBound(1);
                int colEnd = values.GetUpperBound(1);

                bool hasNum = false, hasText = false, hasDate = false;
                for (int i = rowStart; i <= rowEnd; i++)
                {
                    for (int j = colStart; j <= colEnd; j++)
                    {
                        var item = values[i, j];
                        if (item == null) continue;
                        if (item is double || item is int || item is long || item is decimal || item is float)
                            hasNum = true;
                        else if (item is DateTime)
                            hasDate = true;
                        else if (item is string)
                            hasText = true;
                        else if (item is bool)
                            continue; // bool 不计入 text/num
                        else
                            hasText = true; // 其他类型按文本处理
                    }
                }

                if (hasDate && !hasNum && !hasText)
                    return "该列是日期，求和无意义。建议用 COUNTA 计数，或按日期分组。";
                if (hasNum && hasText)
                    return "该列同时包含数字和文本，无法直接求和。建议：求和(忽略文本)、计数(全部 COUNTA)、计数(仅数字 COUNT)";
                if (hasText && !hasNum)
                    return "该列是文本，无法直接求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "GenerateRangeSuggestion failed", ex);
            }
            return null;
        }

        /// <summary>
        /// write_formula 失败时根据公式和错误生成澄清提示
        /// </summary>
        private string GenerateFormulaSuggestion(string formula, string error)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(error)) return null;

            if (formula.Contains("SUM") && (error.Contains("类型") || error.Contains("type") || error.Contains("mismatch")))
                return "目标区域包含文本，无法求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
            if (formula.Contains("AVERAGE") && (error.Contains("除数为零") || error.Contains("zero") || error.Contains("divisor")))
                return "目标区域没有数字，无法计算平均值。建议检查数据源。";
            if (error.Contains("范围") || error.Contains("range"))
                return "目标范围无效。请检查地址格式，例如 A1、B1:B10、A:A。";
            return null;
        }

        // ============= 从 Orchestrator 移植的 5 个私有方法 =============

        private ToolResult ExecuteFillDown(string fromAddress, int rowCount)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteFillDown: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "fill_formula_down", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var app = _excelApp;
                var fromRange = app.Range[fromAddress];
                var toRange = fromRange.Resize[rowCount, fromRange.Columns.Count];
                fromRange.AutoFill(toRange, XlAutoFillType.xlFillDefault);
                return new ToolResult { Name = "fill_formula_down", Success = true };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteFillDown failed", ex);
                return new ToolResult { Name = "fill_formula_down", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteReplaceFormula(string rangeAddress, string find, string replace)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteReplaceFormula: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "replace_formula", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var app = _excelApp;
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
                Logger.Instance.Error("ToolDispatcher", "ExecuteReplaceFormula failed", ex);
                return new ToolResult { Name = "replace_formula", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCleanData(string rangeAddress, string[] operations)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteCleanData: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "clean_data", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var cleaner = new DataCleaner(_excelApp);
                var results = new List<string>();

                foreach (var op in operations)
                {
                    switch (op.ToLower())
                    {
                        case "unify_date":
                            var r1 = cleaner.UnifyDateFormats(rangeAddress);
                            results.Add($"unify_date: {(r1.Success ? (r1.Data as dynamic)?.convertedCount + "个" : "失败-" + r1.Error)}");
                            break;
                        case "remove_duplicates":
                            var r2 = cleaner.RemoveDuplicates(rangeAddress);
                            results.Add($"remove_duplicates: {(r2.Success ? (r2.Data as dynamic)?.removedCount + "行" : "失败-" + r2.Error)}");
                            break;
                        case "highlight_missing":
                            var r3 = cleaner.HighlightMissingValues(rangeAddress);
                            results.Add($"highlight_missing: {(r3.Success ? (r3.Data as dynamic)?.highlightedCount + "个" : "失败-" + r3.Error)}");
                            break;
                        case "trim_spaces":
                            var r4 = cleaner.TrimSpaces(rangeAddress);
                            results.Add($"trim_spaces: {(r4.Success ? (r4.Data as dynamic)?.trimmedCount + "个" : "失败-" + r4.Error)}");
                            break;
                        case "text_to_number":
                            var r5 = cleaner.TextToNumber(rangeAddress);
                            results.Add($"text_to_number: {(r5.Success ? (r5.Data as dynamic)?.convertedCount + "个" : "失败-" + r5.Error)}");
                            break;
                    }
                }

                return new ToolResult { Name = "clean_data", Success = true, Data = new { operations = results } };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCleanData failed", ex);
                return new ToolResult { Name = "clean_data", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCreateChart(string dataRange, string chartType, string title, string xLabel, string yLabel)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteCreateChart: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "create_chart", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.CreateChart(dataRange, chartType, title, xLabel, yLabel);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCreateChart failed", ex);
                return new ToolResult { Name = "create_chart", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCreatePivot(
            string sourceRange,
            string destSheet,
            string pivotName,
            string[] rowFields,
            string[] colFields,
            string[] valFields,
            string valFunc)
        {
            if (_excelApp == null)
            {
                Logger.Instance.Warning("ToolDispatcher", "ExecuteCreatePivot: _excelApp is null, excelApp not available for this tool");
                return new ToolResult { Name = "create_pivot_table", Success = false, Error = "excelApp not available for this tool" };
            }
            try
            {
                var tool = new ChartTool(_excelApp);
                return tool.CreatePivotTable(sourceRange, destSheet, pivotName,
                    rowFields?.Length > 0 ? rowFields : null,
                    colFields?.Length > 0 ? colFields : null,
                    valFields?.Length > 0 ? valFields : null,
                    valFunc);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "ExecuteCreatePivot failed", ex);
                return new ToolResult { Name = "create_pivot_table", Success = false, Error = ex.Message };
            }
        }

        // ============= 参数提取（从 Orchestrator 移植）=============

        protected internal T GetArg<T>(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null) return default;
            try
            {
                var val = args[key];

                if (val is JsonElement je)
                {
                    if (typeof(T) == typeof(string))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetString();
                    }
                    if (typeof(T) == typeof(int))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetInt32();
                    }
                    if (typeof(T) == typeof(bool))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetBoolean();
                    }
                    if (typeof(T) == typeof(double))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetDouble();
                    }
                    return (T)Convert.ChangeType(je.GetRawText(), typeof(T));
                }

                return (T)Convert.ChangeType(val, typeof(T));
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", $"GetArg<{typeof(T).Name}>(\"{key}\") failed: " + ex.Message);
                return default;
            }
        }

        protected internal string[] GetArgArray(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key)) return new string[0];
            try
            {
                if (args[key] is JsonElement je)
                {
                    var list = new List<string>();
                    foreach (var item in je.EnumerateArray())
                    {
                        list.Add(item.GetString());
                    }
                    return list.ToArray();
                }
                if (args[key] is object[] arr)
                {
                    return Array.ConvertAll(arr, x => x?.ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Debug("ToolDispatcher", $"GetArgArray(\"{key}\") failed: " + ex.Message);
            }
            return new string[0];
        }
    }
}
