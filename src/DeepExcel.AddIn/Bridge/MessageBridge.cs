using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Sidecar;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Config;
using DeepExcel.AddIn.Security;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 桥接层 - 负责UI(WebView2)与C#核心之间的消息路由
    /// 协议：
    ///   收到: { type, payload }
    ///   发送: { type, payload }
    /// </summary>
    public class MessageBridge : IDisposable
    {
        private readonly Microsoft.Office.Interop.Excel.Application _excelApp;
        private readonly IExcelActions _excelActions;
        private readonly PythonSidecar _sidecar;
        private readonly ToolDispatcher _toolDispatcher;
        private Action<string> _sendToUi;
        private string _pendingClarifyQuestion;

        public MessageBridge(Microsoft.Office.Interop.Excel.Application excelApp, Control uiControl)
        {
            _excelApp = excelApp;

            // 初始化各层（保留原 ExcelActionsImpl 创建逻辑）
            var workbookAnalyzer = new WorkbookAnalyzer(excelApp);
            var rangeAnalyzer = new RangeAnalyzer();
            var snapshotManager = new SnapshotManager(excelApp);
            var vbaExecutor = new VBAExecutor(excelApp, snapshotManager);
            var pythonExecutor = new PythonExecutor(excelApp, snapshotManager);

            _excelActions = new ExcelActionsImpl(
                excelApp, workbookAnalyzer, rangeAnalyzer, vbaExecutor, pythonExecutor, snapshotManager);

            _toolDispatcher = new ToolDispatcher(_excelActions, _excelApp);

            // 创建 Python sidecar（替换 Orchestrator），在 SetSendToUi 中启动
            _sidecar = new PythonSidecar(_excelActions, _excelApp, uiControl);
            _sidecar.OnStreamDelta += OnStreamDelta;
            _sidecar.OnToolCall += OnToolCall;
            _sidecar.OnToolUse += OnToolUse;
            _sidecar.OnClarify += OnClarify;
            _sidecar.OnStreamEnd += OnStreamEndFromSidecar;
            _sidecar.OnError += OnSidecarError;
        }

        /// <summary>
        /// 设置UI消息发送回调
        /// </summary>
        public void SetSendToUi(Action<string> sendToUi)
        {
            _sendToUi = sendToUi;
            _sidecar.Start();

            // 启动后立即发送 config（API key + base_url + model）
            SendConfigToSidecar();
        }

        /// <summary>
        /// 从 AppConfig / SecurityManager 读取当前 provider 配置，发给 sidecar
        /// </summary>
        private void SendConfigToSidecar()
        {
            Logger.Instance.Info("MessageBridge", "SendConfigToSidecar called");
            try
            {
                var cfg = ConfigManager.Instance.Current;
                var providerKey = cfg.CurrentProvider;
                Logger.Instance.Info("MessageBridge", $"SendConfigToSidecar: providerKey={providerKey}, providers count={cfg.Providers?.Count ?? 0}");
                if (!cfg.Providers.ContainsKey(providerKey))
                {
                    Logger.Instance.Warning("MessageBridge", $"SendConfigToSidecar: provider '{providerKey}' not found in Providers, skipping");
                    return;
                }

                var provider = cfg.Providers[providerKey];
                var apiKey = SecurityManager.Instance.GetApiKey(providerKey);
                if (string.IsNullOrEmpty(apiKey))
                {
                    // SecurityManager 没存，尝试用 AppConfig 里的
                    apiKey = provider.ApiKey ?? "";
                }

                // DeepSeek 需要用 Anthropic 兼容端点
                var baseUrl = provider.BaseUrl ?? "https://api.anthropic.com";
                if (providerKey == "deepseek" && !baseUrl.EndsWith("/anthropic"))
                {
                    baseUrl = "https://api.deepseek.com/anthropic";
                }

                var model = cfg.CurrentModel;
                if (string.IsNullOrEmpty(model) && provider.Models != null && provider.Models.Length > 0)
                {
                    model = provider.Models[0];
                }

                Logger.Instance.Info("MessageBridge", $"Sending config to sidecar: provider={providerKey}, model={model}, baseUrl={baseUrl}, hasKey={!string.IsNullOrEmpty(apiKey)}");
                _sidecar.UpdateConfig(baseUrl, model, apiKey);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "SendConfigToSidecar failed", ex);
            }
        }

        /// <summary>
        /// 处理来自UI的消息
        /// </summary>
        public string HandleMessage(string json)
        {
            try
            {
                var msg = JsonSerializer.Deserialize<Message>(json);
                if (msg == null) return MakeError("Invalid message");

                return msg.Type switch
                {
                    "user_message" => HandleUserMessage(msg),
                    "cancel" => HandleCancel(),
                    "ping" => MakeResponse("pong", new { }),
                    "get_selection" => MakeResponse("selection", _excelActions.GetSelection()),
                    "read_range" => HandleReadRange(msg),
                    "read_workbook" => MakeResponse("workbook", _excelActions.ReadWorkbook()),
                    "execute_vba" => HandleExecuteVba(msg),
                    "execute_tool" => HandleExecuteTool(msg),
                    _ => MakeError($"Unknown message type: {msg.Type}")
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleMessage error", ex);
                return MakeError($"Handle error: {ex.Message}");
            }
        }

        private string HandleUserMessage(Message msg)
        {
            try
            {
                var content = msg.Payload?.GetProperty("content").GetString();

                // 如果有待回答的 clarify，把这条消息当 clarify_answer
                if (_pendingClarifyQuestion != null)
                {
                    _sidecar.SendClarifyAnswer(content);
                    _pendingClarifyQuestion = null;
                    return MakeResponse("ack", new { received = true, kind = "clarify_answer" });
                }

                // 正常用户消息：附带 Excel 上下文
                var context = BuildContext();
                var sessionId = Guid.NewGuid().ToString();
                _sidecar.SendUserMessage(content, sessionId, context);

                return MakeResponse("ack", new { received = true });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleUserMessage failed", ex);
                return MakeError($"User message error: {ex.Message}");
            }
        }

        private object BuildContext()
        {
            try
            {
                // 只放元信息，避免二维数组（Object[,]）导致 System.Text.Json 序列化失败
                // 实际数据由 sidecar 通过 read_selection / read_range 工具按需读取
                object selectionInfo = null;
                try
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
                catch { }

                return new
                {
                    workbook = _excelActions.ReadWorkbook(),
                    selection = selectionInfo ?? new { address = "" },
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "BuildContext failed", ex);
                return new { };
            }
        }

        private static string SafeGetWorksheetName(Range range)
        {
            try { return range.Worksheet.Name; }
            catch { return ""; }
        }

        private string HandleReadRange(Message msg)
        {
            var address = msg.Payload?.GetProperty("address").GetString();
            var result = _excelActions.ReadRange(address);
            return MakeResponse("range_data", result);
        }

        private string HandleExecuteVba(Message msg)
        {
            var code = msg.Payload?.GetProperty("code").GetString();
            var result = _excelActions.ExecuteVBA(code);
            return MakeResponse("vba_result", result);
        }

        private string HandleExecuteTool(Message msg)
        {
            var name = msg.Payload?.GetProperty("name").GetString();
            var argsEl = msg.Payload?.GetProperty("arguments");
            var args = new Dictionary<string, object>();
            if (argsEl.HasValue)
            {
                foreach (var prop in argsEl.Value.EnumerateObject())
                {
                    args[prop.Name] = prop.Value.Clone();
                }
            }
            var result = _toolDispatcher.Execute(name, args);
            return MakeResponse("tool_result", result);
        }

        // 回调方法 - 发往UI
        private void OnStreamDelta(string delta)
        {
            Logger.Instance.Debug("MessageBridge", "OnStreamDelta: " + (delta == null ? "null" : ("len=" + delta.Length)));
            SendToUi("stream_delta", new { delta });
        }

        private void OnToolCall(string callId, string name, Dictionary<string, object> args)
        {
            // 工具实际执行的回调（OnToolUse 已经通知过 UI，这里不再重复发送）
            Logger.Instance.Info("MessageBridge", $"OnToolCall: {name} (call_id={callId})");
        }

        private void OnToolUse(string name, Dictionary<string, object> args)
        {
            // SDK 通知即将调用工具，转发给 UI 显示进度（让用户看到 agent 正在做什么）
            // 去掉 mcp__excel__ 前缀，让显示更美观
            var displayName = name?.Replace("mcp__excel__", "") ?? name;
            SendToUi("tool_call", new { call_id = "", name = displayName, arguments = args });
        }

        private void OnClarify(string question, List<string> options)
        {
            _pendingClarifyQuestion = question;
            SendToUi("clarify", new { question, options });
        }

        private void OnStreamEndFromSidecar(int inputTokens, int outputTokens)
        {
            SendToUi("stream_end", new { input_tokens = inputTokens, output_tokens = outputTokens });
        }

        private void OnSidecarError(string error)
        {
            SendToUi("error", new { message = "Sidecar: " + error });
        }

        public void Cancel()
        {
            try
            {
                _sidecar?.SendCancel();
                Logger.Instance.Info("MessageBridge", "Cancel requested");
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "Cancel failed", ex);
            }
        }

        private string HandleCancel()
        {
            Cancel();
            NotifyStreamEnd();
            return MakeResponse("cancelled", new { });
        }

        public void NotifyStreamEnd()
        {
            SendToUi("stream_end", new { });
        }

        private void SendToUi(string type, object payload)
        {
            var json = JsonSerializer.Serialize(new { type, payload });
            _sendToUi?.Invoke(json);
        }

        private string MakeResponse(string type, object payload)
        {
            return JsonSerializer.Serialize(new { type, payload });
        }

        private string MakeError(string message)
        {
            return JsonSerializer.Serialize(new
            {
                type = "error",
                payload = new { message }
            });
        }

        public void Dispose()
        {
            try { _sidecar?.Stop(); }
            catch (Exception ex) { Logger.Instance.Error("MessageBridge", "Dispose failed", ex); }
        }
    }

    /// <summary>
    /// Excel操作实现 - 桥接层用
    /// </summary>
    internal class ExcelActionsImpl : IExcelActions
    {
        private readonly Microsoft.Office.Interop.Excel.Application _app;
        private readonly WorkbookAnalyzer _workbookAnalyzer;
        private readonly RangeAnalyzer _rangeAnalyzer;
        private readonly VBAExecutor _vbaExecutor;
        private readonly PythonExecutor _pythonExecutor;
        private readonly SnapshotManager _snapshotManager;

        public ExcelActionsImpl(
            Microsoft.Office.Interop.Excel.Application app,
            WorkbookAnalyzer workbookAnalyzer,
            RangeAnalyzer rangeAnalyzer,
            VBAExecutor vbaExecutor,
            PythonExecutor pythonExecutor,
            SnapshotManager snapshotManager)
        {
            _app = app;
            _workbookAnalyzer = workbookAnalyzer;
            _rangeAnalyzer = rangeAnalyzer;
            _vbaExecutor = vbaExecutor;
            _pythonExecutor = pythonExecutor;
            _snapshotManager = snapshotManager;
        }

        public object GetSelection()
        {
            try
            {
                var sel = _app.Selection;
                if (sel is Range range)
                {
                    return _rangeAnalyzer.Analyze(range);
                }
                return new { error = "No range selected" };
            }
            catch (Exception ex)
            {
                return new { error = ex.Message };
            }
        }

        public object ReadRange(string address)
        {
            try
            {
                var range = _app.Range[address];
                return _rangeAnalyzer.Analyze(range);
            }
            catch (Exception ex)
            {
                return new { error = ex.Message };
            }
        }

        public object ReadWorkbook()
        {
            return _workbookAnalyzer.Analyze();
        }

        public object ReadWorksheet(string name)
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return new { error = "No active workbook" };

                var sheet = wb.Worksheets[name] as Worksheet;
                if (sheet == null) return new { error = $"Sheet not found: {name}" };

                return new
                {
                    name = sheet.Name,
                    index = sheet.Index,
                    usedRange = sheet.UsedRange.Address
                };
            }
            catch (Exception ex)
            {
                return new { error = ex.Message };
            }
        }

        public ToolResult ExecuteVBA(string code, string macroName = null)
        {
            var result = _vbaExecutor.Execute(code, macroName ?? "DeepExcel_TempMacro");
            return new ToolResult
            {
                Name = result.Name,
                Success = result.Success,
                Data = result.Data,
                Error = result.Error,
            };
        }

        public ToolResult ExecutePython(string code)
        {
            var result = _pythonExecutor.Execute(code);
            return new ToolResult
            {
                Name = result.Name,
                Success = result.Success,
                Data = result.Data,
                Error = result.Error,
            };
        }

        public ToolResult WriteFormula(string address, string formula)
        {
            try
            {
                var range = _app.Range[address];
                range.Formula = formula;
                return new ToolResult { Name = "write_formula", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "write_formula", Success = false, Error = ex.Message };
            }
        }

        public ToolResult WriteValue(string address, object value)
        {
            try
            {
                var range = _app.Range[address];
                range.Value = value;
                return new ToolResult { Name = "write_value", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "write_value", Success = false, Error = ex.Message };
            }
        }

        public string CreateSnapshot()
        {
            return _snapshotManager.CreateSnapshot();
        }

        public bool Rollback(string snapshotId)
        {
            return _snapshotManager.Rollback(snapshotId);
        }

        // ============= Sheet 管理 =============

        public ToolResult AddSheet(string name)
        {
            try
            {
                var ws = (Worksheet)_app.Worksheets.Add();
                try { ws.Name = name; }
                catch
                {
                    // 名称冲突或非法字符，保留默认名
                }
                return new ToolResult { Name = "add_sheet", Success = true, Data = new { sheet_name = ws.Name } };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "add_sheet", Success = false, Error = ex.Message };
            }
        }

        public ToolResult DeleteSheet(string name)
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return new ToolResult { Name = "delete_sheet", Success = false, Error = "No active workbook" };
                if (wb.Worksheets.Count <= 1)
                    return new ToolResult { Name = "delete_sheet", Success = false, Error = "工作簿至少要保留一个工作表" };

                var sheet = wb.Worksheets[name] as Worksheet;
                if (sheet == null) return new ToolResult { Name = "delete_sheet", Success = false, Error = $"找不到工作表: {name}" };

                // 删除前禁用 Excel 的确认弹窗
                _app.DisplayAlerts = false;
                try { sheet.Delete(); }
                finally { _app.DisplayAlerts = true; }

                return new ToolResult { Name = "delete_sheet", Success = true };
            }
            catch (Exception ex)
            {
                try { _app.DisplayAlerts = true; } catch { }
                return new ToolResult { Name = "delete_sheet", Success = false, Error = ex.Message };
            }
        }

        public ToolResult RenameSheet(string oldName, string newName)
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return new ToolResult { Name = "rename_sheet", Success = false, Error = "No active workbook" };

                var sheet = wb.Worksheets[oldName] as Worksheet;
                if (sheet == null) return new ToolResult { Name = "rename_sheet", Success = false, Error = $"找不到工作表: {oldName}" };

                sheet.Name = newName;  // 名称冲突会自动抛 COMException
                return new ToolResult { Name = "rename_sheet", Success = true, Data = new { sheet_name = newName } };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "rename_sheet", Success = false, Error = ex.Message };
            }
        }

        // ============= 格式化 =============

        public ToolResult SetNumberFormat(string address, string format)
        {
            try
            {
                var range = _app.Range[address];
                range.NumberFormat = format;
                return new ToolResult { Name = "set_number_format", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "set_number_format", Success = false, Error = ex.Message };
            }
        }

        public ToolResult SetColumnWidth(string address, double width, bool autoFit)
        {
            try
            {
                var range = _app.Range[address];
                var columns = range.Columns;
                if (autoFit)
                {
                    columns.AutoFit();
                }
                else
                {
                    columns.ColumnWidth = width;
                }
                return new ToolResult { Name = "set_column_width", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "set_column_width", Success = false, Error = ex.Message };
            }
        }

        // ============= 数据操作 =============

        public ToolResult SortData(string rangeAddress, string sortColumn, bool descending)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                // sortColumn 是列字母（如 "A"）或列序号字符串（如 "1"）
                Range sortKey;
                if (int.TryParse(sortColumn, out int colIdx))
                {
                    sortKey = (Range)range.Columns[colIdx];
                }
                else
                {
                    // 字母列：在 range 内找对应列
                    var fullCol = _app.Range[sortColumn + "1"];
                    sortKey = (Range)_app.Intersect(range, fullCol.EntireColumn);
                }

                _app.DisplayAlerts = false;
                try
                {
                    range.Sort(
                        Key1: sortKey,
                        Order1: descending ? XlSortOrder.xlDescending : XlSortOrder.xlAscending,
                        Header: XlYesNoGuess.xlYes);
                }
                finally { _app.DisplayAlerts = true; }

                return new ToolResult { Name = "sort_data", Success = true };
            }
            catch (Exception ex)
            {
                try { _app.DisplayAlerts = true; } catch { }
                return new ToolResult { Name = "sort_data", Success = false, Error = ex.Message };
            }
        }

        public ToolResult FilterData(string rangeAddress, int columnIndex, string criteria)
        {
            try
            {
                var range = _app.Range[rangeAddress];
                _app.DisplayAlerts = false;
                try
                {
                    range.AutoFilter(
                        Field: columnIndex,
                        Criteria1: criteria);
                }
                finally { _app.DisplayAlerts = true; }

                return new ToolResult { Name = "filter_data", Success = true };
            }
            catch (Exception ex)
            {
                try { _app.DisplayAlerts = true; } catch { }
                return new ToolResult { Name = "filter_data", Success = false, Error = ex.Message };
            }
        }

        // ============= 单元格操作 =============

        public ToolResult MergeCells(string address)
        {
            try
            {
                var range = _app.Range[address];
                _app.DisplayAlerts = false;
                try { range.Merge(); }
                finally { _app.DisplayAlerts = true; }
                return new ToolResult { Name = "merge_cells", Success = true };
            }
            catch (Exception ex)
            {
                try { _app.DisplayAlerts = true; } catch { }
                return new ToolResult { Name = "merge_cells", Success = false, Error = ex.Message };
            }
        }

        public ToolResult UnmergeCells(string address)
        {
            try
            {
                var range = _app.Range[address];
                range.UnMerge();
                return new ToolResult { Name = "unmerge_cells", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "unmerge_cells", Success = false, Error = ex.Message };
            }
        }

        public ToolResult SetCellStyle(string address, string fontName, double? fontSize, bool? bold, bool? italic, string fontColor, string bgColor, string hAlign, string vAlign, bool? wrapText)
        {
            try
            {
                var range = _app.Range[address];
                var font = range.Font;
                if (!string.IsNullOrEmpty(fontName)) font.Name = fontName;
                if (fontSize.HasValue) font.Size = fontSize.Value;
                if (bold.HasValue) font.Bold = bold.Value;
                if (italic.HasValue) font.Italic = italic.Value;

                // 颜色：支持 RGB 十六进制（如 "#FF0000"）或颜色名
                // ★ 只有 ParseColor 返回 >= 0 才设置（-1 表示无效颜色，不修改）
                if (!string.IsNullOrEmpty(fontColor))
                {
                    int colorVal = ParseColor(fontColor);
                    Diagnostics.Logger.Instance.Info("ExcelActionsImpl", $"SetCellStyle fontColor='{fontColor}' -> colorVal={colorVal}");
                    if (colorVal >= 0) font.Color = colorVal;
                }
                if (!string.IsNullOrEmpty(bgColor))
                {
                    int colorVal = ParseColor(bgColor);
                    Diagnostics.Logger.Instance.Info("ExcelActionsImpl", $"SetCellStyle bgColor='{bgColor}' -> colorVal={colorVal}");
                    if (colorVal >= 0)
                    {
                        range.Interior.Color = colorVal;
                        range.Interior.Pattern = XlPattern.xlPatternSolid;
                    }
                }

                // 对齐
                if (!string.IsNullOrEmpty(hAlign))
                {
                    range.HorizontalAlignment = ParseHAlign(hAlign);
                }
                if (!string.IsNullOrEmpty(vAlign))
                {
                    range.VerticalAlignment = ParseVAlign(vAlign);
                }
                if (wrapText.HasValue)
                {
                    range.WrapText = wrapText.Value;
                }

                return new ToolResult { Name = "set_cell_style", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "set_cell_style", Success = false, Error = ex.Message };
            }
        }

        public ToolResult CopyRange(string sourceAddress, string destAddress)
        {
            try
            {
                var src = _app.Range[sourceAddress];
                var dst = _app.Range[destAddress];
                src.Copy(dst);
                return new ToolResult { Name = "copy_range", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "copy_range", Success = false, Error = ex.Message };
            }
        }

        public ToolResult ClearRange(string address, string clearType)
        {
            try
            {
                var range = _app.Range[address];
                switch ((clearType ?? "all").ToLower())
                {
                    case "contents":
                        range.ClearContents();
                        break;
                    case "formats":
                        range.ClearFormats();
                        break;
                    case "all":
                    default:
                        range.Clear();
                        break;
                }
                return new ToolResult { Name = "clear_range", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "clear_range", Success = false, Error = ex.Message };
            }
        }

        // ============= 行列操作 =============

        public ToolResult InsertRows(int row, int count)
        {
            try
            {
                var ws = (Worksheet)_app.ActiveSheet;
                var range = ws.Range["A" + row];
                for (int i = 0; i < count; i++)
                {
                    range.Insert(XlInsertShiftDirection.xlShiftDown);
                }
                return new ToolResult { Name = "insert_rows", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "insert_rows", Success = false, Error = ex.Message };
            }
        }

        public ToolResult DeleteRows(int row, int count)
        {
            try
            {
                var ws = (Worksheet)_app.ActiveSheet;
                for (int i = 0; i < count; i++)
                {
                    ((Range)ws.Rows[row]).Delete();
                }
                return new ToolResult { Name = "delete_rows", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "delete_rows", Success = false, Error = ex.Message };
            }
        }

        public ToolResult InsertColumns(int column, int count)
        {
            try
            {
                var ws = (Worksheet)_app.ActiveSheet;
                var colLetter = ColumnIndexToLetter(column);
                var range = ws.Range[colLetter + "1"];
                for (int i = 0; i < count; i++)
                {
                    range.Insert(XlInsertShiftDirection.xlShiftToRight);
                }
                return new ToolResult { Name = "insert_columns", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "insert_columns", Success = false, Error = ex.Message };
            }
        }

        public ToolResult DeleteColumns(int column, int count)
        {
            try
            {
                var ws = (Worksheet)_app.ActiveSheet;
                for (int i = 0; i < count; i++)
                {
                    ((Range)ws.Columns[column]).Delete();
                }
                return new ToolResult { Name = "delete_columns", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "delete_columns", Success = false, Error = ex.Message };
            }
        }

        // ============= 视图 =============

        public ToolResult FreezePanes(string address)
        {
            try
            {
                var ws = (Worksheet)_app.ActiveSheet;
                _app.ActiveWindow.Split = false;  // 先解除现有冻结
                var cell = ws.Range[address];
                cell.Activate();
                _app.ActiveWindow.FreezePanes = true;
                return new ToolResult { Name = "freeze_panes", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "freeze_panes", Success = false, Error = ex.Message };
            }
        }

        // ============= 高级 =============

        public ToolResult ApplyConditionalFormat(string address, string ruleType, object ruleArgs)
        {
            try
            {
                // ★ 打印参数实际值（之前 0x800A03EC 崩溃无法定位 address 内容）
                string argsDesc = "{}";
                if (ruleArgs is System.Collections.Generic.Dictionary<string, object> d)
                {
                    argsDesc = string.Join(",", d.Keys);
                }
                Diagnostics.Logger.Instance.Info("ExcelActionsImpl", $"ApplyConditionalFormat: address=[{address}], ruleType=[{ruleType}], ruleArgs keys=[{argsDesc}]");

                // ★ address 容错：模型可能传列名"销售额"而非 A1 地址
                Range range = null;
                try
                {
                    range = _app.Range[address];
                }
                catch (System.Runtime.InteropServices.COMException ce) when (ce.ErrorCode == unchecked((int)0x800A03EC))
                {
                    // 尝试作为列名解析：扫描表头行找到匹配列，转为列字母地址
                    string resolved = TryResolveColumnNameToRange(address);
                    if (resolved != null)
                    {
                        Diagnostics.Logger.Instance.Info("ExcelActionsImpl", $"ApplyConditionalFormat: 列名解析 [{address}] -> [{resolved}]");
                        range = _app.Range[resolved];
                    }
                    else
                    {
                        throw new System.Runtime.InteropServices.COMException($"address '{address}' 不是有效的 A1 地址，且未在表头中找到匹配列名。原始错误: {ce.Message}", ce.ErrorCode);
                    }
                }

                // 解析 ruleArgs（Dictionary<string, object>）
                var args = ruleArgs as System.Collections.Generic.Dictionary<string, object>;

                // ★ 用 dynamic 后期绑定避免 COM 类型注册问题（"没有注册类" 异常）
                // 某些 Excel 安装中 FormatConditions.AddColorScale/AddDatabar 返回的
                // ColorScale/Databar COM 对象需要 PIA 类型注册，强类型转换会失败。
                // 后期绑定通过 IDispatch 调用，不需要类型注册。
                dynamic fcs = range.FormatConditions;

                switch ((ruleType ?? "").ToLower())
                {
                    case "color_scale":
                        {
                            // 三色色阶（绿-黄-红）
                            dynamic cs = fcs.AddColorScale(3);
                            dynamic criteria = cs.ColorScaleCriteria;
                            criteria(1).Type = 1;  // xlConditionValueLowestValue
                            criteria(1).FormatColor.Color = 0x63BE7B;  // 绿（OLE 颜色 BGR）
                            criteria(2).Type = 5;  // xlConditionValuePercentile
                            criteria(2).FormatColor.Color = 0xFFEB84;  // 黄
                            criteria(3).Type = 2;  // xlConditionValueHighestValue
                            criteria(3).FormatColor.Color = 0xF8696B;  // 红
                        }
                        break;

                    case "data_bar":
                        {
                            dynamic db = fcs.AddDatabar();
                            db.BarColor.Color = 0x638EC6;  // 蓝色数据条
                        }
                        break;

                    case "highlight_rules":
                        {
                            // 高亮重复值
                            fcs.AddUniqueValues();
                            // 取最后添加的规则，设置高亮色
                            dynamic lastRule = fcs(fcs.Count);
                            lastRule.Interior.Color = 0xFFC7CE;  // 浅红
                        }
                        break;

                    case "cell_value":
                        {
                            // 基于单元格值（如 >100 高亮）
                            string op = args != null && args.ContainsKey("operator") ? args["operator"]?.ToString() : "greater";
                            object val1 = args != null && args.ContainsKey("value") ? args["value"] : 0;
                            int xlOp = ParseConditionalOperatorInt(op);
                            fcs.Add(
                                Type: 1,  // xlCellValue
                                Operator: xlOp,
                                Formula1: val1?.ToString());
                            dynamic lastRule = fcs(fcs.Count);
                            lastRule.Interior.Color = 0xFFC7CE;
                        }
                        break;

                    case "top10":
                        {
                            // 前 N 项高亮
                            int n = 10;
                            if (args != null && args.ContainsKey("n"))
                            {
                                int.TryParse(args["n"]?.ToString(), out n);
                            }
                            fcs.AddTop10(Rank: n, Percent: false);
                            dynamic lastRule = fcs(fcs.Count);
                            lastRule.Interior.Color = 0xFFEB9C;  // 浅黄
                        }
                        break;

                    case "above_average":
                        {
                            // 高于平均值的单元格高亮
                            fcs.AddAboveAverage();
                            dynamic lastRule = fcs(fcs.Count);
                            lastRule.Interior.Color = 0xC6EFCE;  // 浅绿
                        }
                        break;

                    default:
                        return new ToolResult { Name = "apply_conditional_format", Success = false, Error = $"未知规则类型: {ruleType}（支持 color_scale/data_bar/highlight_rules/cell_value/top10/above_average）" };
                }

                return new ToolResult { Name = "apply_conditional_format", Success = true };
            }
            catch (Exception ex)
            {
                Diagnostics.Logger.Instance.Error("ExcelActionsImpl", "ApplyConditionalFormat failed: " + ex.ToString());
                return new ToolResult { Name = "apply_conditional_format", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 尝试把列名（如"销售额"）解析为整列 A1 地址（如"E:E"）。
        /// 扫描活动 sheet 的第 1 行（表头），找到文本匹配的列，返回列字母:列字母。
        /// 找不到返回 null。用于 apply_conditional_format 等工具的 address 容错。
        /// </summary>
        private string TryResolveColumnNameToRange(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName)) return null;
            // 排除明显是 A1 地址的输入（含数字或冒号）
            string trimmed = columnName.Trim();
            // 去除可能的"列"后缀（如"销售额列"→"销售额"）
            string name = trimmed;
            if (name.EndsWith("列")) name = name.Substring(0, name.Length - 1);

            try
            {
                var ws = (Worksheet)_app.ActiveSheet;
                var usedRange = ws.UsedRange;
                if (usedRange == null) return null;
                int colCount = usedRange.Columns.Count;
                // 只扫第 1 行
                Range headerRow = (Range)usedRange.Rows[1];
                for (int c = 1; c <= colCount; c++)
                {
                    Range cellVal = (Range)headerRow[1, c];
                    string text = "";
                    try { text = cellVal?.Text?.ToString() ?? ""; } catch { }
                    if (string.IsNullOrEmpty(text))
                    {
                        try { text = cellVal?.Value2?.ToString() ?? ""; } catch { }
                    }
                    if (!string.IsNullOrEmpty(text) &&
                        (text.Trim() == name.Trim() ||
                         text.Trim().IndexOf(name.Trim(), StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        string colLetter = ColumnIndexToLetter(c);
                        return $"{colLetter}:{colLetter}";
                    }
                }
            }
            catch (Exception ex)
            {
                Diagnostics.Logger.Instance.Warning("ExcelActionsImpl", "TryResolveColumnNameToRange failed: " + ex.Message);
            }
            return null;
        }

        /// <summary>
        /// 解析条件格式操作符为 int（避免 COM 枚举类型注册问题）
        /// xlBetween=1, xlNotBetween=2, xlEqual=3, xlNotEqual=4, xlGreater=5, xlLess=6, xlGreaterEqual=7, xlLessEqual=8
        /// </summary>
        private static int ParseConditionalOperatorInt(string s)
        {
            switch ((s ?? "").ToLower())
            {
                case "between": return 1;
                case "not_between": return 2;
                case "equal": case "=": return 3;
                case "not_equal": case "!=": return 4;
                case "greater": case ">": return 5;
                case "less": case "<": return 6;
                case "greater_equal": case ">=": return 7;
                case "less_equal": case "<=": return 8;
                default: return 5;
            }
        }

        public ToolResult WriteTable(string address, string tableName)
        {
            try
            {
                var range = _app.Range[address];
                var wb = _app.ActiveWorkbook;
                if (wb == null) return new ToolResult { Name = "write_table", Success = false, Error = "No active workbook" };

                var ws = (Worksheet)_app.ActiveSheet;
                var table = ws.ListObjects.Add(
                    SourceType: XlListObjectSourceType.xlSrcRange,
                    Source: range,
                    XlListObjectHasHeaders: XlYesNoGuess.xlYes);

                if (!string.IsNullOrEmpty(tableName))
                {
                    try { table.Name = tableName; }
                    catch { /* 名称冲突，用默认名 */ }
                }

                return new ToolResult { Name = "write_table", Success = true, Data = new { table_name = table.Name } };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "write_table", Success = false, Error = ex.Message };
            }
        }

        // ============= 辅助方法 =============

        /// <summary>
        /// 解析颜色字符串：支持 "#RRGGBB" / "RRGGBB" / 颜色名（中英文）
        /// 返回 -1 表示无效颜色（调用方不应设置颜色）
        /// Excel OLE 颜色格式：0x00BBGGRR（BGR 顺序，不是 RGB）
        /// </summary>
        private static int ParseColor(string colorStr)
        {
            if (string.IsNullOrEmpty(colorStr)) return -1;

            string s = colorStr.Trim().ToLower();

            // "none"/"default"/"auto"/"无" 表示不设置
            if (s == "none" || s == "default" || s == "auto" || s == "无" || s == "")
                return -1;

            // 英文颜色名映射（返回 Excel OLE 颜色 0x00BBGGRR）
            switch (s)
            {
                case "red": return RgbToOle(0xFF, 0x00, 0x00);
                case "green": return RgbToOle(0x00, 0x80, 0x00);
                case "blue": return RgbToOle(0x00, 0x00, 0xFF);
                case "yellow": return RgbToOle(0xFF, 0xFF, 0x00);
                case "white": return RgbToOle(0xFF, 0xFF, 0xFF);
                case "black": return RgbToOle(0x00, 0x00, 0x00);
                case "orange": return RgbToOle(0xFF, 0xA5, 0x00);
                case "gray":
                case "grey": return RgbToOle(0x80, 0x80, 0x80);
                case "pink": return RgbToOle(0xFF, 0xC0, 0xCB);
                case "purple": return RgbToOle(0x80, 0x00, 0x80);
                case "cyan": return RgbToOle(0x00, 0xFF, 0xFF);
                case "magenta": return RgbToOle(0xFF, 0x00, 0xFF);
                // 浅色系列（英文）
                case "light blue":
                case "lightblue": return RgbToOle(0xCC, 0xCC, 0xFF);
                case "light green":
                case "lightgreen": return RgbToOle(0xCC, 0xFF, 0xCC);
                case "light red":
                case "lightred": return RgbToOle(0xFF, 0xCC, 0xCC);
                case "light yellow":
                case "lightyellow": return RgbToOle(0xFF, 0xFF, 0xCC);
                case "light gray":
                case "lightgray":
                case "light grey":
                case "lightgrey": return RgbToOle(0xD3, 0xD3, 0xD3);
            }

            // 中文颜色名映射（含"浅色"前缀）
            switch (s)
            {
                case "红色": return RgbToOle(0xFF, 0x00, 0x00);
                case "绿色": return RgbToOle(0x00, 0x80, 0x00);
                case "蓝色": return RgbToOle(0x00, 0x00, 0xFF);
                case "黄色": return RgbToOle(0xFF, 0xFF, 0x00);
                case "白色": return RgbToOle(0xFF, 0xFF, 0xFF);
                case "黑色": return RgbToOle(0x00, 0x00, 0x00);
                case "橙色": return RgbToOle(0xFF, 0xA5, 0x00);
                case "灰色": return RgbToOle(0x80, 0x80, 0x80);
                case "粉色":
                case "粉红色": return RgbToOle(0xFF, 0xC0, 0xCB);
                case "紫色": return RgbToOle(0x80, 0x00, 0x80);
                case "青色": return RgbToOle(0x00, 0xFF, 0xFF);
                // 浅色系列
                case "浅红":
                case "浅红色": return RgbToOle(0xFF, 0xCC, 0xCC);
                case "浅绿":
                case "浅绿色": return RgbToOle(0xCC, 0xFF, 0xCC);
                case "浅蓝":
                case "浅蓝色": return RgbToOle(0xCC, 0xCC, 0xFF);
                case "浅黄":
                case "浅黄色": return RgbToOle(0xFF, 0xFF, 0xCC);
                case "浅灰":
                case "浅灰色": return RgbToOle(0xD3, 0xD3, 0xD3);
                case "浅紫":
                case "浅紫色": return RgbToOle(0xE6, 0xE6, 0xFA);
                // 深色系列
                case "深红":
                case "深红色": return RgbToOle(0x8B, 0x00, 0x00);
                case "深绿":
                case "深绿色": return RgbToOle(0x00, 0x64, 0x00);
                case "深蓝":
                case "深蓝色": return RgbToOle(0x00, 0x00, 0x8B);
                // 天蓝、海军蓝等
                case "天蓝":
                case "天蓝色": return RgbToOle(0x87, 0xCE, 0xEB);
                case "海军蓝": return RgbToOle(0x00, 0x00, 0x80);
            }

            // 十六进制 #RRGGBB 或 RRGGBB
            var hex = colorStr.TrimStart('#');
            if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int rgb))
            {
                int r = (rgb >> 16) & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = rgb & 0xFF;
                return RgbToOle(r, g, b);
            }

            // ★ 无效颜色返回 -1，而不是 0（0 是黑色，会误设背景为黑）
            return -1;
        }

        /// <summary>
        /// RGB 转 Excel OLE 颜色（0x00BBGGRR）
        /// </summary>
        private static int RgbToOle(int r, int g, int b)
        {
            return (b << 16) | (g << 8) | r;
        }

        private static XlHAlign ParseHAlign(string s)
        {
            switch ((s ?? "").ToLower())
            {
                case "left": return XlHAlign.xlHAlignLeft;
                case "center": return XlHAlign.xlHAlignCenter;
                case "right": return XlHAlign.xlHAlignRight;
                case "justify": return XlHAlign.xlHAlignJustify;
                default: return XlHAlign.xlHAlignGeneral;
            }
        }

        private static XlVAlign ParseVAlign(string s)
        {
            switch ((s ?? "").ToLower())
            {
                case "top": return XlVAlign.xlVAlignTop;
                case "center": return XlVAlign.xlVAlignCenter;
                case "bottom": return XlVAlign.xlVAlignBottom;
                case "justify": return XlVAlign.xlVAlignJustify;
                default: return XlVAlign.xlVAlignCenter;
            }
        }

        private static XlFormatConditionOperator ParseConditionalOperator(string s)
        {
            switch ((s ?? "").ToLower())
            {
                case "between": return XlFormatConditionOperator.xlBetween;
                case "not_between": return XlFormatConditionOperator.xlNotBetween;
                case "equal": case "=": return XlFormatConditionOperator.xlEqual;
                case "not_equal": case "!=": return XlFormatConditionOperator.xlNotEqual;
                case "greater": case ">": return XlFormatConditionOperator.xlGreater;
                case "less": case "<": return XlFormatConditionOperator.xlLess;
                case "greater_equal": case ">=": return XlFormatConditionOperator.xlGreaterEqual;
                case "less_equal": case "<=": return XlFormatConditionOperator.xlLessEqual;
                default: return XlFormatConditionOperator.xlGreater;
            }
        }

        /// <summary>
        /// 列序号 -> 列字母（1 -> A, 27 -> AA）
        /// </summary>
        private static string ColumnIndexToLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + (col % 26)) + result;
                col /= 26;
            }
            return result;
        }
    }
}
