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
    }
}
