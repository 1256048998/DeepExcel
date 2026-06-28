using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// Python sidecar 子进程包装：JSON Lines over stdin/stdout 通信
    /// </summary>
    public class PythonSidecar : IDisposable
    {
        private Process _process;
        private readonly IExcelActions _excel;
        private readonly Control _uiControl;
        private readonly ToolDispatcher _dispatcher;
        private readonly object _writeLock = new object();

        /// <summary>
        /// 共享的 JSON 序列化选项：注册 2D 数组转换器，避免 System.Text.Json
        /// 遇到 object[,] / string[,] 时抛 NotSupportedException 导致 Excel 崩溃。
        /// </summary>
        private static readonly JsonSerializerOptions _jsonOptions = BuildJsonOptions();

        private static JsonSerializerOptions BuildJsonOptions()
        {
            var opts = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
            };
            opts.Converters.Add(new Object2DArrayConverter());
            opts.Converters.Add(new String2DArrayConverter());
            return opts;
        }

        // 事件：UI 线程触发
        public event Action<string> OnStreamDelta;
        public event Action<string, string, Dictionary<string, object>> OnToolCall;
        public event Action<string, Dictionary<string, object>> OnToolUse;
        public event Action<string, List<string>> OnClarify;
        public event Action<int, int> OnStreamEnd;
        public event Action<string> OnError;

        public PythonSidecar(IExcelActions excel, Microsoft.Office.Interop.Excel.Application excelApp, Control uiControl)
        {
            _excel = excel;
            _uiControl = uiControl;
            _dispatcher = new ToolDispatcher(excel, excelApp);
        }

        /// <summary>
        /// 启动 Python sidecar 子进程
        /// </summary>
        public void Start()
        {
            var psi = new ProcessStartInfo
            {
                FileName = GetPythonPath(),
                Arguments = $"\"{GetSidecarPath()}\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };

            // 透传环境变量（ANTHROPIC_API_KEY 等）
            // psi.EnvironmentVariables 已自动继承当前进程

            _process = Process.Start(psi);
            _process.OutputDataReceived += OnStdoutLine;
            _process.ErrorDataReceived += OnStderrLine;
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();

            Logger.Instance.Info("PythonSidecar", "Started, pid=" + _process.Id);
        }

        /// <summary>
        /// 停止 sidecar 子进程
        /// </summary>
        public void Stop()
        {
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    try { WriteLine(@"{""type"":""cancel""}"); }
                    catch (Exception wex) { Logger.Instance.Warning("PythonSidecar", "Failed to send cancel on stop", wex); }
                    _process.WaitForExit(2000);
                    if (!_process.HasExited) _process.Kill();
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("PythonSidecar", "Stop failed", ex);
            }
        }

        public void Dispose() => Stop();

        // ============= 发送消息（C# → Python）=============

        public void SendUserMessage(string text, string sessionId, object context)
        {
            var msg = new { type = SidecarProtocol.TypeUserMessage, text, session_id = sessionId, context };
            WriteLine(JsonSerializer.Serialize(msg, _jsonOptions));
        }

        public void SendCancel() => WriteLine(@"{""type"":""cancel""}");

        public void UpdateConfig(string baseUrl, string model, string apiKey)
        {
            var msg = new { type = SidecarProtocol.TypeConfig, base_url = baseUrl, model, api_key = apiKey };
            WriteLine(JsonSerializer.Serialize(msg, _jsonOptions));
        }

        public void SendClarifyAnswer(string answer)
        {
            var msg = new { type = SidecarProtocol.TypeClarifyAnswer, answer };
            WriteLine(JsonSerializer.Serialize(msg, _jsonOptions));
        }

        public void SendToolResult(string callId, bool success, object data, string error, string suggestion, object context)
        {
            var msg = new
            {
                type = SidecarProtocol.TypeToolResult,
                call_id = callId,
                success,
                data,
                error,
                suggestion,
                context,
            };
            string json;
            try
            {
                json = JsonSerializer.Serialize(msg, _jsonOptions);
            }
            catch (Exception ex)
            {
                // 序列化失败（通常是 data 含二维数组 Object[,] 等 System.Text.Json 不支持的类型）
                // 回退为安全表示，避免 Excel 崩溃。
                // ★ 必须设 success=false 并提供明确 error，否则模型看到 success=true 但 data 是 error 会困惑，
                // 导致无限重试同一工具（实测 DeepSeek 会一直循环 read_range）。
                Logger.Instance.Error("PythonSidecar", "SendToolResult serialize failed (falling back to safe)", ex);
                var safeMsg = new
                {
                    type = SidecarProtocol.TypeToolResult,
                    call_id = callId,
                    success = false,
                    data = (object)new { },
                    error = "工具执行成功但结果数据无法序列化: " + ex.Message,
                    suggestion,
                    context = (object)new { },
                };
                json = JsonSerializer.Serialize(safeMsg, _jsonOptions);
            }
            WriteLine(json);
        }

        // ============= 接收消息（Python → C#）=============

        private void OnStdoutLine(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Data)) return;

            // ★ 诊断日志：记录所有从 Python sidecar 收到的 stdout 消息类型
            // 用于排查"C# 没收到 stream_delta/stream_end"的问题
            string msgType = "?";
            try
            {
                using var peekDoc = JsonDocument.Parse(e.Data);
                msgType = peekDoc.RootElement.GetProperty("type").GetString() ?? "?";
            }
            catch { /* 解析失败的日志在下面 catch 块里打 */ }
            Logger.Instance.Info("PythonSidecar", $"OnStdoutLine: type={msgType}, len={e.Data.Length}");

            try
            {
                using var doc = JsonDocument.Parse(e.Data);
                var root = doc.RootElement;
                var type = root.GetProperty("type").GetString();

                switch (type)
                {
                    case SidecarProtocol.TypeStreamDelta:
                        var text = root.GetProperty("text").GetString();
                        Logger.Instance.Info("PythonSidecar", $"OnStreamDelta: text_len={text?.Length ?? 0}");
                        SafeBeginInvoke(() => OnStreamDelta?.Invoke(text));
                        break;

                    case SidecarProtocol.TypeToolCall:
                        HandleToolCall(root);
                        break;

                    case SidecarProtocol.TypeToolUse:
                        // SDK 通知即将调用工具（不同于 tool_call，tool_call 是实际请求 C# 执行）
                        var useToolName = root.GetProperty("tool").GetString();
                        var useArgs = new Dictionary<string, object>();
                        if (root.TryGetProperty("args", out var useArgsEl))
                        {
                            foreach (var prop in useArgsEl.EnumerateObject())
                            {
                                useArgs[prop.Name] = prop.Value.Clone();
                            }
                        }
                        Logger.Instance.Info("PythonSidecar", $"OnToolUse: tool={useToolName}");
                        SafeBeginInvoke(() => OnToolUse?.Invoke(useToolName, useArgs));
                        break;

                    case SidecarProtocol.TypeClarify:
                        var q = root.GetProperty("question").GetString();
                        var opts = root.GetProperty("options").EnumerateArray()
                            .Select(x => x.GetString()).ToList();
                        SafeBeginInvoke(() => OnClarify?.Invoke(q, opts));
                        break;

                    case SidecarProtocol.TypeStreamEnd:
                        var inTok = root.GetProperty("input_tokens").GetInt32();
                        var outTok = root.GetProperty("output_tokens").GetInt32();
                        Logger.Instance.Info("PythonSidecar", $"OnStreamEnd: in={inTok}, out={outTok}");
                        SafeBeginInvoke(() => OnStreamEnd?.Invoke(inTok, outTok));
                        break;

                    default:
                        Logger.Instance.Warning("PythonSidecar", $"OnStdoutLine: UNKNOWN type={type}");
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("PythonSidecar", "Parse stdout failed: " + e.Data, ex);
            }
        }

        private void OnStderrLine(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Data)) return;
            // sidecar 的 stderr 是诊断日志（如 [ipc] route_message、[sidecar] client.query() done），
            // 不是真正的错误，不应转发给前端显示给用户。
            // 只记录到日志文件供开发诊断用。
            Logger.Instance.Info("PythonSidecar", "stderr: " + e.Data);
        }

        /// <summary>
        /// 安全地在 UI 线程执行委托：如果控件句柄未创建（如 sidecar 在面板打开前就发送消息），
        /// 直接跳过 UI 通知，避免 InvalidOperationException 导致 Excel 崩溃。
        /// </summary>
        private void SafeBeginInvoke(Action action)
        {
            try
            {
                if (_uiControl != null && _uiControl.IsHandleCreated && !_uiControl.IsDisposed)
                {
                    _uiControl.BeginInvoke(action);
                }
                else
                {
                    // ★ 诊断日志：记录被丢弃的事件（之前是静默丢弃，导致 stream_end 收到但 UI 不更新）
                    Logger.Instance.Warning("PythonSidecar",
                        $"SafeBeginInvoke DROPPED: uiControlNull={_uiControl == null}, " +
                        $"IsHandleCreated={_uiControl?.IsHandleCreated}, IsDisposed={_uiControl?.IsDisposed}");
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("PythonSidecar", "SafeBeginInvoke failed: " + ex.Message);
            }
        }

        private async void HandleToolCall(JsonElement msg)
        {
            var callId = msg.GetProperty("call_id").GetString();
            var toolName = msg.GetProperty("tool").GetString();
            var args = new Dictionary<string, object>();
            if (msg.TryGetProperty("args", out var argsEl))
            {
                foreach (var prop in argsEl.EnumerateObject())
                {
                    args[prop.Name] = prop.Value.Clone();
                }
            }

            Logger.Instance.Info("PythonSidecar", $"HandleToolCall START: tool={toolName}, call_id={callId}");

            // 通知 UI（fire-and-forget）
            SafeBeginInvoke(() => OnToolCall?.Invoke(callId, toolName, args));

            // ★ 关键：切回 UI 线程执行 Excel COM 操作（STA 要求）
            // BuildExcelSnapshot 也需要 COM 访问（ReadWorkbook + GetSelection），
            // 所以一并放在 Invoke 块内执行。
            ToolResult result = null;
            object context = null;
            try
            {
                if (_uiControl.InvokeRequired)
                {
                    Logger.Instance.Info("PythonSidecar", $"HandleToolCall: InvokeRequired=true, switching to UI thread");
                    var tuple = ((ToolResult, object))_uiControl.Invoke(new Func<(ToolResult, object)>(() =>
                    {
                        var r = _dispatcher.Execute(toolName, args);
                        var ctx = _dispatcher.BuildExcelSnapshot();
                        return (r, ctx);
                    }));
                    result = tuple.Item1;
                    context = tuple.Item2;
                }
                else
                {
                    Logger.Instance.Info("PythonSidecar", $"HandleToolCall: InvokeRequired=false, executing directly");
                    result = _dispatcher.Execute(toolName, args);
                    context = _dispatcher.BuildExcelSnapshot();
                }
                _dispatcher.LogResult(toolName, result);
                Logger.Instance.Info("PythonSidecar", $"HandleToolCall: execute done, success={result?.Success}, context_type={context?.GetType().Name}");
            }
            catch (Exception ex)
            {
                result = new ToolResult { Name = toolName, Success = false, Error = ex.Message };
                Logger.Instance.Error("PythonSidecar", "HandleToolCall failed: " + toolName, ex);
            }

            // 回写 tool_result
            Logger.Instance.Info("PythonSidecar", $"HandleToolCall: sending tool_result, call_id={callId}");
            SendToolResult(
                callId: callId,
                success: result.Success,
                data: result.Data,
                error: result.Error,
                suggestion: result.Suggestion,
                context: context);
            Logger.Instance.Info("PythonSidecar", $"HandleToolCall END: tool={toolName}");
        }

        // ============= 工具方法 =============

        private void WriteLine(string json)
        {
            lock (_writeLock)
            {
                if (_process == null || _process.HasExited)
                {
                    Logger.Instance.Warning("PythonSidecar", $"WriteLine SKIPPED: process exited (json_len={json.Length})");
                    return;
                }
                try
                {
                    _process.StandardInput.WriteLine(json);
                    _process.StandardInput.Flush();
                }
                catch (Exception ex)
                {
                    Logger.Instance.Error("PythonSidecar", "WriteLine failed", ex);
                }
            }
        }

        /// <summary>
        /// 查找 Python 解释器路径
        /// 优先级：1) 内嵌 python/python.exe  2) 系统 PATH 中的 python
        /// </summary>
        public static string GetPythonPath()
        {
            // 1. 内嵌 Python（打包后）
            var addInDir = Path.GetDirectoryName(typeof(PythonSidecar).Assembly.Location);
            var embeddedPy = Path.Combine(addInDir, "python", "python.exe");
            if (File.Exists(embeddedPy)) return embeddedPy;

            // 2. 系统 PATH（开发期）
            var systemPy = Environment.GetEnvironmentVariable("DEEPEXCEL_PYTHON_PATH");
            if (!string.IsNullOrEmpty(systemPy) && File.Exists(systemPy)) return systemPy;

            // 3. 默认 python（在 PATH 中）
            return "python";
        }

        /// <summary>
        /// sidecar.py 路径
        /// </summary>
        public static string GetSidecarPath()
        {
            var addInDir = Path.GetDirectoryName(typeof(PythonSidecar).Assembly.Location);
            return Path.Combine(addInDir, "sidecar", "sidecar.py");
        }
    }
}
