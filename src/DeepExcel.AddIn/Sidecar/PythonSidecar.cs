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
        /// ★ 暴露 ToolDispatcher 供 MessageBridge 注入 Attachments 引用。
        /// read_attachment 工具需要访问 session 的附件映射来查找文件路径。
        /// </summary>
        public ToolDispatcher Dispatcher => _dispatcher;
        // ★ 防止 OnProcessExited 重复触发（sidecar 正常 Stop 时也会触发 Exited 事件）
        private int _exited = 0;
        // ★ 标记是否主动 Stop，主动 Stop 时不发 error 通知前端
        private bool _stopping = false;

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

        // 事件：UI 线程触发，sender = 触发事件的 sidecar 实例（用于多会话精准路由）
        public event Action<PythonSidecar, string> OnStreamDelta;
        public event Action<PythonSidecar, string, string, Dictionary<string, object>> OnToolCall;
        public event Action<PythonSidecar, string, Dictionary<string, object>> OnToolUse;
        public event Action<PythonSidecar, string, List<string>> OnClarify;
        public event Action<PythonSidecar, int, int> OnStreamEnd;
        public event Action<PythonSidecar, string> OnError;
        // ★ AI Native 权限确认：PreToolUse hook 请求用户确认高风险工具
        public event Action<PythonSidecar, string, string, Dictionary<string, object>> OnPermissionRequest;

        public PythonSidecar(IExcelActions excel, Microsoft.Office.Interop.Excel.Application excelApp, Control uiControl, DeepExcel.AddIn.Security.SecurityGateway securityGateway = null)
        {
            _excel = excel;
            _uiControl = uiControl;
            // ★ AI Native 改造后：ToolDispatcher 不再需要 SecurityGateway（权限确认由 PreToolUse hook 处理）
            _dispatcher = new ToolDispatcher(excel, excelApp);
        }

        /// <summary>
        /// 启动 Python sidecar 子进程。
        /// ★ 幂等：如果进程还在运行，直接返回；如果已退出，清理后重启。
        /// </summary>
        public void Start()
        {
            // ★ 幂等检查：进程还活着就不重启
            if (_process != null && !_process.HasExited)
            {
                Logger.Instance.Info("PythonSidecar", "Start: process already running, skip");
                return;
            }

            // 清理旧进程引用，重置退出标志
            if (_process != null)
            {
                try { _process.Dispose(); } catch { }
                _process = null;
            }
            _exited = 0;
            _stopping = false;

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

            // ★ 强制 Python 全局 UTF-8 模式：系统语言为英文时 ANSI 代码页 cp1252 不支持中文，
            // 导致 Python 子进程的 stdin/stdout/管道默认用 cp1252，中文变 "?"
            // PYTHONUTF8=1 让 Python 所有 IO 默认 UTF-8，配合 sidecar.py 的 reconfigure 双保险
            psi.EnvironmentVariables["PYTHONUTF8"] = "1";

            _process = Process.Start(psi);
            _process.OutputDataReceived += OnStdoutLine;
            _process.ErrorDataReceived += OnStderrLine;
            // ★ C-4 修复：订阅 Exited 事件，sidecar 进程崩溃时主动发 stream_end + error 给 UI，
            // 否则前端永久转圈 5 分钟（依赖 IsBusy 超时兜底太慢）
            _process.EnableRaisingEvents = true;
            _process.Exited += OnProcessExited;
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();

            Logger.Instance.Info("PythonSidecar", "Started, pid=" + _process.Id);
        }

        /// <summary>
        /// ★ C-4 修复：sidecar 进程退出/崩溃时触发，发送 error + stream_end 让 UI 恢复。
        /// ★ 防重复：用 Interlocked.Exchange 确保只触发一次（Stop 也会触发 Exited 事件）。
        /// ★ 主动 Stop 时不发 error（避免正常关闭时 UI 闪红）。
        /// ★ Restart 场景：旧进程 Kill 后 Exited 事件可能延迟触发，此时 _process 已指向新进程，
        ///   必须用 ReferenceEquals(sender, _process) 过滤，否则会误报"异常退出"。
        /// </summary>
        private void OnProcessExited(object sender, EventArgs e)
        {
            // ★ 忽略旧进程的延迟事件（Restart 场景：旧进程已 Kill，新进程已 Start）
            if (!ReferenceEquals(sender, _process))
            {
                Logger.Instance.Info("PythonSidecar", "OnProcessExited: ignoring stale process event (restart)");
                return;
            }

            // ★ 防重复触发
            if (System.Threading.Interlocked.Exchange(ref _exited, 1) == 1) return;

            try
            {
                int exitCode = -1;
                try { exitCode = _process?.ExitCode ?? -1; } catch { }
                Logger.Instance.Info("PythonSidecar", $"OnProcessExited: sidecar exited, code={exitCode}, stopping={_stopping}");

                // ★ 主动 Stop 时不发 error（避免正常关闭时 UI 闪烁）
                if (_stopping)
                {
                    Logger.Instance.Info("PythonSidecar", "OnProcessExited: skipping error notify (active stop)");
                    return;
                }

                // 异常退出才通知前端
                SafeBeginInvoke(() =>
                {
                    try
                    {
                        OnError?.Invoke(this, "AI 助手进程异常退出 (code=" + exitCode + ")，请重新发送消息");
                        OnStreamEnd?.Invoke(this, 0, 0);
                    }
                    catch (Exception ex) { Logger.Instance.Warning("PythonSidecar", "OnProcessExited notify failed: " + ex.Message); }
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("PythonSidecar", "OnProcessExited failed", ex);
            }
        }

        /// <summary>
        /// 停止 sidecar 子进程
        /// </summary>
        public void Stop()
        {
            try
            {
                _stopping = true;  // ★ 标记主动停止，OnProcessExited 不发 error
                if (_process != null && !_process.HasExited)
                {
                    // ★ 关键：先取消订阅 Exited 事件，防止旧进程 Kill 后的延迟事件
                    // 在 Restart() 重置 _stopping=false 后触发，导致误报"异常退出"。
                    // 这比仅靠 _stopping 标志更可靠（事件可能在标志重置后才触发）。
                    try { _process.Exited -= OnProcessExited; } catch { }
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

        /// <summary>
        /// ★ 重启 sidecar 子进程（用于新建对话/继续历史对话时清除 AI 上下文）。
        /// Stop 后重新 Start，进程是全新的，SDK 内部 messages 完全清空。
        /// </summary>
        public void Restart()
        {
            Logger.Instance.Info("PythonSidecar", "Restarting sidecar process...");
            Stop();
            // 重置 stopping 标志，让 Start 能正常启动
            _stopping = false;
            Start();
        }

        // ============= 发送消息（C# → Python）=============

        public void SendUserMessage(string text, string sessionId, object context)
        {
            var msg = new { type = SidecarProtocol.TypeUserMessage, text, session_id = sessionId, context };
            WriteLine(JsonSerializer.Serialize(msg, _jsonOptions));
        }

        public void SendCancel() => WriteLine(@"{""type"":""cancel""}");

        /// <summary>
        /// ★ AI Native 权限确认：将用户的允许/拒绝决定发回 sidecar，PreToolUse hook 据此放行或阻止工具执行。
        /// </summary>
        public void SendPermissionResponse(string requestId, string decision)
        {
            var msg = new
            {
                type = SidecarProtocol.TypePermissionResponse,
                request_id = requestId,
                decision = decision
            };
            WriteLine(JsonSerializer.Serialize(msg, _jsonOptions));
            Logger.Instance.Info("PythonSidecar", $"SendPermissionResponse: req_id={requestId}, decision={decision}");
        }

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

        /// <summary>
        /// ★ 把历史对话发给 sidecar，让 ClaudeSDKClient 恢复上下文。
        /// 只发 user/assistant 文本对（工具调用不发给 SDK，避免重新触发工具执行）。
        /// </summary>
        public void SendRestoreHistory(List<DeepExcel.AddIn.Collaboration.HistoryMessage> messages)
        {
            if (messages == null || messages.Count == 0) return;
            var msg = new { type = SidecarProtocol.TypeRestoreHistory, messages };
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
                        SafeBeginInvoke(() => OnStreamDelta?.Invoke(this, text));
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
                        SafeBeginInvoke(() => OnToolUse?.Invoke(this, useToolName, useArgs));
                        break;

                    case SidecarProtocol.TypeClarify:
                        var q = root.GetProperty("question").GetString();
                        var opts = root.GetProperty("options").EnumerateArray()
                            .Select(x => x.GetString()).ToList();
                        SafeBeginInvoke(() => OnClarify?.Invoke(this, q, opts));
                        break;

                    case SidecarProtocol.TypeStreamEnd:
                        var inTok = root.GetProperty("input_tokens").GetInt32();
                        var outTok = root.GetProperty("output_tokens").GetInt32();
                        Logger.Instance.Info("PythonSidecar", $"OnStreamEnd: in={inTok}, out={outTok}");
                        SafeBeginInvoke(() => OnStreamEnd?.Invoke(this, inTok, outTok));
                        break;

                    case SidecarProtocol.TypePermissionRequest:
                        // ★ AI Native 权限确认：PreToolUse hook 请求用户确认高风险工具
                        var permReqId = root.GetProperty("request_id").GetString();
                        var permTool = root.GetProperty("tool").GetString();
                        var permArgs = new Dictionary<string, object>();
                        if (root.TryGetProperty("args", out var permArgsEl))
                        {
                            foreach (var prop in permArgsEl.EnumerateObject())
                            {
                                permArgs[prop.Name] = prop.Value.ValueKind switch
                                {
                                    JsonValueKind.String => prop.Value.GetString(),
                                    JsonValueKind.Number => prop.Value.GetDouble(),
                                    JsonValueKind.True => true,
                                    JsonValueKind.False => false,
                                    _ => prop.Value.GetRawText()
                                };
                            }
                        }
                        Logger.Instance.Info("PythonSidecar", $"OnPermissionRequest: tool={permTool}, req_id={permReqId}");
                        SafeBeginInvoke(() => OnPermissionRequest?.Invoke(this, permReqId, permTool, permArgs));
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
            // ★ C-1 修复：整个方法体包在 try-catch 中，包括 SendToolResult，
            // 防止任何异常逃逸到 AppDomain.UnhandledException 导致 Excel 崩溃。
            // async void 的异常无法被调用方捕获，必须内部消化。
            string callId = null;
            string toolName = null;
            try
            {
                callId = msg.GetProperty("call_id").GetString();
                toolName = msg.GetProperty("tool").GetString();
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
                SafeBeginInvoke(() => OnToolCall?.Invoke(this, callId, toolName, args));

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
                        // ★ C-2 修复：用 BeginInvoke + AsyncWaitHandle 替代同步 Invoke，
                        // 避免 Excel 模态对话框（另存为）阻塞 UI 线程时，sidecar stdout 线程同步等待导致死锁。
                        // 设 60 秒超时，超时返回失败让 LLM 能继续推理。
                        IAsyncResult ar = null;
                        var invokeDelegate = new Func<(ToolResult, object)>(() =>
                        {
                            try
                            {
                                var r = _dispatcher.Execute(toolName, args);
                                var ctx = _dispatcher.BuildExcelSnapshot();
                                return (r, ctx);
                            }
                            catch (Exception innerEx)
                            {
                                Logger.Instance.Error("PythonSidecar", "UI thread execute failed: " + toolName, innerEx);
                                return (new ToolResult { Name = toolName, Success = false, Error = innerEx.Message }, null);
                            }
                        });
                        ar = _uiControl.BeginInvoke(invokeDelegate, null);
                        // 等待 UI 线程，最多 60 秒（避免 Excel 模态对话框永久阻塞）
                        if (!ar.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(60)))
                        {
                            Logger.Instance.Error("PythonSidecar", "HandleToolCall TIMEOUT (60s) waiting for UI thread: " + toolName);
                            result = new ToolResult { Name = toolName, Success = false, Error = "UI 线程 60 秒未响应（可能 Excel 弹出模态对话框）" };
                        }
                        else
                        {
                            var tuple = ((ToolResult, object))_uiControl.EndInvoke(ar);
                            result = tuple.Item1;
                            context = tuple.Item2;
                        }
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

                // ★ C-1 修复：SendToolResult 也在 try-catch 内，防止序列化失败逃逸
                try
                {
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
                catch (Exception sendEx)
                {
                    Logger.Instance.Error("PythonSidecar", "SendToolResult failed (tool=" + toolName + ")", sendEx);
                    // 尝试发一个最小化的失败结果，让 sidecar 不卡死
                    try
                    {
                        SendToolResult(callId, false, null, "SendToolResult 异常: " + sendEx.Message, null, null);
                    }
                    catch { /* 已经尽力了，不能再让异常逃逸 */ }
                }
            }
            catch (Exception outerEx)
            {
                // 最外层兜底：确保任何未预期异常都不会逃逸到 async void 顶层
                Logger.Instance.Error("PythonSidecar",
                    $"HandleToolCall OUTER CRASH: tool={toolName}, call_id={callId}, ex={outerEx.GetType().Name}: {outerEx.Message}", outerEx);
                try
                {
                    if (callId != null)
                    {
                        SendToolResult(callId, false, null, "工具执行外部异常: " + outerEx.Message, null, null);
                    }
                }
                catch { }
            }
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
