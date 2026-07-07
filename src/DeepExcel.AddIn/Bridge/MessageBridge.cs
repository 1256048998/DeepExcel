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
        private readonly SecurityGateway _securityGateway;
        private readonly Control _uiControl;
        // 保留 ToolDispatcher（仅用于内部，前端已无直接调用入口）
        private readonly ToolDispatcher _toolDispatcher;

        // ★ 按工作簿隔离会话：key = workbook FullName
        // 每个工作簿有独立的 PythonSidecar（AI 对话上下文）+ 附件列表，
        // 解决"在A工作簿聊的内容出现在B工作簿里"的串扰问题。
        private readonly Dictionary<string, WorkbookSession> _sessions = new Dictionary<string, WorkbookSession>();

        // ★ C-5 修复：共享 JsonSerializerOptions，注册 2D 数组转换器
        private static readonly JsonSerializerOptions _jsonOptions = BuildJsonOptions();

        private static JsonSerializerOptions BuildJsonOptions()
        {
            var opts = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                // ★ 序列化为 camelCase，与前端 TypeScript 惯例一致。
                // Conversation 类的 Id/Title/CreatedAt/UpdatedAt/WorkbookName
                // 会序列化为 id/title/createdAt/updatedAt/workbookName，匹配前端 ConversationSummary 类型。
                // 有 [JsonPropertyName] 特性的属性不受影响（特性优先）。
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };
            opts.Converters.Add(new DeepExcel.AddIn.Sidecar.Object2DArrayConverter());
            opts.Converters.Add(new DeepExcel.AddIn.Sidecar.String2DArrayConverter());
            return opts;
        }

        /// <summary>
        /// 当前活动工作簿对应的 key，用于 ThisAddIn 的 IsBusy 判断
        /// </summary>
        public bool IsActiveWorkbookBusy
        {
            get
            {
                var session = GetOrCreateActiveSession();
                return session != null && session.IsBusy;
            }
        }

        public MessageBridge(Microsoft.Office.Interop.Excel.Application excelApp, Control uiControl)
        {
            _excelApp = excelApp;
            _uiControl = uiControl;

            // 初始化各层
            var workbookAnalyzer = new WorkbookAnalyzer(excelApp);
            var rangeAnalyzer = new RangeAnalyzer();
            var snapshotManager = new SnapshotManager(excelApp);
            var vbaExecutor = new VBAExecutor(excelApp, snapshotManager);
            var pythonExecutor = new PythonExecutor(excelApp, snapshotManager);

            _excelActions = new ExcelActionsImpl(
                excelApp, workbookAnalyzer, rangeAnalyzer, vbaExecutor, pythonExecutor, snapshotManager);

            var securityManager = SecurityManager.Instance;
            _securityGateway = new SecurityGateway(securityManager);
            _toolDispatcher = new ToolDispatcher(_excelActions, _excelApp, _securityGateway);
        }

        /// <summary>
        /// ★ 获取或创建当前活动工作簿的会话。
        /// 用 workbook FullName 做 key（未保存的"工作簿1"用 Name 兜底）。
        /// 返回 null 表示没有活动工作簿。
        /// </summary>
        private WorkbookSession GetOrCreateActiveSession()
        {
            try
            {
                var wb = _excelApp.ActiveWorkbook;
                if (wb == null) return null;

                string key = GetWorkbookKey(wb);
                string name = wb.Name ?? "未知";

                if (_sessions.TryGetValue(key, out var existing) && existing != null)
                {
                    existing.Touch(); // 更新 LRU 时间
                    return existing;
                }

                // LRU 回收：超过上限时先回收最久未使用的空闲会话
                EnforceSessionLimit();

                // 新建 session
                Logger.Instance.Info("MessageBridge", $"Creating new session for workbook: {name}");
                var session = new WorkbookSession(key, name, _excelActions, _excelApp, _uiControl, _securityGateway);
                session.Sidecar.OnStreamDelta += OnStreamDelta;
                session.Sidecar.OnToolCall += OnToolCall;
                session.Sidecar.OnToolUse += OnToolUse;
                session.Sidecar.OnClarify += OnClarify;
                session.Sidecar.OnStreamEnd += OnStreamEndFromSidecar;
                session.Sidecar.OnError += OnSidecarError;

                _sessions[key] = session;

                // ★ 注入附件映射引用给 ToolDispatcher，read_attachment 工具通过此映射查找附件路径。
                // 引用 session.Attachments，session 增删附件时自动同步（同一对象引用）。
                session.Sidecar.Dispatcher.Attachments = session.Attachments;

                // ★ 不自动加载历史对话到当前对话。
                // 用户打开面板默认是新对话；要看历史需点"历史对话"按钮主动选择"继续"。
                // 启动 sidecar 并发 config
                session.Sidecar.Start();
                SendConfigToSession(session);

                return session;
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "GetOrCreateActiveSession failed", ex);
                return null;
            }
        }

        /// <summary>
        /// ★ 安全获取工作簿唯一 key：优先 FullName（已保存），否则 Name（未保存的工作簿1）。
        /// </summary>
        private static string GetWorkbookKey(Workbook wb)
        {
            try
            {
                string fullName = wb.FullName;
                if (!string.IsNullOrEmpty(fullName) && (fullName.Contains("\\") || fullName.Contains("/")))
                    return fullName;
                return wb.Name ?? "workbook_" + wb.GetHashCode();
            }
            catch
            {
                return "workbook_" + wb.GetHashCode();
            }
        }

        /// <summary>
        /// ★ 找 sidecar 属于哪个 session（sidecar 回消息时用）。
        /// 因为 sidecar 是 session 级的，每个 session 有唯一 sidecar 实例。
        /// </summary>
        private WorkbookSession FindSessionBySidecar(PythonSidecar sidecar)
        {
            if (sidecar == null) return null;
            foreach (var kvp in _sessions)
            {
                if (kvp.Value.Sidecar == sidecar) return kvp.Value;
            }
            return null;
        }

        // ★ 按工作簿路由的 UI 消息事件：(workbookKey, jsonMessage)
        // ThisAddIn 监听此事件，根据 CTP 所属窗口匹配到对应工作簿再分发给对应 WebView。
        // 这样 A 工作簿的 AI 回复不会跑到 B 工作簿的面板里。
        public event Action<string, string> OnSessionUiMessage;

        // 旧的单通道回调（保留兼容，会慢慢去掉）
        private Action<string> _sendToUi;

        /// <summary>
        /// 单通道回调（已弃用，用 OnSessionUiMessage 事件替代）
        /// </summary>
        [Obsolete("Use OnSessionUiMessage event", false)]
        public void SetSendToUi(Action<string> sendToUi)
        {
            _sendToUi = sendToUi;
            Logger.Instance.Warning("MessageBridge", "SetSendToUi called (deprecated)");
        }

        /// <summary>
        /// ★ 从 AppConfig / SecurityManager 读取配置，发给指定 session 的 sidecar
        /// </summary>
        private void SendConfigToSession(WorkbookSession session)
        {
            Logger.Instance.Info("MessageBridge", $"SendConfigToSession: {session.WorkbookName}");
            try
            {
                var cfg = ConfigManager.Instance.Current;
                var providerKey = cfg.CurrentProvider;
                if (!cfg.Providers.ContainsKey(providerKey))
                {
                    Logger.Instance.Warning("MessageBridge", $"SendConfigToSession: provider '{providerKey}' not found");
                    return;
                }

                var provider = cfg.Providers[providerKey];
                var apiKey = SecurityManager.Instance.GetApiKey(providerKey);
                if (string.IsNullOrEmpty(apiKey))
                {
                    apiKey = provider.ApiKey ?? "";
                }

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

                Logger.Instance.Info("MessageBridge",
                    $"SendConfigToSession: wb={session.WorkbookName}, model={model}, baseUrl={baseUrl}, hasKey={!string.IsNullOrEmpty(apiKey)}");
                session.Sidecar.UpdateConfig(baseUrl, model, apiKey);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "SendConfigToSession failed", ex);
            }
        }

        /// <summary>
        /// ★ 把历史对话发给 sidecar，让 ClaudeSDKClient 恢复上下文。
        /// 只发 user/assistant 文本对，工具调用不发（避免重新触发执行）。
        /// 在用户点击"继续"历史对话时调用。
        /// </summary>
        private void SendHistoryToSidecar(WorkbookSession session)
        {
            try
            {
                if (session.ConversationMessages == null || session.ConversationMessages.Count == 0)
                {
                    return;
                }

                // 过滤：只发 user 和 assistant 文本消息（跳过 tool 和 clarify）
                var filtered = new List<DeepExcel.AddIn.Collaboration.HistoryMessage>();
                foreach (var m in session.ConversationMessages)
                {
                    if (m.role == "user" && !string.IsNullOrEmpty(m.content))
                    {
                        filtered.Add(new DeepExcel.AddIn.Collaboration.HistoryMessage
                        {
                            role = "user",
                            content = m.content
                        });
                    }
                    else if (m.role == "assistant" && !string.IsNullOrEmpty(m.content) && m.type != "clarify")
                    {
                        filtered.Add(new DeepExcel.AddIn.Collaboration.HistoryMessage
                        {
                            role = "assistant",
                            content = m.content
                        });
                    }
                }

                if (filtered.Count > 0)
                {
                    Logger.Instance.Info("MessageBridge",
                        $"SendHistoryToSidecar: {session.WorkbookName}, sending {filtered.Count} messages");
                    session.Sidecar.SendRestoreHistory(filtered);
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "SendHistoryToSidecar failed", ex);
            }
        }

        /// <summary>
        /// 对所有活动 session 刷新 config（用户改设置时调用）
        /// </summary>
        public void RefreshConfigForAllSessions()
        {
            foreach (var kvp in _sessions)
            {
                try { SendConfigToSession(kvp.Value); }
                catch { }
            }
        }

        /// <summary>
        /// 处理来自UI的消息。根据当前活动工作簿路由到对应 session。
        /// </summary>
        public string HandleMessage(string json)
        {
            try
            {
                var msg = JsonSerializer.Deserialize<Message>(json);
                if (msg == null) return MakeError("Invalid message");

                // 这些消息不需要 session 上下文
                switch (msg.Type)
                {
                    case "ping":
                        return MakeResponse("pong", new { });
                    case "get_selection":
                        return MakeResponse("selection", _excelActions.GetSelection());
                    case "read_workbook":
                        return MakeResponse("workbook", _excelActions.ReadWorkbook());
                    case "read_range":
                        return HandleReadRange(msg);
                    case "list_snapshots":
                        return HandleListSnapshots();
                    case "rollback_snapshot":
                        return HandleRollbackSnapshot(msg);
                    case "delete_snapshot":
                        return HandleDeleteSnapshot(msg);
                    case "get_model_config":
                        return HandleGetModelConfig();
                    case "save_model_config":
                        return HandleSaveModelConfig(msg);
                    case "test_api_key":
                        return HandleTestApiKey(msg);
                }

                // 以下消息需要 session 上下文
                var session = GetOrCreateActiveSession();
                if (session == null) return MakeError("没有活动工作簿");

                switch (msg.Type)
                {
                    case "user_message":
                        return HandleUserMessage(session, msg);
                    case "cancel":
                        return HandleCancel(session);
                    // ★ 附件管理
                    case "list_attachments":
                        return MakeResponse("attachments", new { list = session.GetAttachmentList() });
                    case "upload_attachment":
                        return HandleUploadAttachment(session, msg);
                    case "delete_attachment":
                        return HandleDeleteAttachment(session, msg);
                    // ★ 多对话历史管理
                    case "get_current_messages":
                        return MakeResponse("current_messages", new { messages = session.ConversationMessages });
                    case "new_conversation":
                        return HandleNewConversation(session);
                    case "list_conversations":
                        return MakeResponse("conversations", new { list = session.ListConversations() });
                    case "continue_conversation":
                        return HandleContinueConversation(session, msg);
                    case "delete_conversation":
                        return HandleDeleteConversation(session, msg);
                    default:
                        return MakeError($"Unknown or blocked message type: {msg.Type}");
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleMessage error", ex);
                return MakeError($"Handle error: {ex.Message}");
            }
        }

        /// <summary>
        /// ★ 返回当前配置（API Key 脱敏）给前端模型配置弹窗
        /// </summary>
        private string HandleGetModelConfig()
        {
            try
            {
                var cfg = ConfigManager.Instance.Current;
                var safe = SecurityManager.Instance.GetSafeConfig(cfg);
                return MakeResponse("model_config", safe);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleGetModelConfig failed", ex);
                return MakeError("加载配置失败: " + ex.Message);
            }
        }

        /// <summary>
        /// ★ 保存模型配置（provider/model/apiKey/baseUrl/maxTurns），立即对所有 session 生效
        /// 占位符 "***keep***" 表示用户未修改 API Key，跳过保存
        /// </summary>
        private string HandleSaveModelConfig(Message msg)
        {
            try
            {
                var payload = msg.Payload.Value;
                var provider = payload.GetProperty("provider").GetString();
                var model = payload.GetProperty("model").GetString();
                var apiKey = payload.GetProperty("apiKey").GetString();
                var baseUrl = payload.GetProperty("baseUrl").GetString();
                int maxTurns = payload.GetProperty("maxTurns").GetInt32();

                if (string.IsNullOrEmpty(provider))
                {
                    return MakeError("provider 不能为空");
                }

                var cfg = ConfigManager.Instance.Current;
                if (!cfg.Providers.ContainsKey(provider))
                {
                    return MakeError($"未知的 provider: {provider}");
                }

                // 1. 更新 BaseUrl（如果非空且与默认不同）
                if (!string.IsNullOrEmpty(baseUrl))
                {
                    cfg.Providers[provider].BaseUrl = baseUrl;
                }

                // 2. 保存 API Key（跳过占位符和空值）
                if (!string.IsNullOrEmpty(apiKey) && apiKey != "***keep***")
                {
                    ConfigManager.Instance.UpdateApiKey(provider, apiKey);
                }

                // 3. 切换 provider + model
                ConfigManager.Instance.SwitchProvider(provider, model);

                // 4. 更新 MaxTurns
                if (cfg.General == null) cfg.General = new Config.GeneralSettings();
                if (maxTurns > 0 && maxTurns <= 200)
                {
                    cfg.General.MaxTurns = maxTurns;
                }

                // 5. 持久化
                ConfigManager.Instance.Save();

                // 6. 立即对所有 session 生效
                RefreshConfigForAllSessions();

                Logger.Instance.Info("MessageBridge",
                    $"HandleSaveModelConfig: provider={provider}, model={model}, maxTurns={maxTurns}, apiKeyChanged={apiKey != "***keep***" && !string.IsNullOrEmpty(apiKey)}");

                return MakeResponse("config_saved", new { success = true });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSaveModelConfig failed", ex);
                return MakeError("保存配置失败: " + ex.Message);
            }
        }

        /// <summary>
        /// ★ 测试 API Key 连接（不保存任何数据）
        /// 向 baseUrl 发一个 Anthropic Messages API 的 1-token 请求
        /// </summary>
        private string HandleTestApiKey(Message msg)
        {
            try
            {
                var payload = msg.Payload.Value;
                var provider = payload.GetProperty("provider").GetString();
                var apiKey = payload.GetProperty("apiKey").GetString();
                var baseUrl = payload.GetProperty("baseUrl").GetString();
                var model = payload.GetProperty("model").GetString();

                if (string.IsNullOrEmpty(apiKey) || apiKey == "***keep***")
                {
                    return MakeResponse("api_test_result", new { success = false, error = "请先输入 API Key" });
                }
                if (string.IsNullOrEmpty(baseUrl))
                {
                    return MakeResponse("api_test_result", new { success = false, error = "Base URL 为空" });
                }

                // 用 Task.Run 在后台线程执行，避免阻塞 UI 线程
                var sw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    using var client = new System.Net.Http.HttpClient();
                    client.Timeout = TimeSpan.FromSeconds(15);

                    // 构造 Anthropic Messages API 最小请求
                    var testUrl = baseUrl.TrimEnd('/') + "/v1/messages";
                    var body = new
                    {
                        model = model,
                        max_tokens = 1,
                        messages = new[] { new { role = "user", content = "hi" } }
                    };
                    var json = System.Text.Json.JsonSerializer.Serialize(body);
                    var content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");

                    var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, testUrl);
                    request.Content = content;
                    request.Headers.Add("x-api-key", apiKey);
                    request.Headers.Add("anthropic-version", "2023-06-01");

                    var response = client.SendAsync(request).Result;
                    sw.Stop();

                    if (response.IsSuccessStatusCode)
                    {
                        Logger.Instance.Info("MessageBridge", $"HandleTestApiKey: success, latency={sw.ElapsedMilliseconds}ms");
                        return MakeResponse("api_test_result", new { success = true, latencyMs = sw.ElapsedMilliseconds, error = (string)null });
                    }
                    else
                    {
                        var errBody = response.Content.ReadAsStringAsync().Result;
                        Logger.Instance.Warning("MessageBridge", $"HandleTestApiKey: HTTP {response.StatusCode}, body={errBody}");
                        return MakeResponse("api_test_result", new
                        {
                            success = false,
                            latencyMs = sw.ElapsedMilliseconds,
                            error = $"HTTP {response.StatusCode}: {errBody}"
                        });
                    }
                }
                catch (System.Net.Http.HttpRequestException hex)
                {
                    sw.Stop();
                    Logger.Instance.Warning("MessageBridge", "HandleTestApiKey network error: " + hex.Message);
                    return MakeResponse("api_test_result", new { success = false, latencyMs = sw.ElapsedMilliseconds, error = "网络错误: " + hex.Message });
                }
                catch (AggregateException aex)
                {
                    sw.Stop();
                    var inner = aex.InnerException?.Message ?? aex.Message;
                    Logger.Instance.Warning("MessageBridge", "HandleTestApiKey error: " + inner);
                    return MakeResponse("api_test_result", new { success = false, latencyMs = sw.ElapsedMilliseconds, error = inner });
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleTestApiKey failed", ex);
                return MakeError("测试连接失败: " + ex.Message);
            }
        }

        private string HandleUserMessage(WorkbookSession session, Message msg)
        {
            try
            {
                var content = msg.Payload?.GetProperty("content").GetString();

                // 如果有待回答的 clarify，把这条消息当 clarify_answer（session 级隔离）
                if (session.PendingClarifyQuestion != null)
                {
                    session.Sidecar.SendClarifyAnswer(content);
                    session.PendingClarifyQuestion = null;
                    return MakeResponse("ack", new { received = true, kind = "clarify_answer" });
                }

                // 正常用户消息：附带 Excel 上下文 + 附件列表
                var context = session.BuildContext(_excelActions);
                var sessionId = session.NextSessionId();
                session.IsBusy = true;
                session.Sidecar.SendUserMessage(content, sessionId, context);

                // ★ 追加到历史
                session.AppendUserMessage(content);

                return MakeResponse("ack", new { received = true });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleUserMessage failed", ex);
                return MakeError($"User message error: {ex.Message}");
            }
        }

        private string HandleCancel(WorkbookSession session)
        {
            try
            {
                // ★ P-2 修复：cancel 清理 session 级 pending clarify
                if (session.PendingClarifyQuestion != null)
                {
                    Logger.Instance.Info("MessageBridge", $"HandleCancel: clearing pending clarify for {session.WorkbookName}");
                    session.PendingClarifyQuestion = null;
                }
                session.Sidecar.SendCancel();
                session.IsBusy = false;
                return MakeResponse("cancelled", new { });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleCancel failed", ex);
                return MakeError($"Cancel error: {ex.Message}");
            }
        }

        // ============= 多对话历史消息处理 =============

        /// <summary>★ 新建对话：存当前对话，清空内存，重启 sidecar 清 AI 上下文</summary>
        private string HandleNewConversation(WorkbookSession session)
        {
            try
            {
                session.NewConversation();

                // ★ 重启 sidecar 进程，彻底清除 AI 上下文
                // 不重启的话 SDK 内部还记着上一轮对话的 messages
                try
                {
                    session.Sidecar.Restart();
                    SendConfigToSession(session);
                }
                catch (Exception ex)
                {
                    Logger.Instance.Warning("MessageBridge", $"HandleNewConversation: sidecar restart failed: {ex.Message}");
                }

                return MakeResponse("new_conversation", new { conversationId = session.CurrentConversationId });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleNewConversation failed", ex);
                return MakeError($"New conversation error: {ex.Message}");
            }
        }

        /// <summary>★ 继续历史对话：加载历史 messages，重启 sidecar 并注入历史上下文</summary>
        private string HandleContinueConversation(WorkbookSession session, Message msg)
        {
            try
            {
                var conversationId = msg.Payload?.GetProperty("conversation_id").GetString();
                if (string.IsNullOrEmpty(conversationId))
                {
                    return MakeError("conversation_id is required");
                }

                var messages = session.ContinueConversation(conversationId);
                if (messages == null)
                {
                    return MakeError($"Conversation not found: {conversationId}");
                }

                // ★ 重启 sidecar 并注入历史上下文
                try
                {
                    session.Sidecar.Restart();
                    SendConfigToSession(session);
                    // 把历史 user/assistant 文本发给新 sidecar，让它"记得"之前的对话
                    SendHistoryToSidecar(session);
                }
                catch (Exception ex)
                {
                    Logger.Instance.Warning("MessageBridge", $"HandleContinueConversation: sidecar restart failed: {ex.Message}");
                }

                return MakeResponse("continue_conversation", new
                {
                    conversationId = conversationId,
                    messages = messages
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleContinueConversation failed", ex);
                return MakeError($"Continue conversation error: {ex.Message}");
            }
        }

        /// <summary>★ 删除指定历史对话</summary>
        private string HandleDeleteConversation(WorkbookSession session, Message msg)
        {
            try
            {
                var conversationId = msg.Payload?.GetProperty("conversation_id").GetString();
                if (string.IsNullOrEmpty(conversationId))
                {
                    return MakeError("conversation_id is required");
                }

                bool ok = session.DeleteConversation(conversationId);
                return MakeResponse("delete_conversation", new { success = ok, conversationId });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleDeleteConversation failed", ex);
                return MakeError($"Delete conversation error: {ex.Message}");
            }
        }

        /// <summary>★ 上传附件</summary>
        private string HandleUploadAttachment(WorkbookSession session, Message msg)
        {
            try
            {
                var fileName = msg.Payload?.GetProperty("file_name").GetString();
                var fileBase64 = msg.Payload?.GetProperty("file_base64").GetString();
                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(fileBase64))
                    return MakeError("缺少 file_name 或 file_base64 参数");

                var info = session.AddAttachment(fileName, fileBase64);
                return MakeResponse("uploaded", info);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleUploadAttachment failed", ex);
                return MakeError($"上传失败: {ex.Message}");
            }
        }

        /// <summary>★ 删除附件</summary>
        private string HandleDeleteAttachment(WorkbookSession session, Message msg)
        {
            try
            {
                var fileName = msg.Payload?.GetProperty("file_name").GetString();
                if (string.IsNullOrEmpty(fileName)) return MakeError("缺少 file_name 参数");
                bool ok = session.RemoveAttachment(fileName);
                return MakeResponse("deleted", new { success = ok, file_name = fileName });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleDeleteAttachment failed", ex);
                return MakeError($"删除失败: {ex.Message}");
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

        /// <summary>★ 历史版本：列出所有快照</summary>
        private string HandleListSnapshots()
        {
            try
            {
                var snapshots = _excelActions.ListSnapshots();
                // 转成前端友好的 camelCase 对象数组
                var list = new List<object>();
                foreach (var s in snapshots)
                {
                    list.Add(new
                    {
                        id = s.Id,
                        workbookName = s.WorkbookName,
                        createdAt = s.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"),
                        timestamp = ((DateTimeOffset)DateTime.SpecifyKind(s.CreatedAt, DateTimeKind.Local)).ToUnixTimeSeconds(),
                        reason = s.Reason,
                    });
                }
                return MakeResponse("snapshots", new { list });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleListSnapshots failed", ex);
                return MakeError("列出快照失败: " + ex.Message);
            }
        }

        /// <summary>★ 历史版本：回滚到指定快照</summary>
        private string HandleRollbackSnapshot(Message msg)
        {
            try
            {
                var snapshotId = msg.Payload?.GetProperty("snapshot_id").GetString();
                if (string.IsNullOrEmpty(snapshotId))
                {
                    return MakeError("缺少 snapshot_id 参数");
                }
                Logger.Instance.Info("MessageBridge", "HandleRollbackSnapshot: " + snapshotId);
                bool ok = _excelActions.Rollback(snapshotId);
                return MakeResponse("rollback_result", new { success = ok, snapshot_id = snapshotId });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleRollbackSnapshot failed", ex);
                return MakeError("回滚失败: " + ex.Message);
            }
        }

        /// <summary>★ 历史版本：删除指定快照</summary>
        private string HandleDeleteSnapshot(Message msg)
        {
            try
            {
                var snapshotId = msg.Payload?.GetProperty("snapshot_id").GetString();
                if (string.IsNullOrEmpty(snapshotId))
                {
                    return MakeError("缺少 snapshot_id 参数");
                }
                bool ok = _excelActions.DeleteSnapshot(snapshotId);
                return MakeResponse("delete_snapshot_result", new { success = ok, snapshot_id = snapshotId });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleDeleteSnapshot failed", ex);
                return MakeError("删除快照失败: " + ex.Message);
            }
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

        // 回调方法 - sidecar 发回的消息，通过 sender 精准找到对应 session 再发给 UI
        private void OnStreamDelta(PythonSidecar sender, string delta)
        {
            var session = FindSessionBySidecar(sender);
            if (session != null)
            {
                if (!session.IsBusy) session.IsBusy = true;
                SendToSessionUi(session.WorkbookKey, "stream_delta", new { delta });
                // ★ 追加到历史（流式增量）
                session.AppendAssistantDelta(delta);
            }
            Logger.Instance.Debug("MessageBridge", "OnStreamDelta: len=" + (delta?.Length ?? 0));
        }

        private void OnToolCall(PythonSidecar sender, string callId, string name, Dictionary<string, object> args)
        {
            Logger.Instance.Info("MessageBridge", $"OnToolCall: {name} (call_id={callId})");
        }

        private void OnToolUse(PythonSidecar sender, string name, Dictionary<string, object> args)
        {
            var session = FindSessionBySidecar(sender);
            if (session != null)
            {
                var displayName = name?.Replace("mcp__excel__", "") ?? name;
                SendToSessionUi(session.WorkbookKey, "tool_call",
                    new { call_id = "", name = displayName, arguments = args });
                // ★ 追加到历史
                session.AppendToolCall(displayName);
            }
        }

        private void OnClarify(PythonSidecar sender, string question, List<string> options)
        {
            var session = FindSessionBySidecar(sender);
            if (session != null)
            {
                session.PendingClarifyQuestion = question;
                session.IsBusy = false;
                SendToSessionUi(session.WorkbookKey, "clarify", new { question, options });
                // ★ 追加到历史
                session.AppendClarify(question, options?.ToArray());
            }
        }

        private void OnStreamEndFromSidecar(PythonSidecar sender, int inputTokens, int outputTokens)
        {
            var session = FindSessionBySidecar(sender);
            if (session != null)
            {
                session.IsBusy = false;
                SendToSessionUi(session.WorkbookKey, "stream_end",
                    new { input_tokens = inputTokens, output_tokens = outputTokens });
                // ★ stream_end 时持久化对话历史到磁盘
                session.OnStreamEnd();
            }
        }

        private void OnSidecarError(PythonSidecar sender, string error)
        {
            var session = FindSessionBySidecar(sender);
            if (session != null)
            {
                session.IsBusy = false;
                SendToSessionUi(session.WorkbookKey, "error", new { message = "Sidecar: " + error });
            }
            else
            {
                // 找不到 session，用旧通道兜底（至少让用户看到错误）
                SendToUi("error", new { message = "Sidecar: " + error });
            }
        }

        /// <summary>
        /// 找当前处于 busy 状态的 session（流式响应中）。
        /// 因为用户在一个时刻只能操作一个活动工作簿，正常只有一个 session 在流式输出。
        /// </summary>
        private WorkbookSession FindBusySession()
        {
            WorkbookSession firstBusy = null;
            foreach (var kvp in _sessions)
            {
                if (kvp.Value.IsBusy)
                {
                    if (firstBusy == null) firstBusy = kvp.Value;
                    // 优先匹配当前活动工作簿
                    try
                    {
                        var activeWb = _excelApp.ActiveWorkbook;
                        if (activeWb != null)
                        {
                            string activeKey = GetWorkbookKey(activeWb);
                            if (kvp.Key == activeKey) return kvp.Value;
                        }
                    }
                    catch { }
                }
            }
            return firstBusy;
        }

        public void Cancel()
        {
            try
            {
                var session = GetOrCreateActiveSession();
                session?.Sidecar.SendCancel();
                if (session != null) session.IsBusy = false;
                Logger.Instance.Info("MessageBridge", "Cancel requested");
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "Cancel failed", ex);
            }
        }

        private void SendToSessionUi(string workbookKey, string type, object payload)
        {
            var json = JsonSerializer.Serialize(new { type, payload }, _jsonOptions);
            try
            {
                OnSessionUiMessage?.Invoke(workbookKey, json);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "SendToSessionUi event handler failed", ex);
            }
            // ★ 不再走 _sendToUi 兜底广播：OnSessionUiMessage 已稳定，
            // 双通道会导致同一 WebView 收到两份消息，前端 React 状态累加后消息会渲染两遍。
            // _sendToUi 仅保留给 OnSidecarError 找不到 session 时的兜底分支使用。
        }

        private void SendToUi(string type, object payload)
        {
            var json = JsonSerializer.Serialize(new { type, payload }, _jsonOptions);
            try { _sendToUi?.Invoke(json); } catch { }
        }

        private string MakeResponse(string type, object payload)
        {
            // ★ C-5 修复：用 _jsonOptions 序列化
            return JsonSerializer.Serialize(new { type, payload }, _jsonOptions);
        }

        private string MakeError(string message)
        {
            // ★ C-5 修复：用 _jsonOptions 序列化
            return JsonSerializer.Serialize(new
            {
                type = "error",
                payload = new { message }
            }, _jsonOptions);
        }

        // ============= 性能优化：预启动 =============

        /// <summary>
        /// 预热：预启动当前活动工作簿的 sidecar 进程。
        /// 在 Excel 启动完成后后台调用，用户打开面板发送消息时 sidecar 已经准备好了。
        /// 非阻塞：sidecar 启动是异步的，这里只是触发启动。
        /// </summary>
        public void PreWarmActiveSession()
        {
            try
            {
                // 只是创建 session + 启动 sidecar，不等结果
                var session = GetOrCreateActiveSession();
                if (session != null)
                {
                    Logger.Instance.Info("MessageBridge",
                        $"Pre-warmed sidecar for workbook: {session.WorkbookName}");
                }
            }
            catch (Exception ex)
            {
                // 预热失败不影响主流程，只记日志
                Logger.Instance.Warning("MessageBridge", "PreWarmActiveSession failed: " + ex.Message);
            }
        }

        // ============= 工作簿生命周期管理 =============

        /// <summary>
        /// 最大会话数上限。每个会话对应一个 Python 子进程，
        /// 过多会占用大量内存。超过上限时按 LRU 回收最久未使用的会话。
        /// </summary>
        private const int MaxSessions = 8;

        /// <summary>
        /// 工作簿关闭时调用：清理对应会话，释放 sidecar 进程。
        /// </summary>
        public void OnWorkbookClose(string workbookKey)
        {
            if (string.IsNullOrEmpty(workbookKey)) return;
            try
            {
                if (_sessions.TryGetValue(workbookKey, out var session))
                {
                    session.Dispose();
                    _sessions.Remove(workbookKey);
                    Logger.Instance.Info("MessageBridge", $"Session closed (workbook closed): {workbookKey}");
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "OnWorkbookClose failed", ex);
            }
        }

        /// <summary>
        /// 工作簿另存为后调用：更新会话 key（FullName 变了），
        /// 避免对话历史丢失。
        /// </summary>
        public void OnWorkbookAfterSave(string oldKey, string newKey, string newName)
        {
            if (string.IsNullOrEmpty(oldKey) || string.IsNullOrEmpty(newKey) || oldKey == newKey) return;
            try
            {
                if (_sessions.TryGetValue(oldKey, out var session))
                {
                    _sessions.Remove(oldKey);
                    session.UpdateKey(newKey, newName);
                    _sessions[newKey] = session;
                    Logger.Instance.Info("MessageBridge",
                        $"Session key updated (after save): {oldKey} -> {newKey}");
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "OnWorkbookAfterSave failed", ex);
            }
        }

        /// <summary>
        /// LRU 回收：会话数超过上限时，回收最久未使用的空闲会话。
        /// 在创建新会话前调用。
        /// </summary>
        private void EnforceSessionLimit()
        {
            if (_sessions.Count < MaxSessions) return;

            // 找最久未使用的非 busy 会话
            WorkbookSession oldestIdle = null;
            string oldestKey = null;
            DateTime oldestTime = DateTime.MaxValue;

            foreach (var kvp in _sessions)
            {
                var s = kvp.Value;
                if (s.IsBusy) continue; // 忙碌中的会话不回收
                if (s.LastUsedTime < oldestTime)
                {
                    oldestTime = s.LastUsedTime;
                    oldestIdle = s;
                    oldestKey = kvp.Key;
                }
            }

            if (oldestIdle != null && oldestKey != null)
            {
                Logger.Instance.Info("MessageBridge",
                    $"LRU evicting idle session: {oldestKey} (last used {oldestTime})");
                oldestIdle.Dispose();
                _sessions.Remove(oldestKey);
            }
        }

        public void Dispose()
        {
            foreach (var kvp in _sessions)
            {
                try { kvp.Value.Dispose(); }
                catch (Exception ex) { Logger.Instance.Error("MessageBridge", $"Dispose session {kvp.Value.WorkbookName} failed", ex); }
            }
            _sessions.Clear();
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

        /// <summary>
        /// ★ 公共方法：解析可能含 sheet 限定的地址（如 "Sheet3!A1"）为 Range。
        /// 当 sheet 不存在时返回 null 并设置 error/suggestion，让调用方给用户友好提示。
        /// 避免 _app.Range[address] 抛晦涩的 COMException 0x800A03EC。
        /// </summary>
        /// <param name="address">地址，可能含 sheet 前缀（如 "Sheet3!A1" 或 "A1"）</param>
        /// <param name="toolName">工具名（用于错误提示）</param>
        /// <param name="error">错误信息（sheet 不存在时填充）</param>
        /// <param name="suggestion">建议（如"请先调 add_sheet"）</param>
        /// <returns>解析成功返回 Range，失败返回 null</returns>
        private Range TryResolveRange(string address, string toolName, out string error, out string suggestion)
        {
            error = null;
            suggestion = null;
            if (string.IsNullOrEmpty(address))
            {
                error = "address 不能为空";
                suggestion = "请传入有效的单元格地址（如 A1 或 Sheet3!A1）";
                return null;
            }

            // 检测 sheet 前缀
            string sheetName = null;
            if (address.Contains("!"))
            {
                int bang = address.IndexOf('!');
                sheetName = address.Substring(0, bang);
            }

            if (sheetName != null)
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null)
                {
                    error = "没有活动工作簿";
                    return null;
                }
                bool sheetExists = false;
                foreach (Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        sheetExists = true;
                        break;
                    }
                }
                if (!sheetExists)
                {
                    Logger.Instance.Warning("ExcelActions", $"{toolName}: sheet '{sheetName}' not found, address={address}");
                    error = $"工作表 '{sheetName}' 不存在";
                    suggestion = $"请先调用 add_sheet(name=\"{sheetName}\") 创建该工作表";
                    return null;
                }
            }

            try
            {
                return _app.Range[address];
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ExcelActions", $"{toolName}: Range[{address}] failed: {ex.Message}");
                error = ex.Message;
                suggestion = "地址格式可能错误（正确示例：A1 或 Sheet3!A1:G100）";
                return null;
            }
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
                var range = TryResolveRange(address, "read_range", out string error, out string suggestion);
                if (range == null)
                {
                    return new { error, suggestion };
                }
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
            // ★ 注入上下文：工作簿路径、活动 sheet 名等，供 AI 代码使用
            // 之前不传 context，AI 代码访问 workbook_path 会 KeyError
            var context = new Dictionary<string, object>();
            try
            {
                var wb = _app?.ActiveWorkbook;
                if (wb != null)
                {
                    context["workbook_path"] = wb.FullName ?? wb.Name ?? "";
                    context["workbook_name"] = wb.Name ?? "";
                    var activeSheet = wb.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    context["active_sheet"] = activeSheet?.Name ?? "";
                }
            }
            catch (Exception ex) { Logger.Instance.Warning("MessageBridge", "Build python context failed: " + ex.Message); }

            var result = _pythonExecutor.Execute(code, context);
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
                // ★ H2 修复：公式黑名单，阻止 LLM 通过 WEBSERVICE/SHELL 等函数外泄数据或执行命令
                if (!string.IsNullOrEmpty(formula))
                {
                    string upper = formula.ToUpperInvariant();
                    string[] blocked = {
                        "WEBSERVICE", "FILTERXML",  // HTTP 请求外泄数据
                        "CALL",  // 调用 DLL 函数
                        "REGISTER.ID", "REGISTER",  // 注册 DLL
                        "EXEC",  // 执行程序（旧版）
                        "HYPERLINK",  // 创建可疑链接（虽不直接外泄，但可能钓鱼）
                    };
                    foreach (var bad in blocked)
                    {
                        // 匹配 =WEBSERVICE( 或 =XXX.WEBSERVICE( 等
                        if (upper.Contains("=" + bad + "(") || upper.Contains("." + bad + "("))
                        {
                            Logger.Instance.Warning("MessageBridge", "WriteFormula blocked by H2 blacklist: " + bad + " in " + formula);
                            return new ToolResult
                            {
                                Name = "write_formula",
                                Success = false,
                                Error = "公式包含被禁用的函数: " + bad + "（不允许网络请求/外部调用）",
                                Suggestion = "请用其他公式实现，或改用 read_range + write_value 手动计算"
                            };
                        }
                    }
                }

                var range = TryResolveRange(address, "write_formula", out string wfrError, out string wfrSuggestion);
                if (range == null)
                {
                    return new ToolResult { Name = "write_formula", Success = false, Error = wfrError, Suggestion = wfrSuggestion };
                }
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
                var range = TryResolveRange(address, "write_value", out string wvError, out string wvSuggestion);
                if (range == null)
                {
                    return new ToolResult { Name = "write_value", Success = false, Error = wvError, Suggestion = wvSuggestion };
                }
                // ★ 如果 value 是以 = 开头的字符串，Excel 会把它当公式解析（如 "=张三" 或 "=\"张三\""）。
                // write_value 的语义是写纯文本值，所以先设单元格为文本格式，避免被当公式。
                if (value is string s && !string.IsNullOrEmpty(s) && s[0] == '=')
                {
                    try { range.NumberFormat = "@"; } catch { }
                }
                range.Value = value;
                return new ToolResult { Name = "write_value", Success = true };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", $"WriteValue failed: address={address}, value={value}, error={ex.Message}");
                return new ToolResult
                {
                    Name = "write_value",
                    Success = false,
                    Error = ex.Message,
                    Suggestion = "地址格式可能错误（正确示例：A1 或 Sheet3!A1），或工作表不存在（先调 add_sheet 创建）",
                };
            }
        }

        /// <summary>
        /// ★ 批量写入二维数组到指定起始单元格。
        /// address 是左上角单元格（如 "A1" 或 "Sheet3!A1"），values 是二维数组（values[row][col]）。
        /// 内部用 Range.Value = 2DArray 一次性写入，比逐个 write_value 快 100 倍。
        /// </summary>
        public ToolResult WriteRange(string address, object[][] values)
        {
            try
            {
                if (values == null || values.Length == 0)
                {
                    return new ToolResult { Name = "write_range", Success = false, Error = "values 为空" };
                }

                int rowCount = values.Length;
                int colCount = values[0]?.Length ?? 0;
                if (colCount == 0)
                {
                    return new ToolResult { Name = "write_range", Success = false, Error = "values 第一行为空" };
                }

                Logger.Instance.Info("MessageBridge", $"WriteRange: address={address}, rows={rowCount}, cols={colCount}");

                var range = TryResolveRange(address, "write_range", out string wrErr, out string wrSug);
                if (range == null)
                {
                    return new ToolResult { Name = "write_range", Success = false, Error = wrErr, Suggestion = wrSug };
                }

                // 扩展到完整区域
                var fullRange = range.Resize[rowCount, colCount];

                // 构造 1-based 2D 数组（Excel COM 要求）
                var arr = new object[rowCount, colCount];
                for (int r = 0; r < rowCount; r++)
                {
                    var row = values[r];
                    for (int c = 0; c < colCount; c++)
                    {
                        arr[r, c] = row != null && c < row.Length ? row[c] : null;
                    }
                }

                fullRange.Value = arr;
                return new ToolResult
                {
                    Name = "write_range",
                    Success = true,
                    Data = new { address, rows_written = rowCount, cols_written = colCount },
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", $"WriteRange failed: address={address}, error={ex.Message}");
                return new ToolResult
                {
                    Name = "write_range",
                    Success = false,
                    Error = ex.Message,
                    Suggestion = "地址格式可能错误（正确示例：A1 或 Sheet3!A1），或工作表不存在（先调 add_sheet 创建）",
                };
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

        /// <summary>
        /// ★ 列出所有历史快照（前端历史版本 UI 调用）
        /// </summary>
        public System.Collections.Generic.List<DeepExcel.AddIn.Executor.SnapshotMeta> ListSnapshots()
        {
            return _snapshotManager.ListSnapshots();
        }

        /// <summary>
        /// ★ 删除单个快照（前端历史版本 UI 调用）
        /// </summary>
        public bool DeleteSnapshot(string snapshotId)
        {
            return _snapshotManager.DeleteSnapshot(snapshotId);
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
                var range = TryResolveRange(address, "set_number_format", out string err, out string sug);
                if (range == null)
                {
                    return new ToolResult { Name = "set_number_format", Success = false, Error = err, Suggestion = sug };
                }
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
                var range = TryResolveRange(address, "set_column_width", out string err, out string sug);
                if (range == null)
                {
                    return new ToolResult { Name = "set_column_width", Success = false, Error = err, Suggestion = sug };
                }
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

        public ToolResult SortData(string rangeAddress, string sortColumn, bool descending, bool hasHeader = false)
        {
            try
            {
                Logger.Instance.Info("ExcelActions",
                    $"SortData: range={rangeAddress}, sortColumn={sortColumn}, descending={descending}, hasHeader={hasHeader}");

                var range = TryResolveRange(rangeAddress, "sort_data", out string sortErr, out string sortSug);
                if (range == null)
                {
                    return new ToolResult { Name = "sort_data", Success = false, Error = sortErr, Suggestion = sortSug };
                }

                // ★ 排序前检查合并单元格：有合并单元格的区域无法正确排序，
                // 会导致数据错位（用户反馈的"B列数据换到A列"可能就是这个原因）。
                // 检测到合并单元格时直接报错，让用户先取消合并再排序。
                object merged = null;
                try { merged = range.MergeCells; } catch { }
                if (merged is bool bm && bm)
                {
                    return new ToolResult
                    {
                        Name = "sort_data",
                        Success = false,
                        Error = "排序区域包含合并单元格，无法直接排序。请先取消合并单元格（unmerge_cells）后再排序。",
                        Suggestion = "请先调用 unmerge_cells 工具取消该区域的合并，然后再排序。"
                    };
                }
                if (merged == null)
                {
                    // 混合情况（部分合并）：检查是否真的有合并单元格
                    // 简单检测：检查第一行、最后一行、第一列、最后一列
                    bool hasMerged = false;
                    try
                    {
                        int rows = range.Rows.Count;
                        int cols = range.Columns.Count;
                        // 检查四角
                        var corners = new Range[]
                        {
                            (Range)range.Cells[1, 1],
                            (Range)range.Cells[1, cols],
                            (Range)range.Cells[rows, 1],
                            (Range)range.Cells[rows, cols],
                        };
                        foreach (var corner in corners)
                        {
                            try
                            {
                                if (corner.MergeCells is bool cm && cm)
                                {
                                    hasMerged = true;
                                    break;
                                }
                            }
                            catch { }
                        }
                        // 再检查中间一些单元格
                        if (!hasMerged && rows > 4 && cols > 2)
                        {
                            for (int r = 2; r < rows && r < 6; r++)
                            {
                                for (int c = 1; c <= cols && c <= 3; c++)
                                {
                                    try
                                    {
                                        var cell = (Range)range.Cells[r, c];
                                        if (cell.MergeCells is bool cm2 && cm2)
                                        {
                                            hasMerged = true;
                                            break;
                                        }
                                    }
                                    catch { }
                                }
                                if (hasMerged) break;
                            }
                        }
                    }
                    catch { }

                    if (hasMerged)
                    {
                        return new ToolResult
                        {
                            Name = "sort_data",
                            Success = false,
                            Error = "排序区域包含合并单元格，无法直接排序。请先取消合并单元格（unmerge_cells）后再排序。",
                            Suggestion = "请先调用 unmerge_cells 工具取消该区域的合并，然后再排序。"
                        };
                    }
                }

                // sortColumn 是列字母（如 "F"）或列序号字符串（如 "1"）或列名（如"销售额"）
                // ★★★ 必须转成 range 内部的相对列号（1-based），不是工作表绝对列号！
                // 错误做法：直接用 ColumnLetterToNumber("F")=6 作为 range.Cells[r, 6] 的列索引，
                // 当 range=E2:F14（2列）时，range.Cells[2, 6] 会跑到 J 列去（绝对位置），
                // 导致 Sort 报错"排序引用无效，请确保它在所要排序的数据内"。
                // 正确做法：用绝对列号减去 range 起始列号，得到相对列号。
                int absoluteColNum;
                if (int.TryParse(sortColumn, out int colIdx))
                {
                    absoluteColNum = colIdx;
                }
                else
                {
                    absoluteColNum = ColumnLetterToNumber(sortColumn);
                }

                // 计算 range 起始列号，把绝对列号转成 range 内部相对列号（1-based）
                int rangeStartCol = ((Range)range.Cells[1, 1]).Column;
                int sortColNum = absoluteColNum - rangeStartCol + 1;

                int totalCols = range.Columns.Count;

                // ★ 校验：sortColNum 必须在 range 范围内
                if (sortColNum < 1 || sortColNum > totalCols)
                {
                    Logger.Instance.Error("ExcelActions",
                        $"SortData: sort column out of range! absoluteCol={absoluteColNum}, rangeStartCol={rangeStartCol}, relativeCol={sortColNum}, totalCols={totalCols}");
                    return new ToolResult
                    {
                        Name = "sort_data",
                        Success = false,
                        Error = $"排序列 {sortColumn}（绝对列号 {absoluteColNum}）不在排序区域 {rangeAddress}（列范围 {rangeStartCol}-{rangeStartCol + totalCols - 1}）内。",
                        Suggestion = $"请确保 sort_column 在 range_address 范围内。range 起始列是 {ColumnIndexToLetter(rangeStartCol)}，请传该范围内的列字母。"
                    };
                }

                // ★ 智能检测：AI 传 has_header=false 时，检查第一行是否疑似表头。
                // 判定标准：第一行是文本，且第二行及以后至少有一列是数字。
                // 这是"表头+数据"的典型特征，纯文本表（如姓名/籍贯对照）不会被误判。
                // 检测到疑似表头时自动纠正为 has_header=true，继续执行（不拒绝，保证可用性）。
                // ★★★ 必须用 Value2 判断类型，不能用 Text！
                // Text 返回格式化后的字符串（如"1,234.56"/"¥100"/"####"），double.TryParse 会失败。
                // Value2 对数字返回 double，对文本返回 string，不经过格式化。
                bool headerAutoCorrected = false;
                if (!hasHeader && range.Rows.Count >= 2)
                {
                    bool firstRowAllText = true;
                    bool laterRowsHaveNumber = false;
                    int checkCols = Math.Min(totalCols, 10);
                    for (int c = 1; c <= checkCols; c++)
                    {
                        try
                        {
                            var firstCell = (Range)range.Cells[1, c];
                            var firstVal = firstCell.Value2;
                            // 第一行必须是文本（非空、非数字）
                            if (firstVal == null) { firstRowAllText = false; break; }
                            if (firstVal is double) { firstRowAllText = false; break; }
                            string firstText = firstVal.ToString();
                            if (string.IsNullOrEmpty(firstText)) { firstRowAllText = false; break; }
                            // 排除数字字符串（如"2024"）
                            if (double.TryParse(firstText, out _)) { firstRowAllText = false; break; }

                            // 检查后续行该列是否有数字（用 Value2 判断类型）
                            for (int r = 2; r <= range.Rows.Count; r++)
                            {
                                try
                                {
                                    var cell = (Range)range.Cells[r, c];
                                    var val = cell.Value2;
                                    if (val is double)
                                    {
                                        laterRowsHaveNumber = true;
                                        break;
                                    }
                                    if (val is string s && double.TryParse(s, out _))
                                    {
                                        laterRowsHaveNumber = true;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }

                    if (firstRowAllText && laterRowsHaveNumber)
                    {
                        Logger.Instance.Warning("ExcelActions",
                            $"SortData: SUSPECTED HEADER DETECTED! has_header=false but row1 is text and later rows have numbers. Auto-correcting to has_header=true");
                        hasHeader = true;
                        headerAutoCorrected = true;
                    }
                }

                // Key1 必须指向数据行的单元格（不是表头行）。
                int keyRow = hasHeader ? 2 : 1;
                Range sortKey = (Range)range.Cells[keyRow, sortColNum];

                // ★ 防御性检查：排序前记录每列第一行的值，排序后验证列顺序未被破坏。
                var colHeadersBefore = new string[totalCols + 1];
                for (int c = 1; c <= totalCols; c++)
                {
                    // ★ 用 Value2 而非 Text 做比对。Text 受格式和列宽影响（如"####"），会误报列交换。
                    try { colHeadersBefore[c] = ((Range)range.Cells[1, c]).Value2?.ToString() ?? ""; }
                    catch { colHeadersBefore[c] = ""; }
                }

                Logger.Instance.Info("ExcelActions",
                    $"SortData: sortKey address={sortKey.Address}, keyRow={keyRow}, sortColNum={sortColNum}, range rows={range.Rows.Count}, cols={totalCols}");
                for (int c = 1; c <= totalCols; c++)
                {
                    Logger.Instance.Info("ExcelActions", $"SortData: col{c} header before = '{colHeadersBefore[c]}'");
                }

                _app.DisplayAlerts = false;
                try
                {
                    // ★★★ 关键修复：Orientation 必须传 xlSortColumns！
                    // xlSortColumns=1: 以列为排序依据，重排行（正常排序行为）
                    // xlSortRows=2: 以行为排序依据，重排列（会导致列交换！）
                    // 不传 Orientation 时 Excel 会复用上次设置，可能恰好是 xlSortRows 导致列交换。
                    range.Sort(
                        Key1: sortKey,
                        Order1: descending ? XlSortOrder.xlDescending : XlSortOrder.xlAscending,
                        Header: hasHeader ? XlYesNoGuess.xlYes : XlYesNoGuess.xlNo,
                        Orientation: XlSortOrientation.xlSortColumns,
                        SortMethod: XlSortMethod.xlPinYin,
                        DataOption1: XlSortDataOption.xlSortNormal);
                }
                finally { _app.DisplayAlerts = true; }

                // ★ 排序后验证：检查列顺序是否被破坏（防止列交换）
                // ★ 用 Value2 而非 Text 做比对。Text 受格式和列宽影响（如"####"），会误报列交换。
                bool columnsSwapped = false;
                for (int c = 1; c <= totalCols; c++)
                {
                    try
                    {
                        string headerAfter = ((Range)range.Cells[1, c]).Value2?.ToString() ?? "";
                        if (headerAfter != colHeadersBefore[c])
                        {
                            columnsSwapped = true;
                            Logger.Instance.Error("ExcelActions",
                                $"SortData: COLUMN SWAP DETECTED! col{c} was '{colHeadersBefore[c]}', now '{headerAfter}'");
                        }
                    }
                    catch { }
                }

                if (columnsSwapped)
                {
                    Logger.Instance.Error("ExcelActions", "SortData: columns were swapped, sort orientation was wrong");
                    return new ToolResult
                    {
                        Name = "sort_data",
                        Success = false,
                        Error = "排序时检测到列顺序被破坏（列交换），排序未正确执行。请重试。"
                    };
                }

                Logger.Instance.Info("ExcelActions", "SortData: sort completed successfully, columns verified");
                if (headerAutoCorrected)
                {
                    return new ToolResult
                    {
                        Name = "sort_data",
                        Success = true,
                        Warning = "检测到第一行疑似表头（第一行文本+后续行数字），已自动按 has_header=true 排序（表头未参与）。下次请明确传入 has_header=true。"
                    };
                }
                return new ToolResult { Name = "sort_data", Success = true };
            }
            catch (Exception ex)
            {
                try { _app.DisplayAlerts = true; } catch { }
                Logger.Instance.Error("ExcelActions", "SortData failed", ex);
                return new ToolResult { Name = "sort_data", Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// 列字母转列序号（A=1, B=2, ..., Z=26, AA=27, ...）
        /// </summary>
        private static int ColumnLetterToNumber(string columnLetter)
        {
            int result = 0;
            foreach (char c in columnLetter.ToUpper())
            {
                if (c >= 'A' && c <= 'Z')
                {
                    result = result * 26 + (c - 'A' + 1);
                }
            }
            return result;
        }

        public ToolResult FilterData(string rangeAddress, int columnIndex, string criteria)
        {
            try
            {
                var range = TryResolveRange(rangeAddress, "filter_data", out string fErr, out string fSug);
                if (range == null)
                {
                    return new ToolResult { Name = "filter_data", Success = false, Error = fErr, Suggestion = fSug };
                }
                var ws = range.Worksheet;
                _app.DisplayAlerts = false;
                try
                {
                    // ★ 避免 toggle：先检查 AutoFilter 是否已开启
                    // 如果工作表已开启 AutoFilter 且范围匹配，直接设置条件；
                    // 如果未开启，则调用 AutoFilter 开启。
                    // 直接调用 range.AutoFilter() 在已开启时会关闭（toggle），
                    // 导致"过滤了反而显示全部"的反直觉行为。
                    bool filterMode = false;
                    try { filterMode = ws.AutoFilterMode; } catch { }

                    if (filterMode)
                    {
                        // 已开启筛选：直接修改指定列的筛选条件
                        // 通过 AutoFilter.Filters 访问对应字段的筛选
                        try
                        {
                            var af = ws.AutoFilter;
                            if (af != null && af.Range != null)
                            {
                                // 确保筛选范围包含目标范围
                                string afAddr = af.Range.Address;
                                string rangeAddr = range.Address;
                                if (afAddr == rangeAddr || afAddr.Contains(rangeAddr))
                                {
                                    // 范围匹配或更大，直接重新应用筛选以确保条件生效
                                    range.AutoFilter(
                                        Field: columnIndex,
                                        Criteria1: criteria,
                                        Operator: XlAutoFilterOperator.xlAnd,
                                        VisibleDropDown: true);
                                }
                                else
                                {
                                    // 范围不同：先关旧的，再开新的
                                    ws.AutoFilterMode = false;
                                    range.AutoFilter(
                                        Field: columnIndex,
                                        Criteria1: criteria);
                                }
                            }
                            else
                            {
                                range.AutoFilter(
                                    Field: columnIndex,
                                    Criteria1: criteria);
                            }
                        }
                        catch
                        {
                            // 兜底：直接调用
                            range.AutoFilter(
                                Field: columnIndex,
                                Criteria1: criteria);
                        }
                    }
                    else
                    {
                        // 未开启筛选：正常开启
                        range.AutoFilter(
                            Field: columnIndex,
                            Criteria1: criteria);
                    }
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
                var range = TryResolveRange(address, "merge_cells", out string mErr, out string mSug);
                if (range == null)
                {
                    return new ToolResult { Name = "merge_cells", Success = false, Error = mErr, Suggestion = mSug };
                }
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
                var range = TryResolveRange(address, "unmerge_cells", out string uErr, out string uSug);
                if (range == null)
                {
                    return new ToolResult { Name = "unmerge_cells", Success = false, Error = uErr, Suggestion = uSug };
                }
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
                var range = TryResolveRange(address, "set_cell_style", out string stErr, out string stSug);
                if (range == null)
                {
                    return new ToolResult { Name = "set_cell_style", Success = false, Error = stErr, Suggestion = stSug };
                }
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
                var src = TryResolveRange(sourceAddress, "copy_range[source]", out string srcErr, out string srcSug);
                if (src == null)
                {
                    return new ToolResult { Name = "copy_range", Success = false, Error = srcErr, Suggestion = srcSug };
                }
                var dst = TryResolveRange(destAddress, "copy_range[dest]", out string dstErr, out string dstSug);
                if (dst == null)
                {
                    return new ToolResult { Name = "copy_range", Success = false, Error = dstErr, Suggestion = dstSug };
                }
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
                var range = TryResolveRange(address, "clear_range", out string cErr, out string cSug);
                if (range == null)
                {
                    return new ToolResult { Name = "clear_range", Success = false, Error = cErr, Suggestion = cSug };
                }
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
                // ★ 必须用 EntireRow.Insert() 插入整行，所有列才会一起下移。
                // 错误做法：ws.Range["A"+row].Insert(xlShiftDown) 只会下移 A 列，
                // 其他列（如 B 列）数据保持原位，导致列错位。
                var range = ws.Range["A" + row];
                for (int i = 0; i < count; i++)
                {
                    range.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
                }
                Logger.Instance.Info("ExcelActions",
                    $"InsertRows: inserted {count} row(s) at row {row} (entire row shift)");
                return new ToolResult { Name = "insert_rows", Success = true };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ExcelActions", "InsertRows failed", ex);
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
                // ★ 必须用 EntireColumn.Insert() 插入整列，所有行才会一起右移。
                // 错误做法：range.Insert(xlShiftToRight) 只会右移第 1 行，
                // 其他行数据保持原位，导致行错位。
                for (int i = 0; i < count; i++)
                {
                    range.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
                }
                Logger.Instance.Info("ExcelActions",
                    $"InsertColumns: inserted {count} column(s) at col {colLetter} (entire column shift)");
                return new ToolResult { Name = "insert_columns", Success = true };
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ExcelActions", "InsertColumns failed", ex);
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
                // freeze_panes 的 address 不含 sheet 前缀（只对当前活动 sheet 操作）
                // 用 TryResolveRange 检测地址格式有效性
                var cell = TryResolveRange(address, "freeze_panes", out string fErr, out string fSug);
                if (cell == null)
                {
                    return new ToolResult { Name = "freeze_panes", Success = false, Error = fErr, Suggestion = fSug };
                }
                _app.ActiveWindow.Split = false;  // 先解除现有冻结
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
                range = TryResolveRange(address, "apply_conditional_format", out string acfErr, out string acfSug);
                if (range == null)
                {
                    // 尝试作为列名解析：扫描表头行找到匹配列，转为列字母地址
                    string resolved = TryResolveColumnNameToRange(address);
                    if (resolved != null)
                    {
                        Diagnostics.Logger.Instance.Info("ExcelActionsImpl", $"ApplyConditionalFormat: 列名解析 [{address}] -> [{resolved}]");
                        range = TryResolveRange(resolved, "apply_conditional_format", out acfErr, out acfSug);
                    }
                }
                if (range == null)
                {
                    return new ToolResult
                    {
                        Name = "apply_conditional_format",
                        Success = false,
                        Error = acfErr ?? $"address '{address}' 不是有效的 A1 地址，且未在表头中找到匹配列名",
                        Suggestion = acfSug ?? "请传入 A1 格式地址（如 A1:A100）或表头列名（如\"销售额\"）",
                    };
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
                    // ★ 用 Value2 获取原始文本做列名匹配。Text 受格式影响（如"####"）会匹配失败。
                    string text = "";
                    try { text = cellVal?.Value2?.ToString() ?? ""; } catch { }
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
