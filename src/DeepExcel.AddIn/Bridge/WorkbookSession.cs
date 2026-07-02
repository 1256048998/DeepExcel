using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Security;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// ★ 工作簿级会话：每个工作簿（FullName 唯一标识）拥有独立的
    /// PythonSidecar（AI 对话上下文）+ 附件列表 + 状态标志。
    /// 这样用户在 A 工作簿聊天、切换到 B 工作簿，两边上下文完全隔离。
    /// </summary>
    public class WorkbookSession : IDisposable
    {
        /// <summary>工作簿完整路径（字典 key）</summary>
        public string WorkbookKey { get; private set; }

        /// <summary>工作簿显示名（不带路径）</summary>
        public string WorkbookName { get; private set; }

        /// <summary>最后使用时间（LRU 回收用）</summary>
        public DateTime LastUsedTime { get; private set; } = DateTime.Now;

        /// <summary>该工作簿的 AI sidecar 进程（独立对话上下文）</summary>
        public PythonSidecar Sidecar { get; }

        /// <summary>附件列表（文件名 -> 绝对路径）</summary>
        public Dictionary<string, string> Attachments { get; } = new Dictionary<string, string>();

        /// <summary>该 session 下一次 user_message 的 session_id（每次递增）</summary>
        private int _sessionSeq;

        /// <summary>待回答的澄清问题（session 级，避免跨工作簿串扰）</summary>
        public string PendingClarifyQuestion;

        /// <summary>
        /// ★ 当前对话 ID。新建对话时生成新 GUID，继续历史对话时用历史的 ID。
        /// null 表示尚未开始对话（首次用户消息时自动创建）。
        /// </summary>
        public string CurrentConversationId { get; private set; }

        /// <summary>
        /// ★ 当前对话的消息列表（in-memory）。
        /// 每次 user_message/tool_call/stream_end 时追加，stream_end 时持久化。
        /// 前端通过 get_current_messages 拉取这份列表恢复 UI。
        /// </summary>
        public List<DeepExcel.AddIn.Collaboration.HistoryMessage> ConversationMessages { get; }
            = new List<DeepExcel.AddIn.Collaboration.HistoryMessage>();

        /// <summary>该 session 是否正忙（流式响应中）</summary>
        private bool _isBusy;
        private DateTime _busySince;
        private static readonly TimeSpan BusyTimeout = TimeSpan.FromMinutes(5);

        public bool IsBusy
        {
            get
            {
                if (_isBusy && (DateTime.Now - _busySince) > BusyTimeout)
                {
                    Logger.Instance.Warning("WorkbookSession",
                        $"IsBusy timeout ({WorkbookName}), auto-resetting");
                    _isBusy = false;
                }
                return _isBusy;
            }
            set
            {
                _isBusy = value;
                if (value) _busySince = DateTime.Now;
            }
        }

        /// <summary>附件存储目录（按 session 隔离在 %LOCALAPPDATA%\DeepExcel\Attachments\{workbookKeyHash}）</summary>
        public string AttachmentsDir { get; }

        public WorkbookSession(string workbookKey, string workbookName, IExcelActions excelActions,
            Microsoft.Office.Interop.Excel.Application excelApp, Control uiControl, SecurityGateway securityGateway)
        {
            WorkbookKey = workbookKey;
            WorkbookName = workbookName;

            Sidecar = new PythonSidecar(excelActions, excelApp, uiControl, securityGateway);

            // 附件目录：用 workboookKey 的 hash 做目录名，避免路径含特殊字符
            int hash = Math.Abs(workbookKey.GetHashCode());
            AttachmentsDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DeepExcel", "Attachments", hash.ToString("X8"));
            Directory.CreateDirectory(AttachmentsDir);
        }

        /// <summary>生成下一个 session_id（每次用户消息递增）</summary>
        public string NextSessionId()
        {
            _sessionSeq++;
            return $"wb_{_sessionSeq}";
        }

        /// <summary>附件大小限制：单文件最大 10MB</summary>
        private const long MaxAttachmentSize = 10 * 1024 * 1024; // 10 MB

        /// <summary>
        /// 允许的附件文件类型白名单。
        /// 只允许文本类和常见文档格式，禁止可执行文件。
        /// </summary>
        private static readonly HashSet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // 文本/代码
            ".txt", ".md", ".csv", ".json", ".xml", ".yaml", ".yml",
            ".py", ".cs", ".js", ".ts", ".html", ".css", ".sql",
            ".r", ".m", ".cpp", ".c", ".h", ".java", ".go", ".rs",
            ".sh", ".bat", ".ps1",
            // 文档
            ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".pdf", ".rtf",
            // 图片（仅参考用，不读取内容）
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg",
            // 其他
            ".log", ".ini", ".conf", ".config", ".env",
        };

        /// <summary>
        /// 保存上传的附件。返回 { fileName, filePath, size } 元信息。
        /// fileBase64 是文件内容的 base64 编码。
        /// </summary>
        public Dictionary<string, object> AddAttachment(string fileName, string fileBase64)
        {
            // 文件名安全化
            string safeName = Path.GetFileName(fileName) ?? "unnamed";

            // 文件类型白名单校验
            string ext = Path.GetExtension(safeName) ?? "";
            if (!AllowedExtensions.Contains(ext))
            {
                throw new ArgumentException(
                    $"不支持的文件类型: {ext}。支持的类型: 文本、代码、Office 文档、图片。");
            }

            byte[] bytes;
            try
            {
                bytes = Convert.FromBase64String(fileBase64);
            }
            catch (FormatException)
            {
                throw new ArgumentException("文件内容不是有效的 base64 编码");
            }

            // 大小限制校验
            if (bytes.Length > MaxAttachmentSize)
            {
                throw new ArgumentException(
                    $"文件过大 ({bytes.Length / 1024 / 1024}MB)，最大支持 {MaxAttachmentSize / 1024 / 1024}MB");
            }

            // 防止重名覆盖，加序号
            string targetPath = Path.Combine(AttachmentsDir, safeName);
            if (File.Exists(targetPath))
            {
                string nameWithoutExt = Path.GetFileNameWithoutExtension(safeName);
                string extWithDot = Path.GetExtension(safeName);
                int n = 1;
                while (File.Exists(Path.Combine(AttachmentsDir, $"{nameWithoutExt}_{n}{extWithDot}")))
                    n++;
                safeName = $"{nameWithoutExt}_{n}{extWithDot}";
                targetPath = Path.Combine(AttachmentsDir, safeName);
            }

            File.WriteAllBytes(targetPath, bytes);

            Attachments[safeName] = targetPath;

            Logger.Instance.Info("WorkbookSession", $"[{WorkbookName}] attachment added: {safeName}, size={bytes.Length}");

            return new Dictionary<string, object>
            {
                ["fileName"] = safeName,
                ["filePath"] = targetPath,
                ["size"] = bytes.Length,
            };
        }

        /// <summary>删除附件</summary>
        public bool RemoveAttachment(string fileName)
        {
            if (!Attachments.ContainsKey(fileName)) return false;
            try
            {
                string path = Attachments[fileName];
                if (File.Exists(path)) File.Delete(path);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("WorkbookSession", $"Delete attachment failed: {fileName}, {ex.Message}");
            }
            Attachments.Remove(fileName);
            return true;
        }

        /// <summary>获取附件列表（前端展示用）</summary>
        public List<Dictionary<string, object>> GetAttachmentList()
        {
            var list = new List<Dictionary<string, object>>();
            foreach (var kvp in Attachments)
            {
                long size = 0;
                try
                {
                    if (File.Exists(kvp.Value))
                        size = new FileInfo(kvp.Value).Length;
                }
                catch { }
                list.Add(new Dictionary<string, object>
                {
                    ["fileName"] = kvp.Key,
                    ["size"] = size,
                });
            }
            return list;
        }

        /// <summary>构建该 session 的 Excel 上下文（含附件信息）</summary>
        public object BuildContext(IExcelActions excelActions)
        {
            try
            {
                object workbook = null;
                object selection = null;
                try { workbook = excelActions.ReadWorkbook(); } catch { }
                try { selection = excelActions.GetSelection(); } catch { }

                var attachmentList = new List<Dictionary<string, object>>();
                foreach (var kvp in Attachments)
                {
                    long size = 0;
                    try { if (File.Exists(kvp.Value)) size = new FileInfo(kvp.Value).Length; } catch { }
                    attachmentList.Add(new Dictionary<string, object>
                    {
                        ["name"] = kvp.Key,
                        ["size"] = size,
                        ["path"] = kvp.Value,
                    });
                }

                return new
                {
                    workbookName = WorkbookName,
                    workbook = workbook,
                    selection = selection,
                    attachments = attachmentList,
                    timestamp = DateTime.Now.ToString("o"),
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("WorkbookSession", "BuildContext failed: " + ex.Message);
                return new { workbookName = WorkbookName, attachments = new List<object>() };
            }
        }

        /// <summary>更新工作簿 key 和名称（另存为后调用）</summary>
        public void UpdateKey(string newKey, string newName)
        {
            WorkbookKey = newKey;
            if (!string.IsNullOrEmpty(newName))
                WorkbookName = newName;
            LastUsedTime = DateTime.Now;
        }

        /// <summary>更新最后使用时间（每次用户消息/工具调用时调用，用于 LRU）</summary>
        public void Touch()
        {
            LastUsedTime = DateTime.Now;
        }

        // ============= 多对话历史管理 =============

        /// <summary>
        /// 列出该工作簿的所有历史对话（不含 messages，只含元信息）。
        /// 前端用于展示历史对话列表。
        /// </summary>
        public List<Collaboration.Conversation> ListConversations()
        {
            var convs = Collaboration.ConversationHistory.LoadAll(WorkbookKey);
            // 不返回 messages（避免传输大量数据），只返回元信息
            return convs.Select(c => new Collaboration.Conversation
            {
                Id = c.Id,
                Title = c.Title,
                CreatedAt = c.CreatedAt,
                UpdatedAt = c.UpdatedAt,
                WorkbookName = c.WorkbookName,
                Messages = null  // 列表不返回 messages
            }).ToList();
        }

        /// <summary>
        /// ★ 新建对话：把当前对话存盘（如有），清空内存，生成新 conversationId。
        /// sidecar 进程重启以清除 AI 上下文。
        /// </summary>
        public void NewConversation()
        {
            // 把当前对话存盘
            Logger.Instance.Info("WorkbookSession",
                $"[{WorkbookName}] NewConversation: msgCount={ConversationMessages.Count}, currentId={CurrentConversationId ?? "null"}");
            if (ConversationMessages.Count > 0 && !string.IsNullOrEmpty(CurrentConversationId))
            {
                SaveCurrentConversation();
                Logger.Instance.Info("WorkbookSession", $"[{WorkbookName}] NewConversation: saved previous conversation");
            }
            else
            {
                Logger.Instance.Info("WorkbookSession", $"[{WorkbookName}] NewConversation: skip save (empty or no id)");
            }

            // 清空内存
            ConversationMessages.Clear();
            CurrentConversationId = Guid.NewGuid().ToString("N");
            PendingClarifyQuestion = null;
            IsBusy = false;

            Logger.Instance.Info("WorkbookSession", $"[{WorkbookName}] new conversation: {CurrentConversationId}");
        }

        /// <summary>
        /// ★ 继续历史对话：加载指定对话的 messages 到内存，设为当前对话。
        /// 返回该对话的 messages 列表（前端用于恢复显示）。
        /// </summary>
        public List<Collaboration.HistoryMessage> ContinueConversation(string conversationId)
        {
            var convs = Collaboration.ConversationHistory.LoadAll(WorkbookKey);
            var target = convs.FirstOrDefault(c => c.Id == conversationId);
            if (target == null)
            {
                Logger.Instance.Warning("WorkbookSession",
                    $"[{WorkbookName}] continue conversation failed: id={conversationId} not found");
                return null;
            }

            // 先把当前对话存盘（如有）
            if (ConversationMessages.Count > 0 && !string.IsNullOrEmpty(CurrentConversationId)
                && CurrentConversationId != conversationId)
            {
                SaveCurrentConversation();
            }

            // 加载目标对话到内存
            ConversationMessages.Clear();
            if (target.Messages != null)
            {
                foreach (var m in target.Messages)
                {
                    m.streaming = false;  // 恢复时强制 false
                    ConversationMessages.Add(m);
                }
            }
            CurrentConversationId = target.Id;
            PendingClarifyQuestion = null;
            IsBusy = false;

            Logger.Instance.Info("WorkbookSession",
                $"[{WorkbookName}] continue conversation: {conversationId}, msgs={ConversationMessages.Count}");

            return target.Messages ?? new List<Collaboration.HistoryMessage>();
        }

        /// <summary>删除指定历史对话</summary>
        public bool DeleteConversation(string conversationId)
        {
            try
            {
                Collaboration.ConversationHistory.DeleteConversation(WorkbookKey, conversationId);
                // 如果删的是当前对话，也清空内存
                if (CurrentConversationId == conversationId)
                {
                    ConversationMessages.Clear();
                    CurrentConversationId = null;
                }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("WorkbookSession", $"DeleteConversation failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>追加用户消息到当前对话</summary>
        public void AppendUserMessage(string content)
        {
            // 首次用户消息时自动创建 conversationId
            if (string.IsNullOrEmpty(CurrentConversationId))
            {
                CurrentConversationId = Guid.NewGuid().ToString("N");
            }

            ConversationMessages.Add(new Collaboration.HistoryMessage
            {
                role = "user",
                content = content
            });
        }

        /// <summary>追加助手消息到历史（流式增量追加）</summary>
        public void AppendAssistantDelta(string delta)
        {
            var last = ConversationMessages.Count > 0
                ? ConversationMessages[ConversationMessages.Count - 1]
                : null;
            if (last != null && last.role == "assistant" && last.streaming)
            {
                last.content += delta;
            }
            else
            {
                ConversationMessages.Add(new Collaboration.HistoryMessage
                {
                    role = "assistant",
                    content = delta,
                    streaming = true
                });
            }
        }

        /// <summary>追加工具调用消息到历史</summary>
        public void AppendToolCall(string toolName)
        {
            var last = ConversationMessages.Count > 0
                ? ConversationMessages[ConversationMessages.Count - 1]
                : null;
            if (last != null && last.role == "tool" && last.toolGroup != null)
            {
                var list = new List<string>(last.toolGroup) { toolName };
                last.toolGroup = list.ToArray();
            }
            else
            {
                ConversationMessages.Add(new Collaboration.HistoryMessage
                {
                    role = "tool",
                    content = "",
                    toolGroup = new[] { toolName }
                });
            }
        }

        /// <summary>追加 clarify 消息到历史</summary>
        public void AppendClarify(string question, string[] options)
        {
            ConversationMessages.Add(new Collaboration.HistoryMessage
            {
                role = "assistant",
                content = question,
                type = "clarify",
                options = options
            });
        }

        /// <summary>
        /// stream_end 时调用：重置 streaming 标志，持久化当前对话到磁盘。
        /// </summary>
        public void OnStreamEnd()
        {
            foreach (var m in ConversationMessages)
            {
                if (m.streaming) m.streaming = false;
            }
            if (!string.IsNullOrEmpty(CurrentConversationId) && ConversationMessages.Count > 0)
            {
                SaveCurrentConversation();
            }
        }

        /// <summary>把当前内存对话保存到磁盘</summary>
        private void SaveCurrentConversation()
        {
            try
            {
                Logger.Instance.Info("WorkbookSession",
                    $"[{WorkbookName}] SaveCurrentConversation: id={CurrentConversationId}, msgCount={ConversationMessages.Count}");
                var convs = Collaboration.ConversationHistory.LoadAll(WorkbookKey);

                // 找已有或新建
                var existing = convs.FirstOrDefault(c => c.Id == CurrentConversationId);
                if (existing == null)
                {
                    existing = new Collaboration.Conversation
                    {
                        Id = CurrentConversationId,
                        CreatedAt = DateTime.Now,
                        WorkbookName = WorkbookName
                    };
                    convs.Add(existing);
                }

                // 标题：首条用户消息前 30 字
                var firstUser = ConversationMessages.FirstOrDefault(m => m.role == "user");
                existing.Title = firstUser?.content?.Length > 30
                    ? firstUser.content.Substring(0, 30) + "..."
                    : (firstUser?.content ?? "新对话");
                existing.UpdatedAt = DateTime.Now;
                existing.Messages = new List<Collaboration.HistoryMessage>(ConversationMessages);

                Collaboration.ConversationHistory.SaveAll(WorkbookKey, WorkbookName, convs);
                Logger.Instance.Info("WorkbookSession",
                    $"[{WorkbookName}] SaveCurrentConversation: saved, title='{existing.Title}', totalConvs={convs.Count}");
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("WorkbookSession", $"SaveCurrentConversation failed: {ex.Message}");
            }
        }

        /// <summary>清空当前对话内存（不删磁盘历史）</summary>
        public void ClearCurrentMessages()
        {
            ConversationMessages.Clear();
            CurrentConversationId = null;
            PendingClarifyQuestion = null;
            IsBusy = false;
        }

        public void Dispose()
        {
            try
            {
                Sidecar?.Dispose();
            }
            catch { }
            // 不删附件文件，用户可能下次打开还想要
        }
    }
}
