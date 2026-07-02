using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Collaboration
{
    /// <summary>
    /// ★ 对话历史持久化：每个工作簿一个 JSON 文件，包含该工作簿的所有对话列表。
    /// 存储路径：%LOCALAPPDATA%\DeepExcel\history\{workbook-hash}.json
    /// 文件结构：{ workbookKey, workbookName, conversations: [{ id, title, createdAt, updatedAt, messages: [...] }] }
    /// 每个对话最多保留 MaxMessagesPerConversation 条消息（滚动覆盖）。
    /// 每个工作簿最多保留 MaxConversations 个对话（超出时删最旧的）。
    /// </summary>
    public static class ConversationHistory
    {
        private static readonly string HistoryDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DeepExcel", "history");

        private const int MaxMessagesPerConversation = 200;
        private const int MaxConversations = 20;
        private const long MaxTotalSizeBytes = 50 * 1024 * 1024;

        private static string GetFileName(string workbookKey)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(workbookKey));
                return BitConverter.ToString(bytes, 0, 16).Replace("-", "").ToLowerInvariant() + ".json";
            }
        }

        private static string GetFilePath(string workbookKey)
        {
            return Path.Combine(HistoryDir, GetFileName(workbookKey));
        }

        /// <summary>加载工作簿的所有对话。无历史返回空列表。</summary>
        public static List<Conversation> LoadAll(string workbookKey)
        {
            try
            {
                var path = GetFilePath(workbookKey);
                if (!File.Exists(path)) return new List<Conversation>();

                var json = File.ReadAllText(path);
                var data = JsonSerializer.Deserialize<HistoryFile>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
                return data?.Conversations ?? new List<Conversation>();
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ConversationHistory", $"LoadAll failed: {ex.Message}");
                return new List<Conversation>();
            }
        }

        /// <summary>保存工作簿的所有对话（覆盖写）。</summary>
        public static void SaveAll(string workbookKey, string workbookName, List<Conversation> conversations)
        {
            try
            {
                Directory.CreateDirectory(HistoryDir);

                // 限制对话数量：超出时删最旧的
                if (conversations != null && conversations.Count > MaxConversations)
                {
                    conversations = conversations
                        .OrderByDescending(c => c.UpdatedAt ?? c.CreatedAt)
                        .Take(MaxConversations)
                        .ToList();
                }

                // 限制每个对话的消息数量
                if (conversations != null)
                {
                    foreach (var c in conversations)
                    {
                        if (c.Messages != null && c.Messages.Count > MaxMessagesPerConversation)
                        {
                            c.Messages = c.Messages
                                .Skip(c.Messages.Count - MaxMessagesPerConversation)
                                .ToList();
                        }
                    }
                }

                var data = new HistoryFile
                {
                    WorkbookKey = workbookKey,
                    WorkbookName = workbookName,
                    SavedAt = DateTime.Now,
                    Conversations = conversations ?? new List<Conversation>()
                };

                var path = GetFilePath(workbookKey);
                var json = JsonSerializer.Serialize(data, new JsonSerializerOptions
                {
                    WriteIndented = false,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
                File.WriteAllText(path, json, Encoding.UTF8);

                CleanupTotalSize();
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ConversationHistory", $"SaveAll failed: {ex.Message}");
            }
        }

        /// <summary>删除工作簿的所有对话历史。</summary>
        public static void DeleteAll(string workbookKey)
        {
            try
            {
                var path = GetFilePath(workbookKey);
                if (File.Exists(path)) File.Delete(path);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ConversationHistory", $"DeleteAll failed: {ex.Message}");
            }
        }

        /// <summary>删除指定对话。</summary>
        public static void DeleteConversation(string workbookKey, string conversationId)
        {
            try
            {
                var convs = LoadAll(workbookKey);
                var removed = convs.RemoveAll(c => c.Id == conversationId);
                if (removed > 0)
                {
                    var wbName = convs.FirstOrDefault()?.WorkbookName ?? "";
                    SaveAll(workbookKey, wbName, convs);
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ConversationHistory", $"DeleteConversation failed: {ex.Message}");
            }
        }

        private static void CleanupTotalSize()
        {
            try
            {
                var dir = new DirectoryInfo(HistoryDir);
                if (!dir.Exists) return;

                var files = dir.GetFiles("*.json").OrderBy(f => f.LastWriteTime).ToList();
                long totalSize = files.Sum(f => f.Length);

                if (totalSize <= MaxTotalSizeBytes) return;

                foreach (var f in files)
                {
                    if (totalSize <= MaxTotalSizeBytes * 0.8) break;
                    try
                    {
                        totalSize -= f.Length;
                        f.Delete();
                        Logger.Instance.Info("ConversationHistory", $"Cleanup: deleted old history {f.Name}");
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ConversationHistory", $"CleanupTotalSize failed: {ex.Message}");
            }
        }
    }

    /// <summary>单个对话</summary>
    public class Conversation
    {
        public string Id { get; set; }
        public string Title { get; set; }           // 首条用户消息的前 N 字
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public string WorkbookName { get; set; }    // 冗余字段，历史列表展示用
        public List<HistoryMessage> Messages { get; set; } = new List<HistoryMessage>();
    }

    /// <summary>历史文件结构</summary>
    public class HistoryFile
    {
        public string WorkbookKey { get; set; }
        public string WorkbookName { get; set; }
        public DateTime SavedAt { get; set; }
        public List<Conversation> Conversations { get; set; }
    }

    /// <summary>
    /// 历史消息（前端展示用 + sidecar 恢复用）。
    /// 字段与前端 Message 类型对应。
    /// </summary>
    public class HistoryMessage
    {
        public string role { get; set; }       // user / assistant / tool
        public string content { get; set; }
        public string type { get; set; }        // clarify 等特殊类型（可选）
        public string[] options { get; set; }   // clarify 选项（可选）
        public string[] toolGroup { get; set; } // 工具调用组（可选）
        public bool streaming { get; set; }
    }
}
