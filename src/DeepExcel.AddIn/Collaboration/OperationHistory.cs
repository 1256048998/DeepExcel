using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Collaboration
{
    /// <summary>
    /// 操作历史管理器 - 记录所有操作并支持回放
    /// </summary>
    public class OperationHistoryManager
    {
        private static readonly string HistoryDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DeepExcel", "history");

        private static OperationHistoryManager _instance;
        public static OperationHistoryManager Instance => _instance ??= new OperationHistoryManager();

        private readonly List<OperationRecord> _currentSession = new();
        private readonly Dictionary<string, List<OperationRecord>> _savedSessions = new();

        public event Action<OperationRecord> OnOperationRecorded;
        public event Action<string, List<OperationRecord>> OnSessionSaved;

        private OperationHistoryManager()
        {
            try
            {
                Directory.CreateDirectory(HistoryDir);
                LoadSavedSessions();
            }
            catch { }
        }

        /// <summary>
        /// 记录操作
        /// </summary>
        public void Record(OperationRecord record)
        {
            record.Timestamp = DateTime.Now;
            record.SessionId = GetCurrentSessionId();
            _currentSession.Add(record);
            OnOperationRecorded?.Invoke(record);
        }

        /// <summary>
        /// 获取当前会话的操作记录
        /// </summary>
        public List<OperationRecord> GetCurrentSessionHistory()
        {
            return _currentSession.ToList();
        }

        /// <summary>
        /// 获取所有保存的会话
        /// </summary>
        public List<SessionSummary> GetSavedSessions()
        {
            return _savedSessions.Select(kv => new SessionSummary
            {
                SessionId = kv.Key,
                OperationCount = kv.Value.Count,
                FirstOperation = kv.Value.FirstOrDefault()?.Timestamp ?? DateTime.MinValue,
                LastOperation = kv.Value.LastOrDefault()?.Timestamp ?? DateTime.MinValue,
                WorkbookName = kv.Value.FirstOrDefault()?.WorkbookName
            }).ToList();
        }

        /// <summary>
        /// 获取指定会话的操作记录
        /// </summary>
        public List<OperationRecord> GetSessionHistory(string sessionId)
        {
            if (_savedSessions.ContainsKey(sessionId))
                return _savedSessions[sessionId].ToList();
            return new List<OperationRecord>();
        }

        /// <summary>
        /// 保存当前会话
        /// </summary>
        public bool SaveCurrentSession(string name = null)
        {
            if (_currentSession.Count == 0) return false;

            try
            {
                var sessionId = GetCurrentSessionId();
                var fileName = $"{sessionId}.json";
                if (!string.IsNullOrEmpty(name))
                    fileName = $"{name}-{sessionId}.json";

                var path = Path.Combine(HistoryDir, fileName);
                var json = JsonSerializer.Serialize(_currentSession, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);

                _savedSessions[sessionId] = _currentSession.ToList();
                OnSessionSaved?.Invoke(sessionId, _currentSession.ToList());
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 加载保存的会话
        /// </summary>
        public bool LoadSession(string sessionId)
        {
            try
            {
                var files = Directory.GetFiles(HistoryDir, $"{sessionId}*.json");
                if (files.Length == 0) return false;

                var json = File.ReadAllText(files[0]);
                var records = JsonSerializer.Deserialize<List<OperationRecord>>(json);
                if (records != null)
                {
                    _savedSessions[sessionId] = records;
                    return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 删除保存的会话
        /// </summary>
        public bool DeleteSession(string sessionId)
        {
            try
            {
                var files = Directory.GetFiles(HistoryDir, $"{sessionId}*.json");
                foreach (var file in files)
                {
                    File.Delete(file);
                }
                _savedSessions.Remove(sessionId);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 导出操作记录（用于分享）
        /// </summary>
        public string ExportHistory(string sessionId, ExportFormat format = ExportFormat.Json)
        {
            var records = GetSessionHistory(sessionId);
            if (records.Count == 0) return "";

            switch (format)
            {
                case ExportFormat.Json:
                    return JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true });
                case ExportFormat.Markdown:
                    return ToMarkdown(records);
                case ExportFormat.Csv:
                    return ToCsv(records);
                default:
                    return "";
            }
        }

        /// <summary>
        /// 导入操作记录
        /// </summary>
        public bool ImportHistory(string json)
        {
            try
            {
                var records = JsonSerializer.Deserialize<List<OperationRecord>>(json);
                if (records != null && records.Count > 0)
                {
                    var sessionId = records.First().SessionId ?? Guid.NewGuid().ToString();
                    _savedSessions[sessionId] = records;
                    return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 清空当前会话
        /// </summary>
        public void ClearCurrentSession()
        {
            _currentSession.Clear();
        }

        /// <summary>
        /// 获取当前会话ID
        /// </summary>
        private string GetCurrentSessionId()
        {
            return System.Diagnostics.Process.GetCurrentProcess().Id.ToString();
        }

        private void LoadSavedSessions()
        {
            try
            {
                var files = Directory.GetFiles(HistoryDir, "*.json");
                foreach (var file in files)
                {
                    var json = File.ReadAllText(file);
                    var records = JsonSerializer.Deserialize<List<OperationRecord>>(json);
                    if (records != null && records.Count > 0)
                    {
                        var sessionId = records.First().SessionId ?? Path.GetFileNameWithoutExtension(file);
                        _savedSessions[sessionId] = records;
                    }
                }
            }
            catch { }
        }

        private string ToMarkdown(List<OperationRecord> records)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("# DeepExcel 操作历史");
            sb.AppendLine($"生成时间: {DateTime.Now}");
            sb.AppendLine($"操作数: {records.Count}");
            sb.AppendLine();
            sb.AppendLine("| 时间 | 工作簿 | 操作类型 | 工具 | 结果 |");
            sb.AppendLine("|------|--------|----------|------|------|");
            foreach (var r in records)
            {
                sb.AppendLine($"| {r.Timestamp:HH:mm:ss} | {r.WorkbookName} | {r.OperationType} | {r.ToolName} | {(r.Success ? "✓" : "✗")} |");
            }
            return sb.ToString();
        }

        private string ToCsv(List<OperationRecord> records)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("时间,工作簿,操作类型,工具名,参数,结果,错误");
            foreach (var r in records)
            {
                var args = r.Arguments != null ? JsonSerializer.Serialize(r.Arguments) : "";
                sb.AppendLine($"{r.Timestamp:yyyy-MM-dd HH:mm:ss},{r.WorkbookName},{r.OperationType},{r.ToolName},{args},{(r.Success ? "成功" : "失败")},{r.Error ?? ""}");
            }
            return sb.ToString();
        }
    }

    public enum ExportFormat
    {
        Json,
        Markdown,
        Csv
    }

    public class OperationRecord
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString();

        [JsonPropertyName("sessionId")]
        public string SessionId { get; set; }

        [JsonPropertyName("timestamp")]
        public DateTime Timestamp { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("operationType")]
        public string OperationType { get; set; }  // "tool_call" | "model_call" | "user_message"

        [JsonPropertyName("toolName")]
        public string ToolName { get; set; }

        [JsonPropertyName("arguments")]
        public Dictionary<string, object> Arguments { get; set; }

        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("error")]
        public string Error { get; set; }

        [JsonPropertyName("data")]
        public object Data { get; set; }

        [JsonPropertyName("tokenUsage")]
        public TokenUsage TokenUsage { get; set; }
    }

    public class TokenUsage
    {
        public int InputTokens { get; set; }
        public int OutputTokens { get; set; }
    }

    public class SessionSummary
    {
        public string SessionId { get; set; }
        public int OperationCount { get; set; }
        public DateTime FirstOperation { get; set; }
        public DateTime LastOperation { get; set; }
        public string WorkbookName { get; set; }
    }
}
