using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace DeepExcel.AddIn.Diagnostics
{
    /// <summary>
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warning = 2,
        Error = 3
    }

    /// <summary>
    /// 日志条目
    /// </summary>
    public class LogEntry
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public LogLevel Level { get; set; }
        public string Category { get; set; }
        public string Message { get; set; }
        public Exception Exception { get; set; }
        public Dictionary<string, object> Properties { get; set; } = new();

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append($"[{Timestamp:HH:mm:ss.fff}] [{Level}] [{Category}] {Message}");
            if (Exception != null)
            {
                sb.Append($"\n  Exception: {Exception.GetType().Name}: {Exception.Message}");
                sb.Append($"\n  StackTrace: {Exception.StackTrace}");
            }
            if (Properties != null && Properties.Count > 0)
            {
                sb.Append($"\n  Properties: {JsonSerializer.Serialize(Properties)}");
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// 日志器 - 同时输出到UI（事件）和文件
    /// </summary>
    public class Logger
    {
        private static readonly string LogDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DeepExcel", "logs");
        private static readonly string LogFile = Path.Combine(LogDir, $"deepexcel-{DateTime.Now:yyyyMMdd}.log");

        private static Logger _instance;
        public static Logger Instance => _instance ??= new Logger();

        public LogLevel MinLevel { get; set; } = LogLevel.Info;
        public event Action<LogEntry> OnLogEntry;

        private readonly ConcurrentQueue<LogEntry> _buffer = new();
        private readonly object _fileLock = new();
        private const int MaxFileSizeBytes = 10 * 1024 * 1024;  // 10MB
        private const int MaxFiles = 5;

        private Logger()
        {
            try
            {
                Directory.CreateDirectory(LogDir);
            }
            catch { }
        }

        public void Debug(string category, string message, Dictionary<string, object> properties = null)
            => Log(LogLevel.Debug, category, message, null, properties);

        public void Info(string category, string message, Dictionary<string, object> properties = null)
            => Log(LogLevel.Info, category, message, null, properties);

        public void Warning(string category, string message, Exception ex = null, Dictionary<string, object> properties = null)
            => Log(LogLevel.Warning, category, message, ex, properties);

        public void Error(string category, string message, Exception ex = null, Dictionary<string, object> properties = null)
            => Log(LogLevel.Error, category, message, ex, properties);

        public void Log(LogLevel level, string category, string message, Exception ex = null, Dictionary<string, object> properties = null)
        {
            if (level < MinLevel) return;

            var entry = new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = level,
                Category = category,
                Message = message,
                Exception = ex,
                Properties = properties ?? new()
            };

            // 推送到UI
            try { OnLogEntry?.Invoke(entry); } catch { }

            // 写入文件
            _ = Task.Run(() => WriteToFile(entry));
        }

        private void WriteToFile(LogEntry entry)
        {
            try
            {
                lock (_fileLock)
                {
                    CheckRotation();
                    File.AppendAllText(LogFile, entry.ToString() + Environment.NewLine);
                }
            }
            catch { }
        }

        private void CheckRotation()
        {
            try
            {
                if (File.Exists(LogFile))
                {
                    var size = new FileInfo(LogFile).Length;
                    if (size > MaxFileSizeBytes)
                    {
                        var backupName = LogFile + "." + DateTime.Now.ToString("HHmmss") + ".bak";
                        File.Move(LogFile, backupName);

                        // 清理旧文件
                        var dir = Path.GetDirectoryName(LogFile);
                        var files = new DirectoryInfo(dir).GetFiles("deepexcel-*.log.*.bak");
                        if (files.Length > MaxFiles)
                        {
                            Array.Sort(files, (a, b) => a.CreationTime.CompareTo(b.CreationTime));
                            for (int i = 0; i < files.Length - MaxFiles; i++)
                            {
                                try { files[i].Delete(); } catch { }
                            }
                        }
                    }
                }
            }
            catch { }
        }

        public string LogDirectory => LogDir;
        public string CurrentLogFile => LogFile;

        public List<LogEntry> GetRecentEntries(int count = 100, LogLevel? minLevel = null)
        {
            var effectiveLevel = minLevel ?? MinLevel;
            var list = new List<LogEntry>();
            try
            {
                if (File.Exists(LogFile))
                {
                    var lines = File.ReadAllLines(LogFile);
                    int start = Math.Max(0, lines.Length - count);
                    for (int i = start; i < lines.Length; i++)
                    {
                        var entry = ParseLogLine(lines[i]);
                        if (entry != null && entry.Level >= effectiveLevel)
                            list.Add(entry);
                    }
                }
            }
            catch { }
            return list;
        }

        private LogEntry ParseLogLine(string line)
        {
            try
            {
                if (line.Length < 12) return null;
                // 解析 [HH:mm:ss.fff] [Level] [Category] Message
                var parts = line.Split(new[] { "] [" }, StringSplitOptions.None);
                if (parts.Length < 4) return null;

                var entry = new LogEntry
                {
                    Timestamp = DateTime.Today.Add(TimeSpan.Parse(parts[0].TrimStart('['))),
                    Level = (LogLevel)Enum.Parse(typeof(LogLevel), parts[1]),
                    Category = parts[2].TrimEnd(']')
                };
                var rest = string.Join("] [", parts.Skip(3));
                entry.Message = rest;

                return entry;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 性能监控 - 追踪每次工具调用的耗时和Token消耗
    /// </summary>
    public class MetricsCollector
    {
        private static MetricsCollector _instance;
        public static MetricsCollector Instance => _instance ??= new MetricsCollector();

        public event Action<MetricEntry> OnMetric;

        private readonly List<MetricEntry> _entries = new();
        private readonly object _lock = new();

        public void RecordToolCall(string tool, double durationMs, bool success, int inputTokens = 0, int outputTokens = 0)
        {
            var entry = new MetricEntry
            {
                Type = "tool",
                Name = tool,
                DurationMs = durationMs,
                Success = success,
                InputTokens = inputTokens,
                OutputTokens = outputTokens,
                Timestamp = DateTime.Now
            };
            lock (_lock) _entries.Add(entry);
            OnMetric?.Invoke(entry);
        }

        public void RecordModelCall(string model, double durationMs, int inputTokens, int outputTokens, bool success)
        {
            var entry = new MetricEntry
            {
                Type = "model",
                Name = model,
                DurationMs = durationMs,
                Success = success,
                InputTokens = inputTokens,
                OutputTokens = outputTokens,
                Timestamp = DateTime.Now
            };
            lock (_lock) _entries.Add(entry);
            OnMetric?.Invoke(entry);
        }

        public MetricsSummary GetSummary()
        {
            lock (_lock)
            {
                var summary = new MetricsSummary();
                summary.TotalToolCalls = _entries.Count(e => e.Type == "tool");
                summary.TotalModelCalls = _entries.Count(e => e.Type == "model");
                summary.FailedCalls = _entries.Count(e => !e.Success);
                summary.TotalInputTokens = _entries.Sum(e => e.InputTokens);
                summary.TotalOutputTokens = _entries.Sum(e => e.OutputTokens);
                summary.AvgToolDurationMs = _entries.Where(e => e.Type == "tool").Select(e => e.DurationMs).DefaultIfEmpty(0).Average();
                return summary;
            }
        }
    }

    public class MetricEntry
    {
        public string Type { get; set; }      // "tool" | "model"
        public string Name { get; set; }
        public double DurationMs { get; set; }
        public bool Success { get; set; }
        public int InputTokens { get; set; }
        public int OutputTokens { get; set; }
        public DateTime Timestamp { get; set; }
    }

    public class MetricsSummary
    {
        public int TotalToolCalls { get; set; }
        public int TotalModelCalls { get; set; }
        public int FailedCalls { get; set; }
        public int TotalInputTokens { get; set; }
        public int TotalOutputTokens { get; set; }
        public double AvgToolDurationMs { get; set; }
    }
}
