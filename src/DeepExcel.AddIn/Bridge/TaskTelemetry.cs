using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DeepExcel.AddIn.Account;
using DeepExcel.AddIn.Config;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Per-task telemetry.
    ///
    /// This exists to answer one question that is currently unanswerable: what
    /// fraction of user instructions produce a usable result without manual
    /// correction. Every capability decision after this is supposed to be judged
    /// against that number, and it cannot be judged before it can be measured.
    ///
    /// Nothing recorded here can identify a workbook or its contents. Tool names
    /// and an outcome are all that leave the machine, and the server drops
    /// anything else regardless.
    /// </summary>
    public partial class MessageBridge
    {
        private sealed class TaskTrace
        {
            public string SessionId;
            public string UserRequest;
            public Stopwatch Clock;
            public List<string> Tools;
            /// <summary>Full calls, kept so a successful run can become a skill.</summary>
            public List<Skills.RecordedStep> Steps;
            public string Provider;
            public string Model;
        }

        /// <summary>
        /// The last successful run, offered as a skill.
        ///
        /// Only one is kept: the offer is made right after the task finishes, so
        /// anything older has already been declined.
        /// </summary>
        private Skills.RecordedStep[] _lastSuccessfulSteps;
        private string _lastSuccessfulRequest;

        private readonly Dictionary<string, TaskTrace> _taskTraces =
            new Dictionary<string, TaskTrace>(StringComparer.OrdinalIgnoreCase);
        private readonly object _traceLock = new object();

        /// <summary>workbookKey -> 最近一次 run_summary 的 outcome，stream_end 时取走。</summary>
        private readonly Dictionary<string, string> _lastRunOutcome =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 侧车 run_summary.outcome → 任务轨迹 outcome。没收到 run_summary（旧侧车）按成功记，
        /// 与以前的行为一致；达到最大轮次不算可用结果。
        /// </summary>
        internal static string TraceOutcomeFromRunSummary(string outcome)
        {
            switch (outcome)
            {
                case null:
                case "":
                case "success":
                    return "success";
                case "interrupted":
                    return "cancelled";
                default:
                    return "error";
            }
        }

        /// <summary>Created lazily so a local-only install never allocates it.</summary>
        private TelemetryReporter _telemetry;

        private TelemetryReporter Telemetry
        {
            get
            {
                if (_telemetry == null && AccountSession != null)
                {
                    _telemetry = new TelemetryReporter(AccountSession);
                    ReportSessionStart();
                }
                return _telemetry;
            }
        }

        private void ReportSessionStart()
        {
            try
            {
                _telemetry.Record("session_start", new Dictionary<string, object>
                {
                    ["client_version"] = typeof(MessageBridge).Assembly.GetName().Version.ToString(),
                    ["os_version"] = Environment.OSVersion.VersionString,
                    ["office_version"] = SafeOfficeVersion(),
                    ["office_bitness"] = IntPtr.Size == 8 ? "64" : "32",
                    ["host"] = "excel"
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "session_start telemetry failed: " + ex.Message);
            }
        }

        private string SafeOfficeVersion()
        {
            try { return _excelApp?.Version; }
            catch (Exception) { return null; }
        }

        internal void BeginTaskTrace(string workbookKey, string sessionId, string userRequest = null)
        {
            if (string.IsNullOrEmpty(workbookKey))
            {
                return;
            }
            string provider = null, model = null;
            try
            {
                var cfg = ConfigManager.Instance.Current;
                provider = cfg.CurrentProvider;
                model = cfg.CurrentModel;
            }
            catch (Exception) { }

            lock (_traceLock)
            {
                _taskTraces[workbookKey] = new TaskTrace
                {
                    SessionId = sessionId,
                    UserRequest = userRequest,
                    Clock = Stopwatch.StartNew(),
                    Tools = new List<string>(),
                    Steps = new List<Skills.RecordedStep>(),
                    Provider = provider,
                    Model = model
                };
            }
        }

        internal void RecordTraceTool(
            string workbookKey, string toolName, Dictionary<string, object> args = null)
        {
            if (string.IsNullOrEmpty(workbookKey) || string.IsNullOrEmpty(toolName))
            {
                return;
            }
            lock (_traceLock)
            {
                if (_taskTraces.TryGetValue(workbookKey, out var trace) && trace.Tools.Count < 60)
                {
                    trace.Tools.Add(toolName);
                    trace.Steps.Add(new Skills.RecordedStep
                    {
                        Tool = toolName,
                        Arguments = args == null
                            ? new Dictionary<string, object>()
                            : new Dictionary<string, object>(args)
                    });
                }
            }
        }

        /// <summary>
        /// Whether the last run is worth offering as a skill.
        ///
        /// A single-step task is not: describing it again is faster than naming,
        /// saving and finding a skill for it. The value only appears once a task
        /// has enough steps that repeating it by hand is tedious.
        /// </summary>
        public const int MinimumStepsForSkill = 3;

        internal object GetSkillCandidate()
        {
            var steps = _lastSuccessfulSteps;
            if (steps == null || steps.Length < MinimumStepsForSkill)
            {
                return null;
            }
            return new
            {
                step_count = steps.Length,
                tools = steps.Select(s => s.Tool).ToArray(),
                request = _lastSuccessfulRequest
            };
        }

        /// <param name="outcome">success, error, cancelled or clarify.</param>
        internal void CompleteTaskTrace(
            string workbookKey, string outcome, int inputTokens = 0, int outputTokens = 0)
        {
            TaskTrace trace;
            lock (_traceLock)
            {
                if (!_taskTraces.TryGetValue(workbookKey ?? "", out trace))
                {
                    return;
                }
                _taskTraces.Remove(workbookKey);
            }

            if (outcome == "success" && trace.Steps != null && trace.Steps.Count > 0)
            {
                // Only successful runs are offered as skills: saving a failed
                // sequence would bake the failure in.
                _lastSuccessfulSteps = trace.Steps.ToArray();
                _lastSuccessfulRequest = trace.UserRequest;
            }

            var reporter = Telemetry;
            if (reporter == null)
            {
                return;
            }

            try
            {
                reporter.Record("task_complete", new Dictionary<string, object>
                {
                    ["session_id"] = trace.SessionId,
                    ["duration_ms"] = (int)Math.Min(trace.Clock.ElapsedMilliseconds, int.MaxValue),
                    ["tool_sequence"] = trace.Tools,
                    ["turn_count"] = trace.Tools.Count,
                    ["tokens_in"] = inputTokens,
                    ["tokens_out"] = outputTokens,
                    ["provider"] = trace.Provider,
                    ["model"] = trace.Model,
                    ["outcome"] = outcome
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge", "task_complete telemetry failed: " + ex.Message);
            }
        }

        /// <summary>
        /// A tool failure. Only the tool name and a classification code are sent:
        /// raw error text routinely quotes cell values and file paths.
        /// </summary>
        internal void ReportToolError(string toolName, string errorCode)
        {
            var reporter = Telemetry;
            if (reporter == null || string.IsNullOrEmpty(toolName))
            {
                return;
            }
            try
            {
                reporter.Record("tool_error", new Dictionary<string, object>
                {
                    ["tool_name"] = toolName,
                    ["error_code"] = ClassifyError(errorCode)
                });
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Something stopped the assistant from working at all: the sidecar
        /// died, or its engine self-check failed. Recorded synchronously into
        /// the outbox, so it survives Excel going down right after. Only the
        /// code and the version leave the machine.
        /// </summary>
        internal void ReportStartupError(string diagnosticCode)
        {
            var reporter = Telemetry;
            if (reporter == null || string.IsNullOrEmpty(diagnosticCode))
            {
                return;
            }
            try
            {
                reporter.Record("startup_error", new Dictionary<string, object>
                {
                    ["diagnostic_code"] = diagnosticCode,
                    ["client_version"] = typeof(MessageBridge).Assembly.GetName().Version.ToString()
                });
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Maps a message onto a small fixed vocabulary.
        ///
        /// Sending the message itself would defeat the privacy guarantee, and a
        /// free-form string is useless for aggregation anyway -- the point is to
        /// rank failure kinds, not to read individual failures.
        /// </summary>
        internal static string ClassifyError(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return "unknown";
            }
            var text = message.ToLowerInvariant();
            if (text.Contains("timeout") || text.Contains("timed out")) return "timeout";
            if (text.Contains("permission") || text.Contains("denied") || text.Contains("拒绝")) return "permission_denied";
            if (text.Contains("not registered") || text.Contains("regdb")) return "com_not_registered";
            if (text.Contains("out of range") || text.Contains("bounds") || text.Contains("越界")) return "range_out_of_bounds";
            if (text.Contains("not found") || text.Contains("找不到") || text.Contains("不存在")) return "not_found";
            if (text.Contains("hresult") || text.Contains("com exception")) return "com_error";
            if (text.Contains("unauthorized") || text.Contains("401") || text.Contains("api key")) return "auth_failed";
            if (text.Contains("quota") || text.Contains("rate limit") || text.Contains("429")) return "rate_limited";
            if (text.Contains("network") || text.Contains("connection")) return "network";
            if (text.Contains("vba")) return "vba_error";
            if (text.Contains("syntax") || text.Contains("parse")) return "parse_error";
            return "other";
        }
    }
}
