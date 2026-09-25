namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// Sidecar IPC 消息 type 常量（C# ↔ Python 双向）
    /// 协议：每行一个 JSON 对象，以 \n 分隔
    /// </summary>
    public static class SidecarProtocol
    {
        // C# → Python
        public const string TypeUserMessage = "user_message";
        public const string TypeCancel = "cancel";
        public const string TypeToolResult = "tool_result";
        public const string TypeConfig = "config";
        public const string TypeClarifyAnswer = "clarify_answer";
        public const string TypeRestoreHistory = "restore_history";
        public const string TypeSetPermissionMode = "set_permission_mode";

        // Python → C#
        public const string TypeStreamDelta = "stream_delta";
        public const string TypeToolCall = "tool_call";
        public const string TypeToolUse = "tool_use";
        public const string TypeClarify = "clarify";
        public const string TypeStreamEnd = "stream_end";
        public const string TypePermissionRequest = "permission_request";
        /// <summary>面板事件信封（docs/ui-event-protocol.md），宿主原样转发给面板。</summary>
        public const string TypeUiEvent = "ui_event";

        // C# → Python（权限响应）
        public const string TypePermissionResponse = "permission_response";

        /// <summary>
        /// 权限模式（侧车 permission_modes.py）：每步确认 / 本次会话自动应用写入 / 只出方案。
        /// 宿主只转发、不保存：「自动应用写入」只在当前会话有效。
        /// </summary>
        public static bool IsPermissionMode(string mode) =>
            mode == "default" || mode == "accept_writes" || mode == "plan";
    }
}
