using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 通用消息类型 - WebView2 ↔ C# 通信协议
    /// </summary>
    public class Message
    {
        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("payload")]
        public JsonElement? Payload { get; set; }

        [JsonPropertyName("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString();
    }

    /// <summary>
    /// 工具调用请求（来自UI/Agent）
    /// </summary>
    public class ToolCallRequest
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("arguments")]
        public JsonElement? Arguments { get; set; }
    }

    /// <summary>
    /// 工具执行结果（回传给UI/Sidecar）
    /// </summary>
    public class ToolResult
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("data")]
        public object Data { get; set; }

        [JsonPropertyName("error")]
        public string Error { get; set; }

        [JsonPropertyName("suggestion")]
        public string Suggestion { get; set; }   // 失败再问提示；非空时触发模型反问

        [JsonPropertyName("warning")]
        public string Warning { get; set; }      // 成功但有警告（如参数自动纠正）

        [JsonPropertyName("context")]
        public object Context { get; set; }      // Excel 上下文快照

        /// <summary>写入前自动备份的快照 ID（本回合首次写入时创建，同回合共用），可传给 rollback 撤销</summary>
        [JsonPropertyName("backup_snapshot_id")]
        public string BackupSnapshotId { get; set; }

        /// <summary>写后自动体检的结论（公式错误前后对比、新增外部链接、回读样本）；没做体检为 null</summary>
        [JsonPropertyName("verification")]
        public object Verification { get; set; }

        /// <summary>这一步执行前单独存的检查点（面板上「回到这一步之前」用）；和回合备份不同，每步一份</summary>
        [JsonPropertyName("checkpoint_id")]
        public string CheckpointId { get; set; }

        /// <summary>内联 diff：这一步实际改了哪些格（改动数 + 前几处原值→新值）；没比较为 null</summary>
        [JsonPropertyName("changes")]
        public object Changes { get; set; }
    }

    /// <summary>
    /// 流式输出片段
    /// </summary>
    public class StreamDelta
    {
        [JsonPropertyName("delta")]
        public string Delta { get; set; }
    }

    /// <summary>
    /// 操作建议（Agent在执行前推送给UI的预览）
    /// </summary>
    public class ActionPlan
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString();

        [JsonPropertyName("summary")]
        public string Summary { get; set; }   // 简短说明

        [JsonPropertyName("steps")]
        public PlanStep[] Steps { get; set; }  // 详细步骤

        [JsonPropertyName("risk")]
        public string Risk { get; set; }       // 风险等级: low/medium/high

        [JsonPropertyName("requiresConfirmation")]
        public bool RequiresConfirmation { get; set; } = true;
    }

    /// <summary>
    /// 操作步骤
    /// </summary>
    public class PlanStep
    {
        [JsonPropertyName("order")]
        public int Order { get; set; }

        [JsonPropertyName("tool")]
        public string Tool { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("arguments")]
        public object Arguments { get; set; }
    }

    /// <summary>
    /// 用户确认结果
    /// </summary>
    public class ConfirmationResult
    {
        [JsonPropertyName("planId")]
        public string PlanId { get; set; }

        [JsonPropertyName("approved")]
        public bool Approved { get; set; }

        [JsonPropertyName("feedback")]
        public string Feedback { get; set; }   // 用户拒绝时填写的修改意见
    }

    /// <summary>
    /// 工具调用前预览（带确认标记）
    /// </summary>
    public class ToolPreview
    {
        [JsonPropertyName("tool")]
        public string Tool { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("arguments")]
        public object Arguments { get; set; }
    }
}
