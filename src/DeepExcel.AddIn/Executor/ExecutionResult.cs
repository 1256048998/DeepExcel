using System;
using System.Collections.Generic;
using System.IO;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 执行结果
    /// </summary>
    public class ExecutionResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Output { get; set; }
        public string SnapshotId { get; set; }
        public DateTime ExecutedAt { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// 工具执行结果（用于桥接层）
    /// </summary>
    public class ToolResult
    {
        public string Name { get; set; }
        public bool Success { get; set; }
        public object Data { get; set; }
        public string Error { get; set; }
    }
}
