using System;
using System.Text.Json;
using DeepExcel.AddIn.Account;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 「数据与隐私」里的使用统计开关。关掉后不再记录、不再发送，本机还没发出的记录一并删掉
    /// （TelemetryReporter.SetOptOut）；状态是发件箱旁的标记文件，所有 Excel 进程共用。
    /// </summary>
    public partial class MessageBridge
    {
        private string HandleSetUsageStats(Message msg)
        {
            bool enabled = true;
            if (msg.Payload.HasValue && msg.Payload.Value.ValueKind == JsonValueKind.Object &&
                msg.Payload.Value.TryGetProperty("enabled", out var el) &&
                (el.ValueKind == JsonValueKind.True || el.ValueKind == JsonValueKind.False))
            {
                enabled = el.GetBoolean();
            }
            else
            {
                return MakeError("缺少 enabled 参数");
            }

            if (!TelemetryReporter.SetOptOut(!enabled))
            {
                Logger.Instance.Warning("MessageBridge", "set_usage_stats: could not write the preference");
                return MakeError("没能保存这个设置，请稍后重试");
            }
            Logger.Instance.Info("MessageBridge", $"usage statistics {(enabled ? "on" : "off")}");
            return MakeResponse("usage_stats", new { supported = true, enabled = !TelemetryReporter.IsOptedOut() });
        }
    }
}
