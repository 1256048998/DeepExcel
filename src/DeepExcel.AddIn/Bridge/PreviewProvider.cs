using System;
using System.Collections.Generic;
using System.Diagnostics;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Preview;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Builds the change preview shown in the confirmation panel.
    /// </summary>
    public partial class MessageBridge
    {
        /// <summary>
        /// A preview runs while the user waits for the operation they asked
        /// for. Past this it stops being a preview and starts being a hang, so
        /// it degrades to "cannot preview, snapshot taken" instead.
        /// </summary>
        private static readonly TimeSpan PreviewBudget = TimeSpan.FromSeconds(3);

        internal ChangePreview BuildChangePreview(string toolName, Dictionary<string, object> args)
        {
            if (string.IsNullOrEmpty(toolName))
            {
                return null;
            }

            var clock = Stopwatch.StartNew();
            try
            {
                var builder = new PreviewBuilder(new ExcelRangeReader(_excelApp));
                var preview = builder.Build(toolName, args);

                if (preview != null && clock.Elapsed > PreviewBudget)
                {
                    Logger.Instance.Warning("MessageBridge",
                        $"Preview for {toolName} took {clock.ElapsedMilliseconds}ms");
                }
                return preview;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("MessageBridge",
                    $"Preview for {toolName} failed: {ex.Message}");
                // Never let a preview failure block the operation or, worse,
                // present an empty change set that reads as "this is safe".
                return ChangePreview.NotPreviewable(toolName, "预览计算失败，已自动创建快照");
            }
        }
    }
}
