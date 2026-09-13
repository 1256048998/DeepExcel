using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace DeepExcel.Repair
{
    /// <summary>
    /// Packs everything support needs and nothing the user would not want to
    /// send. The allowlist below is deliberate: snapshots, conversation history
    /// and attachments all contain workbook contents and are never included.
    /// </summary>
    public static class DiagnosticBundle
    {
        /// <summary>
        /// Values whose name matches this are replaced before the file is
        /// written. config.json is documented as key-free, but a redaction pass
        /// costs nothing and protects against a future field that is not.
        /// </summary>
        private static readonly Regex SecretField = new Regex(
            "\"([^\"]*(?:key|token|secret|password|credential)[^\"]*)\"\\s*:\\s*\"[^\"]*\"",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static string Write(string installDirectory, Report report)
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var name = string.Format(
                CultureInfo.InvariantCulture,
                "DeepExcel-诊断-{0:yyyyMMdd-HHmmss}.zip",
                DateTime.Now);
            var target = Path.Combine(desktop, name);

            using (var stream = new FileStream(target, FileMode.Create, FileAccess.Write))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteText(archive, "report.txt", BuildReport(installDirectory, report));

                foreach (var log in RecentLogs())
                {
                    // Cap each log so a runaway file cannot produce a bundle
                    // too large to email.
                    WriteText(archive, "logs/" + Path.GetFileName(log), ReadTail(log, 512 * 1024));
                }

                if (File.Exists(AddInIdentity.ConfigPath))
                {
                    WriteText(archive, "config.redacted.json", Redact(ReadTail(AddInIdentity.ConfigPath, 256 * 1024)));
                }
            }

            return target;
        }

        /// <summary>
        /// Only the newest logs. A long-lived install accumulates dozens of
        /// daily files; shipping them all makes the bundle awkward to send and
        /// adds nothing -- a load failure is always in the most recent runs.
        /// </summary>
        private const int MaxLogFiles = 5;

        private static IEnumerable<string> RecentLogs()
        {
            var logDirectory = AddInIdentity.LogDirectory;
            if (!Directory.Exists(logDirectory))
            {
                return new string[0];
            }

            var logs = new List<string>(Directory.GetFiles(logDirectory, "*.log"));
            logs.Sort((left, right) => File.GetLastWriteTimeUtc(right).CompareTo(File.GetLastWriteTimeUtc(left)));
            if (logs.Count > MaxLogFiles)
            {
                logs.RemoveRange(MaxLogFiles, logs.Count - MaxLogFiles);
            }
            return logs;
        }

        private static string BuildReport(string installDirectory, Report report)
        {
            var builder = new StringBuilder();
            builder.AppendLine("DeepExcel 诊断报告");
            builder.AppendLine("生成时间: " + DateTime.Now.ToString("u", CultureInfo.InvariantCulture));
            builder.AppendLine();
            builder.AppendLine("== 环境 ==");
            builder.AppendLine("OS: " + Environment.OSVersion.VersionString);
            builder.AppendLine("64 位操作系统: " + Environment.Is64BitOperatingSystem);
            builder.AppendLine("进程位数: " + ComProbe.BitnessOfCurrentProcess);
            builder.AppendLine("CLR: " + Environment.Version);
            builder.AppendLine(".NET 4.8: " + EnvironmentChecks.IsDotNet48Installed());
            builder.AppendLine("WebView2: " + EnvironmentChecks.IsWebView2Installed());
            builder.AppendLine("安装目录: " + installDirectory);

            var dllPath = Path.Combine(installDirectory, AddInIdentity.AddInDllName);
            if (File.Exists(dllPath))
            {
                builder.AppendLine("加载项版本: " + AddInIdentity.AssemblyVersion(dllPath));
            }
            else
            {
                builder.AppendLine("加载项版本: 主文件缺失");
            }

            builder.AppendLine();
            builder.AppendLine("== 自检结果 ==");
            builder.AppendLine(report.Render());
            builder.AppendLine();
            builder.AppendLine("本报告不含单元格内容、工作簿路径、对话记录或 API Key。");
            return builder.ToString();
        }

        private static string Redact(string json)
        {
            return SecretField.Replace(json, "\"$1\": \"[已脱敏]\"");
        }

        private static string ReadTail(string path, int maxBytes)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (stream.Length > maxBytes)
                    {
                        stream.Seek(stream.Length - maxBytes, SeekOrigin.Begin);
                    }
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                return "[无法读取 " + path + ": " + ex.Message + "]";
            }
        }

        private static void WriteText(ZipArchive archive, string entryName, string content)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }
    }
}
