using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 工作簿身份：会话 key、快照归属、回滚目标三处必须用同一套规则，
    /// 否则回滚会找不到（或找错）工作簿。已保存的用 FullName，未保存的"工作簿1"用 Name。
    /// </summary>
    public static class WorkbookIdentity
    {
        public static string KeyOf(Workbook wb)
        {
            if (wb == null) return null;
            try
            {
                return KeyFrom(wb.FullName, wb.Name) ?? "workbook_" + wb.GetHashCode();
            }
            catch
            {
                return "workbook_" + wb.GetHashCode();
            }
        }

        public static string KeyFrom(string fullName, string name)
        {
            if (!string.IsNullOrEmpty(fullName) && (fullName.Contains("\\") || fullName.Contains("/")))
                return fullName;
            return name;
        }

        public static bool SameKey(string a, string b)
        {
            return !string.IsNullOrEmpty(a) && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// 快照文件格式。SaveCopyAs 按工作簿自身格式写盘、不看扩展名，
    /// 所以 xlsm 工作簿存成 .xlsx 得到的是"内容是 xlsm、名字是 xlsx"的文件，Excel 拒绝打开。
    /// 扩展名必须跟着源格式走。
    /// </summary>
    public static class SnapshotFormats
    {
        public static readonly string[] KnownExtensions =
        {
            ".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".xlt", ".csv",
        };

        // XlFileFormat 常量：只列 SaveCopyAs 可能遇到的工作簿格式
        private static readonly Dictionary<int, string> ExtensionByFileFormat = new Dictionary<int, string>
        {
            [51] = ".xlsx",   // xlOpenXMLWorkbook
            [52] = ".xlsm",   // xlOpenXMLWorkbookMacroEnabled
            [50] = ".xlsb",   // xlExcel12
            [56] = ".xls",    // xlExcel8
            [-4143] = ".xls", // xlWorkbookNormal
            [54] = ".xltx",   // xlOpenXMLTemplate
            [53] = ".xltm",   // xlOpenXMLTemplateMacroEnabled
            [17] = ".xlt",    // xlTemplate
            [6] = ".csv",     // xlCSV
            [62] = ".csv",    // xlCSVUTF8
        };

        public static string ChooseExtension(string fullName, int? fileFormat)
        {
            string ext = null;
            try { ext = Path.GetExtension(fullName ?? "")?.ToLowerInvariant(); } catch { }
            if (!string.IsNullOrEmpty(ext) && KnownExtensions.Contains(ext)) return ext;
            if (fileFormat.HasValue && ExtensionByFileFormat.TryGetValue(fileFormat.Value, out var byFormat))
                return byFormat;
            return ".xlsx";
        }
    }

    /// <summary>快照覆盖范围：整本，或只涉及某几张表</summary>
    public sealed class SnapshotScope
    {
        public bool WholeWorkbook { get; private set; }
        public IReadOnlyList<string> Sheets { get; private set; }

        public static SnapshotScope Whole() => new SnapshotScope { WholeWorkbook = true, Sheets = new string[0] };

        public static SnapshotScope ForSheets(params string[] sheets)
        {
            var clean = (sheets ?? new string[0]).Where(s => !string.IsNullOrEmpty(s)).ToArray();
            if (clean.Length == 0) return Whole();
            return new SnapshotScope { WholeWorkbook = false, Sheets = clean };
        }
    }

    /// <summary>一次备份的结果：成功时 SnapshotId 非空，失败时 Error 说明原因</summary>
    public sealed class SnapshotAttempt
    {
        public string SnapshotId { get; set; }
        public string Error { get; set; }
        public bool Success => !string.IsNullOrEmpty(SnapshotId);
    }

    /// <summary>回滚结果：如实报告恢复了哪些表、哪些没恢复、恢复前的状态存在哪</summary>
    public sealed class RollbackResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string SnapshotId { get; set; }
        public string WorkbookName { get; set; }
        /// <summary>恢复前对当前状态做的备份；恢复本身也可以撤销</summary>
        public string PreRestoreSnapshotId { get; set; }
        public List<string> RestoredSheets { get; } = new List<string>();
        public List<string> RemovedSheets { get; } = new List<string>();
        public List<string> FailedSheets { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();

        public static RollbackResult Fail(string snapshotId, string error)
            => new RollbackResult { Success = false, SnapshotId = snapshotId, Error = error };
    }

    /// <summary>
    /// 回滚计划（纯逻辑，便于单测）。
    ///
    /// 两边都有的表：原地恢复内容（保留工作表对象，其他表对它的引用不会变成 #REF!）；
    /// 只在快照里有的表（被删/被改名）：从快照复制回来；
    /// 只在当前有的表（快照之后新建的）：删除。
    /// 范围限定时只处理范围内的表，用户在其他表上的编辑不受影响。
    /// </summary>
    public sealed class RestorePlan
    {
        public List<string> RestoreInPlace { get; } = new List<string>();
        public List<string> CopyFromSnapshot { get; } = new List<string>();
        public List<string> RemoveFromTarget { get; } = new List<string>();

        public bool IsEmpty => RestoreInPlace.Count == 0 && CopyFromSnapshot.Count == 0 && RemoveFromTarget.Count == 0;

        public static RestorePlan Build(IList<string> snapshotSheets, IList<string> targetSheets, SnapshotScope scope)
        {
            var plan = new RestorePlan();
            var inSnapshot = new HashSet<string>(snapshotSheets ?? new string[0], StringComparer.OrdinalIgnoreCase);
            var inTarget = new HashSet<string>(targetSheets ?? new string[0], StringComparer.OrdinalIgnoreCase);

            IEnumerable<string> candidates;
            if (scope == null || scope.WholeWorkbook)
            {
                candidates = (snapshotSheets ?? new string[0]).Concat(targetSheets ?? new string[0]);
            }
            else
            {
                candidates = scope.Sheets;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in candidates)
            {
                if (string.IsNullOrEmpty(name) || !seen.Add(name)) continue;
                bool s = inSnapshot.Contains(name), t = inTarget.Contains(name);
                if (s && t) plan.RestoreInPlace.Add(Canonical(name, snapshotSheets));
                else if (s) plan.CopyFromSnapshot.Add(Canonical(name, snapshotSheets));
                else if (t) plan.RemoveFromTarget.Add(Canonical(name, targetSheets));
            }
            return plan;
        }

        private static string Canonical(string name, IList<string> source)
        {
            return source?.FirstOrDefault(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase)) ?? name;
        }
    }
}
