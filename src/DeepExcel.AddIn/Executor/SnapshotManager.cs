using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using DeepExcel.AddIn.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 快照管理器 - 操作前备份，按需恢复。
    ///
    /// 恢复的原则（2026-09-24 重写，旧实现会丢数据）：
    /// - 恢复目标是快照来源的那本工作簿，按快照元数据找，不看 ActiveWorkbook。
    ///   旧实现直接关掉"当前活动"工作簿，用户中途切到别的文件时关错文件、未保存内容全丢。
    /// - 不关闭工作簿、不改写磁盘上的原文件。快照以隐藏、禁宏、只读方式打开，
    ///   按表把内容恢复回已打开的工作簿；用户要不要保存由用户决定。
    ///   （关闭工作簿还会触发 WorkbookBeforeClose，把该工作簿的 AI 会话一起销毁。）
    /// - 恢复前先给当前状态再存一份；这份存不下来就什么都不改（fail-closed）。
    /// </summary>
    public class SnapshotManager
    {
        private readonly Application _app;
        private readonly string _snapshotFolder;

        /// <summary>每个工作簿保留最近 N 个快照；每回合自动备份后 20 个大约是最近 20 轮</summary>
        public const int MaxSnapshotsPerWorkbook = 20;
        /// <summary>所有工作簿合计上限，防止无限吃磁盘</summary>
        public const int MaxSnapshotsTotal = 60;

        private static readonly Regex SnapshotIdPattern = new Regex("^[0-9a-f]{32}$", RegexOptions.IgnoreCase);

        public SnapshotManager(Application app, string snapshotFolder = null)
        {
            _app = app;
            _snapshotFolder = snapshotFolder ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DeepExcel", "Snapshots");
            Directory.CreateDirectory(_snapshotFolder);
        }

        // ============================== 创建 ==============================

        /// <summary>给当前活动工作簿做整本快照（手动快照 / 旧调用方）。失败返回 null。</summary>
        public string CreateSnapshot(string reason = null)
        {
            Workbook wb = null;
            try { wb = _app?.ActiveWorkbook; } catch { }
            if (wb == null) return null;
            return TryCreateSnapshot(wb, reason ?? "auto", SnapshotScope.Whole()).SnapshotId;
        }

        /// <summary>给指定 key 的已打开工作簿做快照</summary>
        public SnapshotAttempt TryCreateSnapshotFor(string workbookKey, string reason, SnapshotScope scope)
        {
            var wb = FindOpenWorkbook(workbookKey);
            if (wb == null) return new SnapshotAttempt { Error = "找不到要备份的工作簿（可能已关闭）" };
            return TryCreateSnapshot(wb, reason, scope);
        }

        public SnapshotAttempt TryCreateSnapshot(Workbook wb, string reason, SnapshotScope scope)
            => TryCreateSnapshot(wb, reason, scope, cleanup: true);

        private SnapshotAttempt TryCreateSnapshot(Workbook wb, string reason, SnapshotScope scope, bool cleanup)
        {
            if (wb == null) return new SnapshotAttempt { Error = "没有打开的工作簿" };
            scope = scope ?? SnapshotScope.Whole();

            var snapshotId = Guid.NewGuid().ToString("N");
            string fullName = SafeGet(() => wb.FullName);
            int? fileFormat = null;
            try { fileFormat = (int)wb.FileFormat; } catch { }
            var extension = SnapshotFormats.ChooseExtension(fullName, fileFormat);
            var backupPath = Path.Combine(_snapshotFolder, snapshotId + extension);

            try
            {
                wb.SaveCopyAs(backupPath);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("SnapshotManager", "SaveCopyAs failed: " + ex.Message);
                SafeDelete(backupPath);
                return new SnapshotAttempt { Error = "保存工作簿副本失败（" + ex.Message + "）" };
            }
            if (!File.Exists(backupPath))
            {
                return new SnapshotAttempt { Error = "保存工作簿副本失败（文件未生成）" };
            }

            var meta = new SnapshotMeta
            {
                Id = snapshotId,
                WorkbookName = SafeGet(() => wb.Name) ?? "未知",
                OriginalPath = PathOrNull(fullName),
                WorkbookKey = WorkbookIdentity.KeyOf(wb),
                FileExtension = extension,
                CreatedAt = DateTime.Now,
                Reason = reason ?? "auto",
                WholeWorkbook = scope.WholeWorkbook,
                AffectedSheets = scope.WholeWorkbook ? new List<string>() : scope.Sheets.ToList(),
            };

            // 元数据记录"这份快照属于哪本工作簿"，没有它回滚就无法确定目标。
            // 以前元数据写失败也继续，现在视为备份失败。
            try
            {
                WriteMeta(meta);
            }
            catch (Exception ex)
            {
                SafeDelete(backupPath);
                return new SnapshotAttempt { Error = "写入快照信息失败（" + ex.Message + "）" };
            }

            if (cleanup) CleanupOldSnapshots();
            return new SnapshotAttempt { SnapshotId = snapshotId };
        }

        /// <summary>
        /// 扩大已有快照的覆盖范围（同一回合里后续写入涉及了新的表）。
        /// 返回 false 表示没能记下——调用方应当视为备份失败，否则回滚会漏掉这张表。
        /// </summary>
        public bool ExtendScope(string snapshotId, SnapshotScope scope)
        {
            if (scope == null) return true;
            var meta = GetMeta(snapshotId);
            if (meta == null) return false;
            if (IsWholeWorkbook(meta)) return true;

            if (scope.WholeWorkbook)
            {
                meta.WholeWorkbook = true;
            }
            else
            {
                var sheets = meta.AffectedSheets ?? new List<string>();
                foreach (var s in scope.Sheets)
                {
                    if (!sheets.Any(x => string.Equals(x, s, StringComparison.OrdinalIgnoreCase))) sheets.Add(s);
                }
                meta.AffectedSheets = sheets;
            }
            try
            {
                WriteMeta(meta);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("SnapshotManager", "ExtendScope failed: " + ex.Message);
                return false;
            }
        }

        // ============================== 查询 ==============================

        public SnapshotMeta GetMeta(string snapshotId)
        {
            if (!IsValidSnapshotId(snapshotId)) return null;
            var metaPath = MetaPath(snapshotId);
            if (!File.Exists(metaPath)) return null;
            try { return JsonSerializer.Deserialize<SnapshotMeta>(File.ReadAllText(metaPath)); }
            catch { return null; }
        }

        /// <summary>列出所有历史快照（按时间倒序，最新的在前）</summary>
        public List<SnapshotMeta> ListSnapshots()
        {
            var result = new List<SnapshotMeta>();
            try
            {
                foreach (var file in EnumerateSnapshotFiles())
                {
                    var id = Path.GetFileNameWithoutExtension(file.Name);
                    var meta = GetMeta(id);
                    if (meta == null)
                    {
                        bool hasMeta = File.Exists(MetaPath(id));
                        meta = new SnapshotMeta
                        {
                            Id = id,
                            WorkbookName = "未知工作簿",
                            CreatedAt = file.CreationTime,
                            Reason = hasMeta ? "meta-lost" : "no-meta",
                        };
                    }
                    if (meta.CreatedAt == default(DateTime)) meta.CreatedAt = file.CreationTime;
                    if (string.IsNullOrEmpty(meta.FileExtension)) meta.FileExtension = file.Extension.ToLowerInvariant();
                    result.Add(meta);
                }
                result.Sort((a, b) => b.CreatedAt.CompareTo(a.CreatedAt));
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("SnapshotManager", "ListSnapshots error: " + ex.Message);
            }
            return result;
        }

        // ============================== 恢复 ==============================

        /// <summary>
        /// 把快照恢复回它所属的、当前已打开的工作簿。详见类注释。
        /// </summary>
        public RollbackResult Rollback(string snapshotId)
        {
            if (!IsValidSnapshotId(snapshotId)) return RollbackResult.Fail(snapshotId, "快照 ID 无效");

            var meta = GetMeta(snapshotId);
            var file = FindSnapshotFile(snapshotId, meta);
            if (file == null) return RollbackResult.Fail(snapshotId, "快照文件不存在（可能已被清理）");

            var target = FindTargetWorkbook(meta);
            if (target == null)
            {
                return RollbackResult.Fail(snapshotId, meta == null
                    ? "这个快照缺少归属信息，无法确定该恢复到哪个工作簿。为避免改错文件，未做任何修改。"
                    : $"请先打开工作簿「{meta.WorkbookName}」再恢复。为避免改错文件，不会恢复到其他工作簿。");
            }

            var result = new RollbackResult
            {
                SnapshotId = snapshotId,
                WorkbookName = SafeGet(() => target.Name),
            };

            // 恢复本身也要可撤销：先把当前状态（含用户未保存的修改）存一份
            // 这里不做配额清理：要恢复的若正好是该工作簿最旧的一份，清理会把它删掉
            var pre = TryCreateSnapshot(target, "恢复前的状态", SnapshotScope.Whole(), cleanup: false);
            if (!pre.Success)
            {
                return RollbackResult.Fail(snapshotId, "恢复前需要先备份当前状态，但备份失败，未做任何修改：" + pre.Error);
            }
            result.PreRestoreSnapshotId = pre.SnapshotId;

            var scope = ScopeOf(meta);
            var ui = ExcelUiState.Capture(_app, target);
            Workbook snap = null;
            try
            {
                ui.Suppress();
                snap = OpenHidden(file);

                var plan = RestorePlan.Build(WorksheetNames(snap), WorksheetNames(target), scope);
                Logger.Instance.Info("SnapshotManager",
                    $"Rollback {snapshotId}: whole={scope.WholeWorkbook}, inPlace=[{string.Join(",", plan.RestoreInPlace)}], " +
                    $"copy=[{string.Join(",", plan.CopyFromSnapshot)}], remove=[{string.Join(",", plan.RemoveFromTarget)}]");

                foreach (var name in plan.CopyFromSnapshot) CopySheetIn(target, (Worksheet)snap.Worksheets[name], result);
                foreach (var name in plan.RestoreInPlace)
                    RestoreSheetInPlace((Worksheet)target.Worksheets[name], (Worksheet)snap.Worksheets[name], result);
                foreach (var name in plan.RemoveFromTarget) RemoveSheet(target, name, result);

                if (scope.WholeWorkbook) ReorderLike(target, WorksheetNames(snap), result);
                FixLinksToSnapshot(target, snap, result);

                if (scope.WholeWorkbook)
                {
                    result.Warnings.Add("工作簿级的名称定义和 VBA 模块不在恢复范围内");
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("SnapshotManager", "Rollback failed: " + snapshotId, ex);
                result.Error = "恢复过程中出错（" + ex.Message + "）。恢复前的状态已另存，可从历史版本找回。";
            }
            finally
            {
                if (snap != null)
                {
                    try { snap.Close(false); } catch { }
                }
                ui.Restore();
                CleanupOldSnapshots();
            }

            if (result.Error == null && result.FailedSheets.Count > 0)
            {
                result.Error = "部分工作表没能恢复：" + string.Join("；", result.FailedSheets);
            }
            result.Success = result.Error == null;
            return result;
        }

        private Workbook OpenHidden(string file)
        {
            // 只读 + 不更新链接 + 不进最近使用列表；宏已由 ExcelUiState.Suppress 强制禁用，
            // 否则 xlsm 快照的 Workbook_Open 会在用户的 Excel 里跑一遍。
            var wb = _app.Workbooks.Open(file, UpdateLinks: 0, ReadOnly: true, AddToMru: false, Notify: false);
            try
            {
                foreach (Window w in wb.Windows) w.Visible = false;
            }
            catch { }
            return wb;
        }

        private static void RestoreSheetInPlace(Worksheet dst, Worksheet src, RollbackResult result)
        {
            string name = SafeGet(() => dst.Name);
            try
            {
                if (dst.ProtectContents)
                {
                    result.FailedSheets.Add(name + "（工作表受保护）");
                    return;
                }

                // 原地恢复：保留工作表对象本身，其他表、图表、名称对它的引用都不会断。
                // 先清掉当前表上的对象，再把快照整张表的单元格（值、公式、格式、合并、
                // 条件格式、批注、行高列宽及随单元格的图形）复制过来。
                try { if (dst.AutoFilterMode) dst.AutoFilterMode = false; } catch { }
                foreach (var lo in dst.ListObjects.Cast<ListObject>().ToList())
                {
                    try { lo.Delete(); } catch { }
                }
                try
                {
                    var pivots = (PivotTables)dst.PivotTables();
                    for (int i = pivots.Count; i >= 1; i--)
                    {
                        try { pivots.Item(i).TableRange2.Clear(); } catch { }
                    }
                }
                catch { }
                foreach (var shape in dst.Shapes.Cast<Shape>().ToList())
                {
                    try
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoComment) continue;
                        shape.Delete();
                    }
                    catch { }
                }

                dst.Cells.Clear();
                src.Cells.Copy(dst.Cells);

                try
                {
                    if (dst.Visible != src.Visible) dst.Visible = src.Visible;
                }
                catch { }

                result.RestoredSheets.Add(name);
            }
            catch (Exception ex)
            {
                result.FailedSheets.Add($"{name}（{ex.Message}）");
            }
        }

        private static void CopySheetIn(Workbook target, Worksheet src, RollbackResult result)
        {
            string name = SafeGet(() => src.Name);
            try
            {
                var visibility = src.Visible;
                // 快照以只读方式打开，改它的可见性不会写回磁盘
                if (visibility != XlSheetVisibility.xlSheetVisible) src.Visible = XlSheetVisibility.xlSheetVisible;

                src.Copy(Type.Missing, target.Sheets[target.Sheets.Count]);
                var copied = (Worksheet)target.Sheets[target.Sheets.Count];
                if (!string.Equals(copied.Name, name, StringComparison.Ordinal))
                {
                    try { copied.Name = name; } catch { }
                }
                if (visibility != XlSheetVisibility.xlSheetVisible)
                {
                    try { copied.Visible = visibility; } catch { }
                }
                result.RestoredSheets.Add(name);
            }
            catch (Exception ex)
            {
                result.FailedSheets.Add($"{name}（无法从快照复制回来：{ex.Message}）");
            }
        }

        private static void RemoveSheet(Workbook target, string name, RollbackResult result)
        {
            try
            {
                ((Worksheet)target.Worksheets[name]).Delete();
                result.RemovedSheets.Add(name);
            }
            catch (Exception ex)
            {
                result.FailedSheets.Add($"{name}（快照之后新建的表，没能删除：{ex.Message}）");
            }
        }

        private static void ReorderLike(Workbook target, IList<string> order, RollbackResult result)
        {
            try
            {
                for (int i = 0; i < order.Count; i++)
                {
                    var ws = (Worksheet)target.Worksheets[order[i]];
                    var at = (Worksheet)target.Worksheets[i + 1];
                    if (!string.Equals(ws.Name, at.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        ws.Move(at, Type.Missing);
                    }
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add("工作表顺序没能完全恢复（" + ex.Message + "）");
            }
        }

        /// <summary>
        /// 跨工作簿复制单元格时，引用其他表的公式会变成指向快照文件的外部链接
        /// （=[快照.xlsx]Sheet2!A1）。改回指向工作簿自身；ChangeLink 覆盖公式、名称和图表，
        /// 对未保存的工作簿不可用时退回到在公式里删掉快照文件名。
        /// </summary>
        private static void FixLinksToSnapshot(Workbook target, Workbook snap, RollbackResult result)
        {
            string snapFull = SafeGet(() => snap.FullName);
            string snapName = SafeGet(() => snap.Name);
            if (string.IsNullOrEmpty(snapFull) || !HasLinkTo(target, snapFull)) return;

            try { target.ChangeLink(snapFull, target.FullName, XlLinkType.xlLinkTypeExcelLinks); }
            catch (Exception ex) { Logger.Instance.Info("SnapshotManager", "ChangeLink fallback: " + ex.Message); }

            if (HasLinkTo(target, snapFull) && !string.IsNullOrEmpty(snapName))
            {
                foreach (Worksheet ws in target.Worksheets)
                {
                    try
                    {
                        ws.Cells.Replace("[" + snapName + "]", "", XlLookAt.xlPart, XlSearchOrder.xlByRows, false);
                    }
                    catch { }
                }
            }

            if (HasLinkTo(target, snapFull))
            {
                result.Warnings.Add("部分公式或图表仍引用快照文件，请在「数据 → 编辑链接」中检查");
            }
        }

        private static bool HasLinkTo(Workbook wb, string file)
        {
            try
            {
                if (!(wb.LinkSources(XlLink.xlExcelLinks) is Array links)) return false;
                foreach (var l in links)
                {
                    if (string.Equals(l as string, file, StringComparison.OrdinalIgnoreCase)) return true;
                }
            }
            catch { }
            return false;
        }

        private static List<string> WorksheetNames(Workbook wb)
        {
            var names = new List<string>();
            foreach (Worksheet ws in wb.Worksheets) names.Add(ws.Name);
            return names;
        }

        // ============================== 目标定位 ==============================

        /// <summary>按 key 在已打开的工作簿里找；找不到返回 null，绝不退回到 ActiveWorkbook</summary>
        public Workbook FindOpenWorkbook(string workbookKey)
        {
            if (string.IsNullOrEmpty(workbookKey) || _app == null) return null;
            try
            {
                foreach (Workbook wb in _app.Workbooks)
                {
                    if (WorkbookIdentity.SameKey(WorkbookIdentity.KeyOf(wb), workbookKey)) return wb;
                }
            }
            catch { }
            return null;
        }

        private Workbook FindTargetWorkbook(SnapshotMeta meta)
        {
            if (meta == null) return null;
            // 新快照记录了 key；旧快照只有路径和名字
            return FindOpenWorkbook(meta.WorkbookKey)
                ?? FindOpenWorkbook(meta.OriginalPath)
                ?? (meta.OriginalPath == null ? FindOpenWorkbook(meta.WorkbookName) : null);
        }

        internal static SnapshotScope ScopeOf(SnapshotMeta meta)
        {
            if (meta == null || IsWholeWorkbook(meta)) return SnapshotScope.Whole();
            return SnapshotScope.ForSheets(meta.AffectedSheets.ToArray());
        }

        /// <summary>旧快照没有范围信息，一律按整本处理</summary>
        internal static bool IsWholeWorkbook(SnapshotMeta meta)
            => meta.WholeWorkbook || meta.AffectedSheets == null || meta.AffectedSheets.Count == 0;

        // ============================== 文件 ==============================

        public static bool IsValidSnapshotId(string snapshotId)
            => !string.IsNullOrEmpty(snapshotId) && SnapshotIdPattern.IsMatch(snapshotId);

        private string MetaPath(string snapshotId) => Path.Combine(_snapshotFolder, snapshotId + ".meta.json");

        private void WriteMeta(SnapshotMeta meta)
        {
            File.WriteAllText(MetaPath(meta.Id), JsonSerializer.Serialize(meta));
        }

        internal string FindSnapshotFile(string snapshotId, SnapshotMeta meta)
        {
            if (!IsValidSnapshotId(snapshotId)) return null;
            if (!string.IsNullOrEmpty(meta?.FileExtension) && SnapshotFormats.KnownExtensions.Contains(meta.FileExtension))
            {
                var p = Path.Combine(_snapshotFolder, snapshotId + meta.FileExtension);
                if (File.Exists(p)) return p;
            }
            foreach (var ext in SnapshotFormats.KnownExtensions)
            {
                var p = Path.Combine(_snapshotFolder, snapshotId + ext);
                if (File.Exists(p)) return p;
            }
            return null;
        }

        private IEnumerable<FileInfo> EnumerateSnapshotFiles()
        {
            return new DirectoryInfo(_snapshotFolder).GetFiles()
                .Where(f => SnapshotFormats.KnownExtensions.Contains(f.Extension.ToLowerInvariant())
                            && IsValidSnapshotId(Path.GetFileNameWithoutExtension(f.Name)));
        }

        /// <summary>
        /// 按工作簿分别保留最近 MaxSnapshotsPerWorkbook 个，合计不超过 MaxSnapshotsTotal 个。
        /// 以前所有工作簿共用 20 个的配额，一个工作簿频繁操作会挤掉另一个的全部历史。
        /// </summary>
        private void CleanupOldSnapshots()
        {
            try
            {
                var all = ListSnapshots(); // 最新在前
                var perWorkbook = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < all.Count; i++)
                {
                    var owner = all[i].WorkbookKey ?? all[i].OriginalPath ?? all[i].WorkbookName ?? "";
                    perWorkbook.TryGetValue(owner, out var n);
                    perWorkbook[owner] = ++n;
                    if (n > MaxSnapshotsPerWorkbook || i >= MaxSnapshotsTotal)
                    {
                        DeleteSnapshot(all[i].Id);
                    }
                }
            }
            catch
            {
                // 清理失败不影响主流程
            }
        }

        public void ClearAllSnapshots()
        {
            try
            {
                foreach (var file in new DirectoryInfo(_snapshotFolder).GetFiles())
                {
                    try { file.Delete(); } catch { }
                }
            }
            catch { }
        }

        public bool DeleteSnapshot(string snapshotId)
        {
            if (!IsValidSnapshotId(snapshotId)) return false;
            try
            {
                foreach (var ext in SnapshotFormats.KnownExtensions)
                {
                    SafeDelete(Path.Combine(_snapshotFolder, snapshotId + ext));
                }
                SafeDelete(MetaPath(snapshotId));
                return true;
            }
            catch { return false; }
        }

        private static void SafeDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        private static T SafeGet<T>(Func<T> getter) where T : class
        {
            try { return getter(); } catch { return null; }
        }

        /// <summary>未保存的工作簿（如"工作簿1"）FullName 没有路径分隔符，返回 null</summary>
        private static string PathOrNull(string fullName)
        {
            if (string.IsNullOrEmpty(fullName) || !fullName.Contains("\\") && !fullName.Contains("/")) return null;
            return fullName;
        }

        /// <summary>
        /// 恢复期间临时改动的 Excel 应用状态：关闭屏幕刷新、事件、提示框，强制禁用宏；
        /// 结束后原样还原，并切回用户原来所在的窗口和工作表。
        /// </summary>
        private sealed class ExcelUiState
        {
            private readonly Application _app;
            private readonly Workbook _target;
            private bool _screenUpdating, _enableEvents, _displayAlerts;
            private Microsoft.Office.Core.MsoAutomationSecurity _security;
            private Window _activeWindow;
            private string _targetActiveSheet;
            private bool _suppressed;

            private ExcelUiState(Application app, Workbook target)
            {
                _app = app;
                _target = target;
            }

            public static ExcelUiState Capture(Application app, Workbook target)
            {
                var s = new ExcelUiState(app, target);
                try { s._activeWindow = app.ActiveWindow; } catch { }
                try { s._targetActiveSheet = (target.ActiveSheet as Worksheet)?.Name; } catch { }
                return s;
            }

            public void Suppress()
            {
                _screenUpdating = _app.ScreenUpdating;
                _enableEvents = _app.EnableEvents;
                _displayAlerts = _app.DisplayAlerts;
                _security = _app.AutomationSecurity;
                _suppressed = true;
                _app.ScreenUpdating = false;
                _app.EnableEvents = false;
                _app.DisplayAlerts = false;
                _app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }

            public void Restore()
            {
                if (_targetActiveSheet != null)
                {
                    try { ((Worksheet)_target.Worksheets[_targetActiveSheet]).Activate(); } catch { }
                }
                try { _activeWindow?.Activate(); } catch { }
                if (!_suppressed) return;
                try { _app.AutomationSecurity = _security; } catch { }
                try { _app.DisplayAlerts = _displayAlerts; } catch { }
                try { _app.EnableEvents = _enableEvents; } catch { }
                try { _app.ScreenUpdating = _screenUpdating; } catch { }
            }
        }
    }

    /// <summary>
    /// 快照元数据（与前端 JSON 协议对应）
    /// </summary>
    public class SnapshotMeta
    {
        public string Id { get; set; }
        public string WorkbookName { get; set; }
        /// <summary>创建快照时工作簿的完整路径（未保存的工作簿为 null）</summary>
        public string OriginalPath { get; set; }
        /// <summary>快照所属工作簿的会话 key（与 WorkbookSession.WorkbookKey 同一规则），回滚按它找目标</summary>
        public string WorkbookKey { get; set; }
        /// <summary>快照文件扩展名，跟随源工作簿格式（xlsm 存 xlsm）</summary>
        public string FileExtension { get; set; }
        public DateTime CreatedAt { get; set; }
        public string Reason { get; set; }
        /// <summary>true = 恢复整本；false = 只恢复 AffectedSheets。旧快照两者都缺省，按整本处理。</summary>
        public bool WholeWorkbook { get; set; }
        public List<string> AffectedSheets { get; set; }
    }
}
