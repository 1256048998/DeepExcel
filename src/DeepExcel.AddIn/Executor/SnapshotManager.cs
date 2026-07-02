using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 快照管理器 - 操作前自动备份，失败可回滚。
    /// ★ 新增 list_snapshots：返回历史快照元数据列表，供前端 UI 显示。
    /// </summary>
    public class SnapshotManager
    {
        private readonly Application _app;
        private readonly string _snapshotFolder;
        private const int MaxSnapshots = 20; // 保留最近20个快照

        public SnapshotManager(Application app)
        {
            _app = app;
            _snapshotFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DeepExcel", "Snapshots");
            Directory.CreateDirectory(_snapshotFolder);
        }

        /// <summary>
        /// 创建当前工作簿的快照。
        /// ★ 同时保存元数据 sidecar JSON（snapshotId、workbook 名、时间、触发原因），
        /// 供前端 UI 显示历史版本列表。
        /// </summary>
        public string CreateSnapshot(string reason = null)
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return null;

                var snapshotId = Guid.NewGuid().ToString("N");
                var backupPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");

                wb.SaveCopyAs(backupPath);

                // ★ 写元数据 sidecar JSON
                var meta = new SnapshotMeta
                {
                    Id = snapshotId,
                    WorkbookName = SafeGetWorkbookName(wb),
                    OriginalPath = SafeGetWorkbookPath(wb),  // ★ 记录原路径，回滚时恢复
                    CreatedAt = DateTime.Now,
                    Reason = reason ?? "auto",
                };
                try
                {
                    var metaPath = Path.Combine(_snapshotFolder, $"{snapshotId}.meta.json");
                    File.WriteAllText(metaPath, JsonSerializer.Serialize(meta, new JsonSerializerOptions { WriteIndented = false }));
                }
                catch { /* 元数据失败不影响主流程 */ }

                // 清理旧快照
                CleanupOldSnapshots();

                return snapshotId;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateSnapshot error: {ex}");
                return null;
            }
        }

        /// <summary>
        /// 列出所有历史快照（按时间倒序，最新的在前）。
        /// 同时清理无元数据的孤立 xlsx（防止 .meta.json 丢失但 xlsx 残留）。
        /// </summary>
        public List<SnapshotMeta> ListSnapshots()
        {
            var result = new List<SnapshotMeta>();
            try
            {
                var xlsxFiles = new DirectoryInfo(_snapshotFolder).GetFiles("*.xlsx");
                foreach (var xlsx in xlsxFiles)
                {
                    var id = Path.GetFileNameWithoutExtension(xlsx.Name);
                    var metaPath = Path.Combine(_snapshotFolder, $"{id}.meta.json");
                    SnapshotMeta meta;
                    if (File.Exists(metaPath))
                    {
                        try
                        {
                            meta = JsonSerializer.Deserialize<SnapshotMeta>(File.ReadAllText(metaPath));
                            // 文件 mtime 更可靠时优先用文件时间
                            if (meta.CreatedAt == default(DateTime))
                            {
                                meta.CreatedAt = xlsx.CreationTime;
                            }
                        }
                        catch
                        {
                            meta = new SnapshotMeta
                            {
                                Id = id,
                                WorkbookName = "未知工作簿",
                                CreatedAt = xlsx.CreationTime,
                                Reason = "meta-lost"
                            };
                        }
                    }
                    else
                    {
                        // 无元数据，用文件信息构造
                        meta = new SnapshotMeta
                        {
                            Id = id,
                            WorkbookName = "未知工作簿",
                            CreatedAt = xlsx.CreationTime,
                            Reason = "no-meta"
                        };
                    }
                    result.Add(meta);
                }
                // 按时间倒序（最新在前）
                result.Sort((a, b) => b.CreatedAt.CompareTo(a.CreatedAt));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ListSnapshots error: {ex}");
            }
            return result;
        }

        /// <summary>
        /// 回滚到指定快照。
        /// ★ 修复：不能直接 Workbooks.Open(快照路径)，否则工作簿名会变成快照 GUID。
        /// 正确做法：读元数据拿原路径 → 复制快照覆盖原文件 → 关闭当前工作簿 → 打开原路径。
        /// 如果原路径丢失（旧快照），复制到临时文件用原名打开。
        /// </summary>
        public bool Rollback(string snapshotId)
        {
            if (string.IsNullOrEmpty(snapshotId)) return false;

            var backupPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");
            if (!File.Exists(backupPath)) return false;

            try
            {
                // 读元数据拿原路径
                string originalPath = null;
                string workbookName = null;
                var metaPath = Path.Combine(_snapshotFolder, $"{snapshotId}.meta.json");
                if (File.Exists(metaPath))
                {
                    try
                    {
                        var meta = JsonSerializer.Deserialize<SnapshotMeta>(File.ReadAllText(metaPath));
                        originalPath = meta?.OriginalPath;
                        workbookName = meta?.WorkbookName;
                    }
                    catch { /* 元数据损坏，走兜底 */ }
                }

                // 关闭当前活动工作簿（不保存，避免覆盖快照前的状态）
                var currentWb = _app.ActiveWorkbook;
                string currentPath = null;
                if (currentWb != null)
                {
                    try { currentPath = currentWb.Path; } catch { }
                    try { currentWb.Close(false); } catch { }
                }

                // 决定恢复目标路径
                string targetPath;
                if (!string.IsNullOrEmpty(originalPath) && Directory.Exists(Path.GetDirectoryName(originalPath)))
                {
                    // 优先恢复到原路径
                    targetPath = originalPath;
                    try
                    {
                        File.Copy(backupPath, targetPath, overwrite: true);
                    }
                    catch
                    {
                        // 原路径不可写（被占用等），兜底用临时文件
                        targetPath = null;
                    }
                }
                else
                {
                    targetPath = null;
                }

                if (targetPath == null)
                {
                    // 兜底：复制到临时目录，用原工作簿名（避免 GUID 名）
                    string safeName = !string.IsNullOrEmpty(workbookName) ? workbookName : "restored.xlsx";
                    // 如果同名文件被占用，加时间戳
                    string tempDir = Path.Combine(Path.GetTempPath(), "DeepExcelRestore");
                    Directory.CreateDirectory(tempDir);
                    string tempPath = Path.Combine(tempDir, safeName);
                    if (File.Exists(tempPath))
                    {
                        try { File.Delete(tempPath); } catch { }
                    }
                    File.Copy(backupPath, tempPath, overwrite: true);
                    targetPath = tempPath;
                }

                _app.Workbooks.Open(targetPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Rollback error: {ex}");
                return false;
            }
        }

        /// <summary>
        /// 清理旧快照（包括 .xlsx 和 .meta.json）
        /// </summary>
        private void CleanupOldSnapshots()
        {
            try
            {
                var files = new DirectoryInfo(_snapshotFolder).GetFiles("*.xlsx");
                if (files.Length <= MaxSnapshots) return;

                Array.Sort(files, (a, b) => a.CreationTime.CompareTo(b.CreationTime));

                int toDelete = files.Length - MaxSnapshots;
                for (int i = 0; i < toDelete; i++)
                {
                    try
                    {
                        files[i].Delete();
                        // 同步删除元数据
                        var metaPath = Path.ChangeExtension(files[i].FullName, ".meta.json");
                        if (File.Exists(metaPath)) File.Delete(metaPath);
                    }
                    catch { }
                }
            }
            catch
            {
                // 清理失败不影响主流程
            }
        }

        /// <summary>
        /// 清理所有快照
        /// </summary>
        public void ClearAllSnapshots()
        {
            try
            {
                var files = new DirectoryInfo(_snapshotFolder).GetFiles();
                foreach (var file in files)
                {
                    try { file.Delete(); } catch { }
                }
            }
            catch
            {
                // ignore
            }
        }

        /// <summary>
        /// 删除单个快照
        /// </summary>
        public bool DeleteSnapshot(string snapshotId)
        {
            try
            {
                var xlsxPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");
                var metaPath = Path.Combine(_snapshotFolder, $"{snapshotId}.meta.json");
                if (File.Exists(xlsxPath)) File.Delete(xlsxPath);
                if (File.Exists(metaPath)) File.Delete(metaPath);
                return true;
            }
            catch { return false; }
        }

        private static string SafeGetWorkbookName(Workbook wb)
        {
            try { return wb.Name; }
            catch { return "未知"; }
        }

        /// <summary>
        /// ★ 安全获取工作簿完整路径（FullName），用于回滚时恢复原路径。
        /// 未保存的工作簿（如"工作簿1"）FullName 是 "工作簿1"（无路径），返回 null。
        /// </summary>
        private static string SafeGetWorkbookPath(Workbook wb)
        {
            try
            {
                string fullName = wb.FullName;
                // 未保存的工作簿 FullName 就是 Name（无路径分隔符）
                if (string.IsNullOrEmpty(fullName) || !fullName.Contains("\\") && !fullName.Contains("/"))
                {
                    return null;
                }
                return fullName;
            }
            catch { return null; }
        }
    }

    /// <summary>
    /// 快照元数据（与前端 JSON 协议对应）
    /// </summary>
    public class SnapshotMeta
    {
        public string Id { get; set; }
        public string WorkbookName { get; set; }
        /// <summary>★ 创建快照时工作簿的完整路径，回滚时恢复到该路径，避免工作簿名变成快照 GUID</summary>
        public string OriginalPath { get; set; }
        public DateTime CreatedAt { get; set; }
        public string Reason { get; set; }
    }
}
