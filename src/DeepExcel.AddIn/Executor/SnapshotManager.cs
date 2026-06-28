using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 快照管理器 - 操作前自动备份，失败可回滚
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
        /// 创建当前工作簿的快照
        /// </summary>
        public string CreateSnapshot()
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return null;

                var snapshotId = Guid.NewGuid().ToString("N");
                var backupPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");

                wb.SaveCopyAs(backupPath);

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
        /// 回滚到指定快照
        /// </summary>
        public bool Rollback(string snapshotId)
        {
            if (string.IsNullOrEmpty(snapshotId)) return false;

            var backupPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");
            if (!File.Exists(backupPath)) return false;

            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb != null)
                {
                    wb.Save();
                    wb.Close(false);
                }

                _app.Workbooks.Open(backupPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Rollback error: {ex}");
                return false;
            }
        }

        /// <summary>
        /// 清理旧快照
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
                    try { files[i].Delete(); }
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
                var files = new DirectoryInfo(_snapshotFolder).GetFiles("*.xlsx");
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
    }
}
