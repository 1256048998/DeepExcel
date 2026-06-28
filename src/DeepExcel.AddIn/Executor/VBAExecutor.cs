using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// VBA执行引擎 - 生成并执行VBA代码
    /// 安全网：执行前自动快照，失败自动回滚
    /// </summary>
    public class VBAExecutor
    {
        private readonly Microsoft.Office.Interop.Excel.Application _app;
        private readonly SnapshotManager _snapshots;
        private const string ModuleName = "DeepExcelModule";

        public VBAExecutor(Microsoft.Office.Interop.Excel.Application app, SnapshotManager snapshots)
        {
            _app = app;
            _snapshots = snapshots;
        }

        /// <summary>
        /// 执行VBA代码
        /// </summary>
        public ToolResult Execute(string vbaCode, string macroName = null)
        {
            if (string.IsNullOrWhiteSpace(vbaCode))
            {
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = "VBA代码为空"
                };
            }

            macroName = macroName ?? "DeepExcel_TempMacro";

            // 1. 提取纯过程代码（去掉Sub/End Sub）
            var procCode = ExtractProcedureBody(vbaCode, macroName);
            if (string.IsNullOrEmpty(procCode))
            {
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = "无法解析VBA代码"
                };
            }

            // 2. 执行前快照
            var snapshotId = _snapshots.CreateSnapshot();

            VBProject vbProject = null;
            VBComponent module = null;
            bool moduleCreated = false;

            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null)
                {
                    return new ToolResult
                    {
                        Name = "execute_vba",
                        Success = false,
                        Error = "没有活动工作簿"
                    };
                }

                vbProject = wb.VBProject as VBProject;
                if (vbProject == null)
                {
                    return new ToolResult
                    {
                        Name = "execute_vba",
                        Success = false,
                        Error = "VBA项目不可访问"
                    };
                }

                // 3. 创建或复用模块
                module = FindOrCreateModule(vbProject, out moduleCreated);
                module.CodeModule.AddFromString(procCode);

                // 4. 执行宏
                _app.Run($"{ModuleName}.{macroName}");

                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = true,
                    Data = new { snapshotId, macro = macroName }
                };
            }
            catch (Exception ex)
            {
                // 5. 失败则回滚
                if (!string.IsNullOrEmpty(snapshotId))
                {
                    _snapshots.Rollback(snapshotId);
                }

                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = ex.Message,
                    Data = new { rolledBack = true, snapshotId }
                };
            }
            finally
            {
                // 6. 清理临时代码
                try
                {
                    if (module != null && moduleCreated)
                    {
                        // 清理模块中刚刚添加的代码
                        var codeModule = module.CodeModule;
                        var lineCount = codeModule.CountOfLines;
                        if (lineCount > 0)
                        {
                            // 删除所有代码（保留模块）
                            codeModule.DeleteLines(1, lineCount);
                        }
                    }
                }
                catch
                {
                    // ignore
                }
            }
        }

        private VBComponent FindOrCreateModule(VBProject vbProject, out bool created)
        {
            // 查找是否已存在模块
            foreach (VBComponent comp in vbProject.VBComponents)
            {
                if (comp.Name == ModuleName)
                {
                    created = false;
                    return comp;
                }
            }

            // 不存在则创建
            created = true;
            return vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        }

        /// <summary>
        /// 从原始代码中提取过程体，包装为完整Sub
        /// </summary>
        private string ExtractProcedureBody(string code, string macroName)
        {
            // 如果已经是完整Sub则直接返回
            if (Regex.IsMatch(code, @"^\s*Sub\s+\w+", RegexOptions.Multiline))
            {
                return code;
            }

            // 否则包装为Sub
            return $"Sub {macroName}()\r\n{code}\r\nEnd Sub";
        }
    }
}
