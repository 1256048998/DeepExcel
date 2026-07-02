using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using DeepExcel.AddIn.Diagnostics;

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

            // ★ P0-3 沙箱校验：阻止 LLM 执行 Shell / WScript.Shell / 文件系统 / 网络请求
            var sandboxError = DeepExcel.AddIn.Security.CodeSandbox.ValidateVba(procCode);
            if (sandboxError != null)
            {
                Logger.Instance.Warning("VBAExecutor", "Code blocked by sandbox: " + sandboxError);
                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = sandboxError
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

                // ★ 诊断：记录注入后的实际代码内容 + 模块状态，便于排查"宏被禁用"
                try
                {
                    var lineCount = module.CodeModule.CountOfLines;
                    var firstLines = lineCount > 0 ? module.CodeModule.get_Lines(1, Math.Min(lineCount, 8)) : "";
                    Logger.Instance.Info("VBAExecutor",
                        $"After AddFromString: module={module.Name}, lineCount={lineCount}, firstLines=\n{firstLines}");
                }
                catch (Exception logEx) { Logger.Instance.Warning("VBAExecutor", "Log code failed: " + logEx.Message); }

                // ★ 第一性原理：_app.Run 在 UI 线程同步执行 VBA，期间 WebView2 渲染进程
                // 无法与 UI 线程通信（消息泵被阻塞）。如果 VBA 执行超过几秒，
                // WebView2 进程会因心跳超时崩溃，导致对话面板白屏/消失。
                // 解决：执行前处理一次消息泵，让 WebView2 有机会完成 pending 的渲染。
                Logger.Instance.Info("VBAExecutor", $"VBA Run START: macro={ModuleName}.{macroName}, codeLen={procCode.Length}");
                System.Windows.Forms.Application.DoEvents();

                // 4. 执行宏
                // ★ 临时关闭 Application.ScreenUpdating 提升性能 + 防闪烁
                bool prevScreenUpdating = _app.ScreenUpdating;
                _app.ScreenUpdating = false;
                var runSw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    _app.Run($"{ModuleName}.{macroName}");
                }
                finally
                {
                    _app.ScreenUpdating = prevScreenUpdating;
                }
                runSw.Stop();

                // 执行后立即处理消息泵，恢复 WebView2 心跳
                System.Windows.Forms.Application.DoEvents();
                Logger.Instance.Info("VBAExecutor", $"VBA Run END: elapsed={runSw.ElapsedMilliseconds}ms");

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

                // ★ 友好错误信息：宏被禁用时给出明确指引
                string errMsg = ex.Message;
                if (errMsg.Contains("宏") && (errMsg.Contains("禁用") || errMsg.Contains("不可用")))
                {
                    Logger.Instance.Warning("VBAExecutor", "宏被禁用，检查宏安全设置。错误: " + errMsg);
                    errMsg = "Excel 宏被禁用，无法执行 VBA。请在 Excel 中开启宏权限：\n" +
                             "文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 选\"启用所有宏\"，\n" +
                             "同时在\"开发工具\"中勾选\"信任对 VBA 工程对象模型的访问\"。";
                }

                return new ToolResult
                {
                    Name = "execute_vba",
                    Success = false,
                    Error = errMsg,
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

            // ★ 清理之前因 bug 创建的未命名模块（Module1/Module2/...）
            // 这些模块是旧版本 FindOrCreateModule 没设置 Name 导致的残留
            var toRemove = new List<VBComponent>();
            foreach (VBComponent comp in vbProject.VBComponents)
            {
                // 标准模块类型 (vbext_ct_StdModule=1) 且名以 "Module" 开头
                if (comp.Type == vbext_ComponentType.vbext_ct_StdModule
                    && (comp.Name?.StartsWith("Module") ?? false))
                {
                    toRemove.Add(comp);
                }
            }
            foreach (var comp in toRemove)
            {
                try
                {
                    Logger.Instance.Info("VBAExecutor", $"FindOrCreateModule: removing stale module {comp.Name}");
                    vbProject.VBComponents.Remove(comp);
                }
                catch (Exception ex) { Logger.Instance.Warning("VBAExecutor", "Remove stale module failed: " + ex.Message); }
            }

            // 不存在则创建
            created = true;
            var newModule = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            // ★ 必须设置 Name 属性！否则模块名是默认的 "Module1/Module2/..."，
            // 而 _app.Run("DeepExcelModule.DeepExcel_TempMacro") 会找不到模块，报"宏被禁用"
            newModule.Name = ModuleName;
            Logger.Instance.Info("VBAExecutor", $"FindOrCreateModule: created new module, name={newModule.Name}");
            return newModule;
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
