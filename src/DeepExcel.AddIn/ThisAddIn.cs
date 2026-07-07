using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using DeepExcel.AddIn.Bridge;

namespace DeepExcel.AddIn
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A")]
    [ProgId("DeepExcel.AddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCallbacks))]
    public class ThisAddIn : Object, IDTExtensibility2, IRibbonExtensibility, IRibbonCallbacks, ICustomTaskPaneConsumer
    {
        private Microsoft.Office.Interop.Excel.Application _excelApp;
        private TaskPaneControl _taskPane;
        // ★ 按窗口缓存 CTP：Excel SDI 模式下每个 workbook 有独立窗口，
        //   需要为每个窗口创建独立的 CustomTaskPane，否则切换到其他窗口时面板不可见。
        // ★ 字典 key 用窗口标题（Caption，即 workbook 名）而非 Window COM 对象本身，
        //   因为每次访问 ActiveWindow 返回的 RCW 引用不同，用对象做 key 会导致同一窗口被重复创建 CTP。
        private readonly System.Collections.Generic.Dictionary<string, Microsoft.Office.Core.CustomTaskPane> _ctpsByWindow
            = new System.Collections.Generic.Dictionary<string, Microsoft.Office.Core.CustomTaskPane>();
        // ★ 每个 CTP 对应的 workbook key（FullName），用于会话消息路由
        private readonly System.Collections.Generic.Dictionary<Microsoft.Office.Core.CustomTaskPane, string> _workbookKeyByCtp
            = new System.Collections.Generic.Dictionary<Microsoft.Office.Core.CustomTaskPane, string>();
        private ICTPFactory _ctpFactory;
        private MessageBridge _bridge;
        private IRibbonUI _ribbon;
        // 隐藏的 UI 控件，用于 sidecar 的 UI 线程封送
        private Control _sidecarUiControl;
        // ★ 缓存的 WebView2 environment，所有窗口的 WebView 共用同一个 environment
        // （WebView2 推荐每个进程使用一个 environment，避免 userDataFolder 锁冲突）
        private Microsoft.Web.WebView2.Core.CoreWebView2Environment _webViewEnv;
        // ★ 标记 bridge 是否已设置 SetSendToUi（只设置一次，避免 sidecar 多次启动）
        private bool _bridgeWired;
        // ★ C-2 修复：记录 AccessVBOM 原始值，加载项关闭时恢复，避免全局降级 Office 安全策略
        private int? _originalAccessVbom;
        private bool _vbaSecurityModified;

        public ThisAddIn()
        {
            Log("Constructor called - COM object being created");
            // COM 加载项不读取 .dll.config，需手动处理 binding redirect
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
        }

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var name = new AssemblyName(args.Name);
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                // System.Runtime.CompilerServices.Unsafe 4.0.4.1 → 6.0.0.0 重定向
                if (name.Name == "System.Runtime.CompilerServices.Unsafe")
                {
                    string path = Path.Combine(dir, "System.Runtime.CompilerServices.Unsafe.dll");
                    if (File.Exists(path)) return Assembly.LoadFrom(path);
                }
                // 其他 System.* 依赖也按文件名加载
                if (name.Name.StartsWith("System.") || name.Name.StartsWith("Microsoft."))
                {
                    string path = Path.Combine(dir, name.Name + ".dll");
                    if (File.Exists(path))
                    {
                        var existing = AppDomain.CurrentDomain.GetAssemblies()
                            .FirstOrDefault(a => a.GetName().Name == name.Name);
                        if (existing != null) return existing;
                        return Assembly.LoadFrom(path);
                    }
                }
            }
            catch (Exception ex)
            {
                try { Log("OnAssemblyResolve FAILED for " + args.Name + ": " + ex.Message); } catch { }
            }
            return null;
        }

        private static void Log(string message)
        {
            try
            {
                // ★ H-2 修复：日志改用 %APPDATA%\DeepExcel\logs\，避免写入 DLL 所在目录（可能受 UAC 限制）
                // 并添加日志轮转：超过 5MB 时重命名为 .bak 并新建日志文件
                string logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "DeepExcel", "logs");
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                string logPath = Path.Combine(logDir, "DeepExcel_Load.log");

                // ★ L-4 修复：日志轮转，避免无限增长
                try
                {
                    var fi = new FileInfo(logPath);
                    if (fi.Exists && fi.Length > 5 * 1024 * 1024)
                    {
                        string bakPath = logPath + ".bak";
                        try { if (File.Exists(bakPath)) File.Delete(bakPath); } catch { }
                        File.Move(logPath, bakPath);
                    }
                }
                catch { }

                // ★ H-2 修复：对日志内容做基本转义，防止日志注入（移除换行符）
                string safeMessage = (message ?? "").Replace("\r", " ").Replace("\n", " ");
                File.AppendAllText(logPath, "[" + DateTime.Now + "] " + safeMessage + Environment.NewLine);
            }
            catch { }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                Log("OnConnection started, connectMode=" + connectMode);
                _excelApp = (Microsoft.Office.Interop.Excel.Application)application;
                Log("Application object stored, type=" + (application == null ? "null" : application.GetType().FullName));

                // addInInst 是 COMAddIn 对象，不实现 ICTPFactory
                // ICTPFactory 需从 Excel Application 的 CTPFactory 属性获取（PIA 可能未暴露，用后期绑定）
                Log("OnConnection completed successfully");
            }
            catch (Exception ex)
            {
                Log("OnConnection FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
                throw;
            }
        }

        /// <summary>
        /// 检查 ICTPFactory 是否已通过 ICustomTaskPaneConsumer.CTPFactoryAvailable 获取
        /// </summary>
        private bool TryAcquireCTPFactory()
        {
            if (_ctpFactory != null)
            {
                Log("ICTPFactory available (from ICustomTaskPaneConsumer)");
                return true;
            }
            Log("ICTPFactory not yet available (ICustomTaskPaneConsumer not called by Office)");
            return false;
        }

        // ============= 工作簿事件处理 =============

        /// <summary>
        /// ★ 检查并设置 VBA 安全注册表项：
        /// - AccessVBOM=1：信任对 VBA 工程对象模型的访问（必须，否则 wb.VBProject 不可访问）
        /// ★ C-2 修复：不再写 VBAWarnings=1（保留用户原有宏警告设置），避免全局降级 Office 安全策略。
        /// 仅临时启用 AccessVBOM，加载项关闭时通过 RestoreVbaSecurity 恢复原始值。
        /// 注册表路径：HKCU\Software\Microsoft\Office\{version}\Excel\Security
        /// HKCU 不需要管理员权限，用户级即可。
        /// </summary>
        private void EnsureVbaSecuritySettings()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Office\" + _excelApp.Version + @"\Excel\Security", true))
                {
                    if (key == null)
                    {
                        Log("VBA Security: registry key not found, creating...");
                        using (var createKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                            @"Software\Microsoft\Office\" + _excelApp.Version + @"\Excel\Security"))
                        {
                            // ★ C-2 修复：不写 VBAWarnings，保留用户原有宏警告设置
                            createKey.SetValue("AccessVBOM", 1, Microsoft.Win32.RegistryValueKind.DWord);
                            _originalAccessVbom = 0; // 原值不存在视为 0
                            _vbaSecurityModified = true;
                        }
                        Log("VBA Security: registry key created, AccessVBOM=1 (VBAWarnings not modified)");
                    }
                    else
                    {
                        var accessVbom = key.GetValue("AccessVBOM");

                        if (accessVbom == null || (int)accessVbom != 1)
                        {
                            // ★ C-2 修复：记录原始值用于关闭时恢复
                            _originalAccessVbom = accessVbom == null ? (int?)null : (int)accessVbom;
                            key.SetValue("AccessVBOM", 1, Microsoft.Win32.RegistryValueKind.DWord);
                            _vbaSecurityModified = true;
                            Log("VBA Security: AccessVBOM set to 1 (was " + (accessVbom?.ToString() ?? "null") + "), will restore on shutdown");
                        }
                        else
                        {
                            Log("VBA Security: AccessVBOM already = 1 (no change needed)");
                        }
                        // ★ C-2 修复：不再修改 VBAWarnings，保留用户原有设置
                    }
                }
            }
            catch (Exception ex)
            {
                Log("EnsureVbaSecuritySettings error: " + ex.Message);
            }
        }

        /// <summary>
        /// ★ C-2 修复：恢复 AccessVBOM 原始值。在 OnBeginShutdown 中调用，
        /// 避免加载项卸载后用户的安全设置仍被降级。
        /// </summary>
        private void RestoreVbaSecurity()
        {
            if (!_vbaSecurityModified || _excelApp == null) return;
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Office\" + _excelApp.Version + @"\Excel\Security", true))
                {
                    if (key == null) return;
                    if (_originalAccessVbom.HasValue)
                    {
                        key.SetValue("AccessVBOM", _originalAccessVbom.Value, Microsoft.Win32.RegistryValueKind.DWord);
                        Log("VBA Security: AccessVBOM restored to " + _originalAccessVbom.Value);
                    }
                    else
                    {
                        key.DeleteValue("AccessVBOM", false);
                        Log("VBA Security: AccessVBOM value deleted (was not present originally)");
                    }
                }
                _vbaSecurityModified = false;
            }
            catch (Exception ex)
            {
                Log("RestoreVbaSecurity error: " + ex.Message);
            }
        }

        /// <summary>
        /// 工作簿关闭前：清理对应会话和 CTP 记录。
        /// ★ 工具执行保护：VBA/Python 执行期间 Excel 可能因宏失败误触发此事件，
        ///   此时 _toolExecutionGuard=true，跳过清理避免会话和面板被误销毁。
        /// </summary>
        private void OnWorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            try
            {
                string wbKey = GetWorkbookKey(Wb);
                Log($"WorkbookBeforeClose: " + wbKey);

                // ★ 工具执行保护：执行期间不清理，避免 VBA 宏失败导致的误触发
                if (Sidecar.ToolDispatcher.ExecutionGuardActive)
                {
                    Log("WorkbookBeforeClose SKIPPED (tool execution guard active): " + wbKey);
                    return;
                }

                // 通知 bridge 清理会话
                if (_bridge != null)
                {
                    _bridge.OnWorkbookClose(wbKey);
                }

                // 清理该工作簿对应的 CTP 记录
                var ctpsToRemove = new List<Microsoft.Office.Core.CustomTaskPane>();
                foreach (var kvp in _workbookKeyByCtp)
                {
                    if (kvp.Value == wbKey)
                        ctpsToRemove.Add(kvp.Key);
                }
                foreach (var ctp in ctpsToRemove)
                {
                    try
                    {
                        _workbookKeyByCtp.Remove(ctp);
                        try { ctp.Delete(); } catch { }
                        Log("Removed CTP for workbook: " + wbKey);
                    }
                    catch (Exception ctpEx) { Log("CTP cleanup error: " + ctpEx.Message); }
                }

                // 清理 _ctpsByWindow 中对应窗口的记录
                try
                {
                    string wbName = Wb.Name ?? "";
                    if (_ctpsByWindow.ContainsKey(wbName))
                    {
                        try { _ctpsByWindow[wbName].Delete(); } catch { }
                        _ctpsByWindow.Remove(wbName);
                        Log("Removed window CTP entry: " + wbName);
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                Log("OnWorkbookBeforeClose FAILED: " + ex.Message);
            }
        }

        /// <summary>
        /// 工作簿保存后：如果是另存为（FullName 变了），更新会话 key 和 CTP 对应关系
        /// </summary>
        private void OnWorkbookAfterSave(Workbook Wb, bool Success)
        {
            try
            {
                if (!Success) return;

                string newKey = GetWorkbookKey(Wb);
                string newName = Wb.Name ?? "";
                Log($"WorkbookAfterSave: newKey={newKey}, newName={newName}");

                // 尝试从 CTP 记录中找旧的 key（用窗口标题可能是旧名称）
                string oldKey = null;
                try
                {
                    // 先尝试用 Name 找对应的 CTP
                    foreach (var kvp in _workbookKeyByCtp)
                    {
                        try
                        {
                            var ctp = kvp.Key;
                            if (ctp == null) continue;
                            // 检查这个 CTP 的窗口是否对应这个工作簿
                            string ctpKey = kvp.Value;
                            // 如果 key 包含旧的 workbook，可能是旧的 FullName 或旧的 Name
                            // 简单策略：如果 newKey 是 FullName（有路径），找一个旧 key 可能是旧的 Name 或旧的 FullName
                            if (ctpKey != newKey)
                            {
                                // 检查旧 key 对应的工作簿是不是这个
                                // 用窗口标题匹配：CTP 标题应该和工作簿名一致
                                try
                                {
                                    var pane = ctp.ContentControl as TaskPaneControl;
                                    // 不太好直接关联，用另一种方式：
                                    // 遍历所有窗口，找到该工作簿的窗口，其标题匹配的 CTP 的 key 更新
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // 更好的方法：遍历所有窗口，找到对应这个 Workbook 的窗口，更新其 CTP 的 workbook key
                try
                {
                    foreach (Window wnd in Wb.Windows)
                    {
                        try
                        {
                            string caption = wnd.Caption as string ?? "";
                            if (_ctpsByWindow.TryGetValue(caption, out var ctp) && ctp != null)
                            {
                                // 找到旧 key
                                if (_workbookKeyByCtp.TryGetValue(ctp, out oldKey))
                                {
                                    // 更新 CTP 对应关系
                                    _workbookKeyByCtp.Remove(ctp);
                                    _workbookKeyByCtp[ctp] = newKey;
                                    Log($"Updated CTP workbook key: {oldKey} -> {newKey}");
                                    break; // 找到一个就够了（通常只有一个主窗口）
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // 通知 bridge 更新会话 key
                if (_bridge != null && oldKey != null && oldKey != newKey)
                {
                    _bridge.OnWorkbookAfterSave(oldKey, newKey, newName);
                }
                else if (_bridge != null && oldKey == null)
                {
                    // 没找到旧 key，尝试用 Name 作为旧 key 试试
                    string oldNameKey = Wb.Name ?? "";
                    if (oldNameKey != newKey)
                    {
                        _bridge.OnWorkbookAfterSave(oldNameKey, newKey, newName);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("OnWorkbookAfterSave FAILED: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            Log("OnDisconnection called, mode=" + disconnectMode);
            try
            {
                // 取消事件订阅
                try
                {
                    if (_excelApp != null)
                    {
                        _excelApp.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                        _excelApp.WorkbookAfterSave -= OnWorkbookAfterSave;
                    }
                }
                catch { }

                if (_bridge is IDisposable disposable) disposable.Dispose();
                _bridge = null;

                if (_sidecarUiControl != null && !_sidecarUiControl.IsDisposed)
                {
                    _sidecarUiControl.Dispose();
                    _sidecarUiControl = null;
                }

                // ★ 清理所有窗口的 CTP（SDI 多窗口）
                foreach (var kvp in _ctpsByWindow)
                {
                    try
                    {
                        if (kvp.Value != null) kvp.Value.Delete();
                    }
                    catch (Exception ctpEx) { Log("CTP cleanup error: " + ctpEx.Message); }
                }
                _ctpsByWindow.Clear();
                _workbookKeyByCtp.Clear();
                _taskPane = null;

                if (_ctpFactory != null)
                {
                    Marshal.ReleaseComObject(_ctpFactory);
                    _ctpFactory = null;
                }
            }
            catch (Exception ex) { Log("OnDisconnection cleanup error: " + ex.Message); }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Log("OnAddInsUpdate called");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete called");
            try
            {
                _sidecarUiControl = new Form();
                // 强制创建窗口句柄，否则 sidecar 的 stdout/stderr 回调在工作线程
                // 调用 BeginInvoke 时会抛 InvalidOperationException 导致 Excel 崩溃
                var handle = _sidecarUiControl.Handle;
                Log("Sidecar UI control handle created: " + handle);
                _bridge = new MessageBridge(_excelApp, _sidecarUiControl);
                Log("MessageBridge initialized with sidecar");

                // 监听工作簿事件，实现会话生命周期管理
                _excelApp.WorkbookBeforeClose += OnWorkbookBeforeClose;
                _excelApp.WorkbookAfterSave += OnWorkbookAfterSave;
                Log("Workbook event handlers registered");

                // ★ 检查并尝试设置宏安全：VBA 功能依赖 AccessVBOM=1（信任对 VBA 工程对象模型的访问）
                // 和 VBAWarnings=1（启用所有宏）。如果不设置，execute_vba 会报"宏被禁用"。
                try
                {
                    EnsureVbaSecuritySettings();
                }
                catch (Exception secEx) { Log("EnsureVbaSecuritySettings failed: " + secEx.Message); }

                // ★ 性能优化：预启动当前活动工作簿的 sidecar
                // 原来要等用户打开面板+发第一条消息才启动 Python 进程，
                // 现在 Excel 启动后就后台预热，用户打开面板时已经准备好了。
                try
                {
                    var wb = _excelApp.ActiveWorkbook;
                    if (wb != null)
                    {
                        Log("Pre-warming sidecar for active workbook...");
                        _bridge.PreWarmActiveSession();
                        Log("Sidecar pre-warm initiated");
                    }
                }
                catch (Exception prewarmEx)
                {
                    Log("Sidecar pre-warm failed (non-critical): " + prewarmEx.Message);
                }
            }
            catch (Exception ex)
            {
                Log("MessageBridge init FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown called");
            // ★ C-2 修复：关闭时恢复 AccessVBOM 原始值，避免全局降级 Office 安全策略
            try { RestoreVbaSecurity(); } catch (Exception ex) { Log("RestoreVbaSecurity in OnBeginShutdown failed: " + ex.Message); }
        }

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI called, RibbonID=" + RibbonID);
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                const string resourceName = "DeepExcel.AddIn.Resources.DeepExcelRibbon.xml";
                using (Stream stream = asm.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        Log("GetCustomUI FAILED: resource not found - " + resourceName);
                        return null;
                    }
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string xml = reader.ReadToEnd();
                        Log("GetCustomUI returning XML, length=" + xml.Length);
                        return xml;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("GetCustomUI FAILED: " + ex.GetType().Name + " - " + ex.Message);
                return null;
            }
        }

        #endregion

        #region ICustomTaskPaneConsumer 实现

        /// <summary>
        /// Office 在加载 COM 加载项时主动调用此方法，传入 ICTPFactory
        /// 这比从 Application.CTPFactory 反射获取更可靠（PIA 版本无关）
        /// </summary>
        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            Log("CTPFactoryAvailable called by Office, factory type=" + (CTPFactoryInst == null ? "null" : CTPFactoryInst.GetType().FullName));
            try
            {
                _ctpFactory = CTPFactoryInst;
                Log("ICTPFactory stored from ICustomTaskPaneConsumer");
            }
            catch (Exception ex)
            {
                Log("CTPFactoryAvailable failed: " + ex.GetType().Name + " - " + ex.Message);
            }
        }

        #endregion

        #region IRibbonCallbacks 显式实现

        void IRibbonCallbacks.OnRibbonLoad(object ribbon)
        {
            Log("OnRibbonLoad called, ribbon type=" + (ribbon == null ? "null" : ribbon.GetType().FullName));
            try
            {
                _ribbon = (IRibbonUI)ribbon;
                Log("Ribbon UI stored");
            }
            catch (Exception ex) { Log("OnRibbonLoad cast failed: " + ex.Message); }
        }

        void IRibbonCallbacks.OnTogglePanel(object control)
        {
            Log("OnTogglePanel called");
            try
            {
                // ★ 按钮仅用于"打开面板"，不切换可见性。
                // 关闭面板仅通过面板右上角的关闭按钮（CustomTaskPane 的 X）。
                // ★ SDI 多窗口支持：Excel 2013+ 每个 workbook 是独立窗口，
                //   必须为每个 ActiveWindow 创建独立的 CustomTaskPane，否则切换窗口时面板不可见。
                if (!TryAcquireCTPFactory())
                {
                    Log("ICTPFactory not available, cannot create panel");
                    MessageBox.Show("面板初始化失败（ICTPFactory 不可用），请重启 Excel。",
                        "DeepExcel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 获取当前活动窗口
                Microsoft.Office.Interop.Excel.Window activeWindow = null;
                try { activeWindow = _excelApp.ActiveWindow; }
                catch (Exception wex) { Log("Get ActiveWindow failed: " + wex.Message); }
                if (activeWindow == null)
                {
                    Log("ActiveWindow is null, cannot create panel");
                    MessageBox.Show("请先打开一个工作簿再点击「打开面板」。",
                        "DeepExcel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string windowCaption = activeWindow.Caption as string ?? "";
                Log("ActiveWindow caption=" + windowCaption);

                // ★ 用 Caption（workbook 名）作为字典 key，而非 Window COM 对象本身
                string windowKey = windowCaption;

                // 查找或创建该窗口对应的 CTP
                Microsoft.Office.Core.CustomTaskPane ctp;
                if (!_ctpsByWindow.TryGetValue(windowKey, out ctp) || ctp == null)
                {
                    // ★ 清理字典中已失效的 CTP（用户点 X 关闭后 CTP 对象会失效）
                    PurgeInvalidCtps();

                    Log("Creating CustomTaskPane for window: " + windowKey);
                    // ★ 第三个参数传 activeWindow，让 CTP 绑定到该窗口
                    ctp = _ctpFactory.CreateCTP(
                        "DeepExcel.AddIn.TaskPaneControl",
                        "DeepExcel AI",
                        activeWindow);
                    Log("CustomTaskPane created");

                    ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    ctp.Width = 420;
                    Log("CustomTaskPane docked right, width=420");

                    _ctpsByWindow[windowKey] = ctp;

                    // ★ 记录该 CTP 所属的 workbook key（FullName），用于会话消息路由
                    try
                    {
                        var wb = activeWindow.Parent as Workbook;
                        if (wb != null)
                        {
                            string wbKey = GetWorkbookKey(wb);
                            _workbookKeyByCtp[ctp] = wbKey;
                            Log($"CTP workbook key: {wbKey}");
                        }
                    }
                    catch (Exception wbEx) { Log("Get workbook key for CTP failed: " + wbEx.Message); }

                    // 第一个窗口的 TaskPaneControl 用于 WebView 初始化
                    if (_taskPane == null)
                    {
                        _taskPane = (TaskPaneControl)ctp.ContentControl;
                        Log("TaskPaneControl retrieved from ContentControl (first window)");
                        InitializeWebView();
                    }
                    else
                    {
                        // 后续窗口的 TaskPaneControl 也需要初始化 WebView
                        var newPane = (TaskPaneControl)ctp.ContentControl;
                        InitializeWebViewForPane(newPane);
                    }
                }
                try
                {
                    if (!ctp.Visible)
                    {
                        ctp.Visible = true;
                        Log("CustomTaskPane opened (was hidden) for window: " + windowKey);
                    }
                    else
                    {
                        Log("CustomTaskPane already visible for window: " + windowKey);
                    }

                    // ★ 性能优化：面板打开时预启动 sidecar
                    // 虽然 OnStartupComplete 已经预热过，但如果是新建工作簿
                    // 或预热失败，这里确保 sidecar 在用户输入前就启动好
                    try
                    {
                        if (_bridge != null)
                        {
                            _bridge.PreWarmActiveSession();
                        }
                    }
                    catch (Exception prewarmEx)
                    {
                        Log("Panel-open prewarm failed (non-critical): " + prewarmEx.Message);
                    }
                }
                catch (Exception visEx)
                {
                    // CTP 已失效（被用户关闭），从字典移除并重新创建
                    Log("CTP.Visible threw: " + visEx.Message + " - recreating");
                    _ctpsByWindow.Remove(windowKey);
                    try { ctp.Delete(); } catch { }
                    // 递归一次重新创建
                    ((IRibbonCallbacks)this).OnTogglePanel(control);
                    return;
                }
            }
            catch (Exception ex)
            {
                Log("OnTogglePanel FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
                MessageBox.Show("打开面板失败: " + ex.Message, "DeepExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IRibbonCallbacks.OnShowHelp(object control)
        {
            Log("OnShowHelp called");
            MessageBox.Show(
                "DeepExcel AI Agent\n\n" +
                "1. 点击「打开面板」按钮显示 AI 面板\n" +
                "2. 在面板中输入你的需求\n" +
                "3. AI 会自动调用工具处理 Excel\n\n" +
                "支持的功能：公式生成、VBA/Python 执行、数据清洗、图表创建",
                "DeepExcel 帮助",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        private async void InitializeWebView()
        {
            try
            {
                Log("InitializeWebView started (first window)");

                // 指定用户数据文件夹到 LocalAppData，避免在 Excel 工作目录下创建失败 (E_ACCESSDENIED)
                // 加上进程 ID 后缀，支持多个 Excel 实例同时加载 DeepExcel（WebView2 会独占锁定 userDataFolder）
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "DeepExcel", "WebView2_" + System.Diagnostics.Process.GetCurrentProcess().Id);
                Directory.CreateDirectory(userDataFolder);
                Log("WebView2 userDataFolder=" + userDataFolder);

                var options = new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions
                {
                    AdditionalBrowserArguments = "--disable-features=RendererCodeIntegrity"
                };
                _webViewEnv = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                    null, userDataFolder, options);
                Log("WebView2 environment created (cached for all windows)");

                // 初始化第一个窗口的 WebView
                await InitializePaneWebViewAsync(_taskPane);

                // 设置 bridge 的 SetSendToUi 回调为广播模式（只设置一次，避免 sidecar 重复启动）
                SetupBridgeBroadcast();

                Log("WebView and bridge wiring complete (first window)");
            }
            catch (Exception ex)
            {
                Log("InitializeWebView FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
            }
        }

        /// <summary>
        /// ★ 后续窗口的 WebView 初始化：复用已缓存的 _webViewEnv，不重复创建 environment。
        /// 不再调用 SetSendToUi（已在第一个窗口设置好广播），只绑定本窗口的 WebMessageReceived。
        /// </summary>
        private async void InitializeWebViewForPane(TaskPaneControl pane)
        {
            try
            {
                Log("InitializeWebViewForPane started for additional window");

                if (_webViewEnv == null)
                {
                    Log("InitializeWebViewForPane FAILED: WebView2 environment not ready (first window init failed?)");
                    return;
                }
                if (pane == null)
                {
                    Log("InitializeWebViewForPane FAILED: pane is null");
                    return;
                }

                await InitializePaneWebViewAsync(pane);
                Log("InitializeWebViewForPane complete for additional window");
            }
            catch (Exception ex)
            {
                Log("InitializeWebViewForPane FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
            }
        }

        /// <summary>
        /// 通用 WebView 初始化：用缓存的 _webViewEnv 初始化指定 pane 的 WebView2 控件，
        /// 设置虚拟主机映射、导航到 index.html，并绑定 WebMessageReceived。
        /// 第一个窗口和后续窗口都调用此方法。
        /// </summary>
        private async System.Threading.Tasks.Task InitializePaneWebViewAsync(TaskPaneControl pane)
        {
            if (pane == null || pane.IsDisposed || pane.WebView == null)
            {
                Log("InitializePaneWebViewAsync: pane invalid, skip");
                return;
            }

            await pane.WebView.EnsureCoreWebView2Async(_webViewEnv);
            Log("WebView2 core initialized for pane");

            string assetsPath = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "WebViewAssets", "index.html");
            Log("Loading WebView from: " + assetsPath);

            pane.WebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "deepexcel.local",
                Path.GetDirectoryName(assetsPath),
                Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);

            // ★ 加时间戳查询参数绕过缓存：每次 Excel 启动都加载最新页面
            // 否则 WebView2 会缓存 index.html 和 JS/CSS，导致前端更新不生效
            string cacheBuster = DateTime.Now.ToString("yyyyMMddHHmmss");
            pane.WebView.CoreWebView2.Navigate($"https://deepexcel.local/index.html?v={cacheBuster}");
            Log("WebView navigated for pane (cache-buster=" + cacheBuster + ")");

            // ★ 每个窗口的 WebView 都要监听前端消息，路由到 bridge
            // 用局部变量捕获 pane，避免闭包引用被修改的字段
            var paneRef = pane;
            var webViewRef = pane.WebView;
            pane.WebView.CoreWebView2.WebMessageReceived += (sender, e) =>
            {
                try
                {
                    string json = e.WebMessageAsJson;
                    Log("WebMessageReceived: " + (json == null ? "null" : (json.Length > 300 ? json.Substring(0, 300) + "..." : json)));
                    string response = _bridge.HandleMessage(json);
                    Log("HandleMessage response: " + (response == null ? "null" : (response.Length > 300 ? response.Substring(0, 300) + "..." : response)));
                    if (!string.IsNullOrEmpty(response))
                    {
                        try
                        {
                            if (webViewRef != null && webViewRef.CoreWebView2 != null && !webViewRef.IsDisposed)
                                webViewRef.CoreWebView2.PostWebMessageAsString(response);
                        }
                        catch (Exception ex) { Log("PostWebMessageAsString response error: " + ex.Message); }
                    }
                }
                catch (Exception ex)
                {
                    Log("WebMessageReceived FAILED: " + ex.GetType().Name + " - " + ex.Message);
                    Log("HRESULT: 0x" + Marshal.GetHRForException(ex).ToString("X8"));
                    Log("Stack: " + ex.StackTrace);
                }
            };
        }

        /// <summary>
        /// ★ 设置 bridge 的消息路由：
        /// 1. OnSessionUiMessage 事件：按 workbook key 精准分发给对应 CTP 的 WebView（会话隔离核心）
        /// 2. 旧 SetSendToUi 通道：保留兼容，广播给所有可见 CTP
        /// 只调用一次（第一个窗口初始化时）。
        /// </summary>
        private void SetupBridgeBroadcast()
        {
            if (_bridgeWired)
            {
                Log("SetupBridgeBroadcast: already wired, skip");
                return;
            }
            _bridgeWired = true;

            // ★ 会话级路由：按 workbook key 分发（主通道，解决上下文串扰）
            _bridge.OnSessionUiMessage += (workbookKey, json) =>
            {
                try
                {
                    Log($"OnSessionUiMessage: wb={workbookKey}, len={json?.Length ?? 0}");
                    DispatchToWorkbookPane(workbookKey, json);
                }
                catch (Exception ex) { Log("OnSessionUiMessage handler error: " + ex.Message); }
            };

            // 旧通道（兼容兜底）：广播给所有可见 CTP
            _bridge.SetSendToUi(json =>
            {
                try
                {
                    BroadcastToAllVisiblePanes(json);
                }
                catch (Exception ex) { Log("SetSendToUi broadcast error: " + ex.Message); }
            });

            Log("SetupBridgeBroadcast: bridge wired (session-routed + broadcast fallback)");
        }

        /// <summary>
        /// 按 workbook key 精准分发给对应 CTP 的 WebView
        /// </summary>
        private void DispatchToWorkbookPane(string workbookKey, string json)
        {
            var invalidCtps = new System.Collections.Generic.List<Microsoft.Office.Core.CustomTaskPane>();
            bool delivered = false;

            foreach (var kvp in _workbookKeyByCtp)
            {
                var ctp = kvp.Key;
                if (ctp == null) { invalidCtps.Add(ctp); continue; }

                bool visible;
                try { visible = ctp.Visible; }
                catch { invalidCtps.Add(ctp); continue; }

                // 只发给匹配的 workbook，且可见的 CTP
                if (kvp.Value != workbookKey) continue;
                if (!visible) continue;

                TaskPaneControl pane = null;
                try { pane = (TaskPaneControl)ctp.ContentControl; }
                catch { invalidCtps.Add(ctp); continue; }
                if (pane == null || pane.IsDisposed || pane.WebView == null) continue;

                var webView = pane.WebView;
                try
                {
                    if (webView.InvokeRequired)
                    {
                        webView.BeginInvoke(new Action<string>(msg =>
                        {
                            try
                            {
                                if (webView.CoreWebView2 != null && !webView.IsDisposed)
                                    webView.CoreWebView2.PostWebMessageAsString(msg);
                            }
                            catch (Exception ex) { Log("DispatchToWorkbookPane (UI thread) error: " + ex.Message); }
                        }), json);
                    }
                    else
                    {
                        if (webView.CoreWebView2 != null)
                            webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                    delivered = true;
                }
                catch (Exception ex) { Log("DispatchToWorkbookPane error: " + ex.Message); }
            }

            // 清理失效 CTP
            foreach (var c in invalidCtps)
            {
                _workbookKeyByCtp.Remove(c);
            }

            if (!delivered)
            {
                Log($"DispatchToWorkbookPane: no visible pane for workbook={workbookKey}, using broadcast fallback");
                BroadcastToAllVisiblePanes(json);
            }
        }

        /// <summary>
        /// 广播给所有可见 CTP（旧模式 / 兜底）
        /// </summary>
        private void BroadcastToAllVisiblePanes(string json)
        {
            var invalidKeys = new System.Collections.Generic.List<string>();
            foreach (var kvp in _ctpsByWindow)
            {
                var ctp = kvp.Value;
                if (ctp == null) { invalidKeys.Add(kvp.Key); continue; }
                bool visible;
                try { visible = ctp.Visible; }
                catch { invalidKeys.Add(kvp.Key); continue; }
                if (!visible) continue;

                TaskPaneControl pane = null;
                try { pane = (TaskPaneControl)ctp.ContentControl; }
                catch { invalidKeys.Add(kvp.Key); continue; }
                if (pane == null || pane.IsDisposed || pane.WebView == null) continue;

                var webView = pane.WebView;
                try
                {
                    if (webView.InvokeRequired)
                    {
                        webView.BeginInvoke(new Action<string>(msg =>
                        {
                            try
                            {
                                if (webView.CoreWebView2 != null && !webView.IsDisposed)
                                    webView.CoreWebView2.PostWebMessageAsString(msg);
                            }
                            catch (Exception ex) { Log("Broadcast (UI thread) error: " + ex.Message); }
                        }), json);
                    }
                    else
                    {
                        if (webView.CoreWebView2 != null)
                            webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                }
                catch (Exception ex) { Log("Broadcast to pane error: " + ex.Message); }
            }
            if (invalidKeys.Count > 0)
            {
                foreach (var k in invalidKeys) _ctpsByWindow.Remove(k);
            }
        }

        /// <summary>
        /// 安全获取工作簿唯一 key（和 MessageBridge 中一致的逻辑）
        /// </summary>
        private static string GetWorkbookKey(Workbook wb)
        {
            try
            {
                string fullName = wb.FullName;
                if (!string.IsNullOrEmpty(fullName) && (fullName.Contains("\\") || fullName.Contains("/")))
                    return fullName;
                return wb.Name ?? "workbook_" + wb.GetHashCode();
            }
            catch
            {
                return "workbook_" + wb.GetHashCode();
            }
        }

        /// <summary>
        /// 清理字典中已失效的 CTP（用户点 X 关闭后 CTP 对象访问任何属性都会抛异常）
        /// </summary>
        private void PurgeInvalidCtps()
        {
            var invalidKeys = new System.Collections.Generic.List<string>();
            foreach (var kvp in _ctpsByWindow)
            {
                try
                {
                    var dummy = kvp.Value.Visible;
                }
                catch
                {
                    invalidKeys.Add(kvp.Key);
                }
            }
            foreach (var k in invalidKeys)
            {
                Log("PurgeInvalidCtps: removing invalid CTP for window: " + k);
                _ctpsByWindow.Remove(k);
            }
        }
    }
}
