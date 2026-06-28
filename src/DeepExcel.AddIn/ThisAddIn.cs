using System;
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
        private Microsoft.Office.Core.CustomTaskPane _customTaskPane;
        private ICTPFactory _ctpFactory;
        private MessageBridge _bridge;
        private IRibbonUI _ribbon;
        // Fallback: 当 CustomTaskPane 不可用时使用独立 Form
        private Form _fallbackForm;
        // 隐藏的 UI 控件，用于 sidecar 的 UI 线程封送
        private Control _sidecarUiControl;

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
                string logPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "DeepExcel_Load.log");
                File.AppendAllText(logPath, "[" + DateTime.Now + "] " + message + Environment.NewLine);
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

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            Log("OnDisconnection called, mode=" + disconnectMode);
            try
            {
                if (_bridge is IDisposable disposable) disposable.Dispose();
                _bridge = null;

                if (_sidecarUiControl != null && !_sidecarUiControl.IsDisposed)
                {
                    _sidecarUiControl.Dispose();
                    _sidecarUiControl = null;
                }

                if (_customTaskPane != null)
                {
                    _customTaskPane.Delete();
                    _customTaskPane = null;
                }
                if (_fallbackForm != null && !_fallbackForm.IsDisposed)
                {
                    _fallbackForm.Close();
                    _fallbackForm.Dispose();
                }
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
                // 优先尝试 CustomTaskPane（可停靠在 Excel 右侧）
                if (TryAcquireCTPFactory())
                {
                    if (_customTaskPane == null)
                    {
                        Log("Creating CustomTaskPane via ICTPFactory");
                        _customTaskPane = _ctpFactory.CreateCTP(
                            "DeepExcel.AddIn.TaskPaneControl",
                            "DeepExcel AI",
                            Type.Missing);
                        Log("CustomTaskPane created");

                        _customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                        _customTaskPane.Width = 420;
                        Log("CustomTaskPane docked right, width=420");

                        _taskPane = (TaskPaneControl)_customTaskPane.ContentControl;
                        Log("TaskPaneControl retrieved from ContentControl");

                        InitializeWebView();
                    }
                    _customTaskPane.Visible = !_customTaskPane.Visible;
                    Log("CustomTaskPane visible=" + _customTaskPane.Visible);
                    return;
                }

                // Fallback: CustomTaskPane 不可用，使用独立浮动窗口
                Log("Falling back to floating Form");
                if (_fallbackForm == null || _fallbackForm.IsDisposed)
                {
                    _fallbackForm = new Form
                    {
                        Text = "DeepExcel AI",
                        Width = 440,
                        Height = 700,
                        StartPosition = FormStartPosition.Manual,
                        FormBorderStyle = FormBorderStyle.SizableToolWindow,
                        BackColor = System.Drawing.Color.White
                    };
                    _taskPane = new TaskPaneControl();
                    _taskPane.Dock = DockStyle.Fill;
                    _fallbackForm.Controls.Add(_taskPane);
                    _fallbackForm.FormClosed += (s, e) => Log("Fallback form closed");
                    Log("Fallback form created");
                    InitializeWebView();
                }
                _fallbackForm.Visible = !_fallbackForm.Visible;
                if (_fallbackForm.Visible) _fallbackForm.Activate();
                Log("Fallback form visible=" + _fallbackForm.Visible);
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
                Log("InitializeWebView started");

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
                var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                    null, userDataFolder, options);
                Log("WebView2 environment created");

                await _taskPane.WebView.EnsureCoreWebView2Async(env);
                Log("WebView2 core initialized");

                string assetsPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "WebViewAssets", "index.html");
                Log("Loading WebView from: " + assetsPath);

                _taskPane.WebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "deepexcel.local",
                    Path.GetDirectoryName(assetsPath),
                    Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);

                _taskPane.WebView.CoreWebView2.Navigate("https://deepexcel.local/index.html");
                Log("WebView navigated");

                // 设置消息桥接 - 必须切回 UI 线程访问 WebView2
                _bridge.SetSendToUi(json =>
                {
                    try
                    {
                        if (_taskPane == null || _taskPane.IsDisposed || _taskPane.WebView == null)
                            return;

                        if (_taskPane.WebView.InvokeRequired)
                        {
                            _taskPane.WebView.BeginInvoke(new Action<string>(msg =>
                            {
                                try
                                {
                                    if (_taskPane.WebView.CoreWebView2 != null)
                                        _taskPane.WebView.CoreWebView2.PostWebMessageAsString(msg);
                                }
                                catch (Exception ex) { Log("SendToUi (UI thread) error: " + ex.Message); }
                            }), json);
                        }
                        else
                        {
                            if (_taskPane.WebView.CoreWebView2 != null)
                                _taskPane.WebView.CoreWebView2.PostWebMessageAsString(json);
                        }
                    }
                    catch (Exception ex) { Log("SendToUi error: " + ex.Message); }
                });

                _taskPane.WebView.CoreWebView2.WebMessageReceived += (sender, e) =>
                {
                    try
                    {
                        // 前端 postMessage 发送的是对象，用 WebMessageAsJson 获取 JSON 字符串
                        string json = e.WebMessageAsJson;
                        Log("WebMessageReceived: " + (json == null ? "null" : (json.Length > 300 ? json.Substring(0, 300) + "..." : json)));
                        string response = _bridge.HandleMessage(json);
                        Log("HandleMessage response: " + (response == null ? "null" : (response.Length > 300 ? response.Substring(0, 300) + "..." : response)));
                        if (!string.IsNullOrEmpty(response))
                        {
                            _taskPane.WebView.CoreWebView2.PostWebMessageAsString(response);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("WebMessageReceived FAILED: " + ex.GetType().Name + " - " + ex.Message);
                        Log("HRESULT: 0x" + Marshal.GetHRForException(ex).ToString("X8"));
                        var inner = ex.InnerException;
                        while (inner != null)
                        {
                            Log("  InnerException: " + inner.GetType().Name + " - " + inner.Message);
                            if (inner.InnerException == null)
                            {
                                Log("  Inner Stack: " + inner.StackTrace);
                            }
                            inner = inner.InnerException;
                        }
                        Log("Stack: " + ex.StackTrace);
                    }
                };

                Log("WebView and bridge wiring complete");
            }
            catch (Exception ex)
            {
                Log("InitializeWebView FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
            }
        }
    }
}
