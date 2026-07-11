// DeepExcel Installer - a simple self-contained .exe installer
// Compile: csc /target:exe /out:DeepExcelInstaller.exe DeepExcelInstaller.cs
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace DeepExcelInstaller
{
    class Program
    {
        static string _logFile = "";

        static void Log(string msg)
        {
            Console.WriteLine(msg);
            try
            {
                if (_logFile != "")
                    File.AppendAllText(_logFile, DateTime.Now.ToString("HH:mm:ss") + " " + msg + Environment.NewLine);
            }
            catch { }
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpVerb;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpFile;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpParameters;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIcon;
            public IntPtr hProcess;
        }

        static void Main(string[] args)
        {
            Console.Title = "DeepExcel Installer";
            Console.WriteLine("==========================================");
            Console.WriteLine("  DeepExcel Installer v0.4.1");
            Console.WriteLine("==========================================");
            Console.WriteLine();

            // Check if running as admin
            if (!IsRunningAsAdmin())
            {
                Console.WriteLine("Administrator privileges required to install DeepExcel.");
                Console.WriteLine("Requesting elevation...");
                Console.WriteLine();
                RestartAsAdmin();
                return;
            }

            string scriptDir = AppDomain.CurrentDomain.BaseDirectory;
            // Install to LOCALAPPDATA (clean path, no admin needed for file copy,
            // but we still need admin for HKLM registry writes).
            // Using LOCALAPPDATA avoids path-with-spaces/Chinese-chars issues.
            string installDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DeepExcel");

            // Set up log file
            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DeepExcel", "logs");
            try { Directory.CreateDirectory(logDir); } catch { }
            _logFile = Path.Combine(logDir, "DeepExcel_Installer.log");
            Log("Log file: " + _logFile);
            Log("Script dir: " + scriptDir);
            Log("Install dir: " + installDir);
            Log("");

            // Step 1: Close Excel
            Log("[1/6] Closing Excel...");
            foreach (var proc in Process.GetProcessesByName("EXCEL"))
            {
                try { proc.Kill(); proc.WaitForExit(3000); Log("  Killed EXCEL PID " + proc.Id); } catch { }
            }
            Log("  Done.");
            Log("");

            // Step 2: Copy files to install directory (clean path, no Chinese chars)
            Log("[2/6] Copying files to " + installDir + " ...");
            try
            {
                if (Directory.Exists(installDir))
                {
                    foreach (var f in Directory.GetFiles(installDir, "*", SearchOption.AllDirectories))
                    {
                        try { File.SetAttributes(f, FileAttributes.Normal); } catch { }
                    }
                    Directory.Delete(installDir, true);
                }
                Directory.CreateDirectory(installDir);

                string[] requiredFiles = {
                    "DeepExcel.AddIn.dll", "DeepExcel.AddIn.dll.config",
                    "Microsoft.Bcl.AsyncInterfaces.dll",
                    "Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.WinForms.dll",
                    "System.Buffers.dll", "System.Memory.dll", "System.Numerics.Vectors.dll",
                    "System.Runtime.CompilerServices.Unsafe.dll",
                    "System.Text.Encodings.Web.dll", "System.Text.Json.dll",
                    "System.Threading.Tasks.Extensions.dll", "System.ValueTuple.dll",
                    "WebView2Loader.dll"
                };

                foreach (var f in requiredFiles)
                {
                    string src = Path.Combine(scriptDir, f);
                    if (File.Exists(src))
                    {
                        File.Copy(src, Path.Combine(installDir, f), true);
                        Log("  " + f);
                    }
                    else
                    {
                        Log("  WARNING: " + f + " not found");
                    }
                }

                // Copy directories
                CopyDirectory(Path.Combine(scriptDir, "WebViewAssets"), Path.Combine(installDir, "WebViewAssets"));
                CopyDirectory(Path.Combine(scriptDir, "sidecar"), Path.Combine(installDir, "sidecar"));
                Log("  Directories copied (WebViewAssets, sidecar).");

                // Unblock all files (remove MOTW)
                int unblocked = 0;
                foreach (var f in Directory.GetFiles(installDir, "*", SearchOption.AllDirectories))
                {
                    try { File.Delete(f + ":Zone.Identifier"); unblocked++; } catch { }
                }
                Log("  Unblocked " + unblocked + " file(s).");
            }
            catch (Exception ex)
            {
                Log("  ERROR: " + ex.Message);
                Log("");
                Log("Press Enter to exit...");
                Console.ReadLine();
                return;
            }
            Log("");

            string dllPath = Path.Combine(installDir, "DeepExcel.AddIn.dll");
            string clsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}";
            string progId = "DeepExcel.AddIn";
            string className = "DeepExcel.AddIn.ThisAddIn";
            string friendlyName = "DeepExcel AI AddIn";
            string taskPaneClsid = "{B2C3D4E5-F6A7-404B-9A5F-9B3C1D2E3F4B}";
            string taskPaneProgId = "DeepExcel.AddIn.TaskPaneControl";
            string taskPaneClass = "DeepExcel.AddIn.TaskPaneControl";

            // Step 3: Register COM via regasm (with fallback to manual)
            Log("[3/6] Registering COM component...");
            bool regasmOk = false;
            string regasm = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                @"Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe");
            if (!File.Exists(regasm))
            {
                regasm = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"Microsoft.NET\Framework\v4.0.30319\RegAsm.exe");
            }

            if (File.Exists(regasm))
            {
                Log("  Using regasm: " + regasm);
                var psi = new ProcessStartInfo
                {
                    FileName = regasm,
                    Arguments = "/codebase \"" + dllPath + "\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                try
                {
                    var p = Process.Start(psi);
                    string stdout = p.StandardOutput.ReadToEnd();
                    string stderr = p.StandardError.ReadToEnd();
                    p.WaitForExit();
                    if (p.ExitCode == 0)
                    {
                        Log("  COM registered successfully via regasm.");
                        regasmOk = true;
                    }
                    else
                    {
                        Log("  regasm failed (exit " + p.ExitCode + "): " + stderr.Trim());
                        Log("  Falling back to manual registration...");
                    }
                }
                catch (Exception ex)
                {
                    Log("  regasm exception: " + ex.Message);
                    Log("  Falling back to manual registration...");
                }
            }
            else
            {
                Log("  RegAsm.exe not found, using manual registration...");
            }

            // ALWAYS run manual registration for ThisAddIn too, even if regasm succeeded.
            // Reason: regasm registers CLSID only under HKCU\Software\Classes, but 64-bit
            // Excel reads HKLM\Software\Classes\CLSID. Without HKLM registration, Excel
            // cannot find the COM class and silently skips loading the add-in.
            ManualRegisterCom(clsid, progId, className, dllPath);
            Log("  ThisAddIn COM registration ensured (HKLM + HKCU).");

            // Register TaskPaneControl
            ManualRegisterCom(taskPaneClsid, taskPaneProgId, taskPaneClass, dllPath);
            Log("  TaskPaneControl registered.");
            Log("");

            // Step 4: Register Excel Addin (HKLM + HKCU)
            Log("[4/6] Registering Excel Add-in...");
            string assemblyVersion = "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null";

            // HKLM Excel Addin
            try
            {
                string hklmAddinKey = @"Software\Microsoft\Office\16.0\Excel\Addins\" + progId;
                using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(hklmAddinKey))
                {
                    key.SetValue("Description", friendlyName);
                    key.SetValue("FriendlyName", friendlyName);
                    key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                    key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                    key.SetValue("Location", dllPath);
                }
                Log("  HKLM Excel Addin key created.");
            }
            catch (Exception ex)
            {
                Log("  HKLM write failed: " + ex.Message);
            }

            // HKCU Excel Addin
            string hkcuAddinKey = @"Software\Microsoft\Office\16.0\Excel\Addins\" + progId;
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(hkcuAddinKey))
            {
                key.SetValue("Description", friendlyName);
                key.SetValue("FriendlyName", friendlyName);
                key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                key.SetValue("Location", dllPath);
            }
            Log("  HKCU Excel Addin key created.");

            // Add DoNotDisableAddinList entry (HKCU)
            string doNotDisableKey = @"Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList";
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(doNotDisableKey))
            {
                key.SetValue(progId, 1, Microsoft.Win32.RegistryValueKind.DWord);
            }
            Log("  DoNotDisableAddinList entry added.");

            // Clear DisabledItems (HKCU)
            string disabledKey = @"Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems";
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(disabledKey, true))
                {
                    if (key != null)
                    {
                        int cleared = 0;
                        foreach (var name in key.GetSubKeyNames())
                        {
                            try { key.DeleteSubKeyTree(name, false); cleared++; } catch { }
                        }
                        Log("  Cleared " + cleared + " DisabledItems entr(ies).");
                    }
                }
            }
            catch { }

            // Also check Group Policy that might block add-ins
            try
            {
                string policyKey = @"Software\Policies\Microsoft\Office\16.0\Excel\Options";
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(policyKey))
                {
                    if (key != null)
                    {
                        var disableAll = key.GetValue("DisableAllAddins");
                        if (disableAll != null && disableAll.ToString() == "1")
                        {
                            Log("  WARNING: Group Policy 'DisableAllAddins' is ON!");
                        }
                    }
                }
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(policyKey))
                {
                    if (key != null)
                    {
                        var disableAll = key.GetValue("DisableAllAddins");
                        if (disableAll != null && disableAll.ToString() == "1")
                        {
                            Log("  WARNING: HKLM Group Policy 'DisableAllAddins' is ON!");
                        }
                    }
                }
            }
            catch { }
            Log("");

            // Step 5: Verification
            Log("[5/6] Verification...");
            bool verifyOk = true;

            // Check Excel Addin key
            using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(hkcuAddinKey))
            {
                if (key != null)
                {
                    var lb = key.GetValue("LoadBehavior");
                    Log("  [OK] HKCU Excel Addin key (LoadBehavior=" + lb + ")");
                }
                else
                {
                    Log("  [FAIL] HKCU Excel Addin key missing!");
                    verifyOk = false;
                }
            }
            using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(hkcuAddinKey))
            {
                if (key != null)
                {
                    var lb = key.GetValue("LoadBehavior");
                    Log("  [OK] HKLM Excel Addin key (LoadBehavior=" + lb + ")");
                }
                else
                {
                    Log("  [WARN] HKLM Excel Addin key missing (HKCU should still work)");
                }
            }

            // Check CLSID
            string clsidKeyHkcu = @"Software\Classes\CLSID\" + clsid;
            string clsidKeyHklm = @"Software\Classes\CLSID\" + clsid;
            bool clsidHkcuOk = false, clsidHklmOk = false;
            using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(clsidKeyHkcu + @"\InprocServer32"))
            {
                if (key != null)
                {
                    var codebase = key.GetValue("CodeBase");
                    Log("  [OK] HKCU CLSID (CodeBase=" + codebase + ")");
                    clsidHkcuOk = true;
                }
                else
                {
                    Log("  [WARN] HKCU CLSID InprocServer32 missing");
                }
            }
            using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(clsidKeyHklm + @"\InprocServer32", true))
            {
                if (key != null)
                {
                    var codebase = key.GetValue("CodeBase");
                    if (codebase == null || codebase.ToString() == "")
                    {
                        // CodeBase missing in HKLM — this is the root cause!
                        // 64-bit Excel reads HKLM\Software\Classes\CLSID, and without
                        // CodeBase it cannot find the DLL. Re-write it now.
                        Log("  [WARN] HKLM CLSID CodeBase missing, re-writing...");
                        key.SetValue("CodeBase", dllPath);
                        key.SetValue("Assembly", "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null");
                        key.SetValue("Class", className);
                        key.SetValue("RuntimeVersion", "v4.0.30319");
                        key.SetValue("ThreadingModel", "Both");
                        codebase = key.GetValue("CodeBase");
                        Log("  [OK] HKLM CLSID (CodeBase=" + codebase + ") [re-written]");
                    }
                    else
                    {
                        Log("  [OK] HKLM CLSID (CodeBase=" + codebase + ")");
                    }
                    clsidHklmOk = true;
                }
                else
                {
                    Log("  [INFO] HKLM CLSID InprocServer32 missing, creating...");
                    using (var newKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(clsidKeyHklm + @"\InprocServer32"))
                    {
                        newKey.SetValue(null, "mscoree.dll");
                        newKey.SetValue("Assembly", "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null");
                        newKey.SetValue("Class", className);
                        newKey.SetValue("CodeBase", dllPath);
                        newKey.SetValue("RuntimeVersion", "v4.0.30319");
                        newKey.SetValue("ThreadingModel", "Both");
                    }
                    Log("  [OK] HKLM CLSID InprocServer32 created with CodeBase");
                    clsidHklmOk = true;
                }
            }
            if (!clsidHkcuOk && !clsidHklmOk)
            {
                Log("  [FAIL] No CLSID registration found!");
                verifyOk = false;
            }

            // Check DLL exists
            if (File.Exists(dllPath))
            {
                Log("  [OK] DLL exists: " + dllPath);
            }
            else
            {
                Log("  [FAIL] DLL missing!");
                verifyOk = false;
            }

            // COM instantiation test
            try
            {
                var type = Type.GetTypeFromProgID(progId);
                if (type != null)
                {
                    var obj = Activator.CreateInstance(type);
                    if (obj != null)
                    {
                        Log("  [OK] COM instantiation test passed");
                        Marshal.ReleaseComObject(obj);
                    }
                }
                else
                {
                    Log("  [WARN] ProgID not found (may need Excel restart)");
                }
            }
            catch (Exception ex)
            {
                Log("  [WARN] COM test: " + ex.Message);
            }

            Log("");

            // Step 6: Done
            Log("[6/6] Installation " + (verifyOk ? "successful!" : "completed with warnings."));
            Log("");
            Log("==========================================");
            Log("  DLL Path: " + dllPath);
            Log("  Log File: " + _logFile);
            Log("==========================================");
            Log("");

            // Offer to launch Excel
            Console.Write("Launch Excel now? (Y/N, default Y): ");
            string answer = Console.ReadLine();
            Log("User chose: " + answer);
            if (string.IsNullOrEmpty(answer) || answer.ToUpper() == "Y")
            {
                string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    @"Microsoft Office\root\Office16\EXCEL.EXE");
                if (!File.Exists(excelPath))
                {
                    excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                        @"Microsoft Office\root\Office16\EXCEL.EXE");
                }
                if (File.Exists(excelPath))
                {
                    Log("Launching Excel: " + excelPath);
                    Process.Start(excelPath);
                }
                else
                {
                    Log("Excel not found in default paths.");
                }
            }
            Log("");
            Log("Look for the 'DeepExcel' tab in the Excel ribbon.");
            Log("If not visible: File > Options > Add-ins > Manage: COM Add-ins > Go");
            Log("");
            Log("Press Enter to exit...");
            Console.ReadLine();
        }

        static void ManualRegisterCom(string clsid, string progId, string className, string dllPath)
        {
            string assemblyVersion = "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null";

            // .NET Category GUID: {62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}
            string netCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";

            // Write to both HKLM and HKCU for maximum compatibility
            var roots = new[] {
                new { Root = Microsoft.Win32.Registry.LocalMachine, Prefix = "HKLM" },
                new { Root = Microsoft.Win32.Registry.CurrentUser, Prefix = "HKCU" }
            };

            foreach (var r in roots)
            {
                try
                {
                    string clsidKey = @"Software\Classes\CLSID\" + clsid;

                    // CLSID root
                    using (var key = r.Root.CreateSubKey(clsidKey))
                    {
                        key.SetValue(null, className);
                    }

                    // InprocServer32
                    using (var key = r.Root.CreateSubKey(clsidKey + @"\InprocServer32"))
                    {
                        key.SetValue(null, "mscoree.dll");
                        key.SetValue("Assembly", assemblyVersion);
                        key.SetValue("Class", className);
                        key.SetValue("CodeBase", dllPath);
                        key.SetValue("RuntimeVersion", "v4.0.30319");
                        key.SetValue("ThreadingModel", "Both");
                    }

                    // Version subkey (regasm writes this)
                    using (var key = r.Root.CreateSubKey(clsidKey + @"\InprocServer32\0.2.4.0"))
                    {
                        key.SetValue("Assembly", assemblyVersion);
                        key.SetValue("Class", className);
                        key.SetValue("CodeBase", dllPath);
                        key.SetValue("RuntimeVersion", "v4.0.30319");
                    }

                    // Implemented Categories (.NET Category - important for .NET COM)
                    using (var key = r.Root.CreateSubKey(clsidKey + @"\Implemented Categories\" + netCategory))
                    {
                        // Empty key, just needs to exist
                    }

                    // ProgId subkey under CLSID
                    using (var key = r.Root.CreateSubKey(clsidKey + @"\ProgId"))
                    {
                        key.SetValue(null, progId);
                    }

                    // ProgID mapping
                    string progIdKey = @"Software\Classes\" + progId;
                    using (var key = r.Root.CreateSubKey(progIdKey))
                    {
                        key.SetValue(null, className);
                    }
                    using (var key = r.Root.CreateSubKey(progIdKey + @"\CLSID"))
                    {
                        key.SetValue(null, clsid);
                    }

                    Log("    " + r.Prefix + " registration complete.");
                }
                catch (Exception ex)
                {
                    Log("    " + r.Prefix + " registration failed: " + ex.Message);
                }
            }
        }

        static void CopyDirectory(string source, string target)
        {
            if (!Directory.Exists(source)) return;
            if (!Directory.Exists(target))
                Directory.CreateDirectory(target);

            foreach (var file in Directory.GetFiles(source))
            {
                File.Copy(file, Path.Combine(target, Path.GetFileName(file)), true);
            }

            foreach (var dir in Directory.GetDirectories(source))
            {
                CopyDirectory(dir, Path.Combine(target, Path.GetFileName(dir)));
            }
        }

        static bool IsRunningAsAdmin()
        {
            try
            {
                var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }

        static void RestartAsAdmin()
        {
            var exePath = Process.GetCurrentProcess().MainModule.FileName;
            var info = new SHELLEXECUTEINFO
            {
                cbSize = Marshal.SizeOf(typeof(SHELLEXECUTEINFO)),
                lpVerb = "runas",
                lpFile = exePath,
                nShow = 1 // SW_SHOWNORMAL
            };
            ShellExecuteEx(ref info);
        }
    }
}
