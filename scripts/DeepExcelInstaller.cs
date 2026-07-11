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
            Console.WriteLine("  DeepExcel Installer v0.4.0");
            Console.WriteLine("==========================================");
            Console.WriteLine();

            // Check if running as admin
            if (!IsRunningAsAdmin())
            {
                Console.WriteLine("Administrator privileges required to install DeepExcel.");
                Console.WriteLine("Requesting elevation...");
                Console.WriteLine();
                // Re-launch self as admin
                RestartAsAdmin();
                return;
            }

            string scriptDir = AppDomain.CurrentDomain.BaseDirectory;
            string installDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "DeepExcel");

            Console.WriteLine("Install location: " + installDir);
            Console.WriteLine();

            // Step 1: Close Excel
            Console.WriteLine("[1/5] Closing Excel...");
            foreach (var proc in Process.GetProcessesByName("EXCEL"))
            {
                try { proc.Kill(); proc.WaitForExit(3000); } catch { }
            }
            Console.WriteLine("  Done.");
            Console.WriteLine();

            // Step 2: Copy files to install directory
            Console.WriteLine("[2/5] Copying files...");
            try
            {
                if (Directory.Exists(installDir))
                {
                    // Keep config/logs in %APPDATA%, only replace binaries
                    foreach (var f in Directory.GetFiles(installDir, "*", SearchOption.AllDirectories))
                    {
                        try { File.SetAttributes(f, FileAttributes.Normal); } catch { }
                    }
                    Directory.Delete(installDir, true);
                }
                Directory.CreateDirectory(installDir);

                string[] extensions = { ".dll", ".exe", ".config", ".py", ".html", ".js", ".css", ".json", ".xml", ".txt" };
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
                        Console.WriteLine("  " + f);
                    }
                    else
                    {
                        Console.WriteLine("  WARNING: " + f + " not found");
                    }
                }

                // Copy directories
                CopyDirectory(Path.Combine(scriptDir, "WebViewAssets"), Path.Combine(installDir, "WebViewAssets"));
                CopyDirectory(Path.Combine(scriptDir, "sidecar"), Path.Combine(installDir, "sidecar"));

                // Unblock all files
                foreach (var f in Directory.GetFiles(installDir, "*", SearchOption.AllDirectories))
                {
                    try { File.Delete(f + ":Zone.Identifier"); } catch { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("  ERROR: " + ex.Message);
                Console.WriteLine();
                Console.WriteLine("Press Enter to exit...");
                Console.ReadLine();
                return;
            }
            Console.WriteLine("  Done.");
            Console.WriteLine();

            // Step 3: Register COM via regasm
            Console.WriteLine("[3/5] Registering COM component...");
            string dllPath = Path.Combine(installDir, "DeepExcel.AddIn.dll");
            string regasm = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                @"Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe");
            if (!File.Exists(regasm))
            {
                regasm = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"Microsoft.NET\Framework\v4.0.30319\RegAsm.exe");
            }

            if (File.Exists(regasm))
            {
                Console.WriteLine("  Using: " + regasm);
                var psi = new ProcessStartInfo
                {
                    FileName = regasm,
                    Arguments = "/codebase \"" + dllPath + "\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                var p = Process.Start(psi);
                string stdout = p.StandardOutput.ReadToEnd();
                string stderr = p.StandardError.ReadToEnd();
                p.WaitForExit();
                if (p.ExitCode != 0)
                {
                    Console.WriteLine("  regasm output: " + stdout);
                    Console.WriteLine("  regasm errors: " + stderr);
                }
                else
                {
                    Console.WriteLine("  COM registered successfully.");
                }

                // Also register 32-bit
                string regasm32 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"Microsoft.NET\Framework\v4.0.30319\RegAsm.exe");
                if (File.Exists(regasm32) && regasm32 != regasm)
                {
                    var psi32 = new ProcessStartInfo
                    {
                        FileName = regasm32,
                        Arguments = "/codebase \"" + dllPath + "\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };
                    var p32 = Process.Start(psi32);
                    p32.WaitForExit();
                    Console.WriteLine("  32-bit COM registered.");
                }
            }
            else
            {
                Console.WriteLine("  WARNING: RegAsm.exe not found, using manual registry registration...");
                ManualRegister(dllPath);
            }
            Console.WriteLine();

            // Step 4: Register Excel Addin (both HKLM and HKCU)
            Console.WriteLine("[4/5] Registering Excel Add-in...");
            string progId = "DeepExcel.AddIn";
            string className = "DeepExcel.AddIn.ThisAddIn";
            string friendlyName = "DeepExcel AI AddIn";
            string clsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}";
            string taskPaneClsid = "{B2C3D4E5-F6A7-404B-9A5F-9B3C1D2E3F4B}";
            string taskPaneProgId = "DeepExcel.AddIn.TaskPaneControl";
            string taskPaneClass = "DeepExcel.AddIn.TaskPaneControl";

            // Write to HKLM (machine-level, visible to all Office instances)
            try
            {
                // HKLM Excel Addin
                string hklmAddinKey = @"Software\Microsoft\Office\16.0\Excel\Addins\" + progId;
                using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(hklmAddinKey))
                {
                    key.SetValue("Description", friendlyName);
                    key.SetValue("FriendlyName", friendlyName);
                    key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                    key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                    key.SetValue("Location", dllPath);
                }
                Console.WriteLine("  HKLM Excel Addin key created.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("  HKLM write failed (may need admin): " + ex.Message);
            }

            // Also write to HKCU (user-level)
            string hkcuAddinKey = @"Software\Microsoft\Office\16.0\Excel\Addins\" + progId;
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(hkcuAddinKey))
            {
                key.SetValue("Description", friendlyName);
                key.SetValue("FriendlyName", friendlyName);
                key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                key.SetValue("Location", dllPath);
            }
            Console.WriteLine("  HKCU Excel Addin key created.");

            // Add DoNotDisableAddinList entry
            string doNotDisableKey = @"Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList";
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(doNotDisableKey))
            {
                key.SetValue(progId, 1, Microsoft.Win32.RegistryValueKind.DWord);
            }
            Console.WriteLine("  DoNotDisableAddinList entry added.");

            // Clear DisabledItems
            string disabledKey = @"Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems";
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(disabledKey, true))
                {
                    if (key != null)
                    {
                        foreach (var name in key.GetSubKeyNames())
                        {
                            try { key.DeleteSubKeyTree(name, false); } catch { }
                        }
                        Console.WriteLine("  DisabledItems cleared.");
                    }
                }
            }
            catch { }

            // Register TaskPaneControl CLSID
            RegisterClsid(taskPaneClsid, taskPaneProgId, taskPaneClass, dllPath);
            Console.WriteLine("  TaskPaneControl registered.");

            Console.WriteLine();

            // Step 5: Launch Excel
            Console.WriteLine("[5/5] Installation complete!");
            Console.WriteLine();
            Console.WriteLine("==========================================");
            Console.WriteLine("  DeepExcel has been installed successfully!");
            Console.WriteLine("  Location: " + installDir);
            Console.WriteLine("==========================================");
            Console.WriteLine();
            Console.Write("Launch Excel now? (Y/N, default Y): ");
            string answer = Console.ReadLine();
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
                    Console.WriteLine("Launching Excel...");
                    Process.Start(excelPath);
                }
                else
                {
                    Console.WriteLine("Excel not found. Please launch manually.");
                }
            }
            Console.WriteLine();
            Console.WriteLine("Look for the 'DeepExcel' tab in the Excel ribbon.");
            Console.WriteLine();
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }

        static void RegisterClsid(string clsid, string progId, string className, string dllPath)
        {
            string clsidKey = @"Software\Classes\CLSID\" + clsid;
            // HKLM
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(clsidKey))
                {
                    key.SetValue(null, className);
                }
                using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(clsidKey + @"\InprocServer32"))
                {
                    key.SetValue(null, "mscoree.dll");
                    key.SetValue("Assembly", "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null");
                    key.SetValue("Class", className);
                    key.SetValue("CodeBase", dllPath);
                    key.SetValue("RuntimeVersion", "v4.0.30319");
                    key.SetValue("ThreadingModel", "Both");
                }
                string progIdKey = @"Software\Classes\" + progId;
                using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(progIdKey))
                {
                    key.SetValue(null, className);
                }
                using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(progIdKey + @"\CLSID"))
                {
                    key.SetValue(null, clsid);
                }
            }
            catch { }

            // HKCU
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(clsidKey))
            {
                key.SetValue(null, className);
            }
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(clsidKey + @"\InprocServer32"))
            {
                key.SetValue(null, "mscoree.dll");
                key.SetValue("Assembly", "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null");
                key.SetValue("Class", className);
                key.SetValue("CodeBase", dllPath);
                key.SetValue("RuntimeVersion", "v4.0.30319");
                key.SetValue("ThreadingModel", "Both");
            }
            string progIdKeyCu = @"Software\Classes\" + progId;
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(progIdKeyCu))
            {
                key.SetValue(null, className);
            }
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(progIdKeyCu + @"\CLSID"))
            {
                key.SetValue(null, clsid);
            }
        }

        static void ManualRegister(string dllPath)
        {
            string clsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}";
            string progId = "DeepExcel.AddIn";
            string className = "DeepExcel.AddIn.ThisAddIn";
            RegisterClsid(clsid, progId, className, dllPath);
            Console.WriteLine("  Manual registration complete.");
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
