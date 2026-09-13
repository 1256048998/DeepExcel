using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace DeepExcel.Repair
{
    /// <summary>
    /// Per-user COM registration for the managed add-in, written natively.
    ///
    /// This replaces the PowerShell half of register-user.ps1. Registration used
    /// to run partly from Inno's [Code] and partly from powershell.exe; the
    /// PowerShell hop is both a hard install-time dependency and a common
    /// antivirus false-positive trigger (an unsigned process writing
    /// HKCU\...\Office\Excel\Addins is a textbook heuristic hit).
    /// </summary>
    public static class RegistryRepair
    {
        /// <summary>
        /// Both registry views must be written. 32-bit Excel reads the 32-bit
        /// view, and spelling a literal WOW6432Node path from a 64-bit process
        /// does NOT land there -- that mistake produced REGDB_E_CLASSNOTREG for
        /// several releases.
        /// </summary>
        private static readonly RegistryView[] Views =
        {
            RegistryView.Registry64,
            RegistryView.Registry32
        };

        private sealed class ComClassSpec
        {
            public ComClassSpec(string clsid, string progId, string className)
            {
                Clsid = clsid;
                ProgId = progId;
                ClassName = className;
            }

            public string Clsid { get; }
            public string ProgId { get; }
            public string ClassName { get; }
        }

        private static readonly ComClassSpec[] Classes =
        {
            new ComClassSpec(AddInIdentity.MainClsid, AddInIdentity.MainProgId, AddInIdentity.MainClass),
            new ComClassSpec(AddInIdentity.TaskPaneClsid, AddInIdentity.TaskPaneProgId, AddInIdentity.TaskPaneClass)
        };

        // ------------------------------------------------------------------
        // Verification
        // ------------------------------------------------------------------

        public static void Verify(Report report, string dllPath)
        {
            var expectedCodeBase = AddInIdentity.CodeBase(dllPath);
            var expectedAssembly = AddInIdentity.AssemblyValue(dllPath);
            var version = AddInIdentity.AssemblyVersion(dllPath);

            foreach (var view in Views)
            {
                var bitness = view == RegistryView.Registry64 ? "64" : "32";
                foreach (var spec in Classes)
                {
                    string actualCodeBase, actualAssembly;
                    if (!TryReadInproc(view, spec.Clsid, out actualCodeBase, out actualAssembly))
                    {
                        report.Add(new Finding(
                            spec.Clsid == AddInIdentity.MainClsid
                                ? (view == RegistryView.Registry64 ? Codes.ComClsid64 : Codes.ComClsid32)
                                : Codes.ComTaskPane,
                            Severity.Blocking,
                            string.Format("{0} 位视图缺少 COM 注册（{1}）", bitness, spec.ProgId),
                            true));
                        continue;
                    }

                    if (!string.Equals(actualCodeBase, expectedCodeBase, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(actualAssembly, expectedAssembly, StringComparison.Ordinal))
                    {
                        report.Add(new Finding(
                            Codes.CodeBaseStale,
                            Severity.Blocking,
                            string.Format(
                                "{0} 位视图的 {1} 指向旧版本（{2}），当前应为 {3}",
                                bitness, spec.ProgId, actualAssembly, expectedAssembly),
                            true));
                        continue;
                    }

                    // RegAsm also writes a version-qualified subkey. Some Office
                    // builds read it instead of the unversioned one.
                    string versionedCodeBase, versionedAssembly;
                    if (!TryReadInprocVersion(view, spec.Clsid, version, out versionedCodeBase, out versionedAssembly) ||
                        !string.Equals(versionedCodeBase, expectedCodeBase, StringComparison.OrdinalIgnoreCase))
                    {
                        report.Add(new Finding(
                            Codes.CodeBaseStale,
                            Severity.Blocking,
                            string.Format("{0} 位视图缺少版本化 InprocServer32 子键（{1}）", bitness, version),
                            true));
                    }
                }
            }

            VerifyLoadBehavior(report);
            VerifyHklmResidual(report);
            VerifyResiliency(report);
        }

        private static void VerifyLoadBehavior(Report report)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(AddInIdentity.ExcelAddinsKey))
            {
                var value = key == null ? null : key.GetValue("LoadBehavior");
                if (value == null || Convert.ToInt32(value) != 3)
                {
                    report.Add(new Finding(
                        Codes.LoadBehavior,
                        Severity.Blocking,
                        value == null
                            ? "Excel 加载项注册项缺失"
                            : string.Format("LoadBehavior={0}（应为 3，Excel 已软禁用加载项）", value),
                        true));
                }
            }
        }

        /// <summary>
        /// HKLM wins over HKCU for CLSID resolution. A machine-wide leftover from
        /// an old admin install silently shadows the per-user registration and
        /// points Excel at a DLL path that no longer exists -- the root cause
        /// behind the v0.3.4-v0.4.15 "add-in does not load" reports.
        /// </summary>
        private static void VerifyHklmResidual(Report report)
        {
            foreach (var path in HklmResidualPaths())
            {
                foreach (var view in Views)
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (var key = baseKey.OpenSubKey(path))
                    {
                        if (key != null)
                        {
                            report.Add(new Finding(
                                Codes.HklmResidual,
                                Severity.Blocking,
                                @"存在机器级注册残留 HKLM\" + path + "（优先级高于用户级，会覆盖本次安装）",
                                false));
                            return;
                        }
                    }
                }
            }
        }

        private static IEnumerable<string> HklmResidualPaths()
        {
            yield return @"SOFTWARE\Classes\CLSID\" + AddInIdentity.MainClsid;
            yield return @"SOFTWARE\Classes\CLSID\" + AddInIdentity.TaskPaneClsid;
            yield return @"SOFTWARE\Microsoft\Office\Excel\Addins\" + AddInIdentity.MainProgId;
            yield return @"SOFTWARE\Microsoft\Office\16.0\Excel\Addins\" + AddInIdentity.MainProgId;
        }

        private static void VerifyResiliency(Report report)
        {
            using (var disabled = Registry.CurrentUser.OpenSubKey(AddInIdentity.ResiliencyKey + @"\DisabledItems"))
            {
                if (disabled != null && FindDeepExcelValueNames(disabled).Count > 0)
                {
                    report.Add(new Finding(
                        Codes.DisabledItems,
                        Severity.Blocking,
                        "DeepExcel 被 Excel 加入 DisabledItems 禁用列表",
                        true));
                }
            }

            using (var crashing = Registry.CurrentUser.OpenSubKey(AddInIdentity.ResiliencyKey + @"\CrashingAddinList"))
            {
                if (crashing != null && crashing.GetValue(AddInIdentity.MainProgId) != null)
                {
                    report.Add(new Finding(
                        Codes.CrashingAddinList,
                        Severity.Warning,
                        "DeepExcel 出现在 CrashingAddinList 中",
                        true));
                }
            }
        }

        // ------------------------------------------------------------------
        // Repair
        // ------------------------------------------------------------------

        public static void RegisterAll(string dllPath)
        {
            var assemblyValue = AddInIdentity.AssemblyValue(dllPath);
            var version = AddInIdentity.AssemblyVersion(dllPath);
            var codeBase = AddInIdentity.CodeBase(dllPath);

            foreach (var view in Views)
            {
                foreach (var spec in Classes)
                {
                    RegisterComClass(view, spec, dllPath, codeBase, assemblyValue, version);
                }
            }

            RegisterExcelAddIn(dllPath);
        }

        private static void RegisterComClass(
            RegistryView view, ComClassSpec spec, string dllPath,
            string codeBase, string assemblyValue, string version)
        {
            using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            {
                var clsidPath = @"Software\Classes\CLSID\" + spec.Clsid;
                var inprocPath = clsidPath + @"\InprocServer32";

                // Recreate InprocServer32 so an upgrade cannot leave an older
                // assembly-version subkey that mscoree might still resolve.
                try { baseKey.DeleteSubKeyTree(inprocPath, false); }
                catch (ArgumentException) { /* not present */ }

                using (var key = baseKey.CreateSubKey(clsidPath))
                {
                    key.SetValue("", spec.ClassName, RegistryValueKind.String);
                }

                using (var key = baseKey.CreateSubKey(inprocPath))
                {
                    key.SetValue("", "mscoree.dll", RegistryValueKind.String);
                    key.SetValue("Assembly", assemblyValue, RegistryValueKind.String);
                    key.SetValue("Class", spec.ClassName, RegistryValueKind.String);
                    key.SetValue("CodeBase", codeBase, RegistryValueKind.String);
                    key.SetValue("RuntimeVersion", "v4.0.30319", RegistryValueKind.String);
                    key.SetValue("ThreadingModel", "Both", RegistryValueKind.String);
                }

                using (var key = baseKey.CreateSubKey(inprocPath + "\\" + version))
                {
                    key.SetValue("Assembly", assemblyValue, RegistryValueKind.String);
                    key.SetValue("Class", spec.ClassName, RegistryValueKind.String);
                    key.SetValue("CodeBase", codeBase, RegistryValueKind.String);
                    key.SetValue("RuntimeVersion", "v4.0.30319", RegistryValueKind.String);
                }

                using (baseKey.CreateSubKey(clsidPath + @"\Implemented Categories\" + AddInIdentity.DotNetCategory))
                {
                }

                using (var key = baseKey.CreateSubKey(clsidPath + @"\ProgId"))
                {
                    key.SetValue("", spec.ProgId, RegistryValueKind.String);
                }

                var progIdPath = @"Software\Classes\" + spec.ProgId;
                using (var key = baseKey.CreateSubKey(progIdPath))
                {
                    key.SetValue("", spec.ClassName, RegistryValueKind.String);
                }
                using (var key = baseKey.CreateSubKey(progIdPath + @"\CLSID"))
                {
                    key.SetValue("", spec.Clsid, RegistryValueKind.String);
                }
            }
        }

        private static void RegisterExcelAddIn(string dllPath)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(AddInIdentity.ExcelAddinsKey))
            {
                key.SetValue("Description", "DeepExcel AI AddIn", RegistryValueKind.String);
                key.SetValue("FriendlyName", "DeepExcel AI AddIn", RegistryValueKind.String);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                key.SetValue("Location", dllPath, RegistryValueKind.String);
            }

            // Without DoNotDisableAddinList, Excel flips LoadBehavior to 2 after
            // any startup hiccup and hides the add-in with no user-visible cause.
            using (var key = Registry.CurrentUser.CreateSubKey(AddInIdentity.ResiliencyKey + @"\DoNotDisableAddinList"))
            {
                key.SetValue(AddInIdentity.MainProgId, 1, RegistryValueKind.DWord);
            }

            ClearResiliencyBlocks();
        }

        public static int ClearResiliencyBlocks()
        {
            var cleared = 0;

            using (var crashing = Registry.CurrentUser.OpenSubKey(
                AddInIdentity.ResiliencyKey + @"\CrashingAddinList", true))
            {
                if (crashing != null && crashing.GetValue(AddInIdentity.MainProgId) != null)
                {
                    crashing.DeleteValue(AddInIdentity.MainProgId, false);
                    cleared++;
                }
            }

            using (var disabled = Registry.CurrentUser.OpenSubKey(
                AddInIdentity.ResiliencyKey + @"\DisabledItems", true))
            {
                if (disabled != null)
                {
                    // DisabledItems holds opaque REG_BINARY values, not subkeys.
                    // Deleting the whole key would silently re-enable unrelated
                    // add-ins, so only DeepExcel's own values are removed.
                    foreach (var name in FindDeepExcelValueNames(disabled))
                    {
                        disabled.DeleteValue(name, false);
                        cleared++;
                    }
                }
            }

            return cleared;
        }

        private static List<string> FindDeepExcelValueNames(RegistryKey key)
        {
            var matches = new List<string>();
            foreach (var name in key.GetValueNames())
            {
                var bytes = key.GetValue(name) as byte[];
                if (bytes == null)
                {
                    continue;
                }

                // The blob mixes UTF-16 and ASCII fragments depending on Office
                // build, so both decodings are searched.
                var haystack = Encoding.Unicode.GetString(bytes) + " " + Encoding.ASCII.GetString(bytes);
                if (haystack.IndexOf("DeepExcel", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    matches.Add(name);
                }
            }
            return matches;
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        private static bool TryReadInproc(RegistryView view, string clsid, out string codeBase, out string assembly)
        {
            return TryReadInprocPath(
                view, @"Software\Classes\CLSID\" + clsid + @"\InprocServer32", out codeBase, out assembly);
        }

        private static bool TryReadInprocVersion(
            RegistryView view, string clsid, string version, out string codeBase, out string assembly)
        {
            return TryReadInprocPath(
                view,
                @"Software\Classes\CLSID\" + clsid + @"\InprocServer32\" + version,
                out codeBase, out assembly);
        }

        private static bool TryReadInprocPath(RegistryView view, string path, out string codeBase, out string assembly)
        {
            codeBase = null;
            assembly = null;
            using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            using (var key = baseKey.OpenSubKey(path))
            {
                if (key == null)
                {
                    return false;
                }
                codeBase = key.GetValue("CodeBase") as string;
                assembly = key.GetValue("Assembly") as string;
                return codeBase != null && assembly != null;
            }
        }
    }
}
