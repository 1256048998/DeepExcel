using System;
using System.Collections.Generic;
using System.Security.Principal;
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

        /// <summary>
        /// Whether programs this user starts run elevated.
        ///
        /// Repair.exe is started like any other program -- from the Start menu,
        /// or by an installer that itself runs at the user's level -- so its own
        /// token is the answer: elevated here means UAC is off or this is the
        /// built-in Administrator, and then Excel and WPS are elevated too.
        /// Windows ignores HKCU\Software\Classes for elevated COM clients, which
        /// is why those machines need the HKLM copy (2026-11-24, user machine:
        /// HKCU registration correct, activation 0x80070002; same registration
        /// copied to HKLM, activation succeeded).
        /// </summary>
        public static bool HostsRunElevated
        {
            get
            {
                try
                {
                    using (var identity = WindowsIdentity.GetCurrent())
                    {
                        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
                    }
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

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

            var elevated = HostsRunElevated;
            if (elevated)
            {
                VerifyMachineRegistration(report, expectedCodeBase, expectedAssembly, version);
            }

            VerifyLoadBehavior(report);
            VerifyHklmResidual(report, elevated);
            VerifyResiliency(report);
        }

        /// <summary>
        /// Elevated hosts see only HKLM, so there it has to say exactly what the
        /// HKCU registration says.
        /// </summary>
        private static void VerifyMachineRegistration(
            Report report, string expectedCodeBase, string expectedAssembly, string version)
        {
            foreach (var view in Views)
            {
                foreach (var spec in Classes)
                {
                    var inproc = @"SOFTWARE\Classes\CLSID\" + spec.Clsid + @"\InprocServer32";
                    string codeBase, assembly, versionedCodeBase, versionedAssembly;
                    // Both keys, as for HKCU: some Office builds read the
                    // version-qualified one instead.
                    if (!TryReadInprocPath(RegistryHive.LocalMachine, view, inproc, out codeBase, out assembly) ||
                        !TryReadInprocPath(RegistryHive.LocalMachine, view, inproc + "\\" + version,
                                           out versionedCodeBase, out versionedAssembly) ||
                        !string.Equals(codeBase, expectedCodeBase, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(versionedCodeBase, expectedCodeBase, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(assembly, expectedAssembly, StringComparison.Ordinal))
                    {
                        report.Add(new Finding(
                            Codes.MachineRegistration,
                            Severity.Blocking,
                            "Excel / WPS 以管理员权限运行（已关闭 UAC 或使用内置 Administrator 账户），" +
                            "Windows 不认只登记给当前用户的加载项；需要补上机器级登记",
                            true));
                        return;
                    }
                }
            }
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
        /// A machine-wide leftover from
        /// an old admin install silently shadows the per-user registration and
        /// points Excel at a DLL path that no longer exists -- the root cause
        /// behind the v0.3.4-v0.4.15 "add-in does not load" reports.
        /// </summary>
        private static void VerifyHklmResidual(Report report, bool elevated)
        {
            foreach (var path in HklmResidualPaths(includeClasses: !elevated))
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

        /// <summary>
        /// With elevated hosts the class keys under HKLM are not residue but the
        /// registration they actually use; VerifyMachineRegistration checks
        /// them for staleness instead.
        /// </summary>
        private static IEnumerable<string> HklmResidualPaths(bool includeClasses)
        {
            if (includeClasses)
            {
                yield return @"SOFTWARE\Classes\CLSID\" + AddInIdentity.MainClsid;
                yield return @"SOFTWARE\Classes\CLSID\" + AddInIdentity.TaskPaneClsid;
            }
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

            var hives = HostsRunElevated
                ? new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine }
                : new[] { RegistryHive.CurrentUser };
            foreach (var hive in hives)
            {
                foreach (var view in Views)
                {
                    foreach (var spec in Classes)
                    {
                        RegisterComClass(hive, view, spec, dllPath, codeBase, assemblyValue, version);
                    }
                }
            }

            RegisterExcelAddIn(dllPath);
        }

        private static void RegisterComClass(
            RegistryHive hive, RegistryView view, ComClassSpec spec, string dllPath,
            string codeBase, string assemblyValue, string version)
        {
            using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
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
            return TryReadInprocPath(RegistryHive.CurrentUser, view, path, out codeBase, out assembly);
        }

        private static bool TryReadInprocPath(
            RegistryHive hive, RegistryView view, string path, out string codeBase, out string assembly)
        {
            codeBase = null;
            assembly = null;
            using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
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
