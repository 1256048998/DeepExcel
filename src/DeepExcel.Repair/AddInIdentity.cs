using System;
using System.IO;
using System.Reflection;

namespace DeepExcel.Repair
{
    /// <summary>
    /// COM identities and paths. These MUST stay in sync with ThisAddIn.cs,
    /// TaskPaneControl.cs and deploy/DeepExcel.Setup.iss. A mismatch here is
    /// indistinguishable from "add-in not installed" at runtime.
    /// </summary>
    public static class AddInIdentity
    {
        public const string MainClsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}";
        public const string TaskPaneClsid = "{B2C3D4E5-F6A7-404B-9A5F-9B3C1D2E3F4B}";
        public const string MainProgId = "DeepExcel.AddIn";
        public const string TaskPaneProgId = "DeepExcel.AddIn.TaskPaneControl";
        public const string MainClass = "DeepExcel.AddIn.ThisAddIn";
        public const string TaskPaneClass = "DeepExcel.AddIn.TaskPaneControl";
        public const string DotNetCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";
        public const string AddInDllName = "DeepExcel.AddIn.dll";

        public const string ExcelAddinsKey = @"Software\Microsoft\Office\Excel\Addins\" + MainProgId;
        public const string ResiliencyKey = @"Software\Microsoft\Office\16.0\Excel\Resiliency";

        /// <summary>
        /// Install root. Repair.exe ships next to DeepExcel.AddIn.dll, so its own
        /// location is the authoritative answer and needs no registry lookup.
        /// </summary>
        public static string InstallDirectory
        {
            get
            {
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            }
        }

        public static string AddInDllPath
        {
            get { return Path.Combine(InstallDirectory, AddInDllName); }
        }

        public static string LogDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "DeepExcel", "logs");
            }
        }

        public static string ConfigPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "DeepExcel", "config.json");
            }
        }

        /// <summary>Assembly strong name string written into InprocServer32.</summary>
        public static string AssemblyValue(string dllPath)
        {
            var version = AssemblyName.GetAssemblyName(dllPath).Version.ToString();
            return "DeepExcel.AddIn, Version=" + version + ", Culture=neutral, PublicKeyToken=null";
        }

        public static string AssemblyVersion(string dllPath)
        {
            return AssemblyName.GetAssemblyName(dllPath).Version.ToString();
        }

        /// <summary>
        /// mscoree resolves CodeBase as a file URI. Building it with Uri keeps
        /// escaping identical to what the installer writes, so a comparison
        /// between the two cannot produce a false "stale path" finding.
        /// </summary>
        public static string CodeBase(string dllPath)
        {
            return new Uri(dllPath).AbsoluteUri;
        }
    }
}
