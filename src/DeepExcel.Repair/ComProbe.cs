using System;
using System.Runtime.InteropServices;

namespace DeepExcel.Repair
{
    /// <summary>
    /// Activates the add-in through COM exactly as Excel would.
    ///
    /// This is the only check that proves the whole chain works: registry entry,
    /// mscoree shim, assembly resolution, and the managed constructor. Registry
    /// verification alone has passed on machines where the add-in still failed
    /// to load.
    ///
    /// Compiled into both DeepExcel.Repair.exe (64-bit) and
    /// DeepExcel.Probe32.exe (32-bit). A process can only activate in its own
    /// bitness, so covering both Office bitnesses needs both executables.
    /// </summary>
    public static class ComProbe
    {
        public static int BitnessOfCurrentProcess
        {
            get { return IntPtr.Size * 8; }
        }

        /// <summary>
        /// Returns null on success, or the failure reason.
        /// </summary>
        public static string TryActivate(string progId)
        {
            object instance = null;
            try
            {
                var type = Type.GetTypeFromProgID(progId, throwOnError: true);
                if (type == null)
                {
                    return "ProgID 未注册：" + progId;
                }
                instance = Activator.CreateInstance(type);
                if (instance == null)
                {
                    return "COM 对象创建返回空：" + progId;
                }
                return null;
            }
            catch (Exception ex)
            {
                return string.Format(
                    "{0} 位进程中 COM 激活失败：{1} - {2}",
                    BitnessOfCurrentProcess, ex.GetType().Name, ex.Message);
            }
            finally
            {
                if (instance != null)
                {
                    try { Marshal.ReleaseComObject(instance); }
                    catch (Exception) { /* best effort */ }
                }
            }
        }
    }
}
