using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 用 Windows Job Object 管住一整棵子进程树。Process.Kill 只杀直接子进程：脚本里起的孙进程
    /// （pandas 的并行、multiprocessing、被调用的外部程序）会留下来继续吃 CPU 和内存。
    /// 进程放进 job 后：超时时 Terminate 整棵树一起结束；插件崩溃或忘了清理时，句柄关闭（进程退出）
    /// 也会带走整棵树（KILL_ON_JOB_CLOSE）。
    ///
    /// 创建失败（极老的系统、嵌套 job 不允许）时退化成只管直接子进程，不影响执行。
    /// </summary>
    public sealed class ProcessJob : IDisposable
    {
        private IntPtr _handle;

        private ProcessJob(IntPtr handle)
        {
            _handle = handle;
        }

        public bool Active => _handle != IntPtr.Zero;

        public static ProcessJob Create(long memoryLimitBytes = 0)
        {
            var handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero)
            {
                Logger.Instance.Warning("ProcessJob", "CreateJobObject failed: " + Marshal.GetLastWin32Error());
                return new ProcessJob(IntPtr.Zero);
            }
            var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
            info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | JOB_OBJECT_LIMIT_DIE_ON_UNHANDLED_EXCEPTION;
            if (memoryLimitBytes > 0)
            {
                info.BasicLimitInformation.LimitFlags |= JOB_OBJECT_LIMIT_JOB_MEMORY;
                info.JobMemoryLimit = new UIntPtr((ulong)memoryLimitBytes);
            }
            var length = Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
            var buffer = Marshal.AllocHGlobal(length);
            try
            {
                Marshal.StructureToPtr(info, buffer, false);
                if (!SetInformationJobObject(handle, JobObjectExtendedLimitInformation, buffer, (uint)length))
                {
                    Logger.Instance.Warning("ProcessJob", "SetInformationJobObject failed: " + Marshal.GetLastWin32Error());
                    CloseHandle(handle);
                    return new ProcessJob(IntPtr.Zero);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
            return new ProcessJob(handle);
        }

        /// <summary>
        /// 把刚启动的进程放进来。Process.Start 没法挂起启动，进程在放进来之前的几毫秒里起的子进程
        /// 管不到——脚本的前几行是我们自己的引导代码，不会起进程，这个窗口实际上是空的。
        /// </summary>
        public bool Assign(Process process)
        {
            if (!Active || process == null) return false;
            try
            {
                if (AssignProcessToJobObject(_handle, process.Handle)) return true;
                Logger.Instance.Warning("ProcessJob", "AssignProcessToJobObject failed: " + Marshal.GetLastWin32Error());
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ProcessJob", "Assign failed: " + ex.Message);
            }
            return false;
        }

        /// <summary>结束 job 里的所有进程</summary>
        public void Terminate()
        {
            if (Active) TerminateJobObject(_handle, 1);
        }

        public void Dispose()
        {
            if (_handle == IntPtr.Zero) return;
            CloseHandle(_handle);  // KILL_ON_JOB_CLOSE：还活着的进程一起结束
            _handle = IntPtr.Zero;
        }

        private const int JobObjectExtendedLimitInformation = 9;
        private const uint JOB_OBJECT_LIMIT_JOB_MEMORY = 0x00000200;
        private const uint JOB_OBJECT_LIMIT_DIE_ON_UNHANDLED_EXCEPTION = 0x00000400;
        private const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000;

        [StructLayout(LayoutKind.Sequential)]
        private struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IO_COUNTERS
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateJobObject(IntPtr attributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr job, int infoClass, IntPtr info, uint length);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);
    }

    /// <summary>
    /// 有上限地读完一个输出流：只保留前 limit 个字符，其余照样读走（不读的话子进程写满管道会卡住），
    /// 只计数。脚本 while True: print(...) 以前能把插件内存吃光。
    /// </summary>
    public sealed class BoundedOutput
    {
        public string Text { get; private set; } = "";
        public long TotalChars { get; private set; }
        public bool Truncated => TotalChars > Text.Length;

        public static async Task<BoundedOutput> ReadAsync(StreamReader reader, int limit)
        {
            var result = new BoundedOutput();
            var kept = new StringBuilder();
            var buffer = new char[4096];
            int read;
            while ((read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
            {
                result.TotalChars += read;
                var room = limit - kept.Length;
                if (room > 0) kept.Append(buffer, 0, Math.Min(room, read));
            }
            result.Text = kept.ToString();
            return result;
        }

        /// <summary>给模型看的文本：截断时在末尾注明</summary>
        public string Describe() =>
            Truncated ? Text + $"\n…（输出太长，只保留前 {Text.Length} 个字符，共 {TotalChars} 个）" : Text;
    }
}
