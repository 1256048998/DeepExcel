using System;

namespace DeepExcel.Repair
{
    /// <summary>
    /// DeepExcel.Probe32.exe -- a 32-bit process whose only job is to prove that
    /// 32-bit Excel can activate the add-in.
    ///
    /// DeepExcel.AddIn is AnyCPU and 32-bit Office is still common, but a
    /// 64-bit process cannot activate a COM class in the 32-bit registry view.
    /// This binary exists purely to be that second process; it replaces the old
    /// SysWOW64 powershell.exe hop in the installer.
    ///
    /// Exit codes: 0 activation succeeded, 1 it did not, 2 usage error.
    /// </summary>
    public static class Probe32Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            // Output is normally read back through a pipe by Repair.exe, which
            // decodes as UTF-8. Keep both ends explicit.
            try { Console.OutputEncoding = System.Text.Encoding.UTF8; }
            catch (Exception) { }

            if (args.Length != 1 || !string.Equals(args[0], "--activation-only", StringComparison.OrdinalIgnoreCase))
            {
                Console.Error.WriteLine("用法: DeepExcel.Probe32.exe --activation-only");
                return 2;
            }

            if (IntPtr.Size != 4)
            {
                // A 64-bit build here would silently test the wrong registry
                // view and report a false pass.
                Console.Error.WriteLine("DeepExcel.Probe32.exe 必须以 32 位运行，当前为 64 位。打包有误。");
                return 2;
            }

            var failure = ComProbe.TryActivate(AddInIdentity.MainProgId);
            if (failure != null)
            {
                Console.Error.WriteLine(failure);
                return 1;
            }

            Console.WriteLine("32 位进程 COM 激活成功。");
            return 0;
        }
    }
}
