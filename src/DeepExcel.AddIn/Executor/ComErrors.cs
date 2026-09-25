using System;
using System.Threading;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// Excel COM 错误的分类：瞬时的（Excel 暂时不接电话）有界重试，硬错误附上「下一步怎么修」。
    ///
    /// 瞬时错误只认三种——它们的意思都是「调用被拒绝、没有执行」，重试不会重复副作用：
    /// - 0x80010001 RPC_E_CALL_REJECTED：Excel 在忙（进程外调用时）
    /// - 0x8001010A RPC_E_SERVERCALL_RETRYLATER：同上，要求稍后再试
    /// - 0x800AC472 VBA_E_IGNORE：用户正在编辑单元格或开着对话框，对象模型暂停服务
    /// </summary>
    public static class ComErrors
    {
        public const int CallRejected = unchecked((int)0x80010001);
        public const int RetryLater = unchecked((int)0x8001010A);
        public const int ExcelBusy = unchecked((int)0x800AC472);

        public const int MaxAttempts = 3;

        public static bool IsTransient(Exception ex)
        {
            for (var e = ex; e != null; e = e.InnerException)
            {
                var hr = e.HResult;
                if (hr == CallRejected || hr == RetryLater || hr == ExcelBusy) return true;
            }
            return false;
        }

        /// <summary>瞬时错误最多试 MaxAttempts 次，每次之间让出消息泵、退避一点；其余错误原样抛出</summary>
        public static T Retry<T>(Func<T> call, string what, Action pump = null, Func<int, int> delayMs = null)
        {
            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    return call();
                }
                catch (Exception ex) when (attempt < MaxAttempts && IsTransient(ex))
                {
                    Logger.Instance.Warning("ComErrors", $"{what}: transient 0x{ex.HResult:X8}, retry {attempt}/{MaxAttempts}");
                    pump?.Invoke();
                    Thread.Sleep(delayMs?.Invoke(attempt) ?? 150 * attempt);
                }
            }
        }

        /// <summary>给模型的下一步提示；认不出的错误返回 null</summary>
        public static string Hint(Exception ex)
        {
            if (ex == null) return null;
            if (IsTransient(ex))
            {
                return "Excel 暂时不响应（用户可能正在编辑单元格或开着对话框），重试了 " + MaxAttempts +
                       " 次仍然不行。请告诉用户按 Esc 退出编辑、关掉 Excel 里的对话框，然后再继续";
            }
            var message = ex.Message ?? "";
            switch (unchecked((uint)ex.HResult))
            {
                case 0x800A03EC:
                    return "Excel 拒绝了这次操作。常见原因：地址或公式写错（公式要用英文函数名和逗号分隔）、" +
                           "目标是合并单元格或受保护的工作表、数组公式区域只改了一部分。先 read_range 看一下目标区域再改";
                case 0x80020005:  // DISP_E_TYPEMISMATCH
                    return "参数类型不对（比如该给数字的地方给了文字）。检查参数后重试";
                case 0x800706BA:  // RPC_S_SERVER_UNAVAILABLE
                case 0x80010108:  // RPC_E_DISCONNECTED
                case 0x800706BE:  // RPC_S_CALL_FAILED
                    return "与 Excel 的连接断了（Excel 可能已关闭或崩溃）。停下来告诉用户，不要继续重试";
                case 0x800A01A8:  // 424 需要对象
                case 0x80004003:  // E_POINTER
                    return "要操作的对象不存在（工作表、表格或区域可能已被删除或改名）。先 list_objects 确认名字";
            }
            if (message.IndexOf("protect", StringComparison.OrdinalIgnoreCase) >= 0 || message.Contains("保护"))
            {
                return "工作表或工作簿受保护。请告诉用户需要先撤销保护，不要尝试破解密码";
            }
            return null;
        }

        /// <summary>同一个错误连续出现这么多次，就提示模型换一种做法</summary>
        public const int SameErrorLimit = 3;

        public const string ChangeStrategyHint =
            "同样的错误已经连续出现 3 次，别再用同样的办法重试：换一种做法（换个工具、先读一下目标区域确认现状、把操作拆小），" +
            "或者停下来告诉用户卡在哪里、问用户怎么办";
    }
}
