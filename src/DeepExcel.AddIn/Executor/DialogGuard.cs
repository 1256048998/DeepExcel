using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>执行期间被自动处理掉的一个弹窗</summary>
    public sealed class HandledDialog
    {
        public string Title { get; set; }
        public string Text { get; set; }
        /// <summary>点了哪个按钮（按钮上的字）；没认出按钮时是「关闭」</summary>
        public string Clicked { get; set; }

        public override string ToString()
        {
            var text = string.IsNullOrWhiteSpace(Text) ? "" : "：" + Text;
            return $"「{Title}{text}」→ 点了「{Clicked}」";
        }
    }

    /// <summary>
    /// VBA 执行期间的弹窗兜底。_app.Run 在 Excel 的 UI 线程上同步跑，这时弹出的对话框没人能点
    /// （任务窗格也被一起卡住），Excel 就一直挂着。这里起一个后台线程，只看目标进程在守护开始
    /// 之后新出现的 #32770 对话框：
    ///
    /// - 询问类只点「否 / 取消 / 中止」，绝不点「是 / 重试 / 忽略」——宁可让代码少做一步，
    ///   也不替用户同意覆盖、删除之类的事
    /// - 提示类（只有「确定」）点确定，比如 VBE 的「编译错误」
    /// - 点了什么、弹窗上写了什么，如实记下来交给模型
    ///
    /// 只用 Win32 消息，不碰 COM，所以在后台线程上安全。
    /// </summary>
    public sealed class DialogGuard : IDisposable
    {
        private const string DialogClass = "#32770";
        private const int WM_COMMAND = 0x0111;
        private const int WM_CLOSE = 0x0010;
        private const int BN_CLICKED = 0;

        // 按优先级：先找最保守的按钮。IDYES(6) / IDRETRY(4) / IDIGNORE(5) / IDTRYAGAIN(10) / IDCONTINUE(11) 永远不点
        private static readonly int[] SafeButtons = { 2 /*IDCANCEL*/, 7 /*IDNO*/, 3 /*IDABORT*/, 1 /*IDOK*/, 8 /*IDCLOSE*/ };

        private readonly int _processId;
        private readonly int _pollMs;
        private readonly HashSet<IntPtr> _seen = new HashSet<IntPtr>();
        private readonly List<HandledDialog> _handled = new List<HandledDialog>();
        private readonly object _lock = new object();
        private readonly Thread _thread;
        private volatile bool _stop;

        private DialogGuard(int processId, int pollMs)
        {
            _processId = processId;
            _pollMs = pollMs;
            // 守护开始前就开着的对话框是用户自己的，不碰
            foreach (var hwnd in DialogsOf(processId)) _seen.Add(hwnd);
            _thread = new Thread(Run) { IsBackground = true, Name = "DeepExcel DialogGuard" };
            _thread.Start();
        }

        /// <summary>开始看守某个进程；processId 为 0 时不看守（返回的对象什么也不做）</summary>
        public static DialogGuard Start(int processId, int pollMs = 150) => new DialogGuard(processId, pollMs);

        /// <summary>Excel 主窗口所在的进程（插件在进程内时就是当前进程；测试里是外部 Excel）</summary>
        public static int ProcessOf(int hwnd)
        {
            if (hwnd == 0) return 0;
            GetWindowThreadProcessId(new IntPtr(hwnd), out var pid);
            return pid;
        }

        public IReadOnlyList<HandledDialog> Handled
        {
            get { lock (_lock) return _handled.ToList(); }
        }

        public void Dispose()
        {
            _stop = true;
            if (_thread.IsAlive && Thread.CurrentThread != _thread) _thread.Join(2000);
        }

        public static string Describe(IReadOnlyList<HandledDialog> dialogs)
        {
            if (dialogs == null || dialogs.Count == 0) return null;
            return "执行期间弹出了 " + dialogs.Count + " 个对话框，已自动处理：" +
                   string.Join("；", dialogs.Take(5).Select(d => d.ToString())) +
                   (dialogs.Count > 5 ? $"；另有 {dialogs.Count - 5} 个" : "");
        }

        private void Run()
        {
            while (!_stop)
            {
                try { Sweep(); }
                catch (Exception ex) { Logger.Instance.Warning("DialogGuard", "Sweep failed: " + ex.Message); }
                Thread.Sleep(_pollMs);
            }
        }

        private void Sweep()
        {
            if (_processId == 0) return;
            foreach (var hwnd in DialogsOf(_processId))
            {
                if (_seen.Contains(hwnd)) continue;
                // 刚创建的对话框按钮可能还没建好，等它可见、有子窗口再处理
                var buttons = Buttons(hwnd);
                if (buttons.Count == 0) continue;
                _seen.Add(hwnd);
                var record = new HandledDialog { Title = WindowText(hwnd), Text = DialogText(hwnd) };
                var choice = SafeButtons
                    .Select(id => buttons.FirstOrDefault(b => b.Id == id))
                    .FirstOrDefault(b => b != null);
                record.Clicked = choice != null ? CleanLabel(choice.Text) : "关闭";
                // 先记下再点：弹窗一关，UI 线程上等着它的调用马上返回并读取记录
                lock (_lock) _handled.Add(record);
                Logger.Instance.Warning("DialogGuard", "Dismissing dialog: " + record);
                if (choice != null)
                {
                    PostMessage(hwnd, WM_COMMAND, new IntPtr((BN_CLICKED << 16) | choice.Id), choice.Hwnd);
                }
                else
                {
                    PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                }
            }
        }

        private sealed class Button
        {
            public IntPtr Hwnd;
            public int Id;
            public string Text;
        }

        private static List<IntPtr> DialogsOf(int processId)
        {
            var found = new List<IntPtr>();
            if (processId == 0) return found;
            EnumWindows((hwnd, _) =>
            {
                GetWindowThreadProcessId(hwnd, out var pid);
                if (pid == processId && IsWindowVisible(hwnd) && ClassOf(hwnd) == DialogClass) found.Add(hwnd);
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private static List<Button> Buttons(IntPtr dialog)
        {
            var buttons = new List<Button>();
            EnumChildWindows(dialog, (child, _) =>
            {
                if (ClassOf(child) == "Button" && IsWindowVisible(child))
                {
                    buttons.Add(new Button { Hwnd = child, Id = GetDlgCtrlID(child), Text = WindowText(child) });
                }
                return true;
            }, IntPtr.Zero);
            return buttons;
        }

        private static string DialogText(IntPtr dialog)
        {
            var parts = new List<string>();
            EnumChildWindows(dialog, (child, _) =>
            {
                if (ClassOf(child) == "Static" && IsWindowVisible(child))
                {
                    var text = WindowText(child).Trim();
                    if (text.Length > 0) parts.Add(text);
                }
                return true;
            }, IntPtr.Zero);
            var joined = string.Join(" ", parts).Replace("\r", " ").Replace("\n", " ");
            return joined.Length > 300 ? joined.Substring(0, 300) + "…" : joined;
        }

        /// <summary>按钮上的「&amp;No」去掉快捷键标记</summary>
        private static string CleanLabel(string text) => (text ?? "").Replace("&", "").Trim();

        private static string ClassOf(IntPtr hwnd)
        {
            var sb = new StringBuilder(64);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string WindowText(IntPtr hwnd)
        {
            var length = GetWindowTextLength(hwnd);
            if (length <= 0) return "";
            var sb = new StringBuilder(length + 1);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern int GetDlgCtrlID(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
    }
}
