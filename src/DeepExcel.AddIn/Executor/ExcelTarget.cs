using System;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// 工具要操作的工作簿（会话绑定的那本），而不是「此刻在前台的那本」。
    ///
    /// 以前所有工具都写 _app.ActiveWorkbook / _app.Range[...]：用户在 AI 工作时切到另一本
    /// 工作簿，读会读错文件，写只能靠守卫整体拒绝。ToolDispatcher 在执行工具期间用
    /// Use(boundWorkbook) 指定目标，这里的 Workbook / ActiveSheet / Range 都按目标解析。
    ///
    /// 没有指定目标、或目标就是前台工作簿时，行为和以前完全一样（仍走 _app.Range 等），
    /// 新的解析路径只在用户切走时才生效。只在 UI 线程使用（工具都在 STA 主线程执行）。
    /// </summary>
    public static class ExcelTarget
    {
        [ThreadStatic] private static Workbook _current;

        /// <summary>在 using 范围内把 wb 当作工具的目标工作簿；wb 为 null 时不改变什么。</summary>
        public static IDisposable Use(Workbook wb)
        {
            var previous = _current;
            if (wb != null) _current = wb;
            return new Scope(previous);
        }

        /// <summary>当前是否指定了一个不在前台的目标工作簿</summary>
        public static bool IsRedirected(Application app)
        {
            var target = _current;
            if (target == null) return false;
            try
            {
                var active = app?.ActiveWorkbook;
                return active == null || !WorkbookIdentity.SameKey(WorkbookIdentity.KeyOf(active), WorkbookIdentity.KeyOf(target));
            }
            catch
            {
                return true;
            }
        }

        public static Workbook Workbook(Application app)
        {
            return _current ?? app?.ActiveWorkbook;
        }

        /// <summary>目标工作簿自己的活动表（不在前台时也有）。</summary>
        public static Worksheet ActiveSheet(Application app)
        {
            if (!IsRedirected(app)) return app?.ActiveSheet as Worksheet;
            try { return _current.ActiveSheet as Worksheet; }
            catch { return null; }
        }

        /// <summary>
        /// 解析地址：「Sheet2!A1:C4」「'My Sheet'!B2」「A1」（目标的活动表）、工作簿级命名区域。
        /// 目标就是前台工作簿时等价于 app.Range[address]。
        /// </summary>
        public static Range Range(Application app, string address)
        {
            if (!IsRedirected(app)) return app.Range[address];

            var wb = _current;
            var text = (address ?? "").Trim();
            var bang = text.LastIndexOf('!');
            Worksheet sheet;
            string cell;
            if (bang >= 0)
            {
                var sheetName = text.Substring(0, bang).Trim();
                if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[sheetName.Length - 1] == '\'')
                {
                    sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'");
                }
                sheet = wb.Worksheets[sheetName] as Worksheet;
                cell = text.Substring(bang + 1).Trim();
            }
            else
            {
                sheet = wb.ActiveSheet as Worksheet;
                cell = text;
            }
            if (sheet == null) throw new ArgumentException("找不到工作表：" + address);

            try
            {
                return sheet.Range[cell];
            }
            catch when (bang < 0)
            {
                // 工作簿级命名区域指向别的表时，工作表的 Range 解析不了
                return wb.Names.Item(cell).RefersToRange;
            }
        }

        private sealed class Scope : IDisposable
        {
            private readonly Workbook _previous;
            private bool _disposed;

            public Scope(Workbook previous) { _previous = previous; }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                _current = _previous;
            }
        }
    }
}
