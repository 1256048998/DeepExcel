using System;
using System.Collections.Generic;
using DeepExcel.AddIn.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>
    /// VBA 执行前记下 Excel 的全局开关，执行后把代码改了没改回来的恢复原样。模型写的 VBA 常见
    /// 「开头 Calculation = 手动、EnableEvents = False，出错跳走没走到结尾」，用户的 Excel 从此
    /// 不再自动重算、事件全失效，而且看不出原因。恢复了哪些如实告诉模型。
    ///
    /// ScreenUpdating 执行器自己也会关，恢复但不回报。
    /// </summary>
    public sealed class AppStateGuard
    {
        private readonly Excel.Application _app;
        private readonly bool? _screenUpdating;
        private readonly bool? _enableEvents;
        private readonly bool? _displayAlerts;
        private readonly Excel.XlCalculation? _calculation;
        private readonly object _statusBar;
        private readonly bool _statusBarRead;

        private AppStateGuard(Excel.Application app)
        {
            _app = app;
            _screenUpdating = Try(() => app.ScreenUpdating);
            _enableEvents = Try(() => app.EnableEvents);
            _displayAlerts = Try(() => app.DisplayAlerts);
            // 没有打开的工作簿时读 Calculation 会抛异常
            _calculation = Try(() => app.Calculation);
            try
            {
                _statusBar = app.StatusBar;
                _statusBarRead = true;
            }
            catch { }
        }

        public static AppStateGuard Capture(Excel.Application app) => new AppStateGuard(app);

        /// <summary>恢复被改掉的开关，返回需要回报的项（空列表表示都没动）</summary>
        public List<string> Restore()
        {
            var restored = new List<string>();
            if (_screenUpdating.HasValue && Try(() => _app.ScreenUpdating) != _screenUpdating)
            {
                Set(() => _app.ScreenUpdating = _screenUpdating.Value, null, restored);
            }
            if (_enableEvents.HasValue && Try(() => _app.EnableEvents) != _enableEvents)
            {
                Set(() => _app.EnableEvents = _enableEvents.Value,
                    $"EnableEvents（代码改成了 {Bool(!_enableEvents.Value)}，已改回 {Bool(_enableEvents.Value)}）", restored);
            }
            if (_displayAlerts.HasValue && Try(() => _app.DisplayAlerts) != _displayAlerts)
            {
                Set(() => _app.DisplayAlerts = _displayAlerts.Value,
                    $"DisplayAlerts（代码改成了 {Bool(!_displayAlerts.Value)}，已改回 {Bool(_displayAlerts.Value)}）", restored);
            }
            var calculation = Try(() => _app.Calculation);
            if (_calculation.HasValue && calculation.HasValue && calculation != _calculation)
            {
                Set(() => _app.Calculation = _calculation.Value,
                    $"Calculation（代码改成了{CalcName(calculation.Value)}，已改回{CalcName(_calculation.Value)}）", restored);
            }
            if (_statusBarRead)
            {
                object current = null;
                try { current = _app.StatusBar; } catch { }
                if (!Equals(current, _statusBar))
                {
                    // StatusBar 默认值读出来是 False。经 COM 写 False 会被当成文字「FALSE」显示出来，
                    // 交还给 Excel 要用 XLM 的 MESSAGE(FALSE)
                    Set(() =>
                        {
                            if (_statusBar is string text) _app.StatusBar = text;
                            else _app.ExecuteExcel4Macro("MESSAGE(FALSE)");
                        },
                        "StatusBar（代码留下了状态栏文字，已清除）", restored);
                }
            }
            return restored;
        }

        public static string Describe(List<string> restored) =>
            restored == null || restored.Count == 0
                ? null
                : "代码改了 Excel 全局设置没改回来，已恢复：" + string.Join("、", restored) +
                  "。以后用完这些开关请在结尾（包括出错跳转的分支）改回去";

        private static void Set(Action set, string report, List<string> restored)
        {
            try
            {
                set();
                if (report != null) restored.Add(report);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("AppStateGuard", "Restore failed: " + ex.Message);
            }
        }

        private static T? Try<T>(Func<T> read) where T : struct
        {
            try { return read(); }
            catch { return null; }
        }

        private static string Bool(bool value) => value ? "True" : "False";

        private static string CalcName(Excel.XlCalculation calculation)
        {
            switch (calculation)
            {
                case Excel.XlCalculation.xlCalculationManual: return "手动计算";
                case Excel.XlCalculation.xlCalculationSemiautomatic: return "除模拟运算表外自动";
                default: return "自动计算";
            }
        }
    }
}
