using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>一次试跑需要的东西。全部在 UI 线程上从真实工作簿取好，试跑线程不碰真实 Excel。</summary>
    public sealed class LabRequest
    {
        /// <summary>SaveCopyAs 出来的副本（在 LabProcessRegistry 的目录里）</summary>
        public string CopyPath { get; set; }
        public string Code { get; set; }
        public string MacroName { get; set; }
        public string ActiveSheet { get; set; }
        public string Selection { get; set; }
        public Excel.XlCalculation? Calculation { get; set; }
        /// <summary>用户 Excel 的进程：副本进程和它相同就说明没拉起新实例，绝不能 Quit</summary>
        public int UserExcelPid { get; set; }
        public TimeSpan Timeout { get; set; } = LabRunner.DefaultTimeout;
    }

    /// <summary>试跑结果：给确认面板看的，不回写真实工作簿。</summary>
    public sealed class LabTrial
    {
        /// <summary>代码在副本上真正执行过（包括执行失败）</summary>
        public bool Ran { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }
        public bool TimedOut { get; set; }
        public long DurationMs { get; set; }
        public int ErrorsBefore { get; set; }
        public int ErrorsAfter { get; set; }
        /// <summary>执行器的提示：静态检查警告、被改回的全局开关、自动点掉的弹窗</summary>
        public string ExecutorNote { get; set; }
        /// <summary>副本代表不了真实环境的地方，面板上逐条标出</summary>
        public List<string> NotRepresentative { get; } = new List<string>();
        public LabDiffResult Diff { get; set; } = new LabDiffResult();
        /// <summary>没能试跑时的原因（副本起不来、超时……）</summary>
        public string SkippedReason { get; set; }
        /// <summary>副本 Excel 的进程号（日志与测试用）</summary>
        public int ProcessId { get; set; }
    }

    /// <summary>
    /// 副本试跑：在隐藏的另一个 Excel 进程里对工作簿副本执行同一段 VBA，把前后差异带回来。
    ///
    /// 为什么是 DCOM 新实例而不是 Process.Start：实测 Process.Start 起的 EXCEL.EXE 会显示窗口，
    /// 而按文件路径 BindToMoniker 可能绑到用户已经开着的那个 Excel——试跑跑到真实工作簿上，
    /// 正好是这个功能要防的事。new Excel.Application 总是新进程、UserControl=false、不可见；
    /// 插件在这种实例里保持被动（见 ThisAddIn.IsAutomationInstance）。
    /// </summary>
    public static class LabRunner
    {
        public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(20);

        /// <summary>对比时最多读这么多格；再大的表 COM 往返就不止几秒了</summary>
        public const int MaxSnapshotCells = 300000;

        public static Task<LabTrial> RunAsync(LabRequest request)
        {
            var done = new TaskCompletionSource<LabTrial>();
            var state = new RunState();
            var thread = new Thread(() =>
            {
                try { RunOnThisThread(request, state, trial => done.TrySetResult(trial)); }
                catch (Exception ex)
                {
                    Logger.Instance.Warning("LabRunner", "Trial crashed: " + ex);
                    done.TrySetResult(new LabTrial { SkippedReason = "试跑出错：" + ex.Message });
                }
            })
            { IsBackground = true, Name = "DeepExcel Lab" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return Task.Run(async () =>
            {
                var finished = await Task.WhenAny(done.Task, Task.Delay(request.Timeout)).ConfigureAwait(false);
                if (finished == done.Task) return done.Task.Result;

                // 超时：多半是死循环或等输入。杀掉副本进程，试跑线程上的 COM 调用随之报错返回
                state.TimedOut = true;
                KillLab(state.Pid);
                var afterKill = await Task.WhenAny(done.Task, Task.Delay(5000)).ConfigureAwait(false);
                var trial = afterKill == done.Task ? done.Task.Result : new LabTrial();
                trial.ProcessId = state.Pid;
                trial.TimedOut = true;
                trial.Success = false;
                trial.DurationMs = (long)request.Timeout.TotalMilliseconds;
                trial.Error = $"副本上运行超过 {request.Timeout.TotalSeconds:0} 秒仍未结束，已终止。代码可能有死循环或在等输入";
                return trial;
            });
        }

        private sealed class RunState
        {
            public volatile int Pid;
            public volatile bool TimedOut;
        }

        /// <summary>结果一出来就 publish，关副本进程（Quit 后可能要等好几秒）放在之后，不让用户等</summary>
        private static void RunOnThisThread(LabRequest request, RunState state, Action<LabTrial> publish)
        {
            var trial = new LabTrial();
            var clock = Stopwatch.StartNew();
            LabProcessRegistry.Reap();

            Excel.Application lab = null;
            Excel.Workbook book = null;
            int pid = 0;
            bool ownsProcess = false;
            try
            {
                lab = new Excel.Application();
                pid = DialogGuard.ProcessOf(SafeHwnd(lab));
                if (pid == 0 || pid == request.UserExcelPid || pid == Process.GetCurrentProcess().Id)
                {
                    // 拿到的不是新进程：什么都不做，更不能 Quit（那会关掉用户的 Excel）
                    trial.SkippedReason = "没能启动独立的副本 Excel";
                    return;
                }
                ownsProcess = true;
                state.Pid = pid;
                trial.ProcessId = pid;
                LabProcessRegistry.Register(pid);

                lab.DisplayAlerts = false;
                lab.EnableEvents = false;
                lab.ScreenUpdating = false;
                try { lab.AskToUpdateLinks = false; } catch (COMException) { }
                DisconnectOwnAddIn(lab);

                book = lab.Workbooks.Open(request.CopyPath, UpdateLinks: 0, ReadOnly: false, AddToMru: false);
                Logger.Instance.Info("LabRunner", $"Lab pid={pid} opened copy in {clock.ElapsedMilliseconds}ms");
                MatchUserView(lab, book, request);
                CollectWorkbookCaveats(book, trial.NotRepresentative);
                trial.NotRepresentative.AddRange(CodeCaveats(request.Code));

                var before = ReadWorkbook(book);
                trial.ErrorsBefore = CountFormulaErrors(book);

                trial.Ran = true;
                var runClock = Stopwatch.StartNew();
                var result = new VBAExecutor(lab).Execute(request.Code, request.MacroName);
                trial.DurationMs = runClock.ElapsedMilliseconds;
                trial.Success = result.Success;
                trial.Error = result.Success ? null : result.Error;
                trial.ExecutorNote = result.Warning;

                var after = ReadWorkbook(book);
                trial.ErrorsAfter = CountFormulaErrors(book);
                trial.Diff = LabDiff.Compare(before, after);
                if (trial.Diff.Truncated)
                {
                    trial.NotRepresentative.Add($"工作簿很大，只对比了前 {MaxSnapshotCells:N0} 个单元格");
                }
                return;
            }
            catch (Exception ex) when (state.TimedOut)
            {
                Logger.Instance.Info("LabRunner", "Trial aborted after timeout: " + ex.Message);
                return;
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("LabRunner", "Trial failed: " + ex);
                if (!trial.Ran) trial.SkippedReason = "副本没能打开：" + ex.Message;
                else { trial.Success = false; trial.Error = "试跑出错：" + ex.Message; }
                return;
            }
            finally
            {
                publish(trial);
                if (ownsProcess) Shutdown(lab, book, pid);
                else if (lab != null) { try { Marshal.FinalReleaseComObject(lab); } catch (Exception) { } }
                LabProcessRegistry.DeleteCopyDirectory(System.IO.Path.GetDirectoryName(request.CopyPath));
                Logger.Instance.Info("LabRunner",
                    $"Trial done in {clock.ElapsedMilliseconds}ms pid={pid} ran={trial.Ran} ok={trial.Success} cells={trial.Diff?.AffectedCells}");
            }
        }

        private static readonly (Regex Pattern, string Note)[] CodePatterns =
        {
            (new Regex(@"\bWorkbooks\s*(\(|\.\s*(Open|Add)\b)", RegexOptions.IgnoreCase),
                "代码会打开或切换到其他工作簿，副本 Excel 里没有它们"),
            (new Regex(@"\b(ThisWorkbook|ActiveWorkbook)\s*\.\s*(Path|FullName)\b", RegexOptions.IgnoreCase),
                "代码读取了工作簿所在路径，副本在临时目录里"),
            (new Regex(@"\.\s*(Save|SaveAs|SaveCopyAs)\b", RegexOptions.IgnoreCase),
                "代码会保存文件；试跑里保存的只是副本"),
            (new Regex(@"\b(Rnd|Randomize|Now|Timer|RandBetween)\b|(?<!\bAs\s+)\bDate\b", RegexOptions.IgnoreCase),
                "结果依赖当前时间或随机数，真实执行时会不同"),
        };

        private static readonly Regex StringLiteral = new Regex("\"[^\"]*\"", RegexOptions.Compiled);

        /// <summary>从代码本身看出来的「副本代表不了」的地方。纯函数，便于测试。</summary>
        public static List<string> CodeCaveats(string code)
        {
            var notes = new List<string>();
            if (string.IsNullOrEmpty(code)) return notes;
            // 注释和字符串里的字不算（"=TODAY()" 写进单元格、注释里提到 Now 都不是代码在取时间）
            var body = string.Join("\n", code.Split('\n').Select(line => StringLiteral.Replace(StripComment(line), "\"\"")));
            foreach (var (pattern, note) in CodePatterns)
            {
                if (pattern.IsMatch(body)) notes.Add(note);
            }
            return notes;
        }

        private static string StripComment(string line)
        {
            bool inString = false;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"') inString = !inString;
                else if (line[i] == '\'' && !inString) return line.Substring(0, i);
            }
            return line;
        }

        private static void CollectWorkbookCaveats(Excel.Workbook book, List<string> notes)
        {
            try
            {
                if (book.LinkSources(Excel.XlLink.xlExcelLinks) is Array links && links.Length > 0)
                {
                    notes.Add("工作簿引用了其他文件（外部链接），副本打开时没有更新这些链接");
                }
            }
            catch (COMException) { }
            try
            {
                if (book.Connections.Count > 0)
                {
                    notes.Add("工作簿有数据连接或 Power Query，副本里没有刷新");
                }
            }
            catch (COMException) { }
            try
            {
                if (book.HasVBProject)
                {
                    notes.Add("工作簿自带宏；试跑时关闭了事件，Worksheet_Change 之类的事件宏没有触发");
                }
            }
            catch (COMException) { }
        }

        private static void MatchUserView(Excel.Application lab, Excel.Workbook book, LabRequest request)
        {
            // 计算模式要在有工作簿打开之后才能设
            if (request.Calculation.HasValue)
            {
                try { lab.Calculation = request.Calculation.Value; } catch (COMException) { }
            }
            if (string.IsNullOrEmpty(request.ActiveSheet)) return;
            try
            {
                var sheet = (Excel.Worksheet)book.Worksheets[request.ActiveSheet];
                sheet.Activate();
                if (!string.IsNullOrEmpty(request.Selection)) sheet.Range[request.Selection].Select();
            }
            catch (Exception ex) when (ex is COMException || ex is InvalidCastException)
            {
                // 活动的是图表页之类：代码里的 ActiveSheet 本来就指不到单元格，照样跑
                Logger.Instance.Info("LabRunner", "Could not mirror active sheet: " + ex.Message);
            }
        }

        /// <summary>副本实例里也会加载本插件（被动模式）；断开它，省掉加载时间，也彻底排除它碰侧车</summary>
        private static void DisconnectOwnAddIn(Excel.Application lab)
        {
            try
            {
                foreach (Microsoft.Office.Core.COMAddIn addIn in lab.COMAddIns)
                {
                    var progId = addIn.ProgId ?? "";
                    if (progId.IndexOf("DeepExcel", StringComparison.OrdinalIgnoreCase) >= 0 && addIn.Connect)
                    {
                        addIn.Connect = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Info("LabRunner", "Could not disconnect add-in in lab: " + ex.Message);
            }
        }

        internal static List<SheetContent> ReadWorkbook(Excel.Workbook book)
        {
            var sheets = new List<SheetContent>();
            int budget = MaxSnapshotCells;
            foreach (Excel.Worksheet ws in book.Worksheets)
            {
                var content = new SheetContent { Name = ws.Name };
                sheets.Add(content);
                var used = ws.UsedRange;
                int rows = used.Rows.Count, cols = used.Columns.Count;
                content.Row1 = used.Row;
                content.Col1 = used.Column;
                if (rows == 1 && cols == 1 && used.Formula is string only && only.Length == 0) continue;
                if ((long)rows * cols > budget)
                {
                    rows = Math.Max(0, budget / Math.Max(cols, 1));
                    content.Truncated = true;
                }
                if (rows == 0) continue;
                budget -= rows * cols;
                var range = ws.Range[ws.Cells[content.Row1, content.Col1],
                    ws.Cells[content.Row1 + rows - 1, content.Col1 + cols - 1]];
                content.Cells = ToGrid(range.Formula, rows, cols);
            }
            return sheets;
        }

        private static string[,] ToGrid(object formula, int rows, int cols)
        {
            var grid = new string[rows, cols];
            if (formula is object[,] values)
            {
                int r0 = values.GetLowerBound(0), c0 = values.GetLowerBound(1);
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        grid[r, c] = Convert.ToString(values[r0 + r, c0 + c], CultureInfo.InvariantCulture);
            }
            else
            {
                grid[0, 0] = Convert.ToString(formula, CultureInfo.InvariantCulture);
            }
            return grid;
        }

        internal static int CountFormulaErrors(Excel.Workbook book)
        {
            int total = 0;
            foreach (Excel.Worksheet ws in book.Worksheets)
            {
                try
                {
                    // 没有匹配的格时 SpecialCells 抛异常，不是返回空
                    var errors = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);
                    total += (int)Math.Min(int.MaxValue, Convert.ToInt64(errors.CountLarge));
                }
                catch (COMException) { }
            }
            return total;
        }

        private static int SafeHwnd(Excel.Application app)
        {
            try { return app.Hwnd; } catch (COMException) { return 0; }
        }

        private static void Shutdown(Excel.Application lab, Excel.Workbook book, int pid)
        {
            try { book?.Close(false); } catch (Exception) { }
            try { lab?.Quit(); } catch (Exception) { }
            try { if (book != null) Marshal.FinalReleaseComObject(book); } catch (Exception) { }
            try { if (lab != null) Marshal.FinalReleaseComObject(lab); } catch (Exception) { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Quit 之后进程可能因为别的加载项继续挂着（实测过）：等一会儿不走就结束它
            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    if (!p.WaitForExit(5000)) p.Kill();
                }
            }
            catch (Exception) { }
            LabProcessRegistry.Unregister(pid);
        }

        private static void KillLab(int pid)
        {
            if (pid <= 0) return;
            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    if (!p.HasExited) p.Kill();
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("LabRunner", $"Kill lab pid={pid} failed: {ex.Message}");
            }
        }
    }
}
