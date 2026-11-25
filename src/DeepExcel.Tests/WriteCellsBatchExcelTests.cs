using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Preview;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace DeepExcel.Tests
{
    /// <summary>
    /// write_cells（ToolDispatcher.ExecuteBatch）在真 Excel 里：单测只能假设的几件事——
    /// 一批只存一份快照、后面一步失败时真的整批回到写之前（包括删掉这批新建的表）、
    /// 按表收窄的检查点确实盖住了这批改到的表、合并预览和实际写入一致。
    /// 默认跳过，DEEPEXCEL_EXCEL_TESTS=1 才运行。
    /// </summary>
    public class WriteCellsBatchExcelTests : IDisposable
    {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "DeepExcelBatch_" + Guid.NewGuid().ToString("N"));
        private Excel.Application _app;
        private int _excelPid;
        private Workbook _book;
        private SnapshotManager _snapshots;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private ToolDispatcher Setup()
        {
            Directory.CreateDirectory(_root);
            _app = new Excel.Application { Visible = false, DisplayAlerts = false };
            try { GetWindowThreadProcessId(new IntPtr(_app.Hwnd), out _excelPid); } catch { }

            _book = _app.Workbooks.Add();
            while (_book.Worksheets.Count < 2) _book.Worksheets.Add(After: _book.Worksheets[_book.Worksheets.Count]);
            ((Worksheet)_book.Worksheets[1]).Name = "Data";
            ((Worksheet)_book.Worksheets[2]).Name = "Sum";
            ((Worksheet)_book.Worksheets["Data"]).Activate();
            _book.SaveAs(Path.Combine(_root, "book.xlsx"));

            _snapshots = new SnapshotManager(_app, Path.Combine(_root, "snapshots"));
            var actions = new ExcelActionsImpl(_app, new WorkbookAnalyzer(_app), new RangeAnalyzer(), null, null, _snapshots);
            var key = WorkbookIdentity.KeyOf(_book);
            return new ToolDispatcher(actions, _app) { BoundWorkbookKey = () => key, SlowCheckpointThresholdMs = int.MaxValue };
        }

        public void Dispose()
        {
            if (_app != null)
            {
                try
                {
                    foreach (Workbook wb in _app.Workbooks.Cast<Workbook>().ToList())
                    {
                        try { wb.Close(false); } catch { }
                    }
                    _app.Quit();
                }
                catch { }
                try { Marshal.FinalReleaseComObject(_app); } catch { }
                _app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (_excelPid != 0)
            {
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(_excelPid);
                    if (!p.WaitForExit(5000)) p.Kill();
                }
                catch { }
            }
            try { Directory.Delete(_root, true); } catch { }
        }

        /// <summary>Arguments as the sidecar's tool_call delivers them: JsonElement values.</summary>
        private static Dictionary<string, object> Ops(string opsJson)
        {
            using (var doc = JsonDocument.Parse("{\"ops\":" + opsJson + "}"))
                return new Dictionary<string, object> { ["ops"] = doc.RootElement.GetProperty("ops").Clone() };
        }

        private static Dictionary<string, object> Args(params (string Key, object Value)[] pairs)
            => pairs.ToDictionary(p => p.Key, p => p.Value);

        private Worksheet Sheet(string name) => (Worksheet)_book.Worksheets[name];

        private bool HasSheet(string name)
            => _book.Worksheets.Cast<Worksheet>().Any(ws => ws.Name == name);

        private object Value(string sheet, string address) => Sheet(sheet).Range[address].Value2;

        [ExcelFact]
        public void A_batch_writes_every_step_and_keeps_one_snapshot()
        {
            var d = Setup();

            var r = d.Execute("write_cells", Ops(
                "[{\"type\":\"add_sheet\",\"name\":\"汇总\"}," +
                " {\"type\":\"values\",\"address\":\"汇总!A1:B2\",\"values\":[[1,2],[3,4]]}," +
                " {\"type\":\"formula\",\"address\":\"汇总!C1\",\"formula\":\"=A1+B1\"}," +
                " {\"type\":\"text\",\"address\":\"Sum!D1\",\"value\":\"备注\"}]"));

            Assert.True(r.Success, r.Error);
            Assert.NotNull(r.CheckpointId);
            Assert.Equal(4.0, Convert.ToDouble(Value("汇总", "B2")));
            Assert.Equal(3.0, Convert.ToDouble(Value("汇总", "C1")));
            Assert.Equal("备注", Value("Sum", "D1"));
            Assert.Single(_snapshots.ListSnapshots());
        }

        [ExcelFact]
        public void A_later_step_failing_puts_the_book_back_including_a_sheet_the_batch_added()
        {
            var d = Setup();
            Sheet("Data").Range["A1"].Value2 = "原值";
            Sheet("Sum").Range["B2"].Value2 = 42;
            Assert.True(d.Execute("read_range", Args(("address", "Data!A1"))).Success);   // 先读后写
            Assert.True(d.Execute("read_range", Args(("address", "Sum!B2"))).Success);

            var r = d.Execute("write_cells", Ops(
                "[{\"type\":\"text\",\"address\":\"Data!A1\",\"value\":\"新值\"}," +
                " {\"type\":\"add_sheet\",\"name\":\"汇总\"}," +
                " {\"type\":\"values\",\"address\":\"汇总!A1\",\"values\":[[1]]}," +
                " {\"type\":\"values\",\"address\":\"Sum!B2\",\"values\":[[7]]}," +
                " {\"type\":\"formula\",\"address\":\"汇总!B1\",\"formula\":\"=SUM((\"}]"));

            Assert.False(r.Success);
            Assert.Contains("第 5 步", r.Error);
            Assert.Contains("已撤销", r.Error);
            Assert.Equal("原值", Value("Data", "A1"));
            Assert.Equal(42.0, Convert.ToDouble(Value("Sum", "B2")));
            Assert.False(HasSheet("汇总"), "the sheet the batch added should be gone after the rollback");
        }

        [ExcelFact]
        public void Unqualified_addresses_after_add_sheet_land_on_the_new_sheet_and_roll_back_with_it()
        {
            // Worksheets.Add makes the new sheet active, so a later "C1" means the new sheet's C1 —
            // the same as calling add_sheet and write_value one after the other.
            var d = Setup();
            Sheet("Data").Range["C1"].Value2 = "原值";
            Assert.True(d.Execute("read_range", Args(("address", "Data!C1"))).Success);

            var ok = d.Execute("write_cells", Ops(
                "[{\"type\":\"add_sheet\",\"name\":\"汇总\"}," +
                " {\"type\":\"text\",\"address\":\"C1\",\"value\":\"x\"}]"));
            Assert.True(ok.Success, ok.Error);
            Assert.Equal("x", Value("汇总", "C1"));
            Assert.Equal("原值", Value("Data", "C1"));

            Sheet("Data").Activate();
            var failed = d.Execute("write_cells", Ops(
                "[{\"type\":\"add_sheet\",\"name\":\"明细\"}," +
                " {\"type\":\"text\",\"address\":\"C1\",\"value\":\"y\"}," +
                " {\"type\":\"formula\",\"address\":\"C2\",\"formula\":\"=SUM((\"}]"));
            Assert.False(failed.Success);
            Assert.Contains("已撤销", failed.Error);
            Assert.False(HasSheet("明细"));
            Assert.Equal("原值", Value("Data", "C1"));
            Assert.Equal("x", Value("汇总", "C1"));   // the first batch is not touched by the second's rollback
        }

        [ExcelFact]
        public void The_combined_preview_matches_what_the_batch_writes()
        {
            var d = Setup();
            var ws = Sheet("Data");
            ws.Range["A1"].Value2 = "旧";
            ws.Range["B1"].Value2 = 5;
            ws.Range["C3"].Value2 = "不动";
            Assert.True(d.Execute("read_range", Args(("address", "Data!A1:C3"))).Success);

            const string ops =
                "[{\"type\":\"text\",\"address\":\"A1\",\"value\":\"新\"}," +
                " {\"type\":\"values\",\"address\":\"B1\",\"values\":[[1],[2]]}," +
                " {\"type\":\"formula\",\"address\":\"C1\",\"formula\":\"=B1+B2\"}]";
            // The permission request carries nested arguments as JSON text.
            var preview = new PreviewBuilder(new ExcelRangeReader(_app)).Build("write_cells", Args(("ops", ops)));
            Assert.True(preview.Previewable, preview.NotPreviewableReason);

            var reader = new ExcelRangeReader(_app);
            var before = reader.Read("Data", "A1:C3").ToDictionary(c => c.Address, PreviewBuilder.Describe);
            var r = d.Execute("write_cells", Ops(ops));
            Assert.True(r.Success, r.Error);
            var after = reader.Read("Data", "A1:C3").ToDictionary(c => c.Address, PreviewBuilder.Describe);

            var listed = preview.Changes.ToDictionary(c => c.Address);
            var picture = "\n预览：" + string.Join("，", preview.Changes.Select(c => c.Address + "=" + c.After)) +
                          "\n实际：" + string.Join("，", after.Select(kv => kv.Key + "=" + kv.Value));
            foreach (var address in before.Keys)
            {
                if (listed.ContainsKey(address))
                    Assert.True(before[address] != after[address], $"{address} was listed as changing but did not" + picture);
                else
                    Assert.True(before[address] == after[address], $"{address} changed but the preview did not list it" + picture);
            }
            Assert.Equal(new[] { "A1", "B1", "B2", "C1" }, listed.Keys.OrderBy(k => k).ToArray());
        }
    }
}
