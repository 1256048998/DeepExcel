using System;
using System.Collections.Generic;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    public class ToolDispatcherTests
    {
        [Fact]
        public void Execute_WriteFormula_CallsExcelActions()
        {
            var fake = new FakeExcelActions
            {
                WriteFormulaFn = (a, f) => new ToolResult { Name = "write_formula", Success = true },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var args = new Dictionary<string, object>
            {
                { "address", "A1" },
                { "formula", "=SUM(B:B)" },
            };

            var result = dispatcher.Execute("write_formula", args);

            Assert.True(result.Success);
            Assert.Equal(1, fake.WriteFormulaCalls.Count);
            Assert.Equal("A1", fake.WriteFormulaCalls[0].Item1);
            Assert.Equal("=SUM(B:B)", fake.WriteFormulaCalls[0].Item2);
        }

        [Fact]
        public void Execute_ReadRange_ReturnsDataAndSuggestion()
        {
            var fake = new FakeExcelActions
            {
                ReadRangeFn = (a) => new { cells = new[] { "苹果", "香蕉" }, data_type = "text" },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var args = new Dictionary<string, object> { { "address", "A:A" } };

            var result = dispatcher.Execute("read_range", args);

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
            Assert.NotNull(result.Context);
        }

        [Fact]
        public void Execute_UnknownTool_ReturnsError()
        {
            var fake = new FakeExcelActions();
            var dispatcher = new ToolDispatcher(fake, null);
            var result = dispatcher.Execute("nonexistent", new Dictionary<string, object>());
            Assert.False(result.Success);
            Assert.Contains("未知工具", result.Error);
        }

        [Fact]
        public void Execute_ReadWorkbook_ReturnsData()
        {
            var fake = new FakeExcelActions
            {
                ReadWorkbookFn = () => new { name = "Book1" },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var result = dispatcher.Execute("read_workbook", new Dictionary<string, object>());

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
            Assert.NotNull(result.Context);
            Assert.True(fake.ReadWorkbookCalls >= 2);
        }

        [Fact]
        public void Execute_ReadSelection_ReturnsData()
        {
            var fake = new FakeExcelActions
            {
                GetSelectionFn = () => new { address = "A1:B10" },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var result = dispatcher.Execute("read_selection", new Dictionary<string, object>());

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
            Assert.NotNull(result.Context);
            // ★ 只会调用一次 IExcelActions.GetSelection()：
            // BuildExcelSnapshot 里的 selection 走的是 COM Application.Selection（本测试传 null），
            // 不再经过 IExcelActions，所以这里不能像 read_workbook 那样断言 >= 2。
            Assert.Equal(1, fake.GetSelectionCalls);
        }

        [Fact]
        public void Execute_ExecuteVBA_CallsExcelActions()
        {
            var fake = new FakeExcelActions
            {
                ExecuteVBAFn = (c, m) => new ToolResult { Name = "execute_vba", Success = true, Data = "ok" },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var args = new Dictionary<string, object> { { "code", "MsgBox \"hi\"" } };

            var result = dispatcher.Execute("execute_vba", args);

            Assert.True(result.Success);
            Assert.Equal(1, fake.ExecuteVBACalls.Count);
            Assert.Equal("MsgBox \"hi\"", fake.ExecuteVBACalls[0].Item1);
        }

        [Fact]
        public void Execute_ExecutePython_CallsExcelActions()
        {
            var fake = new FakeExcelActions
            {
                ExecutePythonFn = (c) => new ToolResult { Name = "execute_python", Success = true, Data = "done" },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var args = new Dictionary<string, object> { { "code", "print(1)" } };

            var result = dispatcher.Execute("execute_python", args);

            Assert.True(result.Success);
            Assert.Equal(1, fake.ExecutePythonCalls.Count);
            Assert.Equal("print(1)", fake.ExecutePythonCalls[0]);
        }

        [Fact]
        public void Execute_CreateSnapshot_ReturnsSnapshotId()
        {
            var fake = new FakeExcelActions
            {
                CreateSnapshotFn = () => "snap-123",
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var result = dispatcher.Execute("create_snapshot", new Dictionary<string, object>());

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
            Assert.Equal("snap-123", (result.Data as dynamic)?.snapshot_id);
            Assert.Equal(1, fake.CreateSnapshotCalls);
        }

        [Fact]
        public void Execute_Rollback_CallsExcelActions()
        {
            var fake = new FakeExcelActions
            {
                RollbackFn = (id) => new DeepExcel.AddIn.Executor.RollbackResult { Success = id == "snap-1" },
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var args = new Dictionary<string, object> { { "snapshot_id", "snap-1" } };

            var result = dispatcher.Execute("rollback", args);

            Assert.True(result.Success);
            Assert.Equal(1, fake.RollbackCalls.Count);
            Assert.Equal("snap-1", fake.RollbackCalls[0]);
        }

        [Fact]
        public void Execute_CreateSnapshot_EmptyId_ReturnsFailure()
        {
            var fake = new FakeExcelActions
            {
                CreateSnapshotFn = () => null,
            };

            var dispatcher = new ToolDispatcher(fake, null);
            var result = dispatcher.Execute("create_snapshot", new Dictionary<string, object>());

            Assert.False(result.Success);
        }

        // ★ IExcelActions mock 已抽到共享的 FakeExcelActions.cs
    }
}
