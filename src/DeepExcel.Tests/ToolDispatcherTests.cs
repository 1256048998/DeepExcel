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
        public void Execute_Echo_ReturnsInputText()
        {
            var fake = new FakeExcelActions();
            var dispatcher = new ToolDispatcher(fake, null);
            var args = new Dictionary<string, object> { { "text", "hello" } };

            var result = dispatcher.Execute("echo", args);

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
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
            Assert.True(fake.GetSelectionCalls >= 2);
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
                RollbackFn = (id) => id == "snap-1",
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

        /// <summary>
        /// 手写可配置 IExcelActions mock —— 不引入 Moq 依赖（环境无 Moq 包）。
        /// 方法签名严格匹配 IExcelActions 接口；Func 字段可被测试内联配置，
        /// 未配置的方法走默认实现（success / 空对象），不会抛异常。
        /// </summary>
        private class FakeExcelActions : IExcelActions
        {
            public Func<string, object> ReadRangeFn { get; set; } = _ => new { };
            public Func<string, string, ToolResult> WriteFormulaFn { get; set; } = (a, f) => new ToolResult { Success = true };
            public Func<object> GetSelectionFn { get; set; } = () => null;
            public Func<object> ReadWorkbookFn { get; set; } = () => new { };
            public Func<string, object> ReadWorksheetFn { get; set; } = _ => new { };
            public Func<string, string, ToolResult> ExecuteVBAFn { get; set; } = (c, m) => new ToolResult { Success = true };
            public Func<string, ToolResult> ExecutePythonFn { get; set; } = _ => new ToolResult { Success = true };
            public Func<string, object, ToolResult> WriteValueFn { get; set; } = (a, v) => new ToolResult { Success = true };
            public Func<string> CreateSnapshotFn { get; set; } = () => "snap-1";
            public Func<string, bool> RollbackFn { get; set; } = _ => true;

            public List<(string, string)> WriteFormulaCalls { get; } = new List<(string, string)>();
            public int ReadWorkbookCalls { get; private set; }
            public int GetSelectionCalls { get; private set; }
            public List<(string, string)> ExecuteVBACalls { get; } = new List<(string, string)>();
            public List<string> ExecutePythonCalls { get; } = new List<string>();
            public int CreateSnapshotCalls { get; private set; }
            public List<string> RollbackCalls { get; } = new List<string>();

            public object GetSelection()
            {
                GetSelectionCalls++;
                return GetSelectionFn();
            }

            public object ReadRange(string address) => ReadRangeFn(address);

            public object ReadWorkbook()
            {
                ReadWorkbookCalls++;
                return ReadWorkbookFn();
            }

            public object ReadWorksheet(string name) => ReadWorksheetFn(name);

            public ToolResult ExecuteVBA(string code, string macroName = null)
            {
                ExecuteVBACalls.Add((code, macroName));
                return ExecuteVBAFn(code, macroName);
            }

            public ToolResult ExecutePython(string code)
            {
                ExecutePythonCalls.Add(code);
                return ExecutePythonFn(code);
            }

            public ToolResult WriteFormula(string address, string formula)
            {
                WriteFormulaCalls.Add((address, formula));
                return WriteFormulaFn(address, formula);
            }

            public ToolResult WriteValue(string address, object value) => WriteValueFn(address, value);

            public string CreateSnapshot()
            {
                CreateSnapshotCalls++;
                return CreateSnapshotFn();
            }

            public bool Rollback(string snapshotId)
            {
                RollbackCalls.Add(snapshotId);
                return RollbackFn(snapshotId);
            }
        }
    }
}
