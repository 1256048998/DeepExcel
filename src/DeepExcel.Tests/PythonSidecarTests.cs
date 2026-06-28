using System.Windows.Forms;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;

namespace DeepExcel.Tests
{
    public class PythonSidecarTests
    {
        [Fact]
        public void PythonSidecar_CanBeConstructedWithDeps()
        {
            var excelActions = new FakeExcelActions();
            var uiControl = new UserControl();  // 真实 Control 用于 Invoke
            var sidecar = new PythonSidecar(excelActions, null, uiControl);
            Assert.NotNull(sidecar);
            sidecar.Dispose();
        }

        [Fact]
        public void PythonSidecar_GetSidecarPath_ReturnsValidPath()
        {
            // 静态方法测试，不依赖进程
            var path = PythonSidecar.GetSidecarPath();
            Assert.Contains("sidecar.py", path);
        }

        [Fact]
        public void PythonSidecar_GetPythonPath_ReturnsNonEmpty()
        {
            var py = PythonSidecar.GetPythonPath();
            Assert.False(string.IsNullOrEmpty(py));
        }

        /// <summary>
        /// 手写 IExcelActions mock —— 不引入 Moq 依赖。
        /// 方法签名严格匹配 IExcelActions 接口。
        /// </summary>
        private class FakeExcelActions : IExcelActions
        {
            public object GetSelection() => null;
            public object ReadRange(string address) => new { };
            public object ReadWorkbook() => new { };
            public object ReadWorksheet(string name) => new { };
            public ToolResult ExecuteVBA(string code, string macroName = null) => new ToolResult { Success = true };
            public ToolResult ExecutePython(string code) => new ToolResult { Success = true };
            public ToolResult WriteFormula(string address, string formula) => new ToolResult { Success = true };
            public ToolResult WriteValue(string address, object value) => new ToolResult { Success = true };
            public string CreateSnapshot() => "snap-1";
            public bool Rollback(string snapshotId) => true;
        }
    }
}
