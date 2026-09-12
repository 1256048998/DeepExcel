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

        // ★ IExcelActions mock 已抽到共享的 FakeExcelActions.cs，
        // 避免接口增加方法时每个测试文件都要各自补一遍 stub
    }
}
