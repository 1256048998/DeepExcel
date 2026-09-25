using DeepExcel.AddIn.Executor;
using Xunit;

namespace DeepExcel.Tests
{
    public class VBAExecutorTests
    {
        [Fact]
        public void TryPrepareCode_WrapsStatementsInStableEntryPoint()
        {
            var ok = VBAExecutor.TryPrepareCode("Range(\"A1\").Value = 1", null,
                out var code, out var entry, out var error);

            Assert.True(ok, error);
            Assert.Equal("DeepExcel_TempMacro", entry);
            Assert.Contains("Public Sub DeepExcel_TempMacro()", code);
        }

        [Fact]
        public void TryPrepareCode_UsesActualParameterlessSubName()
        {
            var ok = VBAExecutor.TryPrepareCode("Sub CreateGanttChart()\nEnd Sub", null,
                out var code, out var entry, out var error);

            Assert.True(ok, error);
            Assert.Equal("CreateGanttChart", entry);
            Assert.Contains("Sub CreateGanttChart()", code);
        }

        [Fact]
        public void TryPrepareCode_StripsMarkdownFence()
        {
            var ok = VBAExecutor.TryPrepareCode("```vba\nSub Demo()\nEnd Sub\n```", null,
                out var code, out var entry, out var error);

            Assert.True(ok, error);
            Assert.Equal("Demo", entry);
            Assert.DoesNotContain("```", code);
        }

        [Fact]
        public void TryPrepareCode_RenamesNonAsciiEntryPoint()
        {
            var ok = VBAExecutor.TryPrepareCode("Sub 创建图表()\nRange(\"A1\").Value = \"完成\"\nEnd Sub", null,
                out var code, out var entry, out var error);

            Assert.True(ok, error);
            Assert.Equal("DeepExcel_TempMacro", entry);
            Assert.Contains("Sub DeepExcel_TempMacro()", code);
        }

        [Fact]
        public void TryPrepareCode_RejectsNonAsciiVariableNames()
        {
            var ok = VBAExecutor.TryPrepareCode("Sub Demo()\nDim 数量 As Long\nEnd Sub", null,
                out _, out _, out var error);

            Assert.False(ok);
            Assert.Contains("变量名必须使用英文", error);
        }

        [Fact]
        public void TryPrepareCode_RejectsParameterizedEntryPoint()
        {
            var ok = VBAExecutor.TryPrepareCode("Sub Demo(value As Long)\nEnd Sub", null,
                out _, out _, out var error);

            Assert.False(ok);
            Assert.Contains("不带参数", error);
        }

        [Fact]
        public void TryPrepareCode_RejectsInteractiveDialogsButNotDialogText()
        {
            var blocked = VBAExecutor.TryPrepareCode("Sub Demo()\nMsgBox \"done\"\nEnd Sub", null,
                out _, out _, out var blockedError);
            var allowed = VBAExecutor.TryPrepareCode("Sub Demo()\nRange(\"A1\").Value = \"MsgBox\"\nEnd Sub", null,
                out _, out _, out var allowedError);

            Assert.False(blocked);
            Assert.Contains("交互式弹窗", blockedError);
            Assert.True(allowed, allowedError);
        }

        [Fact]
        public void BuildInvocationWrapper_CapturesVbaRuntimeErrors()
        {
            var wrapper = VBAExecutor.BuildInvocationWrapper("Demo", "InvokeDemo");

            Assert.Contains("On Error GoTo DeepExcel_Error", wrapper);
            Assert.Contains("Call Demo", wrapper);
            Assert.Contains("Err.Description", wrapper);
        }
    }
}
