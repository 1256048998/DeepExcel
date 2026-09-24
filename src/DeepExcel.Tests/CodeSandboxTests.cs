using DeepExcel.AddIn.Security;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// execute_python 不能经 COM 驱动 Excel：与加载项所在的 Excel 主进程并发 COM
    /// 调用会让 Excel 崩溃退出。拦截必须带上指向正确工具的提示，否则模型换个写法再试。
    /// </summary>
    public class CodeSandboxTests
    {
        [Theory]
        [InlineData("import win32com.client\nxl = win32com.client.Dispatch('Excel.Application')")]
        [InlineData("from win32com.client import Dispatch")]
        [InlineData("from win32com.client.gencache import EnsureDispatch")]
        [InlineData("import comtypes.client")]
        [InlineData("import pythoncom")]
        [InlineData("import win32api")]
        [InlineData("import xlwings as xw")]
        [InlineData("from xlwings import Book")]
        public void ValidatePython_BlocksComAutomation(string code)
        {
            var error = CodeSandbox.ValidatePython(code);

            Assert.NotNull(error);
            Assert.Contains("COM", error);
            Assert.Contains("execute_vba", error);
        }

        [Theory]
        [InlineData("total = sum([1, 2, 3])")]
        [InlineData("import re\nre.sub(r'\\s+', ' ', 'a  b')")]
        [InlineData("import math\nmath.sqrt(2)")]
        [InlineData("import json\njson.dumps({'a': 1})")]
        // 标识符里恰好含 com/win 的普通代码不能被误伤
        [InlineData("commission = 0.05\nwinner = 'A'")]
        public void ValidatePython_AllowsPureComputation(string code)
        {
            Assert.Null(CodeSandbox.ValidatePython(code));
        }

        [Fact]
        public void ValidatePython_OpenpyxlKeepsItsOwnHint()
        {
            // COM 提示不能抢走 openpyxl 的"文件被锁定"提示
            var error = CodeSandbox.ValidatePython("import openpyxl");

            Assert.NotNull(error);
            Assert.Contains("锁定", error);
        }
    }
}
