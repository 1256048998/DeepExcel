using System;
using System.Linq;
using System.Text;
using DeepExcel.AddIn.Executor;
using DeepExcel.AddIn.Security;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>VBA 编译前静态检查：模型常犯、会让 Excel 弹窗或卡死的错误。</summary>
    public class VbaStaticCheckerTests
    {
        private static string[] Errors(string code) =>
            VbaStaticChecker.Check(code).Where(i => i.Level == VbaIssueLevel.Error).Select(i => i.ToString()).ToArray();

        [Fact]
        public void Clean_code_has_no_issues()
        {
            var code = "Option Explicit\n\nSub Demo()\n    Dim i As Long\n    For i = 1 To 10\n        Cells(i, 1).Value = \"第\" & i & \"行\" ' 注释里有 \" 也没关系\n    Next i\n    Range(\"B1\").Formula = \"=IF(A1=\"\"是\"\",1,0)\"\nEnd Sub";
            Assert.Empty(VbaStaticChecker.Check(code));
        }

        [Fact]
        public void Backslash_escaped_quotes_in_a_formula_are_rejected()
        {
            var errors = Errors("Range(\"B1\").Formula = \"=IF(A1=\\\"x\\\",1,0)\"");
            Assert.Contains(errors, e => e.Contains("反斜杠转义"));
            // 路径里的反斜杠结尾不是转义
            Assert.Empty(Errors("Dim p As String\np = \"C:\\data\\\" & \"a.xlsx\""));
        }

        [Fact]
        public void Unbalanced_quotes_are_rejected_with_the_line_number()
        {
            var errors = Errors("Sub Demo()\n    Range(\"A1\").Value = \"abc\nEnd Sub");
            Assert.Contains(errors, e => e.StartsWith("第 2 行") && e.Contains("引号不配对"));
            Assert.Empty(Errors("Sub Demo()\n    ' 他说\"你好\n    Range(\"A1\").Value = 1\nEnd Sub"));  // 注释里的单个引号
        }

        [Fact]
        public void Too_many_line_continuations_are_rejected()
        {
            var lines = string.Join("\n", Enumerable.Range(1, 26).Select(i => $"    x = x + {i} _"));
            var errors = Errors("Sub Demo()\n    Dim x As Long\n" + lines + "\n    + 0\nEnd Sub");
            Assert.Contains(errors, e => e.Contains("续行"));
            var ok = string.Join("\n", Enumerable.Range(1, 10).Select(i => $"    x = x + {i} _"));
            Assert.Empty(Errors("Sub Demo()\n    Dim x As Long\n" + ok + "\n    + 0\nEnd Sub"));
        }

        [Fact]
        public void Option_after_a_procedure_and_statements_outside_procedures_are_rejected()
        {
            Assert.Contains(Errors("Sub A()\nEnd Sub\nOption Explicit"), e => e.Contains("Option"));
            var outside = Errors("Dim total As Long\nRange(\"A1\").Value = 1\nSub A()\n    total = 1\nEnd Sub");
            Assert.Contains(outside, e => e.StartsWith("第 2 行") && e.Contains("Sub / Function 外面"));
            Assert.DoesNotContain(outside, e => e.StartsWith("第 1 行"));
            // 没有任何 Sub 的代码会被包进一个 Sub，不算
            Assert.Empty(Errors("Range(\"A1\").Value = 1"));
            // 模块级续行、Type 块
            Assert.Empty(Errors("Private Const X As Long = 1 + _\n    2\nPrivate Type P\n    Name As String\nEnd Type\nSub A()\nEnd Sub"));
        }

        [Fact]
        public void Forms_dialogs_and_self_modification_are_rejected()
        {
            Assert.Contains(Errors("Sub A()\n    UserForm1.Show\nEnd Sub"), e => e.Contains(".Show"));
            Assert.Contains(Errors("Sub A()\n    Application.Dialogs(1).Show\nEnd Sub"), e => e.Contains(".Show"));
            Assert.Contains(Errors("Sub A()\n    ThisWorkbook.VBProject.VBComponents.Add 1\nEnd Sub"), e => e.Contains("VBA 工程"));
            Assert.Empty(Errors("Sub A()\n    Range(\"A1\").Value = \".Show VBProject\"\nEnd Sub"));
        }

        [Fact]
        public void Loops_without_an_exit_are_rejected()
        {
            Assert.Contains(Errors("Sub A()\n    Do\n        x = x + 1\n    Loop\nEnd Sub"), e => e.Contains("循环没有出口"));
            Assert.Contains(Errors("Sub A()\n    Do While True\n        x = x + 1\n    Loop\nEnd Sub"), e => e.Contains("循环没有出口"));
            Assert.Contains(Errors("Sub A()\n    While True\n        x = 1\n    Wend\nEnd Sub"), e => e.Contains("循环没有出口"));
            // End If 不是出口
            Assert.Contains(Errors("Sub A()\n    Do\n        If x > 1 Then\n            x = 0\n        End If\n    Loop\nEnd Sub"), e => e.Contains("循环没有出口"));
            // 有出口的
            Assert.Empty(Errors("Sub A()\n    Do\n        x = x + 1\n        If x > 9 Then Exit Do\n    Loop\nEnd Sub"));
            Assert.Empty(Errors("Sub A()\n    Do\n        x = x + 1\n    Loop Until x > 9\nEnd Sub"));
            Assert.Empty(Errors("Sub A()\n    Do While x < 9\n        x = x + 1\n    Loop\nEnd Sub"));
            Assert.Empty(Errors("Sub A()\n    Do\n        If x > 9 Then Exit Sub\n        x = x + 1\n    Loop\nEnd Sub"));
            // Exit Do 只跳出内层：外层仍然是死循环
            Assert.Contains(Errors("Sub A()\n    Do\n        Do\n            Exit Do\n        Loop\n    Loop\nEnd Sub"), e => e.StartsWith("第 2 行"));
            // On Error GoTo 不是出口
            Assert.Contains(Errors("Sub A()\n    Do\n        On Error GoTo Oops\n    Loop\nOops:\nEnd Sub"), e => e.Contains("循环没有出口"));
        }

        [Fact]
        public void Resume_next_without_reset_is_a_warning_not_an_error()
        {
            var issues = VbaStaticChecker.Check("Sub A()\n    On Error Resume Next\n    x = 1 / 0\nEnd Sub");
            Assert.Single(issues);
            Assert.Equal(VbaIssueLevel.Warning, issues[0].Level);
            Assert.Empty(VbaStaticChecker.Check("Sub A()\n    On Error Resume Next\n    x = 1 / 0\n    On Error GoTo 0\nEnd Sub"));
        }
    }

    /// <summary>VBA 字符串代码页安全。</summary>
    public class VbaEncodingTests
    {
        private static readonly Encoding Western = Encoding.GetEncoding(1252);
        private static readonly Encoding Chinese = Encoding.GetEncoding(936);

        [Fact]
        public void Only_characters_the_code_page_cannot_represent_are_rewritten()
        {
            var code = "Range(\"A1\").Value = \"销售完成\"";
            var western = VbaEncoding.Encode(code, Western);
            Assert.DoesNotContain("销售完成", western.Code);
            Assert.Contains("DeepExcelU(\"9500552E5B8C6210\")", western.Code);
            Assert.True(western.UsesHelper);

            var chinese = VbaEncoding.Encode(code, Chinese);
            Assert.Equal(code, chinese.Code);  // 中文 Windows 上原样保留
            Assert.False(chinese.UsesHelper);
            Assert.Equal("x = \"café\"", VbaEncoding.Encode("x = \"café\"", Western).Code);  // 1252 能表示 é
        }

        [Fact]
        public void Short_runs_use_chrw_and_mixed_text_keeps_its_ascii()
        {
            var result = VbaEncoding.Encode("x = \"第\" & i & \"行 of 10\"", Western);
            Assert.Equal("x = ChrW(31532) & i & (ChrW(34892) & \" of 10\")", result.Code);
            Assert.False(result.UsesHelper);
        }

        [Fact]
        public void Comments_are_left_alone_even_with_a_lone_quote()
        {
            var code = "' 他说\"你好\nRange(\"A1\").Value = 1 ' 设置\"标题";
            Assert.Equal(code.Replace("\n", "\r\n"), VbaEncoding.Encode(code, Western).Code);
        }

        [Fact]
        public void Long_strings_wrap_at_the_joins_instead_of_overflowing_the_line()
        {
            var text = string.Concat(Enumerable.Repeat("销售数据汇总", 60)) + " end";  // 360 个汉字
            var result = VbaEncoding.Encode($"Range(\"A1\").Value = \"{text}\"", Western);
            var lines = result.Code.Split(new[] { "\r\n" }, StringSplitOptions.None);
            Assert.True(lines.Length > 1);
            Assert.All(lines, l => Assert.True(l.Length <= VbaEncoding.MaxLineLength, $"{l.Length} chars"));
            Assert.All(lines.Take(lines.Length - 1), l => Assert.EndsWith(" _", l));
            Assert.False(result.TooLong);
        }

        [Fact]
        public void Consts_become_assignments_or_property_gets()
        {
            var inside = VbaEncoding.Encode("Sub A()\n    Const TITLE As String = \"销售报表\"\n    Range(\"A1\").Value = TITLE\nEnd Sub", Western);
            Assert.Contains("    Dim TITLE As String: TITLE = DeepExcelU(", inside.Code);

            var module = VbaEncoding.Encode("Public Const TITLE = \"销售报表\"\nSub A()\nEnd Sub", Western);
            Assert.Contains("Public Property Get TITLE() As String\r\n    TITLE = DeepExcelU(", module.Code);
            Assert.Contains("End Property", module.Code);

            // 不需要改写的 Const 不动
            Assert.Equal("Const N = 10", VbaEncoding.Encode("Const N = 10", Western).Code);
        }
    }

    /// <summary>真 Excel（本机代码页 1252）：改写后的中文字符串在 VBA 里取回来一字不差。默认跳过。</summary>
    public class VbaEncodingExcelTests : ExcelTestHost
    {
        [ExcelFact]
        public void Chinese_strings_survive_an_ansi_code_page_that_cannot_represent_them()
        {
            NewBook(out var wb);
            var longText = string.Concat(Enumerable.Repeat("应收账款明细汇总表", 40));
            var code = string.Join("\n",
                "Public Const HEADER = \"客户名称\"",
                "Sub Fill()",
                "    Const NOTE As String = \"已核对 ✓\"",
                "    Range(\"A1\").Value = HEADER",
                "    Range(\"A2\").Value = NOTE",
                "    Range(\"A3\").Value = \"" + longText + "\"",
                "    Range(\"A4\").Formula = \"=IF(A1=\"\"客户名称\"\",\"\"是\"\",\"\"否\"\")\"",
                "    ' 注释里有中文和一个单独的引号\"",
                "End Sub");
            var result = new VBAExecutor(App).Execute(code);
            Assert.True(result.Success, result.Error + " / " + result.Suggestion);
            var ws = (Worksheet)wb.Worksheets[1];
            Assert.Equal("客户名称", ws.Range["A1"].Value2);
            Assert.Equal("已核对 ✓", ws.Range["A2"].Value2);
            Assert.Equal(longText, ws.Range["A3"].Value2);
            Assert.Equal("是", ws.Range["A4"].Value2);
        }

        [ExcelFact]
        public void Statically_broken_code_is_returned_without_touching_the_workbook()
        {
            NewBook(out var wb);
            var result = new VBAExecutor(App).Execute("Sub A()\n    Range(\"A1\").Formula = \"=IF(B1=\\\"x\\\",1,0)\"\nEnd Sub");
            Assert.False(result.Success);
            Assert.Contains("反斜杠转义", result.Error);
            Assert.Null(((Worksheet)wb.Worksheets[1]).Range["A1"].Value2);
        }
    }
}
