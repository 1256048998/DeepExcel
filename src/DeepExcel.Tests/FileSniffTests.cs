using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using DeepExcel.AddIn.Perception;
using DeepExcel.AddIn.Sidecar;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>附件文件头嗅探：企业透明加密的密文、设了打开密码的文件在读之前就说清楚。</summary>
    public class FileSniffTests
    {
        private static readonly byte[] ZipHead = { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00 };
        private static readonly byte[] OleHead = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00 };

        private static byte[] Noise(int length, int seed)
        {
            var bytes = new byte[length];
            new Random(seed).NextBytes(bytes);
            bytes[0] = 0x9C;  // 不是 PK / D0
            return bytes;
        }

        [Theory]
        [InlineData(".xlsx")]
        [InlineData(".xlsm")]
        [InlineData(".XLSB")]
        public void Ooxml_files_must_be_zip(string ext)
        {
            Assert.Equal(SniffVerdict.Ok, FileSniff.Classify(ZipHead, ext).Verdict);
            Assert.Equal(SniffVerdict.PasswordProtected, FileSniff.Classify(OleHead, ext).Verdict);
            var encrypted = FileSniff.Classify(Noise(512, 1), ext);
            Assert.Equal(SniffVerdict.LikelyEncrypted, encrypted.Verdict);
            Assert.Contains("企业加密软件", encrypted.Message);
            Assert.Contains("在 Excel 中打开", encrypted.Suggestion);
        }

        [Fact]
        public void Xls_accepts_ole_zip_and_the_html_or_text_that_business_systems_export()
        {
            Assert.True(FileSniff.Classify(OleHead, ".xls").Ok);
            Assert.True(FileSniff.Classify(ZipHead, ".xls").Ok);
            Assert.True(FileSniff.Classify(Encoding.UTF8.GetBytes("<html><table><tr><td>客户</td>"), ".xls").Ok);
            Assert.True(FileSniff.Classify(Encoding.GetEncoding(936).GetBytes("客户\t金额\r\n甲\t10\r\n"), ".xls").Ok);
            Assert.Equal(SniffVerdict.LikelyEncrypted, FileSniff.Classify(Noise(512, 2), ".xls").Verdict);
            Assert.True(FileSniff.Classify(Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><Workbook>"), ".xlsx").Ok);
        }

        [Fact]
        public void Text_attachments_must_look_like_text()
        {
            Assert.True(FileSniff.Classify(Encoding.UTF8.GetBytes("日期,客户,金额\n2024-01-01,甲,10\n"), ".csv").Ok);
            Assert.True(FileSniff.Classify(new byte[] { 0xFF, 0xFE, 0x65, 0x51, 0x00, 0x00 }, ".txt").Ok);  // UTF-16
            Assert.Equal(SniffVerdict.LikelyEncrypted, FileSniff.Classify(Noise(512, 3), ".csv").Verdict);
            Assert.True(FileSniff.Classify(Noise(512, 4), ".png").Ok);  // 不认识的类型不下结论
        }

        [Fact]
        public void Empty_files_are_reported()
        {
            Assert.Equal(SniffVerdict.Empty, FileSniff.Classify(new byte[0], ".xlsx").Verdict);
        }

        [Fact]
        public void Read_attachment_refuses_an_encrypted_file_before_opening_it()
        {
            var path = Path.Combine(Path.GetTempPath(), "deepexcel_sniff_" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.WriteAllBytes(path, Noise(2048, 5));
            try
            {
                var d = new ToolDispatcher(new FakeExcelActions(), null)
                {
                    Attachments = new Dictionary<string, string> { ["客户.xlsx"] = path },
                };
                var r = d.Execute("read_attachment", new Dictionary<string, object> { ["file_name"] = "客户.xlsx" });
                Assert.False(r.Success);
                Assert.Contains("企业加密软件", r.Error);
                Assert.Contains("在 Excel 中打开", r.Suggestion);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }

    /// <summary>真 Excel：设了打开密码的附件不能弹出密码框卡住 Excel。默认跳过。</summary>
    public class AttachmentPasswordExcelTests : ExcelTestHost
    {
        [ExcelFact]
        public void Password_protected_attachments_fail_fast_and_plain_ones_still_open()
        {
            NewBook(out var wb);
            ((Worksheet)wb.Worksheets[1]).Range["A1"].Value2 = "secret";
            var dir = Path.Combine(Path.GetTempPath(), "deepexcel_pw_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            var locked = Path.Combine(dir, "locked.xls");
            var lockedX = Path.Combine(dir, "locked.xlsx");
            var plain = Path.Combine(dir, "plain.xlsx");
            try
            {
                wb.SaveCopyAs(plain);
                wb.SaveAs(lockedX, XlFileFormat.xlOpenXMLWorkbook, Password: "abc");
                wb.SaveAs(locked, XlFileFormat.xlExcel8, Password: "abc");
                // 已经开着的文件再 Open 会直接返回那本工作簿、不问密码，所以先关掉
                wb.Close(false);

                // 加了打开密码的 xlsx 是 OLE 容器：嗅探就拦下
                Assert.Equal(SniffVerdict.PasswordProtected, FileSniff.Sniff(lockedX).Verdict);
                Assert.True(FileSniff.Sniff(locked).Ok);  // 旧版 xls 本来就是 OLE，交给打开时处理
                Assert.True(FileSniff.Sniff(plain).Ok);

                var d = new ToolDispatcher(new FakeExcelActions(), App)
                {
                    Attachments = new Dictionary<string, string>
                    {
                        ["locked.xls"] = locked, ["locked.xlsx"] = lockedX, ["plain.xlsx"] = plain,
                    },
                };
                var clock = Stopwatch.StartNew();
                var r = d.Execute("read_attachment", new Dictionary<string, object> { ["file_name"] = "locked.xls" });
                Assert.False(r.Success);
                Assert.Contains("打开密码", r.Error);
                Assert.True(clock.Elapsed < TimeSpan.FromSeconds(20), "应当立刻报错，而不是等密码框");

                r = d.Execute("read_attachment", new Dictionary<string, object> { ["file_name"] = "locked.xlsx" });
                Assert.False(r.Success);
                Assert.Contains("打开密码", r.Error);

                r = d.Execute("read_attachment", new Dictionary<string, object> { ["file_name"] = "plain.xlsx" });
                Assert.True(r.Success, r.Error);
                Assert.Contains("secret", PythonSidecar.BuildToolResultJson("c", true, r.Data, null, null, null, null, null));
            }
            finally
            {
                foreach (Workbook open in App.Workbooks.Cast<Workbook>().ToList())
                {
                    try { open.Close(false); } catch { }
                }
                try { Directory.Delete(dir, true); } catch { }
            }
        }
    }
}
