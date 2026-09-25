using System;
using System.IO;

namespace DeepExcel.AddIn.Perception
{
    public enum SniffVerdict
    {
        Ok,
        Empty,
        /// <summary>OOXML 扩展名却是 OLE 容器：设了打开密码（加密的 OOXML 就装在 OLE 里），或改了扩展名的旧版 xls</summary>
        PasswordProtected,
        /// <summary>文件头既不是 ZIP / OLE，也不像文本：多半是企业透明加密软件处理过的密文</summary>
        LikelyEncrypted,
    }

    public sealed class SniffResult
    {
        public SniffVerdict Verdict { get; set; }
        public string Message { get; set; }
        public string Suggestion { get; set; }
        public bool Ok => Verdict == SniffVerdict.Ok;
    }

    /// <summary>
    /// 附件读取前嗅探文件头。企业透明加密（按进程解密）下，面板上传的文件落盘后是密文：
    /// 以前会一路交给 Excel 打开，报一个看不懂的「文件格式无效」，或者文本附件读出一堆乱码
    /// 交给模型去猜。现在在读之前就说清楚。
    /// </summary>
    public static class FileSniff
    {
        public const int HeadLength = 512;

        private static readonly byte[] Zip = { 0x50, 0x4B, 0x03, 0x04 };
        private static readonly byte[] ZipEmpty = { 0x50, 0x4B, 0x05, 0x06 };
        private static readonly byte[] Ole = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        private const string EncryptedMessage = "文件可能被企业加密软件保护，DeepExcel 读到的是加密后的内容";
        private const string EncryptedSuggestion = "请在 Excel 中打开这个文件，把需要的内容复制到当前工作簿（或另存一份）后再处理";

        public static SniffResult Sniff(string path)
        {
            byte[] head;
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    head = new byte[Math.Min(HeadLength, stream.Length)];
                    var read = 0;
                    while (read < head.Length)
                    {
                        var n = stream.Read(head, read, head.Length - read);
                        if (n <= 0) break;
                        read += n;
                    }
                    if (read < head.Length) Array.Resize(ref head, read);
                }
            }
            catch (Exception)
            {
                // 读不了文件头就不下结论，交给后面的读取去报真正的错误
                return new SniffResult { Verdict = SniffVerdict.Ok };
            }
            return Classify(head, Path.GetExtension(path));
        }

        public static SniffResult Classify(byte[] head, string extension)
        {
            var ext = (extension ?? "").Trim().ToLowerInvariant();
            if (head == null || head.Length == 0)
            {
                return new SniffResult { Verdict = SniffVerdict.Empty, Message = "文件是空的（0 字节）", Suggestion = "请确认上传的是正确的文件" };
            }

            var zip = StartsWith(head, Zip) || StartsWith(head, ZipEmpty);
            var ole = StartsWith(head, Ole);
            var text = LooksLikeText(head);

            switch (ext)
            {
                case ".xlsx":
                case ".xlsm":
                case ".xlsb":
                    if (zip) return OkResult();
                    if (ole)
                    {
                        return new SniffResult
                        {
                            Verdict = SniffVerdict.PasswordProtected,
                            Message = "文件设置了打开密码（也可能是改了扩展名的旧版 .xls）",
                            Suggestion = "请在 Excel 中输入密码打开，另存一份没有打开密码的副本再上传",
                        };
                    }
                    // 业务系统导出的「Excel」常是 HTML / XML 表格，Excel 能打开
                    if (text) return OkResult();
                    return Encrypted();
                case ".xls":
                    // 旧版二进制；或改了扩展名的 xlsx；或 HTML / 制表符文本冒充的 xls：Excel 都能打开
                    if (ole || zip || text) return OkResult();
                    return Encrypted();
                case ".csv":
                case ".txt":
                case ".tsv":
                case ".json":
                case ".xml":
                case ".md":
                case ".log":
                    if (text) return OkResult();
                    return Encrypted();
                default:
                    return OkResult();
            }
        }

        private static SniffResult OkResult() => new SniffResult { Verdict = SniffVerdict.Ok };

        private static SniffResult Encrypted() => new SniffResult
        {
            Verdict = SniffVerdict.LikelyEncrypted,
            Message = EncryptedMessage,
            Suggestion = EncryptedSuggestion,
        };

        private static bool StartsWith(byte[] head, byte[] magic)
        {
            if (head.Length < magic.Length) return false;
            for (var i = 0; i < magic.Length; i++)
            {
                if (head[i] != magic[i]) return false;
            }
            return true;
        }

        /// <summary>
        /// 像不像文本：UTF-16 BOM 直接算；否则控制字符（除制表、换行、回车、换页）和 NUL 不超过 2%。
        /// UTF-8 / GBK 的多字节字符都是 ≥ 0x80 的字节，不会被算成控制字符。
        /// </summary>
        public static bool LooksLikeText(byte[] head)
        {
            if (head.Length >= 2 && ((head[0] == 0xFF && head[1] == 0xFE) || (head[0] == 0xFE && head[1] == 0xFF)))
            {
                return true;
            }
            var bad = 0;
            foreach (var b in head)
            {
                if (b == 0x09 || b == 0x0A || b == 0x0C || b == 0x0D) continue;
                if (b < 0x20 || b == 0x7F) bad++;
            }
            return bad <= head.Length / 50;
        }
    }
}
