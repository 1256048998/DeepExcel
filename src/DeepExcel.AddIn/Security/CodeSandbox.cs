using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Security
{
    /// <summary>
    /// ★ P0-3 修复：execute_python / execute_vba 沙箱化（黑名单拦截）。
    /// 阻止 LLM 通过代码执行任意系统命令、读写注册表、外泄数据。
    /// 这是软沙箱（语法层拦截），无法完全阻止恶意代码，但能挡住最常见的 RCE 路径。
    /// </summary>
    public static class CodeSandbox
    {
        // Python 危险模式：禁止 import os/subprocess/sys、os.system、subprocess.*、
        /// __import__、eval/exec、open(*, 'w') 写文件、socket 等。
        /// （read_range/write_range 通过 csv 已隔离真实路径，但 execute_python 可直接调用 open）
        /// </summary>
        private static readonly Regex[] PythonBlockedPatterns =
        {
            // 系统/进程操作
            new Regex(@"\bimport\s+(os|subprocess|shutil|ctypes|socket|http|urllib|requests|ftplib|smtplib|telnetlib|paramiko)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bfrom\s+(os|subprocess|shutil|ctypes|socket|http|urllib|requests|ftplib|smtplib|telnetlib|paramiko)\s+import\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bos\.system\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bsubprocess\.(run|Popen|call|check_output|check_call|getoutput)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bshutil\.(rmtree|move|copy|copyfile)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\b__import__\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // 动态执行
            new Regex(@"\beval\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bexec\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bcompile\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // 网络外泄
            new Regex(@"\bsocket\.(socket|create_connection)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\brequests\.(get|post|put|delete|head|patch|request)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\burllib\.", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // 文件写（防止覆盖重要文件，临时 csv 由 write_range 走专用路径）
            new Regex(@"\bopen\s*\([^)]*['""]w", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // 进程/退出
            new Regex(@"\bos\.(exec|spawn|fork|kill|_exit|exit|popen|getpid|getppid)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bsys\.(exit|modules|path)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // 环境/注册表
            new Regex(@"\bos\.(environ|getenv|putenv|chmod|chown|system|unlink|remove)\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bplatform\.", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // 反射窥探
            new Regex(@"\bglobals\s*\(\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\blocals\s*\(\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bvars\s*\(\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bgetattr\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bsetattr\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            // ★ 禁止 openpyxl/pandas 操作工作簿文件：工作簿被 Excel 锁定，save 必然 PermissionError
            // AI 应该用 read_range/write_value/execute_vba 代替
            new Regex(@"\bload_workbook\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bopenpyxl\.", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bimport\s+openpyxl\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bfrom\s+openpyxl\s+import\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bimport\s+pandas\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bfrom\s+pandas\s+import\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bpd\.read_excel\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bpd\.ExcelWriter\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\b\.to_excel\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        };

        // VBA 危险模式：Shell / WScript.Shell / CreateObject(任意) / Kill / Open 等等
        private static readonly Regex[] VbaBlockedPatterns =
        {
            new Regex(@"\bShell\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bCreateObject\s*\(\s*[""']WScript\.Shell[""']", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bCreateObject\s*\(\s*[""']Scripting\.FileSystemObject[""']", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bCreateObject\s*\(\s*[""']ADODB\.Stream[""']", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bCreateObject\s*\(\s*[""']MSXML2\.XMLHTTP[""']", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bCreateObject\s*\(\s*[""']WinHttp\.WinHttpRequest", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bCreateObject\s*\(\s*[""']InternetExplorer\.Application[""']", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bKill\s+\w", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bName\s+\w+\s+As\s+\w+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bOpen\s+[""']([A-Za-z]:[\\/]|[\\/])", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bEnviron\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bSetAttr\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bFileDateTime\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new Regex(@"\bFileLen\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        };

        /// <summary>
        /// 校验 Python 代码，返回第一个命中的危险模式描述（null 表示通过）。
        /// </summary>
        public static string ValidatePython(string code)
        {
            if (string.IsNullOrEmpty(code)) return null;
            foreach (var pattern in PythonBlockedPatterns)
            {
                var match = pattern.Match(code);
                if (match.Success)
                {
                    // ★ 针对性提示：openpyxl/pandas 操作 Excel 文件会 PermissionError
                    string hint = "（系统/进程/网络/文件写入已禁用，请用 read_range/write_value 工具操作数据）";
                    if (match.Value.Contains("openpyxl") || match.Value.Contains("pandas")
                        || match.Value.Contains("load_workbook") || match.Value.Contains("to_excel")
                        || match.Value.Contains("read_excel") || match.Value.Contains("ExcelWriter"))
                    {
                        hint = "（工作簿被 Excel 锁定，openpyxl/pandas 无法读写。请用 read_range 读取数据，write_value/write_formula 写入数据，execute_vba 做复杂格式操作）";
                    }
                    return "Python 代码包含受限操作: " + match.Value.Trim() + hint;
                }
            }
            return null;
        }

        /// <summary>
        /// 校验 VBA 代码，返回第一个命中的危险模式描述（null 表示通过）。
        /// </summary>
        public static string ValidateVba(string code)
        {
            if (string.IsNullOrEmpty(code)) return null;
            foreach (var pattern in VbaBlockedPatterns)
            {
                var match = pattern.Match(code);
                if (match.Success)
                {
                    return "VBA 代码包含受限操作: " + match.Value.Trim() +
                           "（Shell/文件系统/网络请求/注册表已禁用）";
                }
            }
            return null;
        }
    }
}
