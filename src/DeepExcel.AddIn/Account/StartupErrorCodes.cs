using System.Collections.Generic;

namespace DeepExcel.AddIn.Account
{
    /// <summary>
    /// startup_error 遥测事件的诊断码（服务端白名单要求 E-XXX-NNN）。与 DeepExcel.Repair 的
    /// Codes 同一个命名空间：同一个码永远只表示一件事，E-LOAD-001 与修复工具的含义一致。
    /// 只发码和版本号，从不发异常文字——异常文字里常有路径和单元格内容。
    /// </summary>
    public static class StartupErrorCodes
    {
        /// <summary>Excel 加载了插件，但 OnConnection 抛异常（与 Repair 的 LastLoadFailed 相同）</summary>
        public const string LoadFailed = "E-LOAD-001";
        public const string BridgeInitFailed = "E-LOAD-002";
        public const string WebViewInitFailed = "E-LOAD-003";

        /// <summary>侧车进程意外退出（不是用户关闭、也不是重启）</summary>
        public const string SidecarCrashed = "E-SIDE-001";

        /// <summary>侧车启动自检（selfcheck.py）的诊断 → 遥测码</summary>
        public static readonly IReadOnlyDictionary<string, string> Engine = new Dictionary<string, string>
        {
            ["os_too_old"] = "E-ENG-001",
            ["cli_missing"] = "E-ENG-002",
            ["cli_blocked"] = "E-ENG-003",
            ["cli_incompatible"] = "E-ENG-004",
            ["cli_timeout"] = "E-ENG-005",
            ["cli_crashed"] = "E-ENG-006",
            ["connect_failed"] = "E-ENG-007",
        };
    }
}
