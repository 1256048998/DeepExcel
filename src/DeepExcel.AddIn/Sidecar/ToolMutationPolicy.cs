using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DeepExcel.AddIn.Executor;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// 哪些工具会改工作簿、改的是哪几张表。ToolDispatcher 据此在写入前自动备份。
    ///
    /// 用"只读白名单"而不是"写入名单"：没列在这里的工具一律当作写入处理。
    /// 新加一个工具忘了登记，代价是多备份一次；反过来的代价是那次修改无法撤销。
    /// </summary>
    internal static class ToolMutationPolicy
    {
        public static readonly HashSet<string> ReadOnlyTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "echo",
            "read_range", "read_workbook", "read_selection", "read_attachment",
            "screenshot_excel",
            // 导出图表只写图片文件，不改工作簿
            "export_chart",
            // CodeSandbox 禁掉了所有能碰工作簿的库，只能做纯计算
            "execute_python",
            // 这两个自己管理快照：create_snapshot 本身就是备份，rollback 恢复前会先存一份当前状态
            "create_snapshot", "rollback",
        };

        public static bool IsMutating(string toolName) => !ReadOnlyTools.Contains(toolName ?? "");

        /// <summary>
        /// 只改一张表、且不会连带改动其他表的工具，以及它的目标地址参数。
        ///
        /// 不在这里的写入工具按整本处理。典型的是插入/删除行列、删表、改表名：
        /// Excel 会同步改写其他表里引用它的公式，只恢复这一张表会让那些公式指错位置。
        /// </summary>
        private static readonly Dictionary<string, string> SingleSheetAddressArg = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["write_formula"] = "address",
            ["write_value"] = "address",
            ["write_range"] = "address",
            ["write_table"] = "address",
            ["set_number_format"] = "address",
            ["set_column_width"] = "address",
            ["set_cell_style"] = "address",
            ["merge_cells"] = "address",
            ["unmerge_cells"] = "address",
            ["clear_range"] = "address",
            ["apply_conditional_format"] = "address",
            ["freeze_panes"] = "address",
            ["fill_formula_down"] = "from_address",
            ["replace_formula"] = "range_address",
            ["fill_blank_cells"] = "range_address",
            ["highlight_duplicates"] = "range_address",
            ["remove_special_chars"] = "range_address",
            ["clean_amount"] = "range_address",
            ["collapse_spaces"] = "range_address",
            ["rename_columns"] = "range_address",
            ["sort_data"] = "range_address",
            ["filter_data"] = "range_address",
            ["copy_range"] = "dest_address",
        };

        private static readonly Regex A1Reference = new Regex(
            @"^(\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)$",
            RegexOptions.Compiled);

        /// <summary>
        /// 推算这次调用会改哪些表。拿不准就返回整本——范围算小了，回滚会漏掉改动；
        /// 算大了只是多恢复几张表，而恢复前当前状态总会先另存一份。
        /// </summary>
        public static SnapshotScope ResolveScope(string toolName, Func<string, string> getStringArg, Func<string> activeSheetName)
        {
            if (toolName == "add_sheet")
            {
                var name = getStringArg?.Invoke("name");
                return string.IsNullOrWhiteSpace(name) ? SnapshotScope.Whole() : SnapshotScope.ForSheets(name.Trim());
            }

            if (!SingleSheetAddressArg.TryGetValue(toolName ?? "", out var argName)) return SnapshotScope.Whole();

            var address = getStringArg?.Invoke(argName);
            if (!TryGetSheetOfAddress(address, out var sheet, out var qualified)) return SnapshotScope.Whole();
            if (!qualified) sheet = activeSheetName?.Invoke();
            return string.IsNullOrEmpty(sheet) ? SnapshotScope.Whole() : SnapshotScope.ForSheets(sheet);
        }

        /// <summary>
        /// 解析 "A1:B2" / "Sheet1!A1" / "'My Sheet'!A:A" 这类单区域 A1 地址。
        /// 名称、多区域、跨工作簿引用都返回 false（由调用方按整本处理）。
        /// </summary>
        internal static bool TryGetSheetOfAddress(string address, out string sheet, out bool qualified)
        {
            sheet = null;
            qualified = false;
            if (string.IsNullOrWhiteSpace(address)) return false;

            var text = address.Trim();
            var cellPart = text;
            int bang = text.LastIndexOf('!');
            if (bang >= 0)
            {
                var sheetPart = text.Substring(0, bang).Trim();
                cellPart = text.Substring(bang + 1).Trim();
                if (sheetPart.Length >= 2 && sheetPart[0] == '\'' && sheetPart[sheetPart.Length - 1] == '\'')
                {
                    sheetPart = sheetPart.Substring(1, sheetPart.Length - 2).Replace("''", "'");
                }
                if (sheetPart.Length == 0 || sheetPart.IndexOf('[') >= 0 || sheetPart.IndexOf(']') >= 0) return false;
                sheet = sheetPart;
                qualified = true;
            }

            if (!A1Reference.IsMatch(cellPart))
            {
                sheet = null;
                qualified = false;
                return false;
            }
            return true;
        }
    }
}
