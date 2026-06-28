using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Excel操作接口 - 由执行引擎/感知层实现
    /// Bridge层只负责协议转换，不关心具体业务
    /// </summary>
    public interface IExcelActions
    {
        // 感知类
        object GetSelection();
        object ReadRange(string address);
        object ReadWorkbook();
        object ReadWorksheet(string name);

        // 执行类
        ToolResult ExecuteVBA(string code, string macroName = null);
        ToolResult ExecutePython(string code);
        ToolResult WriteFormula(string address, string formula);
        ToolResult WriteValue(string address, object value);

        // Sheet 管理
        ToolResult AddSheet(string name);
        ToolResult DeleteSheet(string name);
        ToolResult RenameSheet(string oldName, string newName);

        // 格式化
        ToolResult SetNumberFormat(string address, string format);
        ToolResult SetColumnWidth(string address, double width, bool autoFit);

        // 数据操作
        ToolResult SortData(string rangeAddress, string sortColumn, bool descending);
        ToolResult FilterData(string rangeAddress, int columnIndex, string criteria);

        // 安全类
        string CreateSnapshot();
        bool Rollback(string snapshotId);
    }
}
