using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// 工作簿分析器 - 读取工作簿级结构信息
    /// </summary>
    public class WorkbookAnalyzer
    {
        private readonly Application _app;

        public WorkbookAnalyzer(Application app)
        {
            _app = app;
        }

        public WorkbookStructure Analyze()
        {
            var wb = _app.ActiveWorkbook;
            if (wb == null) return null;

            return new WorkbookStructure
            {
                Name = SafeGet(() => wb.Name),
                FilePath = SafeGet(() => wb.FullName),
                Worksheets = AnalyzeWorksheets(wb),
                NamedRanges = AnalyzeNamedRanges(wb),
                HasVBAProject = SafeGet(() => wb.HasVBProject, false)
            };
        }

        private WorksheetInfo[] AnalyzeWorksheets(Workbook wb)
        {
            var result = new List<WorksheetInfo>();

            try
            {
                var sheets = wb.Worksheets;
                foreach (Worksheet sheet in sheets)
                {
                    result.Add(new WorksheetInfo
                    {
                        Name = sheet.Name,
                        Index = sheet.Index,
                        UsedRangeAddress = SafeGet(() => sheet.UsedRange.Address),
                        IsVisible = sheet.Visible == XlSheetVisibility.xlSheetVisible
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AnalyzeWorksheets error: {ex}");
            }

            return result.ToArray();
        }

        private NamedRangeInfo[] AnalyzeNamedRanges(Workbook wb)
        {
            var result = new List<NamedRangeInfo>();

            try
            {
                var names = wb.Names;
                foreach (Name name in names)
                {
                    result.Add(new NamedRangeInfo
                    {
                        Name = name.Name,
                        RefersTo = name.RefersTo?.ToString()
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AnalyzeNamedRanges error: {ex}");
            }

            return result.ToArray();
        }

        private string SafeGet(Func<string> getter, string defaultValue = "")
        {
            try { return getter() ?? defaultValue; }
            catch { return defaultValue; }
        }

        private T SafeGet<T>(Func<T> getter, T defaultValue = default)
        {
            try { return getter(); }
            catch { return defaultValue; }
        }
    }
}
