namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// 工作簿结构信息
    /// </summary>
    public class WorkbookStructure
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public WorksheetInfo[] Worksheets { get; set; }
        public NamedRangeInfo[] NamedRanges { get; set; }
        public bool HasVBAProject { get; set; }
        public string ActiveSheet { get; set; }
    }

    public class WorksheetInfo
    {
        public string Name { get; set; }
        public int Index { get; set; }
        public string UsedRangeAddress { get; set; }
        public bool IsVisible { get; set; } = true;
    }

    public class NamedRangeInfo
    {
        public string Name { get; set; }
        public string RefersTo { get; set; }
    }
}
