namespace DeepExcel.AddIn.Perception
{
    /// <summary>
    /// 单元格/区域信息
    /// </summary>
    public class RangeInfo
    {
        public string Address { get; set; }
        public string WorksheetName { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public object[,] Values { get; set; }
        public string[,] Formulas { get; set; }
        public string[,] NumberFormats { get; set; }
        public FormatConditionInfo[] FormatConditions { get; set; }
        public MergeCellInfo[] MergeCells { get; set; }
    }

    public class FormatConditionInfo
    {
        public string Type { get; set; }
        public string Operator { get; set; }
        public string Formula1 { get; set; }
        public string Formula2 { get; set; }
    }

    public class MergeCellInfo
    {
        public string Address { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
    }
}
