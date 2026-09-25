using System;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Perception
{
    /// <summary>一页读取计划：从裁剪后区域的第 Offset 行起读 Rows 行、Columns 列。</summary>
    public sealed class PagePlan
    {
        public int Offset { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
        /// <summary>还有下一页时是下一页的 offset，读完了为 null</summary>
        public int? NextOffset { get; set; }
        public bool ColumnsTruncated { get; set; }
        public string Error { get; set; }
    }

    /// <summary>工作表上的矩形（行列 1 起始，含两端）。</summary>
    public struct SheetBox
    {
        public int Row1, Col1, Row2, Col2;

        public SheetBox(int row1, int col1, int row2, int col2)
        {
            Row1 = row1; Col1 = col1; Row2 = row2; Col2 = col2;
        }

        public int Rows => Row2 - Row1 + 1;
        public int Columns => Col2 - Col1 + 1;
    }

    /// <summary>
    /// read_range 分页（Claude Code 的 Read(offset, limit) 的工作簿版）。
    ///
    /// 以前 read_range("A:A") 会把 104 万行的值、公式、数字格式全读出来再截到 200 行：Excel
    /// 卡住、内存暴涨，模型还不知道下一页怎么读。现在先裁到已用区域，再只读一页，
    /// 并在结果里写明总行数和下一页的 offset。
    /// </summary>
    public static class RangePaging
    {
        public const int DefaultRows = 200;
        public const int MaxRows = 500;
        /// <summary>一页最多这么多格：值 + 公式 + 数字格式序列化后要远低于 SDK 的 1MB 消息上限</summary>
        public const int MaxCells = 10000;
        public const int MaxColumns = 256;

        /// <summary>请求区域与已用区域的交集；没有交集返回 null（整块都是空白）。</summary>
        public static SheetBox? Clip(SheetBox requested, SheetBox used)
        {
            var r1 = Math.Max(requested.Row1, used.Row1);
            var c1 = Math.Max(requested.Col1, used.Col1);
            var r2 = Math.Min(requested.Row2, used.Row2);
            var c2 = Math.Min(requested.Col2, used.Col2);
            if (r1 > r2 || c1 > c2) return null;
            return new SheetBox(r1, c1, r2, c2);
        }

        public static PagePlan Plan(int totalRows, int totalColumns, int offset, int? limit)
        {
            if (offset < 0) offset = 0;
            if (totalRows <= 0 || totalColumns <= 0)
            {
                return new PagePlan { Offset = 0, Rows = 0, Columns = 0 };
            }
            if (offset >= totalRows)
            {
                return new PagePlan
                {
                    Error = $"offset={offset} 超出范围：这块区域一共 {totalRows} 行（offset 从 0 开始，最大 {totalRows - 1}）",
                };
            }

            var columns = Math.Min(totalColumns, MaxColumns);
            var wanted = limit.HasValue && limit.Value > 0 ? Math.Min(limit.Value, MaxRows) : DefaultRows;
            var byCells = Math.Max(1, MaxCells / columns);
            var rows = Math.Min(Math.Min(wanted, byCells), totalRows - offset);
            var next = offset + rows;
            return new PagePlan
            {
                Offset = offset,
                Rows = rows,
                Columns = columns,
                NextOffset = next < totalRows ? next : (int?)null,
                ColumnsTruncated = columns < totalColumns,
            };
        }

        /// <summary>给模型的一句话：读到了哪里、下一页怎么读。读完了且没截列时返回 null。</summary>
        public static string Hint(string address, int totalRows, int totalColumns, PagePlan plan)
        {
            if (plan == null || plan.Rows == 0) return null;
            string hint = null;
            if (plan.NextOffset.HasValue)
            {
                hint = $"共 {totalRows} 行，本页是第 {plan.Offset + 1}–{plan.Offset + plan.Rows} 行，还有 {totalRows - plan.NextOffset.Value} 行未读。" +
                       $"下一页：read_range(address=\"{address}\", offset={plan.NextOffset.Value})";
            }
            if (plan.ColumnsTruncated)
            {
                var cols = $"这块区域有 {totalColumns} 列，只返回了前 {plan.Columns} 列；其余列请用更窄的地址读取。";
                hint = hint == null ? cols : hint + " " + cols;
            }
            return hint;
        }
    }

    /// <summary>read_range 的一页结果：RangeInfo 加分页信息（模型看得到每个字段）。</summary>
    public class RangePage : RangeInfo
    {
        [JsonPropertyName("paging")]
        public PagingInfo Paging { get; set; }

        [JsonPropertyName("hint")]
        public string Hint { get; set; }

        public static RangePage From(RangeInfo info)
        {
            var page = new RangePage();
            if (info == null) return page;
            page.Address = info.Address;
            page.WorksheetName = info.WorksheetName;
            page.RowCount = info.RowCount;
            page.ColumnCount = info.ColumnCount;
            page.Values = info.Values;
            page.Formulas = info.Formulas;
            page.NumberFormats = info.NumberFormats;
            page.FormatConditions = info.FormatConditions;
            page.MergeCells = info.MergeCells;
            return page;
        }
    }

    public class PagingInfo
    {
        /// <summary>模型请求的地址（带表名）</summary>
        [JsonPropertyName("requested")]
        public string Requested { get; set; }

        /// <summary>裁到已用区域后的总行数 / 列数</summary>
        [JsonPropertyName("total_rows")]
        public int TotalRows { get; set; }

        [JsonPropertyName("total_columns")]
        public int TotalColumns { get; set; }

        [JsonPropertyName("offset")]
        public int Offset { get; set; }

        [JsonPropertyName("returned_rows")]
        public int ReturnedRows { get; set; }

        [JsonPropertyName("next_offset")]
        public int? NextOffset { get; set; }

        [JsonPropertyName("columns_truncated")]
        public bool ColumnsTruncated { get; set; }

        /// <summary>请求区域比已用区域大，只读了有内容的部分</summary>
        [JsonPropertyName("clipped_to_used_range")]
        public bool ClippedToUsedRange { get; set; }
    }
}
