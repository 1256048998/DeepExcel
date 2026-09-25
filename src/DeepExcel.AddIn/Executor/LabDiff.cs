using System;
using System.Collections.Generic;
using System.Linq;
using DeepExcel.AddIn.Preview;
using CellRect = DeepExcel.AddIn.Sidecar.CellRect;

namespace DeepExcel.AddIn.Executor
{
    /// <summary>一张表在某一时刻的内容：左上角位置 + 每格的公式文本（常量就是它的文本）。</summary>
    public sealed class SheetContent
    {
        public string Name { get; set; }
        public int Row1 { get; set; } = 1;
        public int Col1 { get; set; } = 1;
        /// <summary>[行, 列]，0 起；null 或空串表示空格</summary>
        public string[,] Cells { get; set; } = new string[0, 0];
        /// <summary>表太大，只读了前面一部分</summary>
        public bool Truncated { get; set; }

        public int Rows => Cells.GetLength(0);
        public int Columns => Cells.GetLength(1);

        public string At(int row, int col)
        {
            int r = row - Row1, c = col - Col1;
            if (r < 0 || c < 0 || r >= Rows || c >= Columns) return "";
            return Cells[r, c] ?? "";
        }
    }

    public sealed class LabDiffResult
    {
        public List<CellChange> Changes { get; } = new List<CellChange>();
        public int AffectedCells { get; set; }
        public List<string> AddedSheets { get; } = new List<string>();
        public List<string> RemovedSheets { get; } = new List<string>();
        /// <summary>每张表上改动的外接矩形：试跑之后用户改到这里，结果就不可信了</summary>
        public List<CellRect> AffectedRects { get; } = new List<CellRect>();
        public bool Truncated { get; set; }
    }

    /// <summary>
    /// 试跑前后的对比。只比公式 / 常量文本（和用户看到的编辑栏一致）：公式没变、只是重算出新值的格
    /// 不算改动——那是依赖关系的结果，不是这段代码写进去的。
    /// </summary>
    public static class LabDiff
    {
        public static LabDiffResult Compare(IList<SheetContent> before, IList<SheetContent> after,
            int maxListed = ChangePreview.MaxListedChanges)
        {
            var result = new LabDiffResult();
            var beforeByName = before.ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
            var afterByName = after.ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
            result.AddedSheets.AddRange(after.Where(s => !beforeByName.ContainsKey(s.Name)).Select(s => s.Name));
            result.RemovedSheets.AddRange(before.Where(s => !afterByName.ContainsKey(s.Name)).Select(s => s.Name));

            foreach (var sheet in after)
            {
                beforeByName.TryGetValue(sheet.Name, out var old);
                old = old ?? new SheetContent { Name = sheet.Name };
                result.Truncated |= sheet.Truncated || old.Truncated;

                int r1 = Math.Min(old.Row1, sheet.Row1), c1 = Math.Min(old.Col1, sheet.Col1);
                int r2 = Math.Max(old.Row1 + old.Rows - 1, sheet.Row1 + sheet.Rows - 1);
                int c2 = Math.Max(old.Col1 + old.Columns - 1, sheet.Col1 + sheet.Columns - 1);
                int minR = int.MaxValue, minC = int.MaxValue, maxR = 0, maxC = 0;

                for (int r = r1; r <= r2; r++)
                {
                    for (int c = c1; c <= c2; c++)
                    {
                        string was = old.At(r, c), now = sheet.At(r, c);
                        if (string.Equals(was, now, StringComparison.Ordinal)) continue;
                        result.AffectedCells++;
                        minR = Math.Min(minR, r); minC = Math.Min(minC, c);
                        maxR = Math.Max(maxR, r); maxC = Math.Max(maxC, c);
                        if (result.Changes.Count >= maxListed) continue;
                        result.Changes.Add(new CellChange
                        {
                            Address = new CellRect(sheet.Name, r, c, r, c).ToA1(),
                            Before = was,
                            After = now,
                            Kind = was.Length == 0 ? ChangeKind.Add : now.Length == 0 ? ChangeKind.Clear : ChangeKind.Overwrite,
                            OverwritesFormula = was.StartsWith("=", StringComparison.Ordinal) &&
                                                !now.StartsWith("=", StringComparison.Ordinal),
                        });
                    }
                }
                if (maxR > 0)
                {
                    result.AffectedRects.Add(new CellRect(sheet.Name, minR, minC, maxR, maxC));
                }
            }
            return result;
        }
    }
}
