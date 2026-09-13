using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DeepExcel.AddIn.Preview
{
    /// <summary>One cell as it currently exists.</summary>
    public sealed class CellSnapshot
    {
        public string Address { get; set; }
        public string Display { get; set; }
        public string Formula { get; set; }

        public bool HasFormula
        {
            get { return !string.IsNullOrEmpty(Formula) && Formula.StartsWith("="); }
        }

        public bool IsEmpty
        {
            get { return string.IsNullOrEmpty(Display) && string.IsNullOrEmpty(Formula); }
        }
    }

    /// <summary>
    /// Reads the current state of a range. Implemented over Excel in
    /// production, and over a dictionary in tests, so the diff logic below can
    /// be exercised without Excel.
    /// </summary>
    public interface IRangeReader
    {
        /// <summary>Cells in row-major order, or an empty list if unreadable.</summary>
        IReadOnlyList<CellSnapshot> Read(string sheet, string address);

        /// <summary>Row count of the used range, for bounding deletes.</summary>
        int UsedRowCount(string sheet);
    }

    /// <summary>
    /// Computes what an operation will change, before it changes it.
    ///
    /// The diff logic is deliberately free of Excel types: this is the part that
    /// must be right, and it has to be testable against the awkward cases
    /// (formula overwrite, no-op writes, ranges larger than the supplied data)
    /// rather than only against whatever workbook is open.
    /// </summary>
    public sealed class PreviewBuilder
    {
        private readonly IRangeReader _reader;

        public PreviewBuilder(IRangeReader reader)
        {
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        }

        public ChangePreview Build(string toolName, IDictionary<string, object> args)
        {
            var level = PreviewPolicy.LevelFor(toolName);
            if (level == ConfirmationLevel.SnapshotAndWarn)
            {
                return ChangePreview.NotPreviewable(toolName, PreviewPolicy.ReasonFor(toolName));
            }
            if (level == ConfirmationLevel.None)
            {
                return null;
            }

            args = args ?? new Dictionary<string, object>();
            var sheet = GetString(args, "sheet") ?? GetString(args, "sheet_name");

            try
            {
                switch (toolName.ToLowerInvariant())
                {
                    case "write_value":
                    case "write_formula":
                        return BuildScalarWrite(toolName, sheet, args);
                    case "write_range":
                    case "write_table":
                        return BuildRangeWrite(toolName, sheet, args);
                    case "clear_range":
                        return BuildClear(toolName, sheet, args);
                    case "delete_rows":
                        return BuildDeleteRows(toolName, sheet, args);
                    case "delete_columns":
                        return BuildDeleteColumns(toolName, sheet, args);
                    default:
                        // Simulable in principle but not implemented yet.
                        // Claiming an empty change set would be a lie the user
                        // would act on, so say so instead.
                        return ChangePreview.NotPreviewable(
                            toolName, "该操作的精确预览尚未实现，已自动创建快照");
                }
            }
            catch (Exception ex)
            {
                return ChangePreview.NotPreviewable(toolName, "预览计算失败：" + ex.GetType().Name);
            }
        }

        // ------------------------------------------------------------------

        private ChangePreview BuildScalarWrite(string toolName, string sheet, IDictionary<string, object> args)
        {
            var address = GetString(args, "address") ?? GetString(args, "cell");
            var after = GetString(args, "value") ?? GetString(args, "formula") ?? "";
            var preview = new ChangePreview { ToolName = toolName };
            if (string.IsNullOrEmpty(address))
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标地址");
            }

            var current = _reader.Read(sheet, address);
            foreach (var cell in current)
            {
                AddChangeIfDifferent(preview, cell, after);
            }
            Finalize(preview);
            return preview;
        }

        private ChangePreview BuildRangeWrite(string toolName, string sheet, IDictionary<string, object> args)
        {
            var address = GetString(args, "address") ?? GetString(args, "range");
            if (string.IsNullOrEmpty(address))
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }

            var values = FlattenValues(args);
            var preview = new ChangePreview { ToolName = toolName };
            var current = _reader.Read(sheet, address);

            for (var i = 0; i < current.Count; i++)
            {
                // A write narrower than the target range leaves the tail alone
                // rather than clearing it; reporting those cells as changed
                // would overstate the blast radius.
                var after = i < values.Count ? values[i] : null;
                if (after == null)
                {
                    continue;
                }
                AddChangeIfDifferent(preview, current[i], after);
            }
            Finalize(preview);
            return preview;
        }

        private ChangePreview BuildClear(string toolName, string sheet, IDictionary<string, object> args)
        {
            var address = GetString(args, "address") ?? GetString(args, "range");
            if (string.IsNullOrEmpty(address))
            {
                return ChangePreview.NotPreviewable(toolName, "缺少目标区域");
            }

            var preview = new ChangePreview { ToolName = toolName };
            foreach (var cell in _reader.Read(sheet, address))
            {
                if (cell.IsEmpty)
                {
                    continue;
                }
                preview.Changes.Add(new CellChange
                {
                    Address = cell.Address,
                    Before = Describe(cell),
                    After = "",
                    Kind = ChangeKind.Clear,
                    OverwritesFormula = cell.HasFormula
                });
            }
            Finalize(preview);
            return preview;
        }

        private ChangePreview BuildDeleteRows(string toolName, string sheet, IDictionary<string, object> args)
        {
            var start = GetInt(args, "start_row") ?? GetInt(args, "row");
            var count = GetInt(args, "count") ?? GetInt(args, "row_count") ?? 1;
            if (start == null || start < 1 || count < 1)
            {
                return ChangePreview.NotPreviewable(toolName, "行号参数无效");
            }

            var preview = new ChangePreview { ToolName = toolName };
            var usedRows = _reader.UsedRowCount(sheet);
            for (var row = start.Value; row < start.Value + count; row++)
            {
                preview.DeletedRows.Add(row);
            }

            // Deleting past the data is usually a miscount rather than intent.
            if (usedRows > 0 && start.Value + count - 1 > usedRows)
            {
                preview.Warnings.Add(
                    $"删除范围超出数据区（数据到第 {usedRows} 行），可能是行号算错了");
            }
            preview.Warnings.Add("删除行会使下方所有行上移，引用这些行的公式可能失效");
            Finalize(preview);
            return preview;
        }

        private ChangePreview BuildDeleteColumns(string toolName, string sheet, IDictionary<string, object> args)
        {
            var columns = GetString(args, "columns") ?? GetString(args, "column");
            if (string.IsNullOrEmpty(columns))
            {
                return ChangePreview.NotPreviewable(toolName, "缺少列参数");
            }

            var preview = new ChangePreview { ToolName = toolName };
            foreach (var part in columns.Split(',', ';'))
            {
                var trimmed = part.Trim();
                if (trimmed.Length > 0)
                {
                    preview.DeletedColumns.Add(trimmed);
                }
            }
            preview.Warnings.Add("删除列会使右侧所有列左移，引用这些列的公式可能失效");
            Finalize(preview);
            return preview;
        }

        // ------------------------------------------------------------------

        private static void AddChangeIfDifferent(ChangePreview preview, CellSnapshot cell, string after)
        {
            var before = Describe(cell);
            if (string.Equals(before, after, StringComparison.Ordinal))
            {
                // Writing the same value is not a change. Counting it would
                // inflate the number the user is being asked to approve.
                return;
            }

            preview.Changes.Add(new CellChange
            {
                Address = cell.Address,
                Before = before,
                After = after,
                Kind = cell.IsEmpty ? ChangeKind.Add : ChangeKind.Overwrite,
                OverwritesFormula = cell.HasFormula
            });
        }

        private static void Finalize(ChangePreview preview)
        {
            preview.AffectedCells = preview.Changes.Count;

            if (preview.FormulasOverwritten > 0)
            {
                preview.Warnings.Insert(0,
                    $"{preview.FormulasOverwritten} 处会覆盖已有公式，覆盖后该单元格不再自动计算");
            }

            // The list is what the user reads; the count is what they approve.
            if (preview.Changes.Count > ChangePreview.MaxListedChanges)
            {
                preview.Changes = preview.Changes.Take(ChangePreview.MaxListedChanges).ToList();
            }
        }

        internal static string Describe(CellSnapshot cell)
        {
            if (cell == null) return "";
            // A formula's text matters more than its current result: that is
            // what is about to be lost.
            return cell.HasFormula ? cell.Formula : (cell.Display ?? "");
        }

        internal static List<string> FlattenValues(IDictionary<string, object> args)
        {
            var flat = new List<string>();
            if (args == null || !args.TryGetValue("values", out var raw) || raw == null)
            {
                return flat;
            }
            Flatten(raw, flat);
            return flat;
        }

        private static void Flatten(object value, List<string> into)
        {
            if (value is string text)
            {
                into.Add(text);
                return;
            }
            if (value is System.Collections.IEnumerable sequence)
            {
                foreach (var item in sequence)
                {
                    Flatten(item, into);
                }
                return;
            }
            into.Add(Stringify(value));
        }

        internal static string Stringify(object value)
        {
            if (value == null) return "";
            if (value is double d)
            {
                return d == Math.Floor(d) && Math.Abs(d) < 1e15
                    ? ((long)d).ToString(CultureInfo.InvariantCulture)
                    : d.ToString("0.####", CultureInfo.InvariantCulture);
            }
            if (value is bool b) return b ? "TRUE" : "FALSE";
            if (value is DateTime dt) return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }

        private static string GetString(IDictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }
            return value as string ?? Stringify(value);
        }

        private static int? GetInt(IDictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }
            if (value is int i) return i;
            if (value is long l) return (int)l;
            if (value is double d) return (int)d;
            return int.TryParse(Stringify(value), out var parsed) ? parsed : (int?)null;
        }
    }
}
