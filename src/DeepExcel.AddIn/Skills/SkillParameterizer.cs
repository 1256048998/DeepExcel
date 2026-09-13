using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Skills
{
    /// <summary>One tool invocation as it actually happened.</summary>
    public sealed class RecordedStep
    {
        public string Tool { get; set; }
        public Dictionary<string, object> Arguments { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// Turns one concrete run into a reusable skill.
    ///
    /// The judgement here is which recorded values are incidental to that run
    /// and should become parameters, and which are part of the task and should
    /// stay fixed. Getting it wrong in either direction makes the skill useless:
    /// parameterise too much and the user has to re-answer everything they
    /// already said; too little and it only ever works on the original data.
    ///
    /// Pure, so the rules can be tested against real argument shapes.
    /// </summary>
    public static class SkillParameterizer
    {
        // Sheet1!A1, 'My Sheet'!$A$1:$F$200, A1:F200, A1
        private static readonly Regex RangeAddress = new Regex(
            @"^(?:(?:'[^']+'|[^!\s]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Whole-column and whole-row forms.
        private static readonly Regex ColumnOrRowAddress = new Regex(
            @"^(?:(?:'[^']+'|[^!\s]+)!)?(?:\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex FilePath = new Regex(
            @"^(?:[a-zA-Z]:\\|\\\\|/)[^\r\n]*$", RegexOptions.Compiled);

        private static readonly string[] DateFormats =
        {
            "yyyy-MM-dd", "yyyy/MM/dd", "yyyy-M-d", "yyyy/M/d",
            "yyyy年M月d日", "yyyy年MM月dd日", "yyyy-MM", "yyyy/MM"
        };

        /// <summary>
        /// Argument names whose value is part of the task, not of this run.
        ///
        /// Turning a sort direction or a chart type into a parameter would make
        /// the user re-specify what they already decided.
        /// </summary>
        private static readonly HashSet<string> NeverParameterize =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ascending", "descending", "order", "direction",
                "has_header", "chart_type", "style", "format",
                "operation", "mode", "type", "position", "orientation",
                "overwrite", "confirm", "dry_run"
            };

        /// <summary>
        /// Code is never parameterised: a VBA body with a placeholder spliced
        /// into it is a template that can silently produce different code than
        /// the one that was reviewed and approved.
        /// </summary>
        private static readonly HashSet<string> NeverParameterizeCode =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "code", "script", "vba", "python" };

        public static Skill Build(string name, string originalRequest, IEnumerable<RecordedStep> recorded)
        {
            var skill = new Skill
            {
                Name = string.IsNullOrWhiteSpace(name) ? "未命名技能" : name.Trim(),
                OriginalRequest = originalRequest
            };

            var steps = (recorded ?? Enumerable.Empty<RecordedStep>())
                .Where(s => s != null && !string.IsNullOrEmpty(s.Tool))
                .ToList();
            if (steps.Count == 0)
            {
                return skill;
            }

            // value -> parameter, so the same range used in three steps becomes
            // one parameter rather than three the user must keep consistent.
            var byValue = new Dictionary<string, SkillParameter>(StringComparer.OrdinalIgnoreCase);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var recordedStep in steps)
            {
                var step = new SkillStep { Tool = recordedStep.Tool };
                foreach (var argument in recordedStep.Arguments ?? new Dictionary<string, object>())
                {
                    var text = Stringify(argument.Value);
                    if (string.IsNullOrEmpty(text))
                    {
                        step.Arguments[argument.Key] = text;
                        continue;
                    }

                    var kind = Classify(argument.Key, text);
                    if (kind == null)
                    {
                        step.Arguments[argument.Key] = text;
                        continue;
                    }

                    if (!byValue.TryGetValue(text, out var parameter))
                    {
                        parameter = new SkillParameter
                        {
                            Name = UniqueName(kind.Value, argument.Key, usedNames),
                            Kind = kind.Value,
                            DefaultValue = text,
                            Label = LabelFor(kind.Value, argument.Key)
                        };
                        byValue[text] = parameter;
                        skill.Parameters.Add(parameter);
                    }
                    step.Arguments[argument.Key] = parameter.Placeholder;
                }
                skill.Steps.Add(step);
            }

            // The original request mentions the same values; substituting there
            // too keeps the intent consistent with the parameters.
            skill.OriginalRequest = ReplaceValuesWithPlaceholders(originalRequest, skill.Parameters);
            return skill;
        }

        /// <summary>Returns the parameter kind, or null to keep the value fixed.</summary>
        internal static ParameterKind? Classify(string argumentName, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            if (NeverParameterizeCode.Contains(argumentName))
            {
                return null;
            }
            if (NeverParameterize.Contains(argumentName))
            {
                return null;
            }
            // Booleans are decisions, not inputs.
            if (bool.TryParse(value, out _))
            {
                return null;
            }

            if (RangeAddress.IsMatch(value) || ColumnOrRowAddress.IsMatch(value))
            {
                return ParameterKind.Range;
            }
            if (FilePath.IsMatch(value))
            {
                return ParameterKind.File;
            }
            if (LooksLikeDate(value))
            {
                return ParameterKind.Date;
            }
            if (IsSheetArgument(argumentName))
            {
                return ParameterKind.Sheet;
            }

            // Numbers are only worth parameterising when the argument name
            // suggests a threshold the user might change. A row count or an
            // index is part of the recorded shape.
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                return LooksLikeThreshold(argumentName) ? ParameterKind.Number : (ParameterKind?)null;
            }

            return null;
        }

        internal static bool LooksLikeDate(string value)
        {
            // Require a separator: a bare 8-digit run is far more often an
            // identifier than a date, and turning an order number into a date
            // parameter would prompt for the wrong thing every run.
            if (value.Length < 6 || value.All(char.IsDigit))
            {
                return false;
            }
            return DateTime.TryParseExact(value, DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _);
        }

        private static bool IsSheetArgument(string argumentName)
        {
            return argumentName != null &&
                   (argumentName.Equals("sheet", StringComparison.OrdinalIgnoreCase) ||
                    argumentName.Equals("sheet_name", StringComparison.OrdinalIgnoreCase) ||
                    argumentName.Equals("worksheet", StringComparison.OrdinalIgnoreCase) ||
                    argumentName.Equals("target_sheet", StringComparison.OrdinalIgnoreCase));
        }

        private static bool LooksLikeThreshold(string argumentName)
        {
            if (argumentName == null) return false;
            var lower = argumentName.ToLowerInvariant();
            return lower.Contains("threshold") || lower.Contains("limit")
                || lower.Contains("min") || lower.Contains("max")
                || lower.Contains("top") || lower.Contains("percent");
        }

        private static string UniqueName(ParameterKind kind, string argumentName, HashSet<string> used)
        {
            var baseName = kind switch
            {
                ParameterKind.Range => "range",
                ParameterKind.Sheet => "sheet",
                ParameterKind.Date => "date",
                ParameterKind.File => "file",
                ParameterKind.Number => argumentName ?? "number",
                _ => argumentName ?? "value"
            };
            baseName = Regex.Replace(baseName, "[^a-zA-Z0-9_]", "_");
            if (used.Add(baseName))
            {
                return baseName;
            }
            for (var i = 2; ; i++)
            {
                var candidate = baseName + i;
                if (used.Add(candidate))
                {
                    return candidate;
                }
            }
        }

        private static string LabelFor(ParameterKind kind, string argumentName)
        {
            switch (kind)
            {
                case ParameterKind.Range: return "数据区域";
                case ParameterKind.Sheet: return "工作表";
                case ParameterKind.Date: return "日期";
                case ParameterKind.File: return "文件";
                default: return argumentName ?? "参数";
            }
        }

        internal static string ReplaceValuesWithPlaceholders(string text, IEnumerable<SkillParameter> parameters)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }
            var result = text;
            // Longest first, so replacing "A1" cannot corrupt "A1:F200".
            foreach (var parameter in parameters.OrderByDescending(p => p.DefaultValue?.Length ?? 0))
            {
                if (string.IsNullOrEmpty(parameter.DefaultValue))
                {
                    continue;
                }
                result = result.Replace(parameter.DefaultValue, parameter.Placeholder);
            }
            return result;
        }

        internal static string Stringify(object value)
        {
            if (value == null) return "";
            if (value is string s) return s;
            if (value is bool b) return b ? "true" : "false";
            if (value is double d)
            {
                return d == Math.Floor(d) && Math.Abs(d) < 1e15
                    ? ((long)d).ToString(CultureInfo.InvariantCulture)
                    : d.ToString("0.####", CultureInfo.InvariantCulture);
            }
            if (value is DateTime dt) return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
    }
}
