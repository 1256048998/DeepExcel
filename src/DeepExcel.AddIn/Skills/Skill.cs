using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Skills
{
    /// <summary>What a captured argument turned into.</summary>
    public enum ParameterKind
    {
        /// <summary>A cell or range address such as A1:F200.</summary>
        Range,
        /// <summary>A worksheet name.</summary>
        Sheet,
        /// <summary>A date, which almost always differs between runs.</summary>
        Date,
        /// <summary>A file path.</summary>
        File,
        /// <summary>A plain number the user may want to vary.</summary>
        Number,
        /// <summary>Free text.</summary>
        Text
    }

    public sealed class SkillParameter
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("kind")]
        public ParameterKind Kind { get; set; }

        /// <summary>The value seen during recording; used as the default.</summary>
        [JsonPropertyName("default_value")]
        public string DefaultValue { get; set; }

        [JsonPropertyName("label")]
        public string Label { get; set; }

        public string Placeholder
        {
            get { return "{{" + Name + "}}"; }
        }
    }

    public sealed class SkillStep
    {
        [JsonPropertyName("tool")]
        public string Tool { get; set; }

        /// <summary>
        /// Arguments with recorded values replaced by <c>{{placeholders}}</c>.
        /// </summary>
        [JsonPropertyName("arguments")]
        public Dictionary<string, string> Arguments { get; set; } = new Dictionary<string, string>();

        public override string ToString()
        {
            return Tool + (Arguments.Count == 0
                ? ""
                : " (" + string.Join(", ", Arguments.Select(a => a.Key + "=" + a.Value).ToArray()) + ")");
        }
    }

    /// <summary>
    /// A successful multi-step task, saved so it can be run again.
    ///
    /// The product is currently a series of one-off conversations: a user who
    /// builds the same report every month describes it from scratch every month,
    /// and nothing accumulates. A skill library is the first thing the user owns
    /// rather than rents -- which is also why it is the most defensible reason
    /// to pay for this later.
    ///
    /// A skill is a skeleton, not a macro. Replay still goes through the agent,
    /// so it can cope with data that does not look exactly like it did during
    /// recording; the steps tell it what was done, not how to blindly repeat it.
    /// </summary>
    public sealed class Skill
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        /// <summary>The instruction that produced this run, kept as intent.</summary>
        [JsonPropertyName("original_request")]
        public string OriginalRequest { get; set; }

        [JsonPropertyName("parameters")]
        public List<SkillParameter> Parameters { get; set; } = new List<SkillParameter>();

        [JsonPropertyName("steps")]
        public List<SkillStep> Steps { get; set; } = new List<SkillStep>();

        [JsonPropertyName("created_at")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [JsonPropertyName("run_count")]
        public int RunCount { get; set; }

        [JsonPropertyName("last_run_at")]
        public DateTime? LastRunAt { get; set; }

        /// <summary>
        /// Renders the skill as an instruction for the agent.
        ///
        /// Deliberately a prompt rather than a direct tool-call replay. Replaying
        /// the recorded calls verbatim breaks the moment the data differs from
        /// the recording -- a different row count, a renamed column - which is
        /// most of the time. Handing the agent the intent plus the steps that
        /// worked lets it adapt while still following the same shape.
        /// </summary>
        public string ToPrompt(IDictionary<string, string> arguments)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var parameter in Parameters)
            {
                values[parameter.Name] =
                    arguments != null && arguments.TryGetValue(parameter.Name, out var supplied)
                    && !string.IsNullOrWhiteSpace(supplied)
                        ? supplied
                        : parameter.DefaultValue;
            }

            var builder = new System.Text.StringBuilder();
            builder.Append("执行技能「").Append(Name).AppendLine("」。");

            if (!string.IsNullOrWhiteSpace(OriginalRequest))
            {
                builder.Append("原始需求：").AppendLine(Substitute(OriginalRequest, values));
            }

            if (values.Count > 0)
            {
                builder.AppendLine("参数：");
                foreach (var parameter in Parameters)
                {
                    builder.Append("  ").Append(parameter.Label ?? parameter.Name)
                           .Append(" = ").AppendLine(values[parameter.Name]);
                }
            }

            builder.AppendLine("上次成功执行的步骤（供参考，实际数据可能已变化，按需调整）：");
            for (var i = 0; i < Steps.Count; i++)
            {
                builder.Append("  ").Append(i + 1).Append(". ")
                       .AppendLine(Substitute(Steps[i].ToString(), values));
            }

            builder.AppendLine(
                "请按上述意图完成任务。若当前数据结构与步骤不符，以意图为准，不要机械重放。");
            return builder.ToString();
        }

        internal static string Substitute(string template, IDictionary<string, string> values)
        {
            if (string.IsNullOrEmpty(template) || values == null)
            {
                return template;
            }
            var result = template;
            foreach (var pair in values)
            {
                result = result.Replace("{{" + pair.Key + "}}", pair.Value ?? "");
            }
            return result;
        }
    }
}
