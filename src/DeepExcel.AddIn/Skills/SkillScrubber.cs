using System;
using System.Text.RegularExpressions;

namespace DeepExcel.AddIn.Skills
{
    /// <summary>
    /// Removes anything that identifies a file on this machine before a skill
    /// leaves it.
    ///
    /// A skill's parameter defaults are values captured from a real run, so a
    /// File parameter holds a real local path and the original request can name
    /// a real workbook. Syncing those verbatim would put someone's file layout
    /// in a database; sharing would hand it to a stranger.
    ///
    /// The server scrubs again. This is not redundancy for its own sake: doing
    /// it here means the data never leaves the machine, and doing it there means
    /// a modified client cannot skip it.
    /// </summary>
    public static class SkillScrubber
    {
        private const string Redacted = "[已移除路径]";

        // Drive-letter and UNC paths.
        private static readonly Regex PathPattern = new Regex(
            @"(?:[a-zA-Z]:\\|\\\\)[^""'\s]*", RegexOptions.Compiled);

        // A bare workbook name still identifies a project, a client or a quarter.
        private static readonly Regex WorkbookPattern = new Regex(
            @"[^\s\\/""']+\.(?:xlsx|xlsm|xls|csv)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static string Scrub(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            var cleaned = PathPattern.Replace(value, Redacted);
            return WorkbookPattern.Replace(cleaned, Redacted);
        }

        /// <summary>
        /// Returns a copy safe to upload. The original is left untouched so the
        /// local skill keeps its useful defaults.
        /// </summary>
        public static Skill ForUpload(Skill skill)
        {
            if (skill == null)
            {
                return null;
            }

            var copy = new Skill
            {
                Id = skill.Id,
                Name = Scrub(skill.Name),
                Description = Scrub(skill.Description),
                OriginalRequest = Scrub(skill.OriginalRequest),
                CreatedAt = skill.CreatedAt,
                RunCount = skill.RunCount,
                LastRunAt = skill.LastRunAt
            };

            foreach (var parameter in skill.Parameters)
            {
                copy.Parameters.Add(new SkillParameter
                {
                    Name = parameter.Name,
                    Kind = parameter.Kind,
                    Label = parameter.Label,
                    // A file default is worthless to a recipient -- their file is
                    // somewhere else -- and harmful to keep, so it never travels.
                    DefaultValue = parameter.Kind == ParameterKind.File
                        ? ""
                        : Scrub(parameter.DefaultValue)
                });
            }

            foreach (var step in skill.Steps)
            {
                var copiedStep = new SkillStep { Tool = step.Tool };
                foreach (var argument in step.Arguments)
                {
                    copiedStep.Arguments[argument.Key] = Scrub(argument.Value);
                }
                copy.Steps.Add(copiedStep);
            }

            return copy;
        }

        /// <summary>
        /// Whether uploading would strip something, so the user can be told
        /// rather than surprised.
        /// </summary>
        public static bool WouldRedact(Skill skill)
        {
            if (skill == null)
            {
                return false;
            }
            if (Scrub(skill.Name) != skill.Name ||
                Scrub(skill.OriginalRequest) != skill.OriginalRequest)
            {
                return true;
            }
            foreach (var parameter in skill.Parameters)
            {
                if (parameter.Kind == ParameterKind.File && !string.IsNullOrEmpty(parameter.DefaultValue))
                {
                    return true;
                }
                if (Scrub(parameter.DefaultValue) != parameter.DefaultValue)
                {
                    return true;
                }
            }
            foreach (var step in skill.Steps)
            {
                foreach (var argument in step.Arguments)
                {
                    if (Scrub(argument.Value) != argument.Value)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
