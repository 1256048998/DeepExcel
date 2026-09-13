using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace DeepExcel.AddIn.Skills
{
    /// <summary>
    /// Local skill library.
    ///
    /// One file per skill rather than a single index: a corrupt write then costs
    /// one skill instead of the whole library, and skills can be shared by
    /// sending a file. Sharing is the point -- a library that cannot leave the
    /// machine that made it does not compound.
    /// </summary>
    public sealed class SkillStore
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        private readonly string _directory;

        public SkillStore(string directory = null)
        {
            _directory = directory ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DeepExcel", "skills");
        }

        public string Directory_ForTests => _directory;

        public List<Skill> List()
        {
            var skills = new List<Skill>();
            try
            {
                if (!Directory.Exists(_directory))
                {
                    return skills;
                }
                foreach (var file in Directory.GetFiles(_directory, "*.json"))
                {
                    var skill = TryRead(file);
                    if (skill != null)
                    {
                        skills.Add(skill);
                    }
                }
            }
            catch (Exception)
            {
            }
            // Most-used first: the library is only useful if the skill someone
            // reaches for is the one they see.
            return skills
                .OrderByDescending(s => s.RunCount)
                .ThenByDescending(s => s.LastRunAt ?? s.CreatedAt)
                .ToList();
        }

        public Skill Get(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return null;
            }
            return TryRead(PathFor(id));
        }

        public bool Save(Skill skill)
        {
            if (skill == null || string.IsNullOrEmpty(skill.Id))
            {
                return false;
            }
            try
            {
                Directory.CreateDirectory(_directory);
                var path = PathFor(skill.Id);
                var temporary = path + ".tmp";
                File.WriteAllText(temporary, JsonSerializer.Serialize(skill, JsonOptions), new UTF8Encoding(false));
                // Write-then-replace: an interrupted save must not leave a
                // truncated file that silently loses the skill.
                if (File.Exists(path))
                {
                    File.Replace(temporary, path, null);
                }
                else
                {
                    File.Move(temporary, path);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool Delete(string id)
        {
            try
            {
                var path = PathFor(id);
                if (File.Exists(path))
                {
                    File.Delete(path);
                    return true;
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        /// <summary>Bumps usage counters after a successful replay.</summary>
        public void RecordRun(string id)
        {
            var skill = Get(id);
            if (skill == null)
            {
                return;
            }
            skill.RunCount++;
            skill.LastRunAt = DateTime.UtcNow;
            Save(skill);
        }

        private Skill TryRead(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    return null;
                }
                var skill = JsonSerializer.Deserialize<Skill>(File.ReadAllText(path));
                // A skill with no steps cannot be replayed; surfacing it in the
                // list would just produce a confusing failure later.
                return skill != null && skill.Steps != null && skill.Steps.Count > 0 ? skill : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private string PathFor(string id)
        {
            // Ids are generated as hex, but a hand-edited file must not be able
            // to escape the directory.
            var safe = new string(id.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_').ToArray());
            if (string.IsNullOrEmpty(safe))
            {
                safe = "invalid";
            }
            return Path.Combine(_directory, safe + ".json");
        }
    }
}
