using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>一个知识技能：技能名、版本、目录里的 .md 文件（文件名 → 内容）</summary>
    public sealed class KnowledgeSkill
    {
        public string Name { get; set; }
        public long Version { get; set; }
        public Dictionary<string, string> Files { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }

    public sealed class KnowledgePackResult
    {
        public bool Accepted => Reason == null;
        /// <summary>拒绝原因的固定代码（no_key / too_large / bad_signature / not_newer ...），接受时为 null</summary>
        public string Reason { get; set; }
        public string Detail { get; set; }
        public long PackVersion { get; set; }
        public List<KnowledgeSkill> Skills { get; set; } = new List<KnowledgeSkill>();

        public static KnowledgePackResult Reject(string reason, string detail)
        {
            return new KnowledgePackResult { Reason = reason, Detail = detail };
        }
    }

    /// <summary>
    /// 服务端下发的知识技能包（侧车 knowledge_skills.py 读取的缓存目录由这里写入）。
    ///
    /// 知识正文会进入模型的上下文，等于给 agent 的指令：服务端被攻破时如果能随便下发，
    /// 就是一条提示词注入通道。所以它和更新清单用同一把离线私钥签名、同一个内置公钥校验，
    /// 服务端只转发已签名的文件（scripts/knowledge_pack.py 生成）。没有内置公钥的构建不同步，
    /// 只用安装包自带的技能。
    ///
    /// 已签名负载：
    ///     { "schema": 1, "product": "DeepExcel", "kind": "knowledge", "pack_version": 3,
    ///       "released_at": "...", "skills": [ { "name": "...", "version": 2,
    ///       "files": { "SKILL.md": "...", "wps.md": "..." } } ] }
    ///
    /// kind 把它和更新清单互相隔开：更新清单没有 kind，知识包没有 channel / url，
    /// 任何一方都不能被当成另一方重放。pack_version 必须严格递增，防止回放旧包。
    /// </summary>
    public static class KnowledgePack
    {
        public const int Schema = 1;
        public const string Kind = "knowledge";
        public const int MaxPackBytes = 1024 * 1024;
        public const int MaxSkills = 64;
        public const int MaxFilesPerSkill = 16;
        public const int MaxFileChars = 60000;
        public const string StampFile = "pack.json";

        private static readonly Regex SkillName = new Regex(@"^[a-z0-9][a-z0-9\-]{0,63}$", RegexOptions.CultureInvariant);
        private static readonly Regex FileName = new Regex(@"^[A-Za-z0-9_\-]{1,64}\.md$", RegexOptions.CultureInvariant);

        /// <summary>与侧车 knowledge_skills.cache_dir() 一致</summary>
        public static string DefaultCacheDir()
        {
            string overrideDir = Environment.GetEnvironmentVariable("DEEPEXCEL_KNOWLEDGE_DIR");
            if (!string.IsNullOrEmpty(overrideDir))
            {
                return overrideDir;
            }
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DeepExcel", "knowledge");
        }

        /// <summary>
        /// 知识包地址从更新源推出来：…/api/v1/updates/latest → …/api/v1/updates/knowledge。
        /// 不另设配置项：更新源没配置的安装，本来就不该去联系任何地址。
        /// </summary>
        public static string UrlForFeed(string feedUrl)
        {
            if (string.IsNullOrWhiteSpace(feedUrl) ||
                !Uri.TryCreate(feedUrl.Trim(), UriKind.Absolute, out Uri uri) ||
                !string.Equals(uri.Scheme, "https", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            string path = uri.AbsolutePath.TrimEnd('/');
            int slash = path.LastIndexOf('/');
            if (slash < 0)
            {
                return null;
            }
            var builder = new UriBuilder(uri) { Path = path.Substring(0, slash + 1) + "knowledge", Query = "" };
            return builder.Uri.AbsoluteUri;
        }

        public static KnowledgePackResult Evaluate(string packJson, IManifestVerifier verifier, long installedPackVersion)
        {
            if (verifier == null || !verifier.IsConfigured)
            {
                return KnowledgePackResult.Reject("no_key", "此版本未内置签名公钥，不同步知识包。");
            }
            if (string.IsNullOrWhiteSpace(packJson))
            {
                return KnowledgePackResult.Reject("malformed", "知识包为空。");
            }
            if (packJson.Length > MaxPackBytes * 2)
            {
                return KnowledgePackResult.Reject("too_large", "知识包过大。");
            }

            Dictionary<string, object> envelope = ParseObject(packJson);
            if (envelope == null)
            {
                return KnowledgePackResult.Reject("malformed", "知识包不是合法 JSON 对象。");
            }
            if (!TryGetInt64(envelope, "schema", out long schema) || schema != Schema)
            {
                return KnowledgePackResult.Reject("unsupported_schema", "知识包格式版本不受支持。");
            }
            if (!string.Equals(GetString(envelope, "key_id"), verifier.KeyId, StringComparison.Ordinal))
            {
                return KnowledgePackResult.Reject("unknown_key", "知识包由未知密钥签名。");
            }
            if (!TryDecodeBase64(GetString(envelope, "payload"), out byte[] payloadBytes) ||
                payloadBytes.Length == 0 || payloadBytes.Length > MaxPackBytes)
            {
                return KnowledgePackResult.Reject("malformed", "知识包负载无法解码或过大。");
            }
            if (!TryDecodeBase64(GetString(envelope, "signature"), out byte[] signature))
            {
                return KnowledgePackResult.Reject("malformed", "知识包签名无法解码。");
            }

            // ---- trust boundary -------------------------------------------
            if (!verifier.Verify(payloadBytes, signature))
            {
                return KnowledgePackResult.Reject("bad_signature", "知识包签名校验失败，已拒绝。");
            }
            // ---- everything below here is authenticated --------------------

            Dictionary<string, object> payload = ParseObject(DecodeUtf8(payloadBytes));
            if (payload == null)
            {
                return KnowledgePackResult.Reject("malformed", "已签名负载不是合法 JSON 对象。");
            }
            if (!TryGetInt64(payload, "schema", out long payloadSchema) || payloadSchema != Schema)
            {
                return KnowledgePackResult.Reject("unsupported_schema", "已签名负载格式版本不受支持。");
            }
            if (!string.Equals(GetString(payload, "product"), UpdateManifest.Product, StringComparison.Ordinal) ||
                !string.Equals(GetString(payload, "kind"), Kind, StringComparison.Ordinal))
            {
                return KnowledgePackResult.Reject("wrong_kind", "已签名文件不是本产品的知识包。");
            }
            if (!TryGetInt64(payload, "pack_version", out long packVersion) || packVersion <= 0)
            {
                return KnowledgePackResult.Reject("malformed", "知识包没有有效的 pack_version。");
            }
            if (packVersion <= installedPackVersion)
            {
                return new KnowledgePackResult
                {
                    Reason = "not_newer",
                    PackVersion = packVersion,
                    Detail = string.Format(CultureInfo.InvariantCulture,
                        "知识包已是最新（本机 {0}，服务端 {1}）。", installedPackVersion, packVersion),
                };
            }

            if (!(payload.TryGetValue("skills", out object rawSkills) && rawSkills is IList skillList) ||
                skillList.Count > MaxSkills)
            {
                return KnowledgePackResult.Reject("malformed", "知识包的 skills 缺失或过多。");
            }

            var result = new KnowledgePackResult { PackVersion = packVersion };
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (object rawSkill in skillList)
            {
                string problem = ReadSkill(rawSkill as Dictionary<string, object>, out KnowledgeSkill skill);
                if (problem != null)
                {
                    return KnowledgePackResult.Reject("malformed", problem);
                }
                if (!seen.Add(skill.Name))
                {
                    return KnowledgePackResult.Reject("malformed", "知识包里有重名技能：" + skill.Name);
                }
                result.Skills.Add(skill);
            }
            return result;
        }

        /// <summary>
        /// 签名只证明内容来自我们；名字和文件名仍然按白名单校验，免得一个签错的包
        /// 把文件写到缓存目录外面。
        /// </summary>
        private static string ReadSkill(Dictionary<string, object> map, out KnowledgeSkill skill)
        {
            skill = null;
            if (map == null)
            {
                return "知识包里有非对象的技能条目。";
            }
            string name = GetString(map, "name");
            if (name == null || !SkillName.IsMatch(name))
            {
                return "技能名不合法：" + Truncate(name);
            }
            if (!TryGetInt64(map, "version", out long version) || version <= 0)
            {
                return "技能 " + name + " 没有有效的 version。";
            }
            if (!(map.TryGetValue("files", out object rawFiles) && rawFiles is Dictionary<string, object> files) ||
                files.Count == 0 || files.Count > MaxFilesPerSkill)
            {
                return "技能 " + name + " 的 files 缺失或过多。";
            }
            if (!files.ContainsKey("SKILL.md"))
            {
                return "技能 " + name + " 缺少 SKILL.md。";
            }
            skill = new KnowledgeSkill { Name = name, Version = version };
            foreach (KeyValuePair<string, object> file in files)
            {
                if (!FileName.IsMatch(file.Key))
                {
                    skill = null;
                    return "技能 " + name + " 的文件名不合法：" + Truncate(file.Key);
                }
                if (!(file.Value is string content) || content.Length > MaxFileChars)
                {
                    skill = null;
                    return "技能 " + name + " 的文件 " + file.Key + " 内容无效或过长。";
                }
                skill.Files[file.Key] = content;
            }
            return null;
        }

        /// <summary>本机已安装的知识包版本；没装过或读不出来时为 0</summary>
        public static long InstalledPackVersion(string cacheDir)
        {
            try
            {
                string text = File.ReadAllText(Path.Combine(cacheDir, StampFile), Encoding.UTF8);
                Dictionary<string, object> stamp = ParseObject(text);
                return TryGetInt64(stamp, "pack_version", out long version) ? version : 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// 整包替换：先写到旁边的临时目录，写完再换名。服务端从包里去掉的技能随之消失
        /// （内置的那份仍在，侧车会退回内置版本）。侧车读取时如果正好撞上换名的一瞬，
        /// 只会少看到缓存、退回内置，不会读到半个技能。
        /// </summary>
        public static void Install(KnowledgePackResult pack, string cacheDir, string keyId)
        {
            if (pack == null || !pack.Accepted)
            {
                throw new ArgumentException("只能安装已校验通过的知识包", nameof(pack));
            }
            string parent = Path.GetDirectoryName(Path.GetFullPath(cacheDir));
            Directory.CreateDirectory(parent);
            string token = Guid.NewGuid().ToString("N").Substring(0, 8);
            string staging = cacheDir + ".staging-" + token;
            string retired = cacheDir + ".old-" + token;

            try
            {
                Directory.CreateDirectory(staging);
                var utf8 = new UTF8Encoding(false);
                foreach (KnowledgeSkill skill in pack.Skills)
                {
                    string dir = Path.Combine(staging, skill.Name);
                    Directory.CreateDirectory(dir);
                    foreach (KeyValuePair<string, string> file in skill.Files)
                    {
                        File.WriteAllText(Path.Combine(dir, file.Key), file.Value, utf8);
                    }
                }
                string stamp = new JavaScriptSerializer().Serialize(new Dictionary<string, object>
                {
                    ["pack_version"] = pack.PackVersion,
                    ["key_id"] = keyId ?? "",
                    ["installed_at"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                });
                File.WriteAllText(Path.Combine(staging, StampFile), stamp, utf8);

                if (Directory.Exists(cacheDir))
                {
                    Directory.Move(cacheDir, retired);
                }
                Directory.Move(staging, cacheDir);
            }
            catch
            {
                // 换名失败时把旧目录放回去，不能让用户落得一个技能都没有
                if (!Directory.Exists(cacheDir) && Directory.Exists(retired))
                {
                    try { Directory.Move(retired, cacheDir); } catch (Exception) { }
                }
                TryDeleteDirectory(staging);
                throw;
            }
            TryDeleteDirectory(retired);
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch (Exception)
            {
            }
        }

        // ------------------------------------------------------------------

        private static Dictionary<string, object> ParseObject(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }
            try
            {
                var serializer = new JavaScriptSerializer { MaxJsonLength = MaxPackBytes * 2 };
                return serializer.DeserializeObject(json) as Dictionary<string, object>;
            }
            catch (ArgumentException)
            {
                return null;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static string DecodeUtf8(byte[] bytes)
        {
            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static bool TryDecodeBase64(string text, out byte[] bytes)
        {
            bytes = null;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            try
            {
                bytes = Convert.FromBase64String(text.Trim());
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static string GetString(Dictionary<string, object> map, string key)
        {
            return map != null && map.TryGetValue(key, out object value) ? value as string : null;
        }

        private static bool TryGetInt64(Dictionary<string, object> map, string key, out long value)
        {
            value = 0;
            if (map == null || !map.TryGetValue(key, out object raw) || raw == null)
            {
                return false;
            }
            if (raw is int i) { value = i; return true; }
            if (raw is long l) { value = l; return true; }
            if (raw is decimal d)
            {
                if (d != decimal.Truncate(d) || d > long.MaxValue || d < long.MinValue) return false;
                value = (long)d;
                return true;
            }
            return false;
        }

        private static string Truncate(string value)
        {
            if (string.IsNullOrEmpty(value)) return "(空)";
            return value.Length <= 64 ? value : value.Substring(0, 64) + "…";
        }
    }
}
