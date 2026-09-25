using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// 知识包：签名校验、防回放、防路径穿越、整包替换安装。
    ///
    /// 向量由真正的 Python 签名工具生成（scripts/knowledge_pack.py testvectors），
    /// 与更新清单一样保证两边在 RSA-PSS 细节上一致。
    /// </summary>
    public class KnowledgePackTests : IDisposable
    {
        private const string VectorFile = "knowledge-test-vectors.json";
        private readonly string _root = Path.Combine(Path.GetTempPath(), "deepexcel-knowledge-" + Guid.NewGuid().ToString("N"));

        private string CacheDir => Path.Combine(_root, "knowledge");

        public void Dispose()
        {
            try { Directory.Delete(_root, true); } catch (Exception) { }
        }

        private static string LocateVectors()
        {
            string assemblyDirectory = Path.GetDirectoryName(
                new Uri(typeof(KnowledgePackTests).Assembly.CodeBase).LocalPath);
            foreach (string candidate in new[]
                     {
                         Path.Combine(assemblyDirectory, "fixtures", VectorFile),
                         Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fixtures", VectorFile),
                     })
            {
                if (File.Exists(candidate)) return candidate;
            }
            throw new FileNotFoundException("缺少知识包测试向量：python scripts/knowledge_pack.py testvectors");
        }

        public static IEnumerable<object[]> Cases()
        {
            using (JsonDocument document = JsonDocument.Parse(File.ReadAllText(LocateVectors())))
            {
                foreach (JsonElement element in document.RootElement.GetProperty("cases").EnumerateArray())
                {
                    yield return new object[] { element.GetProperty("name").GetString() };
                }
            }
        }

        private static (IManifestVerifier Verifier, long Installed, Dictionary<string, (string Expect, string Pack)> Cases) Load()
        {
            using (JsonDocument document = JsonDocument.Parse(File.ReadAllText(LocateVectors())))
            {
                JsonElement root = document.RootElement;
                JsonElement key = root.GetProperty("key");
                var cases = new Dictionary<string, (string, string)>();
                foreach (JsonElement element in root.GetProperty("cases").EnumerateArray())
                {
                    cases[element.GetProperty("name").GetString()] =
                        (element.GetProperty("expect").GetString(), element.GetProperty("manifest").GetRawText());
                }
                return (RsaManifestVerifier.FromBase64(
                            key.GetProperty("modulus_b64").GetString(), key.GetProperty("exponent_b64").GetString()),
                        root.GetProperty("installed_pack_version").GetInt64(),
                        cases);
            }
        }

        private static string Pack(string name) => Load().Cases[name].Pack;

        [Theory]
        [MemberData(nameof(Cases))]
        public void Python_signed_pack_gets_the_expected_verdict(string caseName)
        {
            var (verifier, installed, cases) = Load();
            var (expect, pack) = cases[caseName];

            KnowledgePackResult result = KnowledgePack.Evaluate(pack, verifier, installed);

            Assert.Equal(expect == "" ? null : expect, result.Reason);
        }

        [Fact]
        public void Valid_pack_carries_skills_and_utf8_content()
        {
            var (verifier, installed, _) = Load();
            KnowledgePackResult result = KnowledgePack.Evaluate(Pack("valid"), verifier, installed);

            Assert.True(result.Accepted, result.Detail);
            Assert.Equal(100, result.PackVersion);
            KnowledgeSkill skill = Assert.Single(result.Skills);
            Assert.Equal("cn-sample", skill.Name);
            Assert.Equal(3, skill.Version);
            Assert.Contains("11010519491231002X", skill.Files["SKILL.md"]);
            Assert.Contains("全角「，」", skill.Files["SKILL.md"]);
            Assert.Equal("WPS 差异\n", skill.Files["wps.md"]);
        }

        [Fact]
        public void Unconfigured_verifier_refuses_everything()
        {
            IManifestVerifier none = RsaManifestVerifier.FromBase64("", "");
            Assert.Equal("no_key", KnowledgePack.Evaluate(Pack("valid"), none, 0).Reason);
        }

        [Fact]
        public void Install_writes_skills_and_stamp_then_blocks_replay()
        {
            var (verifier, installed, _) = Load();
            KnowledgePackResult result = KnowledgePack.Evaluate(Pack("valid"), verifier, installed);

            KnowledgePack.Install(result, CacheDir, verifier.KeyId);

            string skillMd = File.ReadAllText(Path.Combine(CacheDir, "cn-sample", "SKILL.md"), Encoding.UTF8);
            Assert.StartsWith("---\nname: cn-sample", skillMd);
            Assert.True(File.Exists(Path.Combine(CacheDir, "cn-sample", "wps.md")));
            Assert.Equal(100, KnowledgePack.InstalledPackVersion(CacheDir));
            // 同一个包再来一次：不是更新的，不装
            Assert.Equal("not_newer",
                KnowledgePack.Evaluate(Pack("valid"), verifier, KnowledgePack.InstalledPackVersion(CacheDir)).Reason);
            // 没有留下临时目录
            Assert.Single(Directory.GetDirectories(_root));
        }

        [Fact]
        public void Install_replaces_the_whole_cache()
        {
            var (verifier, installed, _) = Load();
            Directory.CreateDirectory(Path.Combine(CacheDir, "withdrawn-skill"));
            File.WriteAllText(Path.Combine(CacheDir, "withdrawn-skill", "SKILL.md"), "旧技能");

            KnowledgePack.Install(KnowledgePack.Evaluate(Pack("valid"), verifier, installed), CacheDir, verifier.KeyId);

            Assert.False(Directory.Exists(Path.Combine(CacheDir, "withdrawn-skill")));
            Assert.True(Directory.Exists(Path.Combine(CacheDir, "cn-sample")));
        }

        [Fact]
        public void Install_refuses_a_rejected_pack()
        {
            Assert.Throws<ArgumentException>(() =>
                KnowledgePack.Install(KnowledgePackResult.Reject("bad_signature", "x"), CacheDir, "k"));
            Assert.False(Directory.Exists(CacheDir));
        }

        [Fact]
        public void Missing_or_corrupt_stamp_means_version_zero()
        {
            Assert.Equal(0, KnowledgePack.InstalledPackVersion(CacheDir));
            Directory.CreateDirectory(CacheDir);
            File.WriteAllText(Path.Combine(CacheDir, KnowledgePack.StampFile), "not json");
            Assert.Equal(0, KnowledgePack.InstalledPackVersion(CacheDir));
        }

        [Theory]
        [InlineData("https://api.example.com/api/v1/updates/latest", "https://api.example.com/api/v1/updates/knowledge")]
        [InlineData("https://api.example.com/api/v1/updates/latest?channel=beta", "https://api.example.com/api/v1/updates/knowledge")]
        [InlineData("https://api.example.com:8443/api/v1/updates/latest/", "https://api.example.com:8443/api/v1/updates/knowledge")]
        [InlineData("http://api.example.com/api/v1/updates/latest", null)]
        [InlineData("", null)]
        [InlineData(null, null)]
        [InlineData("not a url", null)]
        public void Pack_url_is_a_sibling_of_the_update_feed(string feed, string expected)
        {
            Assert.Equal(expected, KnowledgePack.UrlForFeed(feed));
        }

        [Fact]
        public async Task Sync_installs_once_and_then_reports_not_newer()
        {
            var (verifier, _, _) = Load();
            string body = Pack("valid");
            int fetches = 0;
            var sync = new KnowledgeSync((url, ct) => { fetches++; return Task.FromResult(body); }, verifier, CacheDir);

            KnowledgePackResult first = await sync.RunOnceAsync("https://x.example/api/v1/updates/knowledge", CancellationToken.None);
            KnowledgePackResult second = await sync.RunOnceAsync("https://x.example/api/v1/updates/knowledge", CancellationToken.None);

            Assert.True(first.Accepted, first.Detail);
            Assert.Equal("not_newer", second.Reason);
            Assert.Equal(2, fetches);
            Assert.True(File.Exists(Path.Combine(CacheDir, "cn-sample", "SKILL.md")));
        }

        [Fact]
        public async Task Sync_does_not_touch_the_network_without_a_key_or_a_feed()
        {
            int fetches = 0;
            Func<string, CancellationToken, Task<string>> fetch = (url, ct) => { fetches++; return Task.FromResult(""); };

            var noKey = new KnowledgeSync(fetch, RsaManifestVerifier.FromBase64("", ""), CacheDir);
            Assert.Equal("no_key", (await noKey.RunOnceAsync("https://x.example/k", CancellationToken.None)).Reason);

            var noFeed = new KnowledgeSync(fetch, Load().Verifier, CacheDir);
            Assert.Equal("not_configured", (await noFeed.RunOnceAsync(null, CancellationToken.None)).Reason);

            Assert.Equal(0, fetches);
        }

        [Fact]
        public async Task Sync_survives_transport_failures_and_unpublished_packs()
        {
            var verifier = Load().Verifier;
            var failing = new KnowledgeSync((url, ct) => throw new IOException("网络断了"), verifier, CacheDir);
            Assert.Equal("transport", (await failing.RunOnceAsync("https://x.example/k", CancellationToken.None)).Reason);

            var unpublished = new KnowledgeSync((url, ct) => Task.FromResult<string>(null), verifier, CacheDir);
            Assert.Equal("not_published", (await unpublished.RunOnceAsync("https://x.example/k", CancellationToken.None)).Reason);

            Assert.False(Directory.Exists(CacheDir));
        }

        [Fact]
        public async Task Tampered_pack_leaves_the_installed_one_alone()
        {
            var (verifier, _, _) = Load();
            await new KnowledgeSync((u, c) => Task.FromResult(Pack("valid")), verifier, CacheDir)
                .RunOnceAsync("https://x.example/k", CancellationToken.None);

            KnowledgePackResult result = await new KnowledgeSync(
                    (u, c) => Task.FromResult(Pack("tampered_payload")), verifier, CacheDir)
                .RunOnceAsync("https://x.example/k", CancellationToken.None);

            Assert.Equal("bad_signature", result.Reason);
            Assert.DoesNotContain("忽略之前",
                File.ReadAllText(Path.Combine(CacheDir, "cn-sample", "SKILL.md"), Encoding.UTF8));
        }
    }
}
