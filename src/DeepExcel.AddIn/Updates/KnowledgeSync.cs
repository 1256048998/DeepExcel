using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>
    /// 拉取并安装服务端下发的知识包（见 <see cref="KnowledgePack"/>）。
    ///
    /// 挂在后台更新检查上跑，节奏也一样：启动 1 分钟后一次，之后每 6 小时一次。
    /// 失败只写日志：知识包只是让技能更新得更快，安装包自带的那一份一直可用。
    /// </summary>
    public sealed class KnowledgeSync
    {
        private static readonly HttpClient SharedHttp = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        private readonly Func<string, CancellationToken, Task<string>> _fetch;
        private readonly IManifestVerifier _verifier;
        private readonly string _cacheDir;

        public KnowledgeSync(
            Func<string, CancellationToken, Task<string>> fetch = null,
            IManifestVerifier verifier = null,
            string cacheDir = null)
        {
            _fetch = fetch ?? FetchAsync;
            _verifier = verifier ?? EmbeddedUpdateKey.Verifier;
            _cacheDir = cacheDir ?? KnowledgePack.DefaultCacheDir();
        }

        public async Task<KnowledgePackResult> RunOnceAsync(string packUrl, CancellationToken cancellationToken)
        {
            if (!_verifier.IsConfigured)
            {
                return KnowledgePackResult.Reject("no_key", "此版本未内置签名公钥，不同步知识包。");
            }
            if (string.IsNullOrEmpty(packUrl))
            {
                return KnowledgePackResult.Reject("not_configured", "未配置更新源，不同步知识包。");
            }

            string body;
            try
            {
                body = await _fetch(packUrl, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex) when (!(ex is OperationCanceledException && cancellationToken.IsCancellationRequested))
            {
                Logger.Instance.Warning("Knowledge", "知识包下载失败：" + ex.Message);
                return KnowledgePackResult.Reject("transport", ex.Message);
            }
            if (body == null)
            {
                // 服务端没发布知识包（404）：正常情况
                return KnowledgePackResult.Reject("not_published", "服务端没有发布知识包。");
            }

            KnowledgePackResult pack = KnowledgePack.Evaluate(
                body, _verifier, KnowledgePack.InstalledPackVersion(_cacheDir));
            if (!pack.Accepted)
            {
                if (pack.Reason != "not_newer")
                {
                    Logger.Instance.Warning("Knowledge", "知识包已拒绝（" + pack.Reason + "）：" + pack.Detail);
                }
                return pack;
            }

            try
            {
                KnowledgePack.Install(pack, _cacheDir, _verifier.KeyId);
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("Knowledge", "知识包安装失败：" + ex.Message);
                return KnowledgePackResult.Reject("install_failed", ex.Message);
            }
            Logger.Instance.Info("Knowledge",
                "知识包已更新到 " + pack.PackVersion + "（" + pack.Skills.Count + " 个技能）");
            return pack;
        }

        /// <summary>https 才拉；404 返回 null；超过上限直接放弃，不读完</summary>
        private static async Task<string> FetchAsync(string url, CancellationToken cancellationToken)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri uri) ||
                !string.Equals(uri.Scheme, "https", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("知识包地址必须是 https：" + url);
            }
            using (HttpResponseMessage response = await SharedHttp
                       .GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                       .ConfigureAwait(false))
            {
                if ((int)response.StatusCode == 404)
                {
                    return null;
                }
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("知识包源返回 " + (int)response.StatusCode);
                }
                long limit = KnowledgePack.MaxPackBytes * 2L;
                if (response.Content.Headers.ContentLength > limit)
                {
                    throw new InvalidOperationException("知识包过大，已拒绝下载。");
                }
                using (Stream stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var buffer = new MemoryStream())
                {
                    var chunk = new byte[81920];
                    int read;
                    while ((read = await stream.ReadAsync(chunk, 0, chunk.Length, cancellationToken)
                               .ConfigureAwait(false)) > 0)
                    {
                        buffer.Write(chunk, 0, read);
                        if (buffer.Length > limit)
                        {
                            throw new InvalidOperationException("知识包过大，已拒绝下载。");
                        }
                    }
                    return new UTF8Encoding(false, true).GetString(buffer.ToArray());
                }
            }
        }
    }
}
