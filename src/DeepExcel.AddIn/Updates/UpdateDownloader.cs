using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>Raised when the update feed or package cannot be retrieved.</summary>
    public sealed class UpdateTransportException : Exception
    {
        public UpdateTransportException(string message) : base(message) { }
        public UpdateTransportException(string message, Exception inner) : base(message, inner) { }
    }

    /// <summary>
    /// Fetches the manifest and the package over HTTPS.
    ///
    /// This layer is deliberately dumb: it moves bytes and enforces limits, and
    /// decides nothing. Whether the bytes are trustworthy is settled by
    /// <see cref="UpdateManifest"/> and <see cref="UpdateStage.VerifyPackage"/>,
    /// so a hostile or merely broken server can waste bandwidth and nothing more.
    /// </summary>
    public sealed class UpdateDownloader : IDisposable
    {
        private readonly HttpClient _http;

        public UpdateDownloader(HttpMessageHandler handler = null)
        {
            _http = handler == null ? new HttpClient() : new HttpClient(handler, disposeHandler: true);
            // Generous, because the package is tens of megabytes and this runs
            // in the background where a slow link costs the user nothing.
            _http.Timeout = TimeSpan.FromMinutes(15);
            _http.DefaultRequestHeaders.TryAddWithoutValidation(
                "User-Agent", "DeepExcel-Updater/" + UpdateService.InstalledVersion);
        }

        /// <summary>Downloads the update feed. The response is untrusted text until verified.</summary>
        public async Task<string> FetchManifestAsync(string feedUrl, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(feedUrl))
            {
                throw new UpdateTransportException("未配置更新源地址。");
            }
            if (!Uri.TryCreate(feedUrl, UriKind.Absolute, out Uri uri) ||
                !string.Equals(uri.Scheme, "https", StringComparison.OrdinalIgnoreCase))
            {
                throw new UpdateTransportException("更新源地址必须是 https：" + feedUrl);
            }

            try
            {
                using (var response = await _http
                           .GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                           .ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new UpdateTransportException(string.Format(
                            CultureInfo.InvariantCulture,
                            "更新源返回 {0}。", (int)response.StatusCode));
                    }
                    if (response.Content.Headers.ContentLength > UpdateManifest.MaxManifestBytes)
                    {
                        throw new UpdateTransportException("更新清单过大，已拒绝下载。");
                    }

                    using (Stream stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var buffer = new MemoryStream())
                    {
                        await CopyCappedAsync(
                            stream, buffer, UpdateManifest.MaxManifestBytes, null, cancellationToken)
                            .ConfigureAwait(false);
                        return new UTF8Encoding(false).GetString(buffer.ToArray());
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                throw new UpdateTransportException("无法连接更新源：" + ex.Message, ex);
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                throw new UpdateTransportException("连接更新源超时。", ex);
            }
        }

        /// <summary>
        /// Downloads the package to <paramref name="destinationPath"/>.
        ///
        /// Written to a .part file and only then moved into place, so an
        /// interrupted download can never be mistaken for a finished one.
        /// </summary>
        public async Task DownloadPackageAsync(
            UpdateRelease release, string destinationPath,
            IProgress<double> progress, CancellationToken cancellationToken)
        {
            if (release == null) throw new ArgumentNullException(nameof(release));

            string partial = destinationPath + UpdateStage.PartialSuffix;
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
            SafeDelete(partial);

            try
            {
                using (var response = await _http
                           .GetAsync(release.Url, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                           .ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new UpdateTransportException(string.Format(
                            CultureInfo.InvariantCulture,
                            "下载更新包失败（HTTP {0}）。", (int)response.StatusCode));
                    }

                    long? advertised = response.Content.Headers.ContentLength;
                    if (advertised.HasValue && advertised.Value != release.Size)
                    {
                        // Cheap and it fails before the bytes are spent; the
                        // digest is still what decides.
                        throw new UpdateTransportException(string.Format(
                            CultureInfo.InvariantCulture,
                            "更新包大小与清单不符（服务器 {0}，清单 {1}）。",
                            advertised.Value, release.Size));
                    }

                    using (Stream source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var target = new FileStream(
                        partial, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1024 * 64,
                        FileOptions.SequentialScan))
                    {
                        await CopyCappedAsync(source, target, release.Size, progress, cancellationToken)
                            .ConfigureAwait(false);
                    }
                }

                SafeDelete(destinationPath);
                File.Move(partial, destinationPath);
            }
            catch (HttpRequestException ex)
            {
                SafeDelete(partial);
                throw new UpdateTransportException("下载更新包失败：" + ex.Message, ex);
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                SafeDelete(partial);
                throw new UpdateTransportException("下载更新包超时。", ex);
            }
            catch (Exception)
            {
                SafeDelete(partial);
                throw;
            }
        }

        /// <summary>
        /// Copies at most <paramref name="limit"/> bytes and fails past it.
        ///
        /// Without the cap a server that keeps sending would fill the user's
        /// disk, and no signature check downstream would ever get the chance to
        /// object.
        /// </summary>
        private static async Task CopyCappedAsync(
            Stream source, Stream destination, long limit,
            IProgress<double> progress, CancellationToken cancellationToken)
        {
            var buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
                       .ConfigureAwait(false)) > 0)
            {
                total += read;
                if (total > limit)
                {
                    throw new UpdateTransportException(string.Format(
                        CultureInfo.InvariantCulture,
                        "服务器发送的数据超过声明的 {0} 字节，已中止。", limit));
                }
                await destination.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                if (progress != null && limit > 0)
                {
                    progress.Report((double)total / limit);
                }
            }
        }

        private static void SafeDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
            }
        }

        public void Dispose()
        {
            _http.Dispose();
        }
    }
}
