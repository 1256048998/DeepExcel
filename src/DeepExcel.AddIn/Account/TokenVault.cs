using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Account
{
    /// <summary>
    /// Persists the refresh token, DPAPI-protected for the current Windows user.
    ///
    /// Deliberately does not reuse SecurityManager.Encrypt: that helper returns
    /// the plaintext when DPAPI fails, which is the right trade-off nowhere and
    /// certainly not for a credential that grants account access for a month.
    /// Both directions here fail closed -- a token that cannot be protected is
    /// not written at all.
    /// </summary>
    public sealed class TokenVault
    {
        private const string FileName = "session.crypt";
        // Binds the ciphertext to this purpose, so a blob lifted from another
        // DPAPI-protected file in the same profile cannot be substituted here.
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("DeepExcel.Account.Session.v1");

        private readonly string _path;

        public TokenVault(string directory = null)
        {
            var root = directory ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DeepExcel", "credentials");
            _path = Path.Combine(root, FileName);
        }

        public string Path_ForTests
        {
            get { return _path; }
        }

        public sealed class PersistedSession
        {
            [JsonPropertyName("server_url")]
            public string ServerUrl { get; set; }

            [JsonPropertyName("refresh_token")]
            public string RefreshToken { get; set; }

            [JsonPropertyName("email")]
            public string Email { get; set; }

            /// <summary>
            /// Last time the server actually confirmed this session. Drives the
            /// offline grace window.
            /// </summary>
            [JsonPropertyName("last_verified_utc")]
            public DateTime LastVerifiedUtc { get; set; }
        }

        public bool TryLoad(out PersistedSession session)
        {
            session = null;
            try
            {
                if (!File.Exists(_path))
                {
                    return false;
                }
                byte[] protectedBytes = File.ReadAllBytes(_path);
                byte[] plain = ProtectedData.Unprotect(
                    protectedBytes, Entropy, DataProtectionScope.CurrentUser);
                session = JsonSerializer.Deserialize<PersistedSession>(
                    Encoding.UTF8.GetString(plain));
                return session != null && !string.IsNullOrEmpty(session.RefreshToken);
            }
            catch (Exception)
            {
                // Unreadable means tampered, corrupt, or written by another
                // Windows user. Treat as "no session" and never fall back to
                // interpreting the bytes as plaintext.
                session = null;
                return false;
            }
        }

        /// <summary>Returns false if the token could not be protected; nothing is written in that case.</summary>
        public bool Save(PersistedSession session)
        {
            if (session == null || string.IsNullOrEmpty(session.RefreshToken))
            {
                return false;
            }
            try
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_path));
                byte[] plain = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(session));
                byte[] protectedBytes = ProtectedData.Protect(
                    plain, Entropy, DataProtectionScope.CurrentUser);

                // Write-then-replace so an interrupted save cannot leave a
                // truncated file that silently signs the user out.
                string temporary = _path + ".tmp";
                File.WriteAllBytes(temporary, protectedBytes);
                if (File.Exists(_path))
                {
                    File.Replace(temporary, _path, null);
                }
                else
                {
                    File.Move(temporary, _path);
                }
                return true;
            }
            catch (Exception)
            {
                try { File.Delete(_path + ".tmp"); } catch (Exception) { }
                return false;
            }
        }

        public void Clear()
        {
            try
            {
                if (File.Exists(_path))
                {
                    File.Delete(_path);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
