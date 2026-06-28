using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.IO;

namespace DeepExcel.AddIn.Security
{
    /// <summary>
    /// 安全管理器 - 负责API Key加密存储和安全操作验证
    /// 使用DPAPI保护敏感数据
    /// </summary>
    public class SecurityManager
    {
        private static readonly string CredentialDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DeepExcel", "credentials");

        private static SecurityManager _instance;
        public static SecurityManager Instance => _instance ??= new SecurityManager();

        private SecurityManager()
        {
            try
            {
                if (!Directory.Exists(CredentialDir))
                    Directory.CreateDirectory(CredentialDir);
            }
            catch { }
        }

        /// <summary>
        /// 加密字符串（使用DPAPI）
        /// </summary>
        public string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            try
            {
                byte[] data = Encoding.UTF8.GetBytes(plainText);
                byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            }
            catch
            {
                return plainText;
            }
        }

        /// <summary>
        /// 解密字符串
        /// </summary>
        public string Decrypt(string encryptedText)
        {
            if (string.IsNullOrEmpty(encryptedText)) return "";
            try
            {
                byte[] data = Convert.FromBase64String(encryptedText);
                byte[] decrypted = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            }
            catch
            {
                return encryptedText;
            }
        }

        /// <summary>
        /// 保存加密的API Key
        /// </summary>
        public bool SaveApiKey(string providerKey, string apiKey)
        {
            try
            {
                var path = GetCredentialPath(providerKey);
                var encrypted = Encrypt(apiKey);
                File.WriteAllText(path, encrypted);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取解密的API Key
        /// </summary>
        public string GetApiKey(string providerKey)
        {
            try
            {
                var path = GetCredentialPath(providerKey);
                if (!File.Exists(path)) return "";
                var encrypted = File.ReadAllText(path);
                return Decrypt(encrypted);
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 删除API Key
        /// </summary>
        public bool DeleteApiKey(string providerKey)
        {
            try
            {
                var path = GetCredentialPath(providerKey);
                if (File.Exists(path)) File.Delete(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 验证API Key是否已保存（仅检查文件存在）
        /// </summary>
        public bool HasApiKey(string providerKey)
        {
            try
            {
                var path = GetCredentialPath(providerKey);
                return File.Exists(path) && !string.IsNullOrEmpty(File.ReadAllText(path));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 生成随机验证码（用于二次验证）
        /// </summary>
        public string GenerateVerificationCode()
        {
            var chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var random = new Random();
            var code = new char[6];
            for (int i = 0; i < 6; i++)
            {
                code[i] = chars[random.Next(chars.Length)];
            }
            return new string(code);
        }

        /// <summary>
        /// 生成操作签名（用于防篡改）
        /// </summary>
        public string GenerateOperationSignature(string operation, string timestamp)
        {
            var data = $"{operation}:{timestamp}:{Environment.MachineName}";
            using var sha256 = SHA256.Create();
            var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(data));
            return Convert.ToBase64String(hash);
        }

        /// <summary>
        /// 验证操作签名
        /// </summary>
        public bool VerifyOperationSignature(string operation, string timestamp, string signature)
        {
            var expected = GenerateOperationSignature(operation, timestamp);
            return expected == signature;
        }

        /// <summary>
        /// 获取安全的配置（不含敏感字段）
        /// </summary>
        public SafeConfig GetSafeConfig(Config.AppConfig config)
        {
            return new SafeConfig
            {
                CurrentProvider = config.CurrentProvider,
                CurrentModel = config.CurrentModel,
                Providers = new Dictionary<string, SafeProvider>(config.Providers.Count),
                General = config.General,
                UI = config.UI
            };
        }

        private string GetCredentialPath(string providerKey)
        {
            return Path.Combine(CredentialDir, $"key_{providerKey}.crypt");
        }
    }

    public class SafeConfig
    {
        public string CurrentProvider { get; set; }
        public string CurrentModel { get; set; }
        public Dictionary<string, SafeProvider> Providers { get; set; }
        public Config.GeneralSettings General { get; set; }
        public Config.UISettings UI { get; set; }
    }

    public class SafeProvider
    {
        public string DisplayName { get; set; }
        public string[] Models { get; set; }
        public bool HasApiKey { get; set; }
    }
}
