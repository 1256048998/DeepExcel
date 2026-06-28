using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace DeepExcel.AddIn.Config
{
    /// <summary>
    /// 应用配置 - 持久化到 %APPDATA%\DeepExcel\config.json
    /// 支持运行时热重载
    /// </summary>
    public class AppConfig
    {
        public string CurrentProvider { get; set; } = "anthropic";
        public string CurrentModel { get; set; } = "claude-3-5-sonnet-20241022";

        public Dictionary<string, ProviderConfig> Providers { get; set; } = new();
        public GeneralSettings General { get; set; } = new();
        public UISettings UI { get; set; } = new();

        public static AppConfig CreateDefault()
        {
            var cfg = new AppConfig();
            cfg.Providers["anthropic"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "Claude (Anthropic)",
                ApiKey = "",
                BaseUrl = "https://api.anthropic.com",
                Models = new[] { "claude-3-5-sonnet-20241022", "claude-3-5-haiku-20241022", "claude-3-opus-20240229" }
            };
            cfg.Providers["deepseek"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "DeepSeek",
                ApiKey = "",
                BaseUrl = "https://api.deepseek.com/anthropic",
                Models = new[] { "deepseek-chat", "deepseek-coder", "deepseek-reasoner" }
            };
            cfg.Providers["openai"] = new ProviderConfig
            {
                Type = "openai",
                DisplayName = "OpenAI",
                ApiKey = "",
                BaseUrl = "https://api.openai.com/v1",
                Models = new[] { "gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo" }
            };
            cfg.Providers["custom"] = new ProviderConfig
            {
                Type = "openai",
                DisplayName = "自定义 (OpenAI兼容)",
                ApiKey = "",
                BaseUrl = "",
                Models = new[] { "custom-model" }
            };
            return cfg;
        }
    }

    public class ProviderConfig
    {
        public string Type { get; set; }       // "anthropic" | "openai"
        public string DisplayName { get; set; }
        public string ApiKey { get; set; }
        public string BaseUrl { get; set; }
        public string[] Models { get; set; }
        public string DefaultModel { get; set; }
        public Dictionary<string, string> Headers { get; set; } = new();
    }

    public class GeneralSettings
    {
        public int MaxRetries { get; set; } = 2;
        public int RequestTimeoutSeconds { get; set; } = 60;
        public bool AutoCreateSnapshot { get; set; } = true;
        public bool RequireConfirmation { get; set; } = true;
        public int MaxConversationHistory { get; set; } = 10;
    }

    public class UISettings
    {
        public string Theme { get; set; } = "light";
        public string Language { get; set; } = "zh-CN";
        public bool ShowTokenUsage { get; set; } = true;
        public bool StreamOutput { get; set; } = true;
    }

    /// <summary>
    /// 配置管理器 - 加载/保存/热重载
    /// </summary>
    public class ConfigManager
    {
        private static readonly string ConfigDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DeepExcel");
        private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.json");

        private AppConfig _config;
        public AppConfig Current => _config;
        public event Action<AppConfig> OnConfigChanged;

        private static ConfigManager _instance;
        public static ConfigManager Instance => _instance ??= new ConfigManager();

        private ConfigManager()
        {
            Load();
        }

        /// <summary>
        /// 加载配置（不存在则创建默认）
        /// </summary>
        public AppConfig Load()
        {
            try
            {
                if (!Directory.Exists(ConfigDir))
                {
                    Directory.CreateDirectory(ConfigDir);
                }

                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    };
                    _config = JsonSerializer.Deserialize<AppConfig>(json, options) ?? AppConfig.CreateDefault();
                }
                else
                {
                    _config = AppConfig.CreateDefault();
                    Save();
                }
            }
            catch
            {
                _config = AppConfig.CreateDefault();
            }
            return _config;
        }

        /// <summary>
        /// 保存配置
        /// </summary>
        public void Save()
        {
            try
            {
                if (!Directory.Exists(ConfigDir))
                {
                    Directory.CreateDirectory(ConfigDir);
                }

                var json = JsonSerializer.Serialize(_config, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(ConfigPath, json);
                OnConfigChanged?.Invoke(_config);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存配置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 切换当前模型提供方
        /// </summary>
        public bool SwitchProvider(string providerKey, string modelName = null)
        {
            if (!_config.Providers.ContainsKey(providerKey)) return false;

            _config.CurrentProvider = providerKey;
            if (!string.IsNullOrEmpty(modelName))
            {
                _config.CurrentModel = modelName;
            }
            else
            {
                _config.CurrentModel = _config.Providers[providerKey].DefaultModel
                    ?? _config.Providers[providerKey].Models[0];
            }
            Save();
            return true;
        }

        /// <summary>
        /// 更新API Key
        /// </summary>
        public void UpdateApiKey(string providerKey, string apiKey)
        {
            if (!_config.Providers.ContainsKey(providerKey))
            {
                _config.Providers[providerKey] = new ProviderConfig();
            }
            _config.Providers[providerKey].ApiKey = apiKey;
            Save();
        }

        /// <summary>
        /// 重新加载配置（响应UI手动编辑）
        /// </summary>
        public void Reload()
        {
            Load();
            OnConfigChanged?.Invoke(_config);
        }

        public string ConfigFilePath => ConfigPath;
    }
}
