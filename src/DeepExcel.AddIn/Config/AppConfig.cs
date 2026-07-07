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
                Models = new[] { "claude-3-5-sonnet-20241022", "claude-3-5-haiku-20241022", "claude-3-opus-20240229" },
                DefaultModel = "claude-3-5-sonnet-20241022",
                SupportsVision = true
            };
            cfg.Providers["deepseek"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "DeepSeek",
                ApiKey = "",
                BaseUrl = "https://api.deepseek.com/anthropic",
                Models = new[] { "deepseek-chat", "deepseek-coder", "deepseek-reasoner" },
                DefaultModel = "deepseek-chat",
                SupportsVision = false
            };
            cfg.Providers["stepfun"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "阶跃星辰 (Step)",
                ApiKey = "",
                BaseUrl = "https://api.stepfun.com/step_plan",
                Models = new[] { "step-3.7-flash", "step-3.5-flash" },
                DefaultModel = "step-3.7-flash",
                SupportsVision = true
            };
            cfg.Providers["openai"] = new ProviderConfig
            {
                Type = "openai",
                DisplayName = "OpenAI",
                ApiKey = "",
                BaseUrl = "https://api.openai.com/v1",
                Models = new[] { "gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo" },
                DefaultModel = "gpt-4o",
                SupportsVision = true
            };
            cfg.Providers["kimi"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "Kimi (月之暗面)",
                ApiKey = "",
                BaseUrl = "https://api.moonshot.cn/anthropic",
                Models = new[] { "kimi-k2.7-code", "kimi-k2.6", "kimi-k2-thinking" },
                DefaultModel = "kimi-k2.7-code",
                SupportsVision = true
            };
            cfg.Providers["qwen"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "通义千问 (阿里)",
                ApiKey = "",
                BaseUrl = "https://dashscope.aliyuncs.com/compatible-mode/anthropic",
                Models = new[] { "qwen3-max", "qwen3-coder-plus", "qwen-plus", "qwen-turbo", "qwen-long" },
                DefaultModel = "qwen3-max",
                SupportsVision = true
            };
            cfg.Providers["zhipu"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "智谱 (GLM)",
                ApiKey = "",
                BaseUrl = "https://api.z.ai/api/anthropic",
                Models = new[] { "glm-4.6", "glm-4.6-air", "glm-4.5" },
                DefaultModel = "glm-4.6",
                SupportsVision = true
            };
            cfg.Providers["minimax"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "Minimax",
                ApiKey = "",
                BaseUrl = "https://api.minimax.io/anthropic",
                Models = new[] { "MiniMax-M2.1", "MiniMax-M2" },
                DefaultModel = "MiniMax-M2.1",
                SupportsVision = false
            };
            cfg.Providers["doubao"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "豆包 (火山引擎)",
                ApiKey = "",
                BaseUrl = "https://ark.cn-beijing.volces.com/api/compatible",
                Models = new[] { "doubao-seed-code", "doubao-seed-1.6", "doubao-seed-1.6-flash" },
                DefaultModel = "doubao-seed-code",
                SupportsVision = true
            };
            cfg.Providers["custom"] = new ProviderConfig
            {
                Type = "openai",
                DisplayName = "自定义 (OpenAI兼容)",
                ApiKey = "",
                BaseUrl = "",
                Models = new[] { "custom-model" },
                DefaultModel = "custom-model",
                SupportsVision = false
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
        /// <summary>★ 该 provider 是否支持 vision（图片识别）。用于附件含图片时自动切换。
        /// anthropic/stepfun=true, deepseek=false。旧 config.json 无此字段时默认 false。</summary>
        public bool SupportsVision { get; set; } = false;
    }

    public class GeneralSettings
    {
        public int MaxRetries { get; set; } = 2;
        public int RequestTimeoutSeconds { get; set; } = 60;
        public bool AutoCreateSnapshot { get; set; } = true;
        public bool RequireConfirmation { get; set; } = true;
        public int MaxConversationHistory { get; set; } = 10;
        /// <summary>★ Claude Agent SDK 控制循环最大轮次（工具调用往返次数）。
        /// 达到限制后 SDK 返回 error_max_turns。默认 20，防止 AI 无限循环调用工具。</summary>
        public int MaxTurns { get; set; } = 20;
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
                    // ★ 迁移：补全新 provider 和字段（旧 config.json 没有 stepfun/supportsVision/maxTurns）
                    MigrateConfig(_config);
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
        /// ★ P0-4 修复：API Key 不再明文存到 config.json，改用 SecurityManager 通过 DPAPI 加密存到
        /// %APPDATA%/DeepExcel/credentials/key_{provider}.crypt 文件，避免密钥被窃取。
        /// config.json 中 ApiKey 字段保留空字符串占位（不存真实 key）。
        /// </summary>
        public void UpdateApiKey(string providerKey, string apiKey)
        {
            if (!_config.Providers.ContainsKey(providerKey))
            {
                _config.Providers[providerKey] = new ProviderConfig();
            }
            // ★ 通过 SecurityManager 加密存储，config.json 不保留明文
            try
            {
                DeepExcel.AddIn.Security.SecurityManager.Instance.SaveApiKey(providerKey, apiKey);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UpdateApiKey: SecurityManager.Save failed: " + ex.Message);
            }
            // config.json 中保留空占位（向后兼容旧读取逻辑）
            _config.Providers[providerKey].ApiKey = "";
            Save();
        }

        /// <summary>
        /// ★ 迁移：补全新 provider 和字段。旧 config.json 可能缺少 stepfun provider、
        /// supportsVision 字段、general.maxTurns 字段。此方法确保旧配置升级后包含所有新字段。
        /// </summary>
        private void MigrateConfig(AppConfig config)
        {
            bool changed = false;

            // 1. 补充 stepfun provider
            if (!config.Providers.ContainsKey("stepfun"))
            {
                config.Providers["stepfun"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "阶跃星辰 (Step)",
                    ApiKey = "",
                    BaseUrl = "https://api.stepfun.com/step_plan",
                    Models = new[] { "step-3.7-flash", "step-3.5-flash" },
                    DefaultModel = "step-3.7-flash",
                    SupportsVision = true
                };
                changed = true;
            }

            // 1b. 补充 5 个国产厂商 provider
            var newProviders = new Dictionary<string, ProviderConfig>
            {
                ["kimi"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "Kimi (月之暗面)",
                    ApiKey = "",
                    BaseUrl = "https://api.moonshot.cn/anthropic",
                    Models = new[] { "kimi-k2.7-code", "kimi-k2.6", "kimi-k2-thinking" },
                    DefaultModel = "kimi-k2.7-code",
                    SupportsVision = true
                },
                ["qwen"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "通义千问 (阿里)",
                    ApiKey = "",
                    BaseUrl = "https://dashscope.aliyuncs.com/compatible-mode/anthropic",
                    Models = new[] { "qwen3-max", "qwen3-coder-plus", "qwen-plus", "qwen-turbo", "qwen-long" },
                    DefaultModel = "qwen3-max",
                    SupportsVision = true
                },
                ["zhipu"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "智谱 (GLM)",
                    ApiKey = "",
                    BaseUrl = "https://api.z.ai/api/anthropic",
                    Models = new[] { "glm-4.6", "glm-4.6-air", "glm-4.5" },
                    DefaultModel = "glm-4.6",
                    SupportsVision = true
                },
                ["minimax"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "Minimax",
                    ApiKey = "",
                    BaseUrl = "https://api.minimax.io/anthropic",
                    Models = new[] { "MiniMax-M2.1", "MiniMax-M2" },
                    DefaultModel = "MiniMax-M2.1",
                    SupportsVision = false
                },
                ["doubao"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "豆包 (火山引擎)",
                    ApiKey = "",
                    BaseUrl = "https://ark.cn-beijing.volces.com/api/compatible",
                    Models = new[] { "doubao-seed-code", "doubao-seed-1.6", "doubao-seed-1.6-flash" },
                    DefaultModel = "doubao-seed-code",
                    SupportsVision = true
                }
            };
            foreach (var kvp in newProviders)
            {
                if (!config.Providers.ContainsKey(kvp.Key))
                {
                    config.Providers[kvp.Key] = kvp.Value;
                    changed = true;
                }
            }

            // 2. 补充 anthropic provider（如果旧 config 删了）
            if (!config.Providers.ContainsKey("anthropic"))
            {
                config.Providers["anthropic"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "Claude (Anthropic)",
                    ApiKey = "",
                    BaseUrl = "https://api.anthropic.com",
                    Models = new[] { "claude-3-5-sonnet-20241022", "claude-3-5-haiku-20241022", "claude-3-opus-20240229" },
                    SupportsVision = true
                };
                changed = true;
            }

            // 3. 为已知 provider 补充 SupportsVision 字段（旧 config 没有此字段时默认 false，需要修正）
            if (config.Providers.ContainsKey("anthropic") && !config.Providers["anthropic"].SupportsVision)
            {
                config.Providers["anthropic"].SupportsVision = true;
                changed = true;
            }
            if (config.Providers.ContainsKey("stepfun") && !config.Providers["stepfun"].SupportsVision)
            {
                config.Providers["stepfun"].SupportsVision = true;
                changed = true;
            }
            if (config.Providers.ContainsKey("openai") && !config.Providers["openai"].SupportsVision)
            {
                config.Providers["openai"].SupportsVision = true;
                changed = true;
            }

            // 4. 为旧 provider 补充 DefaultModel 字段（旧 config 没有此字段）
            if (config.Providers.ContainsKey("anthropic") && string.IsNullOrEmpty(config.Providers["anthropic"].DefaultModel))
            { config.Providers["anthropic"].DefaultModel = "claude-3-5-sonnet-20241022"; changed = true; }
            if (config.Providers.ContainsKey("deepseek") && string.IsNullOrEmpty(config.Providers["deepseek"].DefaultModel))
            { config.Providers["deepseek"].DefaultModel = "deepseek-chat"; changed = true; }
            if (config.Providers.ContainsKey("openai") && string.IsNullOrEmpty(config.Providers["openai"].DefaultModel))
            { config.Providers["openai"].DefaultModel = "gpt-4o"; changed = true; }
            if (config.Providers.ContainsKey("custom") && string.IsNullOrEmpty(config.Providers["custom"].DefaultModel))
            { config.Providers["custom"].DefaultModel = "custom-model"; changed = true; }

            // 4. 补充 GeneralSettings.MaxTurns（旧 config 没有此字段）
            if (config.General == null)
            {
                config.General = new GeneralSettings();
                changed = true;
            }
            if (config.General.MaxTurns <= 0)
            {
                config.General.MaxTurns = 20;
                changed = true;
            }

            if (changed)
            {
                try { Save(); } catch { }
            }
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
