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
        public string CurrentModel { get; set; } = "claude-sonnet-5";
        /// <summary>★ 全局默认厂商（前端 provider 列表排序最前，输入框下拉默认值）。
        /// 全局唯一，用户在 ModelConfigPanel 切换"设为默认"开关时更新。
        /// 旧 config 无此字段默认 null，前端回退到 CurrentProvider。</summary>
        public string DefaultProvider { get; set; } = null;

        /// <summary>★ 内置模型目录版本号。MigrateConfig 只在本地版本落后时才把
        /// LatestModelCatalog 推送到各 provider，避免每次启动都覆盖用户自己整理的模型列表/排序。
        /// 旧 config.json 无此字段默认 0，会执行一次升级后写入当前版本。</summary>
        public int ModelCatalogVersion { get; set; } = 0;

        /// <summary>★ 当前内置模型目录版本。新增/淘汰内置模型时 +1，
        /// 让老用户在下次启动时拿到新目录（用户已自定义的 provider 除外）。</summary>
        public const int CurrentModelCatalogVersion = 2;

        /// <summary>
        /// ★ 合并模型列表：保留用户已有的全部模型和顺序，只把内置目录里缺的追加到末尾。
        /// 永远不删——内置目录滞后于厂商，拿它裁剪用户列表只会把能用的模型弄丢。
        /// </summary>
        internal static string[] MergeModels(string[] existing, string[] catalogModels)
        {
            var merged = new List<string>(existing ?? new string[0]);
            foreach (var model in catalogModels ?? new string[0])
            {
                if (string.IsNullOrWhiteSpace(model)) continue;
                if (!merged.Exists(m => string.Equals(m, model, StringComparison.OrdinalIgnoreCase)))
                    merged.Add(model);
            }
            return merged.ToArray();
        }

        public Dictionary<string, ProviderConfig> Providers { get; set; } = new();
        public GeneralSettings General { get; set; } = new();
        public UISettings UI { get; set; } = new();

        public static AppConfig CreateDefault()
        {
            var cfg = new AppConfig();
            cfg.ModelCatalogVersion = CurrentModelCatalogVersion;
            cfg.Providers["anthropic"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "Claude (Anthropic)",
                ApiKey = "",
                BaseUrl = "https://api.anthropic.com",
                Models = new[] { "claude-sonnet-5", "claude-opus-5", "claude-opus-4.8", "claude-haiku-5", "claude-haiku-4-5-20251001" },
                DefaultModel = "claude-sonnet-5",
                SupportsVision = true
            };
            cfg.Providers["deepseek"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "DeepSeek",
                ApiKey = "",
                BaseUrl = "https://api.deepseek.com/anthropic",
                Models = new[] { "deepseek-v4-pro", "deepseek-v4-flash" },
                DefaultModel = "deepseek-v4-pro",
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
                Models = new[] { "gpt-5.5", "gpt-5.5-pro", "gpt-5" },
                DefaultModel = "gpt-5.5",
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
                Models = new[] { "qwen3.7-max", "qwen3-max", "qwen3-coder-plus" },
                DefaultModel = "qwen3.7-max",
                SupportsVision = true
            };
            cfg.Providers["zhipu"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "智谱 (GLM)",
                ApiKey = "",
                BaseUrl = "https://api.z.ai/api/anthropic",
                Models = new[] { "glm-5.2", "glm-5.1", "glm-4.7-flash" },
                DefaultModel = "glm-5.2",
                SupportsVision = true
            };
            cfg.Providers["minimax"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "Minimax",
                ApiKey = "",
                BaseUrl = "https://api.minimax.io/anthropic",
                Models = new[] { "MiniMax-M2.5", "MiniMax-M2" },
                DefaultModel = "MiniMax-M2.5",
                SupportsVision = false
            };
            cfg.Providers["doubao"] = new ProviderConfig
            {
                Type = "anthropic",
                DisplayName = "豆包 (火山引擎)",
                ApiKey = "",
                BaseUrl = "https://ark.cn-beijing.volces.com/api/compatible",
                Models = new[] { "doubao-seed-2.1-pro", "doubao-seed-2.1", "doubao-seed-1.6" },
                DefaultModel = "doubao-seed-2.1-pro",
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
            cfg.CurrentModel = "claude-sonnet-5";
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
        /// <summary>★ 最近一次测试连接是否成功。前端 provider 列表的圆点据此显示，
        /// 而非 hasApiKey（hasApiKey 只反映 .crypt 文件存在，不代表 key 有效）。
        /// test_api_key 成功置 true，失败/删除 key 时置 false。旧 config 无此字段默认 false。</summary>
        public bool LastTestSuccess { get; set; } = false;
        /// <summary>★ 用户是否自己整理过该 provider 的模型列表（导入/新增/删除/拖拽排序）。
        /// 为 true 时 MigrateConfig 不再用内置目录覆盖 Models/DefaultModel，
        /// 否则用户在"模型优先级"里排好的顺序会在下次启动时被重置。</summary>
        public bool ModelsCustomized { get; set; } = false;
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

        // ★ M-7 修复：使用 Lazy<T> 实现线程安全单例，避免多线程下创建多个实例
        private static readonly Lazy<ConfigManager> _instance = new Lazy<ConfigManager>(() => new ConfigManager());
        public static ConfigManager Instance => _instance.Value;

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
        /// ★ C-1 修复：DPAPI 加密失败时拒绝保存并返回 false，避免明文 key 残留或被无加密使用。
        /// </summary>
        public bool UpdateApiKey(string providerKey, string apiKey)
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
                // ★ C-1 修复：DPAPI 失败则拒绝保存，不更新内存中 ApiKey，不调用 Save()
                // 这样 config.json 永不存明文，且不会出现"内存中持有 key 但磁盘没存"的不一致状态
                return false;
            }
            // config.json 中保留空占位（向后兼容旧读取逻辑）
            _config.Providers[providerKey].ApiKey = "";
            Save();
            return true;
        }

        /// <summary>
        /// ★ 保存某个 provider 的"已选模型"有序列表（数组顺序即优先级，第 0 个为主模型）。
        /// 由前端"模型优先级"列表（导入 / 新增 / 删除 / 拖拽排序）驱动。
        /// 同时把 DefaultModel 对齐到主模型，并标记 ModelsCustomized=true，
        /// 使 MigrateConfig 不再用内置目录覆盖该列表。
        /// </summary>
        /// <returns>false 表示 provider 不存在或模型列表为空</returns>
        public bool UpdateProviderModels(string providerKey, string[] models)
        {
            if (string.IsNullOrEmpty(providerKey) || !_config.Providers.ContainsKey(providerKey)) return false;
            if (models == null || models.Length == 0) return false;

            var p = _config.Providers[providerKey];
            p.Models = models;
            p.DefaultModel = models[0];
            p.ModelsCustomized = true;

            // 当前正在使用的模型如果被移出列表，回落到主模型，避免请求到一个已删除的模型名
            if (_config.CurrentProvider == providerKey &&
                Array.IndexOf(models, _config.CurrentModel) < 0)
            {
                _config.CurrentModel = models[0];
            }
            Save();
            return true;
        }

        /// <summary>
        /// ★ 内置模型目录。
        /// MigrateConfig 只用它**补充**用户还没有的模型，绝不删除用户已有的模型名——
        /// 厂商模型迭代很快，内置目录一定滞后于现实，用它去裁剪用户的列表只会误伤。
        /// 新增模型直接往这里加即可。
        /// </summary>
        private static readonly Dictionary<string, (string[] Models, string DefaultModel)> LatestModelCatalog =
            new Dictionary<string, (string[], string)>
        {
            ["anthropic"] = (new[] { "claude-sonnet-5", "claude-opus-5", "claude-opus-4.8", "claude-haiku-5", "claude-haiku-4-5-20251001" }, "claude-sonnet-5"),
            ["deepseek"] = (new[] { "deepseek-v4-pro", "deepseek-v4-flash" }, "deepseek-v4-pro"),
            ["stepfun"] = (new[] { "step-3.7-flash", "step-3.5-flash" }, "step-3.7-flash"),
            ["openai"] = (new[] { "gpt-5.5", "gpt-5.5-pro", "gpt-5" }, "gpt-5.5"),
            ["kimi"] = (new[] { "kimi-k2.7-code", "kimi-k2.6", "kimi-k2-thinking" }, "kimi-k2.7-code"),
            ["qwen"] = (new[] { "qwen3.7-max", "qwen3-max", "qwen3-coder-plus" }, "qwen3.7-max"),
            ["zhipu"] = (new[] { "glm-5.2", "glm-5.1", "glm-4.7-flash" }, "glm-5.2"),
            ["minimax"] = (new[] { "MiniMax-M2.5", "MiniMax-M2" }, "MiniMax-M2.5"),
            ["doubao"] = (new[] { "doubao-seed-2.1-pro", "doubao-seed-2.1", "doubao-seed-1.6" }, "doubao-seed-2.1-pro"),
        };

        private static Dictionary<string, (string[] Models, string DefaultModel)> GetLatestModelCatalog()
        {
            return LatestModelCatalog;
        }

        /// <summary>
        /// ★ 迁移：补全新 provider 和字段。旧 config.json 可能缺少 stepfun provider、
        /// supportsVision 字段、general.maxTurns 字段。此方法确保旧配置升级后包含所有新字段。
        /// </summary>
        private void MigrateConfig(AppConfig config)
        {
            bool changed = false;

            // ★ C-1 修复：扫描所有 provider，将旧版本 config.json 中残留的明文 ApiKey 迁移到
            // SecurityManager（DPAPI 加密存储），并强制清空 config.json 中的 ApiKey 字段。
            // 这样升级用户即使旧版本在 config.json 存过明文 key，升级后也会被加密保护。
            foreach (var kvp in config.Providers)
            {
                if (!string.IsNullOrEmpty(kvp.Value.ApiKey))
                {
                    try
                    {
                        DeepExcel.AddIn.Security.SecurityManager.Instance.SaveApiKey(kvp.Key, kvp.Value.ApiKey);
                        System.Diagnostics.Debug.WriteLine("MigrateConfig: migrated ApiKey for " + kvp.Key + " to SecurityManager");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("MigrateConfig: SaveApiKey failed for " + kvp.Key + ": " + ex.Message);
                    }
                    kvp.Value.ApiKey = "";
                    changed = true;
                }
            }

            // 1. 补充 stepfun provider
            // ★ M-4 修复：Models 和 DefaultModel 直接从 LatestModelCatalog 读取，保证与 CreateDefault 一致
            var stepCatalog = LatestModelCatalog["stepfun"];
            if (!config.Providers.ContainsKey("stepfun"))
            {
                config.Providers["stepfun"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "阶跃星辰 (Step)",
                    ApiKey = "",
                    BaseUrl = "https://api.stepfun.com/step_plan",
                    Models = stepCatalog.Models,
                    DefaultModel = stepCatalog.DefaultModel,
                    SupportsVision = true
                };
                changed = true;
            }

            // 1b. 补充 5 个国产厂商 provider
            // ★ M-4 修复：Models 和 DefaultModel 从 LatestModelCatalog 读取
            var catalog = LatestModelCatalog;
            var newProviders = new Dictionary<string, ProviderConfig>
            {
                ["kimi"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "Kimi (月之暗面)",
                    ApiKey = "",
                    BaseUrl = "https://api.moonshot.cn/anthropic",
                    Models = catalog["kimi"].Models,
                    DefaultModel = catalog["kimi"].DefaultModel,
                    SupportsVision = true
                },
                ["qwen"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "通义千问 (阿里)",
                    ApiKey = "",
                    BaseUrl = "https://dashscope.aliyuncs.com/compatible-mode/anthropic",
                    Models = catalog["qwen"].Models,
                    DefaultModel = catalog["qwen"].DefaultModel,
                    SupportsVision = true
                },
                ["zhipu"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "智谱 (GLM)",
                    ApiKey = "",
                    BaseUrl = "https://api.z.ai/api/anthropic",
                    Models = catalog["zhipu"].Models,
                    DefaultModel = catalog["zhipu"].DefaultModel,
                    SupportsVision = true
                },
                ["minimax"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "Minimax",
                    ApiKey = "",
                    BaseUrl = "https://api.minimax.io/anthropic",
                    Models = catalog["minimax"].Models,
                    DefaultModel = catalog["minimax"].DefaultModel,
                    SupportsVision = false
                },
                ["doubao"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "豆包 (火山引擎)",
                    ApiKey = "",
                    BaseUrl = "https://ark.cn-beijing.volces.com/api/compatible",
                    Models = catalog["doubao"].Models,
                    DefaultModel = catalog["doubao"].DefaultModel,
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
            var anthropicCatalog = LatestModelCatalog["anthropic"];
            if (!config.Providers.ContainsKey("anthropic"))
            {
                config.Providers["anthropic"] = new ProviderConfig
                {
                    Type = "anthropic",
                    DisplayName = "Claude (Anthropic)",
                    ApiKey = "",
                    BaseUrl = "https://api.anthropic.com",
                    Models = anthropicCatalog.Models,
                    DefaultModel = anthropicCatalog.DefaultModel,
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
            // ★ M-4 修复：从 LatestModelCatalog 读取，保证与 CreateDefault 一致
            foreach (var cat in LatestModelCatalog)
            {
                if (config.Providers.ContainsKey(cat.Key) &&
                    string.IsNullOrEmpty(config.Providers[cat.Key].DefaultModel))
                {
                    config.Providers[cat.Key].DefaultModel = cat.Value.DefaultModel;
                    changed = true;
                }
            }
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

            // 5. ★ 模型目录升级：只做"补充"，不做"覆盖"。
            //    修复的两个历史问题：
            //      a) 以前每次启动都无条件覆盖 Models，用户从厂商拉到的最新模型、
            //         手工排好的优先级顺序，下次打开 Excel 就被内置列表冲掉；
            //      b) 内置目录一旦精简，用户原本能用的模型名会凭空消失。
            //    现在的规则：
            //      - ModelsCustomized=true（用户自己导入/排序过）→ 完全不碰；
            //      - 其余 provider → 把内置目录里缺的模型追加到末尾，已有的一个都不删；
            //      - 版本号只用来控制"每次目录更新最多补充一次"，避免反复追加用户删掉的条目。
            var catalogUpdates = GetLatestModelCatalog();
            bool catalogOutdated = config.ModelCatalogVersion < AppConfig.CurrentModelCatalogVersion;
            foreach (var kvp in catalogUpdates)
            {
                if (!config.Providers.ContainsKey(kvp.Key)) continue;
                var p = config.Providers[kvp.Key];
                var catalogModels = kvp.Value.Models;

                bool isEmpty = p.Models == null || p.Models.Length == 0;
                if (isEmpty)
                {
                    // 列表为空：直接用内置目录兜底，避免下拉框空白
                    p.Models = catalogModels;
                    if (string.IsNullOrEmpty(p.DefaultModel) ||
                        Array.IndexOf(p.Models, p.DefaultModel) < 0)
                        p.DefaultModel = kvp.Value.DefaultModel;
                    changed = true;
                    continue;
                }

                // 用户自己整理过的列表不碰；目录版本没更新也不补
                if (p.ModelsCustomized || !catalogOutdated) continue;

                var merged = AppConfig.MergeModels(p.Models, catalogModels);
                if (merged.Length != p.Models.Length)
                {
                    p.Models = merged;
                    changed = true;
                }
                if (string.IsNullOrEmpty(p.DefaultModel) ||
                    Array.IndexOf(p.Models, p.DefaultModel) < 0)
                {
                    p.DefaultModel = p.Models[0];
                    changed = true;
                }
            }
            if (catalogOutdated)
            {
                config.ModelCatalogVersion = AppConfig.CurrentModelCatalogVersion;
                changed = true;
            }

            // CurrentModel 落在列表外时回落到该 provider 的主模型（不因目录变动而失效）
            if (config.Providers.ContainsKey(config.CurrentProvider))
            {
                var currentProviderConfig = config.Providers[config.CurrentProvider];
                if (currentProviderConfig.Models != null && currentProviderConfig.Models.Length > 0 &&
                    (string.IsNullOrEmpty(config.CurrentModel) ||
                     Array.IndexOf(currentProviderConfig.Models, config.CurrentModel) < 0))
                {
                    config.CurrentModel = currentProviderConfig.DefaultModel ?? currentProviderConfig.Models[0];
                    changed = true;
                }
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
