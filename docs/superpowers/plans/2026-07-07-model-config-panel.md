# 模型配置弹窗 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Excel Ribbon 新增"模型配置"按钮，点击后弹出左右分栏的模型管理界面，支持 10 个内置厂商切换、模型选择、API Key 配置、BaseUrl/MaxTurns 高级设置，保存后立即对所有正在进行的对话生效。

**Architecture:** Ribbon 按钮 → C# `OnShowModelConfig` 通过 WebView2 `PostWebMessageAsString` 发 `{type:'open_model_config'}` → 前端 React `App.tsx` 监听后渲染 `<ModelConfigPanel>` modal → 通过 `sendToHostWithResponse` 与 C# `MessageBridge.HandleMessage` 通信 → `ConfigManager` + `SecurityManager` 持久化 → `RefreshConfigForAllSessions` 立即下发到所有 sidecar。

**Tech Stack:** C# (.NET Framework 4.x, COM interop, System.Text.Json), React 18 + TypeScript + Vite, WebView2

## Global Constraints

- **C# 语言版本**：C# 9.0（`/langversion:9.0`，由 `scripts\_compile_only.ps1` 指定）
- **序列化**：camelCase 输出（`MessageBridge._jsonOptions` 已配 `JsonNamingPolicy.CamelCase`）；2D 数组必须用 `Object2DArrayConverter`/`String2DArrayConverter`（本功能不涉及 2D 数组）
- **API Key 安全**：绝不通过 `get_model_config` 返回明文 API Key；只返回 `hasApiKey` + `apiKeyPreview`（脱敏）；保存通过 `SecurityManager.SaveApiKey` DPAPI 加密
- **CLSID 注册**：每次重新编译 DLL 后必须运行 `scripts\register-user.ps1` 注册 CLSID，否则 Excel 无法加载
- **Excel 关闭**：编译前必须确认 EXCEL.EXE 进程已关闭，否则 DLL 被占用导致编译失败
- **Ribbon DispId**：`IRibbonCallbacks` 的 `DispId` 必须连续（1,2,3,4），否则 COM IDispatch 调用会错位
- **前端样式**：CodeX 简洁风格，紧凑间距，小圆角（6-10px），SVG 图标不用 emoji，避免 AI 风格竖线
- **构建命令**：C# 编译 `powershell -ExecutionPolicy Bypass -File scripts\_compile_only.ps1`；前端构建 `cd src\DeepExcel.UI && npm run build`（自动输出到 `src\DeepExcel.AddIn\WebViewAssets`）
- **C# 文件路径**：所有 C# 文件在 `src\DeepExcel.AddIn\` 下；前端文件在 `src\DeepExcel.UI\src\` 下
- **MigrateConfig 约束**：旧 config.json 升级时必须自动补全 5 个新 provider，不覆盖用户已修改的字段

---

### Task 1: 扩展 AppConfig 新增 5 个国产厂商

**Files:**
- Modify: `src\DeepExcel.AddIn\Config\AppConfig.cs` (CreateDefault 方法 L21-70, MigrateConfig 方法 L240-301)

**Interfaces:**
- Produces: `AppConfig.CreateDefault()` 返回的 `Providers` 字典新增 `kimi`/`qwen`/`zhipu`/`minimax`/`doubao` 5 个 key；`MigrateConfig` 自动补全这 5 个 provider 给旧 config.json

- [ ] **Step 1: 在 `CreateDefault()` 的 `custom` provider 之前插入 5 个新 provider**

打开 `src\DeepExcel.AddIn\Config\AppConfig.cs`，找到 L60 的 `cfg.Providers["custom"]`，在它**之前**插入：

```csharp
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
```

- [ ] **Step 2: 为现有 5 个 provider 补充 `DefaultModel` 字段**

现有 5 个 provider 都没有显式设置 `DefaultModel`。在 `CreateDefault()` 中为每个 provider 加 `DefaultModel`：

- `anthropic`: 加 `DefaultModel = "claude-3-5-sonnet-20241022",`（在 `Models = ...` 行之后）
- `deepseek`: 加 `DefaultModel = "deepseek-chat",`
- `stepfun`: 加 `DefaultModel = "step-3.7-flash",`
- `openai`: 加 `DefaultModel = "gpt-4o",`
- `custom`: 加 `DefaultModel = "custom-model",`

- [ ] **Step 3: 在 `MigrateConfig` 方法中新增 5 个 provider 的迁移逻辑**

找到 `MigrateConfig` 方法（L240），在 stepfun 迁移块（L245-257）**之后**，添加 5 个新 provider 的迁移：

```csharp
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
```

- [ ] **Step 4: 编译验证**

Run: `powershell -ExecutionPolicy Bypass -File scripts\_compile_only.ps1`
Expected: `=== BUILD SUCCESS ===`

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.AddIn/Config/AppConfig.cs
git commit -m "feat(config): add 5 CN providers (Kimi/Qwen/Zhipu/Minimax/Doubao)"
```

---

### Task 2: 扩展 SafeProvider 结构 + 添加脱敏工具方法

**Files:**
- Modify: `src\DeepExcel.AddIn\Security\SecurityManager.cs` (SafeProvider 类 L210-215, GetSafeConfig 方法 L183-193)

**Interfaces:**
- Produces: `SafeProvider` 新增 `Type`/`BaseUrl`/`DefaultModel`/`SupportsVision`/`ApiKeyPreview` 字段；`SecurityManager.MaskApiKey(string key)` 静态方法返回脱敏预览

- [ ] **Step 1: 扩展 `SafeProvider` 类添加新字段**

找到 `SecurityManager.cs` L210-215 的 `SafeProvider` 类，替换为：

```csharp
    public class SafeProvider
    {
        public string DisplayName { get; set; }
        public string Type { get; set; }
        public string BaseUrl { get; set; }
        public string DefaultModel { get; set; }
        public bool SupportsVision { get; set; }
        public string[] Models { get; set; }
        public bool HasApiKey { get; set; }
        public string ApiKeyPreview { get; set; }
    }
```

- [ ] **Step 2: 重写 `GetSafeConfig` 方法填充新字段**

找到 L183-193 的 `GetSafeConfig` 方法，替换为：

```csharp
        public SafeConfig GetSafeConfig(Config.AppConfig config)
        {
            var safe = new SafeConfig
            {
                CurrentProvider = config.CurrentProvider,
                CurrentModel = config.CurrentModel,
                Providers = new Dictionary<string, SafeProvider>(config.Providers.Count),
                General = config.General,
                UI = config.UI
            };
            foreach (var kvp in config.Providers)
            {
                var p = kvp.Value;
                var key = GetApiKey(kvp.Key);
                safe.Providers[kvp.Key] = new SafeProvider
                {
                    DisplayName = p.DisplayName,
                    Type = p.Type,
                    BaseUrl = p.BaseUrl,
                    DefaultModel = p.DefaultModel,
                    SupportsVision = p.SupportsVision,
                    Models = p.Models,
                    HasApiKey = !string.IsNullOrEmpty(key),
                    ApiKeyPreview = MaskApiKey(key)
                };
            }
            return safe;
        }

        /// <summary>
        /// ★ 脱敏 API Key 用于前端显示预览。
        /// 规则：长度 >= 8 时显示前 4 + *** + 后 3；长度 < 8 时全 ***；空时返回空字符串。
        /// </summary>
        public static string MaskApiKey(string key)
        {
            if (string.IsNullOrEmpty(key)) return "";
            if (key.Length < 8) return "***";
            return key.Substring(0, 4) + "***..." + key.Substring(key.Length - 3);
        }
```

- [ ] **Step 3: 编译验证**

Run: `powershell -ExecutionPolicy Bypass -File scripts\_compile_only.ps1`
Expected: `=== BUILD SUCCESS ===`

- [ ] **Step 4: Commit**

```bash
git add src/DeepExcel.AddIn/Security/SecurityManager.cs
git commit -m "feat(security): extend SafeProvider with Type/BaseUrl/Vision/ApiKeyPreview"
```

---

### Task 3: Ribbon 新增"模型配置"按钮

**Files:**
- Modify: `src\DeepExcel.AddIn\Resources\DeepExcelRibbon.xml`
- Modify: `src\DeepExcel.AddIn\IRibbonCallbacks.cs`
- Modify: `src\DeepExcel.AddIn\ThisAddIn.cs` (新增 OnShowModelConfig 实现)

**Interfaces:**
- Produces: Ribbon XML 新增 `btnModelConfig` 按钮；`IRibbonCallbacks` 新增 `[DispId(4)] OnShowModelConfig`；`ThisAddIn` 实现该方法，通过 WebView2 发 `{type:'open_model_config'}` 消息给当前活动窗口的前端

- [ ] **Step 1: 在 Ribbon XML 中新增按钮**

打开 `src\DeepExcel.AddIn\Resources\DeepExcelRibbon.xml`，在 `btnTogglePanel` 和 `btnHelp` 之间插入：

```xml
          <button id="btnModelConfig" label="模型配置" size="large"
                  imageMso="DataSources"
                  onAction="OnShowModelConfig"
                  screentip="模型配置"
                  supertip="配置 AI 模型厂商、API Key 和高级参数。"/>
```

完整 XML 应该是：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="tabDeepExcel" label="DeepExcel" insertAfterMso="TabHome">
        <group id="grpTools" label="AI Agent">
          <button id="btnTogglePanel" label="打开面板" size="large"
                  imageMso="InsertFunction"
                  onAction="OnTogglePanel"
                  screentip="打开 DeepExcel AI 面板"
                  supertip="在面板中输入你的需求，AI 会自动帮你处理 Excel。"/>
          <button id="btnModelConfig" label="模型配置" size="large"
                  imageMso="DataSources"
                  onAction="OnShowModelConfig"
                  screentip="模型配置"
                  supertip="配置 AI 模型厂商、API Key 和高级参数。"/>
          <button id="btnHelp" label="帮助" size="large"
                  imageMso="Info"
                  onAction="OnShowHelp"
                  screentip="使用帮助"
                  supertip="查看 DeepExcel 的使用说明。"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

- [ ] **Step 2: 在 `IRibbonCallbacks.cs` 新增 `OnShowModelConfig` 方法声明**

打开 `src\DeepExcel.AddIn\IRibbonCallbacks.cs`，在 `OnShowHelp` 之后添加 `[DispId(4)]`：

```csharp
using System;
using System.Runtime.InteropServices;

namespace DeepExcel.AddIn
{
    [ComVisible(true)]
    [Guid("C3D4E5F6-A7B8-4C5D-9A7F-1C2D3E4F5A6B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonCallbacks
    {
        [DispId(1)]
        void OnRibbonLoad(object ribbon);

        [DispId(2)]
        void OnTogglePanel(object control);

        [DispId(3)]
        void OnShowHelp(object control);

        [DispId(4)]
        void OnShowModelConfig(object control);
    }
}
```

- [ ] **Step 3: 在 `ThisAddIn.cs` 实现 `OnShowModelConfig`**

找到 `OnShowHelp` 方法（L726-737），在它**之后**添加新方法。该方法需要：如果面板未打开则先打开，然后通过 WebView2 发消息给前端。

```csharp
        void IRibbonCallbacks.OnShowModelConfig(object control)
        {
            Log("OnShowModelConfig called");
            try
            {
                // 获取当前活动窗口
                Microsoft.Office.Interop.Excel.Window activeWindow = null;
                try { activeWindow = _excelApp.ActiveWindow; }
                catch (Exception wex) { Log("Get ActiveWindow failed: " + wex.Message); }
                if (activeWindow == null)
                {
                    MessageBox.Show("请先打开一个工作簿再点击「模型配置」。",
                        "DeepExcel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string windowKey = activeWindow.Caption as string ?? "";

                // 如果该窗口没有 CTP 或 CTP 不可见，先创建/显示面板
                Microsoft.Office.Core.CustomTaskPane ctp = null;
                bool needDelay = false;
                if (!_ctpsByWindow.TryGetValue(windowKey, out ctp) || ctp == null)
                {
                    // 面板未创建，先调 OnTogglePanel 创建
                    Log("OnShowModelConfig: panel not created, creating first");
                    ((IRibbonCallbacks)this).OnTogglePanel(control);
                    needDelay = true;
                }
                else
                {
                    try
                    {
                        if (!ctp.Visible)
                        {
                            ctp.Visible = true;
                            needDelay = true;
                        }
                    }
                    catch { needDelay = true; }
                }

                // 发送 open_model_config 消息（如果面板刚创建/显示，延迟 600ms 等 WebView 初始化）
                if (needDelay)
                {
                    System.Threading.Timer timer = null;
                    timer = new System.Threading.Timer(_ =>
                    {
                        try { SendOpenModelConfigToActivePane(windowKey); }
                        catch (Exception ex) { Log("OnShowModelConfig delayed send failed: " + ex.Message); }
                        finally { timer?.Dispose(); }
                    }, null, 600, System.Threading.Timeout.Infinite);
                }
                else
                {
                    SendOpenModelConfigToActivePane(windowKey);
                }
            }
            catch (Exception ex)
            {
                Log("OnShowModelConfig error: " + ex.Message);
                MessageBox.Show("打开模型配置失败: " + ex.Message, "DeepExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 向活动窗口的 WebView 发送 open_model_config 消息
        /// </summary>
        private void SendOpenModelConfigToActivePane(string windowKey)
        {
            if (!_ctpsByWindow.TryGetValue(windowKey, out var ctp) || ctp == null)
            {
                Log("SendOpenModelConfigToActivePane: no CTP for window=" + windowKey);
                return;
            }
            TaskPaneControl pane = null;
            try { pane = (TaskPaneControl)ctp.ContentControl; }
            catch { }
            if (pane == null || pane.IsDisposed || pane.WebView == null)
            {
                Log("SendOpenModelConfigToActivePane: pane/webview invalid");
                return;
            }

            var webView = pane.WebView;
            var msg = "{\"type\":\"open_model_config\"}";
            try
            {
                if (webView.InvokeRequired)
                {
                    webView.BeginInvoke(new Action<string>(m =>
                    {
                        try
                        {
                            if (webView.CoreWebView2 != null && !webView.IsDisposed)
                                webView.CoreWebView2.PostWebMessageAsString(m);
                        }
                        catch (Exception ex) { Log("SendOpenModelConfig (UI thread) error: " + ex.Message); }
                    }), msg);
                }
                else
                {
                    if (webView.CoreWebView2 != null)
                        webView.CoreWebView2.PostWebMessageAsString(msg);
                }
                Log("SendOpenModelConfigToActivePane: sent open_model_config to window=" + windowKey);
            }
            catch (Exception ex)
            {
                Log("SendOpenModelConfigToActivePane error: " + ex.Message);
            }
        }
```

- [ ] **Step 4: 编译验证**

Run: `powershell -ExecutionPolicy Bypass -File scripts\_compile_only.ps1`
Expected: `=== BUILD SUCCESS ===`

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.AddIn/Resources/DeepExcelRibbon.xml src/DeepExcel.AddIn/IRibbonCallbacks.cs src/DeepExcel.AddIn/ThisAddIn.cs
git commit -m "feat(ribbon): add model config button with WebView2 message dispatch"
```

---

### Task 4: MessageBridge 新增配置消息处理

**Files:**
- Modify: `src\DeepExcel.AddIn\Bridge\MessageBridge.cs` (HandleMessage 方法 L441-457 无 session 分支)

**Interfaces:**
- Consumes: `SecurityManager.Instance.GetSafeConfig(ConfigManager.Instance.Current)` 返回完整脱敏配置；`ConfigManager.Instance.SwitchProvider` / `UpdateApiKey` / `Save`；`RefreshConfigForAllSessions()`
- Produces: `HandleMessage` 新增 3 个 case：`get_model_config`/`save_model_config`/`test_api_key`，返回对应响应类型 `model_config`/`config_saved`/`api_test_result`

- [ ] **Step 1: 在 `HandleMessage` 无 session 分支新增 3 个 case**

找到 `MessageBridge.cs` L441-457 的无 session switch 块，在 `case "delete_snapshot"` 之后、`}` 之前插入：

```csharp
                    case "get_model_config":
                        return HandleGetModelConfig();
                    case "save_model_config":
                        return HandleSaveModelConfig(msg);
                    case "test_api_key":
                        return HandleTestApiKey(msg);
```

- [ ] **Step 2: 实现 `HandleGetModelConfig` 方法**

在 `HandleMessage` 方法**之后**（L496 附近）添加：

```csharp
        /// <summary>
        /// ★ 返回当前配置（API Key 脱敏）给前端模型配置弹窗
        /// </summary>
        private string HandleGetModelConfig()
        {
            try
            {
                var cfg = ConfigManager.Instance.Current;
                var safe = SecurityManager.Instance.GetSafeConfig(cfg);
                return MakeResponse("model_config", safe);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleGetModelConfig failed", ex);
                return MakeError("加载配置失败: " + ex.Message);
            }
        }
```

- [ ] **Step 3: 实现 `HandleSaveModelConfig` 方法**

紧接 `HandleGetModelConfig` 之后添加：

```csharp
        /// <summary>
        /// ★ 保存模型配置（provider/model/apiKey/baseUrl/maxTurns），立即对所有 session 生效
        /// 占位符 "***keep***" 表示用户未修改 API Key，跳过保存
        /// </summary>
        private string HandleSaveModelConfig(Message msg)
        {
            try
            {
                var payload = msg.Payload;
                var provider = payload.GetProperty("provider").GetString();
                var model = payload.GetProperty("model").GetString();
                var apiKey = payload.GetProperty("apiKey").GetString();
                var baseUrl = payload.GetProperty("baseUrl").GetString();
                int maxTurns = payload.GetProperty("maxTurns").GetInt32();

                if (string.IsNullOrEmpty(provider))
                {
                    return MakeError("provider 不能为空");
                }

                var cfg = ConfigManager.Instance.Current;
                if (!cfg.Providers.ContainsKey(provider))
                {
                    return MakeError($"未知的 provider: {provider}");
                }

                // 1. 更新 BaseUrl（如果非空且与默认不同）
                if (!string.IsNullOrEmpty(baseUrl))
                {
                    cfg.Providers[provider].BaseUrl = baseUrl;
                }

                // 2. 保存 API Key（跳过占位符和空值）
                if (!string.IsNullOrEmpty(apiKey) && apiKey != "***keep***")
                {
                    ConfigManager.Instance.UpdateApiKey(provider, apiKey);
                }

                // 3. 切换 provider + model
                ConfigManager.Instance.SwitchProvider(provider, model);

                // 4. 更新 MaxTurns
                if (cfg.General == null) cfg.General = new Config.GeneralSettings();
                if (maxTurns > 0 && maxTurns <= 200)
                {
                    cfg.General.MaxTurns = maxTurns;
                }

                // 5. 持久化
                ConfigManager.Instance.Save();

                // 6. 立即对所有 session 生效
                RefreshConfigForAllSessions();

                Logger.Instance.Info("MessageBridge",
                    $"HandleSaveModelConfig: provider={provider}, model={model}, maxTurns={maxTurns}, apiKeyChanged={apiKey != "***keep***" && !string.IsNullOrEmpty(apiKey)}");

                return MakeResponse("config_saved", new { success = true });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSaveModelConfig failed", ex);
                return MakeError("保存配置失败: " + ex.Message);
            }
        }
```

- [ ] **Step 4: 实现 `HandleTestApiKey` 方法**

紧接 `HandleSaveModelConfig` 之后添加。用 HttpClient 向 Anthropic 兼容端点发最小请求测试：

```csharp
        /// <summary>
        /// ★ 测试 API Key 连接（不保存任何数据）
        /// 向 baseUrl 发一个 Anthropic Messages API 的 1-token 请求
        /// </summary>
        private string HandleTestApiKey(Message msg)
        {
            try
            {
                var payload = msg.Payload;
                var provider = payload.GetProperty("provider").GetString();
                var apiKey = payload.GetProperty("apiKey").GetString();
                var baseUrl = payload.GetProperty("baseUrl").GetString();
                var model = payload.GetProperty("model").GetString();

                if (string.IsNullOrEmpty(apiKey) || apiKey == "***keep***")
                {
                    return MakeResponse("api_test_result", new { success = false, error = "请先输入 API Key" });
                }
                if (string.IsNullOrEmpty(baseUrl))
                {
                    return MakeResponse("api_test_result", new { success = false, error = "Base URL 为空" });
                }

                // 用 Task.Run 在后台线程执行，避免阻塞 UI 线程
                var sw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    using var client = new System.Net.Http.HttpClient();
                    client.Timeout = TimeSpan.FromSeconds(15);

                    // 构造 Anthropic Messages API 最小请求
                    var testUrl = baseUrl.TrimEnd('/') + "/v1/messages";
                    var body = new
                    {
                        model = model,
                        max_tokens = 1,
                        messages = new[] { new { role = "user", content = "hi" } }
                    };
                    var json = System.Text.Json.JsonSerializer.Serialize(body);
                    var content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");

                    var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, testUrl);
                    request.Content = content;
                    request.Headers.Add("x-api-key", apiKey);
                    request.Headers.Add("anthropic-version", "2023-06-01");

                    var response = client.SendAsync(request).Result;
                    sw.Stop();

                    if (response.IsSuccessStatusCode)
                    {
                        Logger.Instance.Info("MessageBridge", $"HandleTestApiKey: success, latency={sw.ElapsedMilliseconds}ms");
                        return MakeResponse("api_test_result", new { success = true, latencyMs = sw.ElapsedMilliseconds, error = (string)null });
                    }
                    else
                    {
                        var errBody = response.Content.ReadAsStringAsync().Result;
                        Logger.Instance.Warning("MessageBridge", $"HandleTestApiKey: HTTP {response.StatusCode}, body={errBody}");
                        return MakeResponse("api_test_result", new
                        {
                            success = false,
                            latencyMs = sw.ElapsedMilliseconds,
                            error = $"HTTP {response.StatusCode}: {errBody}"
                        });
                    }
                }
                catch (System.Net.Http.HttpRequestException hex)
                {
                    sw.Stop();
                    Logger.Instance.Warning("MessageBridge", "HandleTestApiKey network error: " + hex.Message);
                    return MakeResponse("api_test_result", new { success = false, latencyMs = sw.ElapsedMilliseconds, error = "网络错误: " + hex.Message });
                }
                catch (AggregateException aex)
                {
                    sw.Stop();
                    var inner = aex.InnerException?.Message ?? aex.Message;
                    Logger.Instance.Warning("MessageBridge", "HandleTestApiKey error: " + inner);
                    return MakeResponse("api_test_result", new { success = false, latencyMs = sw.ElapsedMilliseconds, error = inner });
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleTestApiKey failed", ex);
                return MakeError("测试连接失败: " + ex.Message);
            }
        }
```

- [ ] **Step 5: 编译验证**

Run: `powershell -ExecutionPolicy Bypass -File scripts\_compile_only.ps1`
Expected: `=== BUILD SUCCESS ===`

- [ ] **Step 6: Commit**

```bash
git add src/DeepExcel.AddIn/Bridge/MessageBridge.cs
git commit -m "feat(bridge): add get/save/test model config message handlers"
```

---

### Task 5: 前端新增厂商图标 + 类型定义

**Files:**
- Create: `src\DeepExcel.UI\src\providerIcons.ts`
- Modify: `src\DeepExcel.UI\src\types.ts`

**Interfaces:**
- Produces: `providerIcons` 映射（key → SVG JSX 函数）；`ProviderInfo`/`ModelConfig` 类型

- [ ] **Step 1: 创建 `providerIcons.ts`**

创建 `src\DeepExcel.UI\src\providerIcons.ts`，包含 10 个厂商的圆形图标 SVG 组件。每个图标是 28x28 圆形背景 + 白色首字母：

```typescript
/**
 * 厂商图标：圆形品牌色背景 + 白色首字母
 * 避免 AI 风格竖线，使用真实品牌色
 */

interface IconProps {
  size?: number
}

const CircleIcon = ({ color, letter, size = 28 }: IconProps & { color: string; letter: string }) => (
  <svg width={size} height={size} viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg">
    <circle cx="14" cy="14" r="13" fill={color} />
    <text
      x="14"
      y="14"
      textAnchor="middle"
      dominantBaseline="central"
      fill="white"
      fontSize="13"
      fontWeight="600"
      fontFamily="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    >
      {letter}
    </text>
  </svg>
)

export const providerIcons: Record<string, (props: IconProps) => JSX.Element> = {
  anthropic: (props) => <CircleIcon color="#D97757" letter="A" {...props} />,
  deepseek: (props) => <CircleIcon color="#4D6BFE" letter="D" {...props} />,
  stepfun: (props) => <CircleIcon color="#7B61FF" letter="S" {...props} />,
  openai: (props) => <CircleIcon color="#000000" letter="O" {...props} />,
  kimi: (props) => <CircleIcon color="#1A1A1A" letter="K" {...props} />,
  qwen: (props) => <CircleIcon color="#615CED" letter="Q" {...props} />,
  zhipu: (props) => <CircleIcon color="#3B82F6" letter="Z" {...props} />,
  minimax: (props) => <CircleIcon color="#2563EB" letter="M" {...props} />,
  doubao: (props) => <CircleIcon color="#FF1A4B" letter="D" {...props} />,
  custom: (props) => <CircleIcon color="#6B7280" letter="+" {...props} />
}

/** 厂商显示顺序（左侧列表按此顺序渲染） */
export const providerOrder = [
  'anthropic', 'deepseek', 'stepfun', 'openai',
  'kimi', 'qwen', 'zhipu', 'minimax', 'doubao', 'custom'
]
```

- [ ] **Step 2: 在 `types.ts` 新增 `ProviderInfo` 和 `ModelConfig` 类型**

打开 `src\DeepExcel.UI\src\types.ts`，在文件末尾追加：

```typescript
// ★ 模型配置弹窗用类型（对应 C# SafeConfig/SafeProvider 结构）

export type ProviderInfo = {
  displayName: string
  type: string
  baseUrl: string
  defaultModel: string
  supportsVision: boolean
  models: string[]
  hasApiKey: boolean
  apiKeyPreview: string
}

export type ModelConfig = {
  currentProvider: string
  currentModel: string
  providers: Record<string, ProviderInfo>
  general: {
    maxRetries: number
    requestTimeoutSeconds: number
    autoCreateSnapshot: boolean
    requireConfirmation: boolean
    maxConversationHistory: number
    maxTurns: number
  }
  ui: {
    theme: string
    language: string
    showTokenUsage: boolean
    streamOutput: boolean
  }
}
```

- [ ] **Step 3: Commit**

```bash
git add src/DeepExcel.UI/src/providerIcons.ts src/DeepExcel.UI/src/types.ts
git commit -m "feat(ui): add provider icons and ModelConfig types"
```

---

### Task 6: 前端新增 `ModelConfigPanel.tsx` 组件

**Files:**
- Create: `src\DeepExcel.UI\src\components\ModelConfigPanel.tsx`

**Interfaces:**
- Consumes: `sendToHostWithResponse` from `bridge.ts`；`providerIcons`/`providerOrder` from `providerIcons.ts`；`ProviderInfo`/`ModelConfig` from `types.ts`
- Produces: `<ModelConfigPanel open={bool} onClose={() => void} />` 组件，发 `get_model_config`/`save_model_config`/`test_api_key` 消息

- [ ] **Step 1: 创建 `ModelConfigPanel.tsx`**

创建 `src\DeepExcel.UI\src\components\ModelConfigPanel.tsx`：

```tsx
import { useState, useEffect } from 'react'
import { sendToHostWithResponse } from '../bridge'
import { providerIcons, providerOrder } from '../providerIcons'
import type { ModelConfig, ProviderInfo } from '../types'

interface Props {
  open: boolean
  onClose: () => void
}

const KEEP_PLACEHOLDER = '***keep***'

export function ModelConfigPanel({ open, onClose }: Props) {
  const [config, setConfig] = useState<ModelConfig | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [selectedProvider, setSelectedProvider] = useState('')
  const [model, setModel] = useState('')
  const [apiKey, setApiKey] = useState(KEEP_PLACEHOLDER)
  const [baseUrl, setBaseUrl] = useState('')
  const [maxTurns, setMaxTurns] = useState(20)
  const [showAdvanced, setShowAdvanced] = useState(false)
  const [showApiKey, setShowApiKey] = useState(false)
  const [testing, setTesting] = useState(false)
  const [testResult, setTestResult] = useState<{ success: boolean; latencyMs?: number; error?: string | null } | null>(null)
  const [saving, setSaving] = useState(false)
  const [saveMsg, setSaveMsg] = useState('')

  // 打开时加载配置
  useEffect(() => {
    if (open) {
      loadConfig()
      setTestResult(null)
      setSaveMsg('')
    }
  }, [open])

  // 加载配置后同步表单
  useEffect(() => {
    if (config) {
      setSelectedProvider(config.currentProvider)
    }
  }, [config])

  // 切换 provider 时更新表单字段
  useEffect(() => {
    if (config && selectedProvider && config.providers[selectedProvider]) {
      const p = config.providers[selectedProvider]
      // 如果选中的是当前 provider，用 currentModel；否则用 defaultModel
      setModel(selectedProvider === config.currentProvider ? config.currentModel : (p.defaultModel || p.models[0] || ''))
      setApiKey(KEEP_PLACEHOLDER)
      setBaseUrl(p.baseUrl || '')
      setMaxTurns(config.general?.maxTurns ?? 20)
      setTestResult(null)
      setSaveMsg('')
    }
  }, [selectedProvider, config])

  async function loadConfig() {
    setLoading(true)
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'get_model_config', payload: {} },
        'model_config'
      )
      if (resp?.type === 'model_config' && resp.payload?.providers) {
        setConfig(resp.payload as ModelConfig)
      } else if (resp?.type === 'error') {
        setError(resp.payload?.message || '加载配置失败')
      } else {
        setError('加载配置失败：未收到响应')
      }
    } catch (e) {
      setError(`加载配置失败: ${e}`)
    } finally {
      setLoading(false)
    }
  }

  async function handleSave() {
    if (!selectedProvider || !model) {
      setSaveMsg('请先选择厂商和模型')
      return
    }
    setSaving(true)
    setSaveMsg('')
    try {
      const resp = await sendToHostWithResponse(
        {
          type: 'save_model_config',
          payload: {
            provider: selectedProvider,
            model,
            apiKey,
            baseUrl,
            maxTurns
          }
        },
        'config_saved'
      )
      if (resp?.type === 'config_saved' && resp.payload?.success) {
        setSaveMsg('✓ 已保存并应用')
        // 重新加载配置以更新 hasApiKey 状态
        await loadConfig()
      } else if (resp?.type === 'error') {
        setSaveMsg('✗ ' + (resp.payload?.message || '保存失败'))
      } else {
        setSaveMsg('✗ 保存失败：未收到响应')
      }
    } catch (e) {
      setSaveMsg(`✗ 保存失败: ${e}`)
    } finally {
      setSaving(false)
    }
  }

  async function handleTest() {
    if (!selectedProvider || !model) {
      setTestResult({ success: false, error: '请先选择厂商和模型' })
      return
    }
    if (!apiKey || apiKey === KEEP_PLACEHOLDER) {
      setTestResult({ success: false, error: '请先输入 API Key（保存后才能测试已存的 key，或直接在输入框填入新 key 测试）' })
      return
    }
    setTesting(true)
    setTestResult(null)
    try {
      const resp = await sendToHostWithResponse(
        {
          type: 'test_api_key',
          payload: {
            provider: selectedProvider,
            apiKey,
            baseUrl,
            model
          }
        },
        'api_test_result',
        20000  // 测试连接超时 20s
      )
      if (resp?.type === 'api_test_result') {
        setTestResult({
          success: resp.payload?.success,
          latencyMs: resp.payload?.latencyMs,
          error: resp.payload?.error
        })
      } else {
        setTestResult({ success: false, error: '测试超时或未收到响应' })
      }
    } catch (e) {
      setTestResult({ success: false, error: `测试失败: ${e}` })
    } finally {
      setTesting(false)
    }
  }

  if (!open) return null

  const currentProviderInfo: ProviderInfo | undefined = config?.providers?.[selectedProvider]

  return (
    <div className="config-overlay" onClick={onClose}>
      <div className="config-panel" onClick={e => e.stopPropagation()}>
        <div className="config-header">
          <h3>模型配置</h3>
          <button className="config-close-btn" onClick={onClose} title="关闭">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
          </button>
        </div>

        {loading && !config ? (
          <div className="config-loading">加载中...</div>
        ) : error ? (
          <div className="config-error">
            <div>{error}</div>
            <button onClick={loadConfig} className="config-retry-btn">重试</button>
          </div>
        ) : config ? (
          <div className="config-body">
            {/* 左侧厂商列表 */}
            <aside className="config-provider-list">
              {providerOrder.filter(k => config.providers[k]).map(key => {
                const p = config.providers[key]
                const Icon = providerIcons[key] || providerIcons.custom
                const isActive = key === selectedProvider
                const isCurrent = key === config.currentProvider
                return (
                  <div
                    key={key}
                    className={`config-provider-item ${isActive ? 'active' : ''}`}
                    onClick={() => setSelectedProvider(key)}
                  >
                    <Icon size={24} />
                    <span className="config-provider-name">{p.displayName}</span>
                    {p.hasApiKey && <span className="config-provider-dot" title="已配置 API Key" />}
                    {isCurrent && <span className="config-provider-current" title="当前使用" />}
                  </div>
                )
              })}
            </aside>

            {/* 右侧配置区 */}
            <main className="config-provider-config">
              {currentProviderInfo ? (
                <>
                  <div className="config-section">
                    <label className="config-label">模型</label>
                    <select
                      className="config-select"
                      value={model}
                      onChange={e => setModel(e.target.value)}
                    >
                      {currentProviderInfo.models.map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </select>
                    {currentProviderInfo.supportsVision && (
                      <span className="config-badge">支持视觉</span>
                    )}
                  </div>

                  <div className="config-section">
                    <label className="config-label">API Key</label>
                    <div className="config-apikey-row">
                      <input
                        type={showApiKey ? 'text' : 'password'}
                        className="config-input"
                        value={apiKey === KEEP_PLACEHOLDER ? '' : apiKey}
                        placeholder={currentProviderInfo.hasApiKey ? `已配置（${currentProviderInfo.apiKeyPreview}）` : '输入 API Key'}
                        onChange={e => setApiKey(e.target.value)}
                      />
                      <button
                        className="config-toggle-btn"
                        onClick={() => setShowApiKey(!showApiKey)}
                        title={showApiKey ? '隐藏' : '显示'}
                        type="button"
                      >
                        {showApiKey ? '隐藏' : '显示'}
                      </button>
                    </div>
                  </div>

                  {/* 高级设置折叠区 */}
                  <div className="config-advanced-toggle" onClick={() => setShowAdvanced(!showAdvanced)}>
                    <svg
                      width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
                      style={{ transform: showAdvanced ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }}
                    >
                      <polyline points="9 18 15 12 9 6"></polyline>
                    </svg>
                    <span>高级设置</span>
                  </div>

                  {showAdvanced && (
                    <div className="config-advanced">
                      <div className="config-section">
                        <label className="config-label">Base URL</label>
                        <input
                          type="text"
                          className="config-input"
                          value={baseUrl}
                          onChange={e => setBaseUrl(e.target.value)}
                          placeholder="https://..."
                        />
                      </div>
                      <div className="config-section">
                        <label className="config-label">MaxTurns（工具调用循环上限）</label>
                        <input
                          type="number"
                          className="config-input config-input-narrow"
                          value={maxTurns}
                          min={1}
                          max={200}
                          onChange={e => setMaxTurns(parseInt(e.target.value) || 20)}
                        />
                      </div>
                    </div>
                  )}

                  {/* 测试结果 */}
                  {testResult && (
                    <div className={`config-test-result ${testResult.success ? 'success' : 'error'}`}>
                      {testResult.success
                        ? `✓ 连接成功（${testResult.latencyMs}ms）`
                        : `✗ ${testResult.error}`}
                    </div>
                  )}

                  {/* 保存消息 */}
                  {saveMsg && (
                    <div className={`config-save-msg ${saveMsg.startsWith('✓') ? 'success' : 'error'}`}>
                      {saveMsg}
                    </div>
                  )}

                  {/* 底部按钮 */}
                  <div className="config-actions">
                    <button
                      className="config-test-btn"
                      onClick={handleTest}
                      disabled={testing}
                      type="button"
                    >
                      {testing ? '测试中...' : '测试连接'}
                    </button>
                    <button
                      className="config-save-btn"
                      onClick={handleSave}
                      disabled={saving}
                      type="button"
                    >
                      {saving ? '保存中...' : '保存并应用'}
                    </button>
                  </div>
                </>
              ) : (
                <div className="config-empty">请选择一个厂商</div>
              )}
            </main>
          </div>
        ) : null}
      </div>
    </div>
  )
}
```

- [ ] **Step 2: Commit**

```bash
git add src/DeepExcel.UI/src/components/ModelConfigPanel.tsx
git commit -m "feat(ui): add ModelConfigPanel with left-right split layout"
```

---

### Task 7: 前端新增弹窗 CSS 样式

**Files:**
- Modify: `src\DeepExcel.UI\src\styles.css` (文件末尾追加)

**Interfaces:**
- Produces: `.config-overlay`/`.config-panel`/`.config-provider-list`/`.config-provider-item` 等样式类

- [ ] **Step 1: 在 `styles.css` 末尾追加模型配置弹窗样式**

```css
/* === 模型配置弹窗（左右分栏） === */
.config-overlay {
  position: absolute;
  inset: 0;
  background: rgba(0, 0, 0, 0.25);
  z-index: 70;
  display: flex;
  align-items: center;
  justify-content: center;
  animation: fadeIn 0.15s ease-out;
}

.config-panel {
  width: 95%;
  max-width: 560px;
  max-height: 80%;
  background: var(--color-bg);
  border-radius: 8px;
  display: flex;
  flex-direction: column;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.18);
  overflow: hidden;
  animation: convPanelIn 0.18s ease-out;
}

.config-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 10px 14px;
  border-bottom: 1px solid var(--color-border);
}

.config-header h3 {
  font-size: 14px;
  font-weight: 600;
  color: var(--color-text);
  margin: 0;
}

.config-close-btn {
  width: 24px;
  height: 24px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: transparent;
  border: 1px solid var(--color-border);
  border-radius: 4px;
  color: var(--color-text-secondary);
  cursor: pointer;
  transition: all 0.15s;
}

.config-close-btn:hover {
  border-color: var(--color-primary);
  color: var(--color-primary);
}

.config-loading,
.config-empty,
.config-error {
  padding: 28px 16px;
  text-align: center;
  color: var(--color-text-secondary);
  font-size: 13px;
}

.config-error {
  color: #b91c1c;
}

.config-retry-btn {
  margin-top: 8px;
  padding: 4px 12px;
  background: var(--color-primary);
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 12px;
}

.config-body {
  flex: 1;
  display: flex;
  overflow: hidden;
}

/* 左侧厂商列表 */
.config-provider-list {
  width: 160px;
  border-right: 1px solid var(--color-border);
  overflow-y: auto;
  padding: 4px 0;
}

.config-provider-item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
  cursor: pointer;
  transition: background 0.15s;
  position: relative;
  border-left: 2px solid transparent;
}

.config-provider-item:hover {
  background: var(--color-bg-secondary);
}

.config-provider-item.active {
  background: var(--color-bg-secondary);
  border-left-color: var(--color-primary);
}

.config-provider-name {
  flex: 1;
  font-size: 12px;
  color: var(--color-text);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.config-provider-dot {
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: #10b981;
  flex-shrink: 0;
}

.config-provider-current {
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: var(--color-primary);
  flex-shrink: 0;
}

/* 右侧配置区 */
.config-provider-config {
  flex: 1;
  padding: 12px 14px;
  overflow-y: auto;
}

.config-section {
  margin-bottom: 12px;
}

.config-label {
  display: block;
  font-size: 12px;
  color: var(--color-text-secondary);
  margin-bottom: 4px;
}

.config-input,
.config-select {
  width: 100%;
  padding: 6px 8px;
  border: 1px solid var(--color-border);
  border-radius: 6px;
  font-size: 13px;
  color: var(--color-text);
  background: var(--color-bg);
  box-sizing: border-box;
}

.config-input:focus,
.config-select:focus {
  outline: none;
  border-color: var(--color-primary);
}

.config-input-narrow {
  width: 100px;
}

.config-apikey-row {
  display: flex;
  gap: 6px;
}

.config-apikey-row .config-input {
  flex: 1;
}

.config-toggle-btn {
  padding: 6px 10px;
  background: var(--color-bg-secondary);
  border: 1px solid var(--color-border);
  border-radius: 6px;
  font-size: 12px;
  color: var(--color-text-secondary);
  cursor: pointer;
  transition: all 0.15s;
}

.config-toggle-btn:hover {
  border-color: var(--color-primary);
  color: var(--color-primary);
}

.config-badge {
  display: inline-block;
  margin-left: 6px;
  padding: 1px 6px;
  background: #ecfdf5;
  color: #047857;
  border-radius: 4px;
  font-size: 11px;
}

.config-advanced-toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 6px 0;
  cursor: pointer;
  font-size: 12px;
  color: var(--color-text-secondary);
  user-select: none;
}

.config-advanced-toggle:hover {
  color: var(--color-primary);
}

.config-advanced {
  padding: 8px 0 4px 14px;
  border-left: 2px solid var(--color-border);
  margin-left: 4px;
}

.config-test-result,
.config-save-msg {
  padding: 6px 10px;
  border-radius: 6px;
  font-size: 12px;
  margin-bottom: 8px;
}

.config-test-result.success,
.config-save-msg.success {
  background: #ecfdf5;
  color: #047857;
  border: 1px solid #a7f3d0;
}

.config-test-result.error,
.config-save-msg.error {
  background: #fef2f2;
  color: #b91c1c;
  border: 1px solid #fecaca;
}

.config-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding-top: 8px;
  border-top: 1px solid var(--color-border);
}

.config-test-btn {
  padding: 6px 14px;
  background: var(--color-bg-secondary);
  border: 1px solid var(--color-border);
  border-radius: 6px;
  font-size: 12px;
  color: var(--color-text);
  cursor: pointer;
  transition: all 0.15s;
}

.config-test-btn:hover:not(:disabled) {
  border-color: var(--color-primary);
  color: var(--color-primary);
}

.config-test-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.config-save-btn {
  padding: 6px 14px;
  background: var(--color-primary);
  border: 1px solid var(--color-primary);
  border-radius: 6px;
  font-size: 12px;
  color: white;
  cursor: pointer;
  transition: all 0.15s;
}

.config-save-btn:hover:not(:disabled) {
  background: var(--color-primary-hover);
  border-color: var(--color-primary-hover);
}

.config-save-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/DeepExcel.UI/src/styles.css
git commit -m "style(ui): add model config panel styles (left-right split)"
```

---

### Task 8: 前端 App.tsx 接入模型配置弹窗

**Files:**
- Modify: `src\DeepExcel.UI\src\App.tsx`

**Interfaces:**
- Consumes: `<ModelConfigPanel>` 组件；监听 `open_model_config` 消息

- [ ] **Step 1: 在 App.tsx 导入 ModelConfigPanel**

找到 `App.tsx` L8 的 `import { ConversationsPanel }`，在下方添加：

```tsx
import { ModelConfigPanel } from './components/ModelConfigPanel'
```

- [ ] **Step 2: 新增 `modelConfigOpen` state**

找到 L32 附近的 `const [conversationsOpen, setConversationsOpen] = useState(false)`，在下方添加：

```tsx
  // ★ 模型配置弹窗（Ribbon 按钮触发）
  const [modelConfigOpen, setModelConfigOpen] = useState(false)
```

- [ ] **Step 3: 在 `onHostMessage` 监听器中处理 `open_model_config` 消息**

找到 L123 附近的 `} else if (data.type === 'connection_ok') {`，在它**之前**插入：

```tsx
      } else if (data.type === 'open_model_config') {
        // ★ Ribbon 按钮触发的打开模型配置弹窗
        setModelConfigOpen(true)
```

- [ ] **Step 4: 在 header 添加"模型配置"按钮（可选，Ribbon 已有入口）**

找到 L287 附近的"版本"按钮，在它**之后**、"附件"按钮**之前**添加一个"模型配置"按钮（让面板内也能打开）：

```tsx
          <button
            className="history-toggle-btn"
            onClick={() => setModelConfigOpen(true)}
            title="模型配置"
            type="button"
          >
            模型
          </button>
```

- [ ] **Step 5: 在组件树底部渲染 ModelConfigPanel**

找到 L337 附近的 `<ConversationsPanel ... />`，在它**之后**添加：

```tsx
      <ModelConfigPanel
        open={modelConfigOpen}
        onClose={() => setModelConfigOpen(false)}
      />
```

- [ ] **Step 6: 前端构建验证**

Run: `cd src\DeepExcel.UI && npm run build`
Expected: 构建成功，输出到 `src\DeepExcel.AddIn\WebViewAssets`

- [ ] **Step 7: Commit**

```bash
git add src/DeepExcel.UI/src/App.tsx src/DeepExcel.AddIn/WebViewAssets
git commit -m "feat(ui): wire ModelConfigPanel into App, listen for open_model_config"
```

---

### Task 9: 集成编译 + 注册 CLSID + Excel 验证

**Files:**
- 无新增/修改文件，纯构建验证

- [ ] **Step 1: 确认 Excel 已关闭**

Run: `tasklist /FI "IMAGENAME eq EXCEL.EXE"`
Expected: `INFO: No tasks are running which match the specified criteria.`
如果 Excel 在运行，先关闭它。

- [ ] **Step 2: 编译 C# DLL**

Run: `powershell -ExecutionPolicy Bypass -File scripts\_compile_only.ps1`
Expected: `=== BUILD SUCCESS ===`

- [ ] **Step 3: 注册 CLSID（必须，否则 Excel 无法加载加载项）**

Run: `powershell -ExecutionPolicy Bypass -File scripts\register-user.ps1`
Expected: `Registration successful!`

- [ ] **Step 4: 启动 Excel 验证**

启动 Excel，检查：
1. DeepExcel tab 下有"打开面板"、"模型配置"、"帮助"三个按钮
2. 点击"模型配置"按钮 → 弹出模型配置弹窗
3. 左侧显示 10 个厂商（含 Kimi/通义/智谱/Minimax/豆包）
4. 点击不同厂商 → 右侧表单更新
5. 输入 API Key → 点"测试连接" → 显示成功/失败
6. 点"保存并应用" → 显示"已保存并应用"
7. 关闭弹窗后，新对话使用新配置

- [ ] **Step 5: 验证旧 config.json 迁移**

检查 `%APPDATA%\DeepExcel\config.json`，应包含 10 个 provider 和 `maxTurns` 字段。

- [ ] **Step 6: Commit 最终版本（如有修复）**

```bash
git add -A
git commit -m "chore: integration verification complete"
```

---

## Self-Review Notes

**Spec coverage 检查**：
- ✅ 10 个内置厂商（Task 1）
- ✅ Ribbon 按钮（Task 3）
- ✅ 左右分栏 UI（Task 6, 7）
- ✅ 模型选择（Task 6）
- ✅ API Key 填写（Task 6）
- ✅ BaseUrl 高级设置（Task 6）
- ✅ MaxTurns 高级设置（Task 6）
- ✅ 保存后立即生效（Task 4 HandleSaveModelConfig 调 RefreshConfigForAllSessions）
- ✅ API Key 安全脱敏（Task 2 MaskApiKey）
- ✅ 测试连接按钮（Task 4 HandleTestApiKey）
- ✅ 真实品牌色图标无 AI 竖线（Task 5）
- ✅ 旧 config 迁移（Task 1 MigrateConfig）

**类型一致性检查**：
- `ProviderInfo` 在 Task 5 定义，Task 6 使用 ✓
- `ModelConfig` 在 Task 5 定义，Task 6 使用 ✓
- `SafeProvider` 在 Task 2 扩展字段（Type/BaseUrl/DefaultModel/SupportsVision/ApiKeyPreview），Task 4 `HandleGetModelConfig` 通过 `GetSafeConfig` 返回 ✓
- `MaskApiKey` 在 Task 2 定义，Task 2 `GetSafeConfig` 内调用 ✓
- `KEEP_PLACEHOLDER = '***keep***'` 在 Task 6 定义，Task 4 `HandleSaveModelConfig` 检查此值 ✓
- `providerOrder` 在 Task 5 定义，Task 6 使用 ✓
- `providerIcons` 在 Task 5 定义，Task 6 使用 ✓

**潜在风险**：
- Task 4 `HandleTestApiKey` 用 `.Result` 同步等待 HttpClient，可能阻塞 UI 线程。但 MessageBridge.HandleMessage 本身是同步的（WebMessageReceived 事件中调用），无法用 async/await。已设 15s 超时兜底。如果实测卡顿，可改为 `Task.Run(() => ...).Wait(15000)` 在后台线程执行。
- Task 3 `SendOpenModelConfigToActivePane` 用 `BeginInvoke` 跨线程访问 WebView，需确保 pane 未 Dispose（已检查 `pane.IsDisposed`）。
- Task 6 `apiKey` 初始值为 `KEEP_PLACEHOLDER`，输入框 `value={apiKey === KEEP_PLACEHOLDER ? '' : apiKey}`，这样用户看到空输入框但有 placeholder 提示已配置。保存时如果用户没输入，发 `KEEP_PLACEHOLDER` 给 C#，C# 跳过保存。这个逻辑在 Task 4 Step 3 已实现。
