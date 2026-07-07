# 模型配置弹窗设计

**日期**: 2026-07-07
**状态**: 已批准，待实施

## 目标

在 Excel Ribbon 的"打开面板"和"帮助"按钮旁新增"模型配置"按钮，点击后弹出模型管理界面。用户可：
1. 选择内置厂商（10 个）
2. 选择该厂商支持的模型
3. 填写/修改 API Key
4. 修改 BaseUrl（高级设置，默认值已填好）
5. 调整 MaxTurns（Claude Agent SDK 循环上限）
6. 保存后立即对所有正在进行的对话生效

参考 Trae Worker 的左右分栏交互模式。

## 非目标

- 不支持新增自定义厂商（用户要求 10 个内置厂商即可）
- 不暴露 RequestTimeoutSeconds/MaxRetries 等全部高级参数（保持简洁）
- 不支持 API Key 明文回显（安全考虑，只返回 masked 预览）

## 架构方案

### 方案 A（已选）：Ribbon 按钮 → WebView2 消息 → 前端 React Modal

```
[Excel Ribbon: 模型配置按钮]
        │ onAction="OnShowModelConfig"
        ▼
[ThisAddIn.cs]
        │ PostWebMessageAsString({type:'open_model_config'})
        ▼
[当前可见 CTP 的 WebView2]
        │ chrome.webview 'message' 事件
        ▼
[App.tsx] 监听 open_model_config
        │ setModelConfigOpen(true)
        ▼
[ModelConfigPanel.tsx] 渲染左右分栏弹窗
        │ sendToHostWithResponse({type:'get_model_config'}, 'model_config')
        ▼
[MessageBridge.HandleMessage] 分发到 ConfigManager
        │ 返回 SafeConfig（API Key 脱敏）
        ▼
[ModelConfigPanel.tsx] 显示配置
        │ 用户编辑后点"保存并应用"
        │ sendToHostWithResponse({type:'save_model_config'}, 'config_saved')
        ▼
[MessageBridge] ConfigManager.SaveApiKey + SwitchProvider
        │ RefreshConfigForAllSessions() 立即生效
        ▼
返回 {success: true}
```

**选择理由**：
- 与现有 ConversationsPanel/HistoryPanel 模式完全一致
- UI 风格统一（React + WebView2）
- 复用现有 modal 基础设施（overlay/panel 样式）
- 符合用户偏好（CodeX 简洁风格，前端 React）

**备选方案（已否决）**：
- C# WinForm 模态对话框：与现有 UI 风格割裂，项目无先例
- 独立 CTP 任务窗格：过度工程化，占用 Excel 空间

## 厂商列表

### 现有 5 个 provider（保留）

| Key | DisplayName | BaseUrl | DefaultModel | Models | SupportsVision |
|---|---|---|---|---|---|
| `anthropic` | Claude (Anthropic) | `https://api.anthropic.com` | `claude-3-5-sonnet-20241022` | claude-3-5-sonnet-20241022, claude-3-5-haiku-20241022, claude-3-opus-20240229 | true |
| `deepseek` | DeepSeek | `https://api.deepseek.com/anthropic` | `deepseek-chat` | deepseek-chat, deepseek-coder, deepseek-reasoner | false |
| `stepfun` | 阶跃星辰 (Step) | `https://api.stepfun.com/step_plan` | `step-3.7-flash` | step-3.7-flash, step-3.5-flash | true |
| `openai` | OpenAI | `https://api.openai.com/v1` | `gpt-4o` | gpt-4o, gpt-4o-mini, gpt-4-turbo, gpt-3.5-turbo | true |
| `custom` | 自定义 (OpenAI兼容) | `` | `custom-model` | custom-model | false |

### 新增 5 个国产厂商

| Key | DisplayName | BaseUrl | DefaultModel | Models | SupportsVision |
|---|---|---|---|---|---|
| `kimi` | Kimi (月之暗面) | `https://api.moonshot.cn/anthropic` | `kimi-k2.7-code` | kimi-k2.7-code, kimi-k2.6, kimi-k2-thinking | true |
| `qwen` | 通义千问 (阿里) | `https://dashscope.aliyuncs.com/compatible-mode/anthropic` | `qwen3-max` | qwen3-max, qwen3-coder-plus, qwen-plus, qwen-turbo, qwen-long | true |
| `zhipu` | 智谱 (GLM) | `https://api.z.ai/api/anthropic` | `glm-4.6` | glm-4.6, glm-4.6-air, glm-4.5 | true |
| `minimax` | Minimax | `https://api.minimax.io/anthropic` | `MiniMax-M2.1` | MiniMax-M2.1, MiniMax-M2 | false |
| `doubao` | 豆包 (火山引擎) | `https://ark.cn-beijing.volces.com/api/compatible` | `doubao-seed-code` | doubao-seed-code, doubao-seed-1.6, doubao-seed-1.6-flash | true |

**注**：百川/零一万物因仅提供 OpenAI 兼容端点（无 Anthropic 端点），不能直接接入 Claude Agent SDK，不内置。

## UI 布局

### 左右分栏布局

```
┌─────────────────────────────────────────────────┐
│ 模型配置                                  [X]   │
├────────────────┬────────────────────────────────┤
│                │                                │
│ ● Claude       │  厂商: Claude (Anthropic)      │
│   DeepSeek     │  ──────────────────────────    │
│   阶跃星辰     │                                │
│   OpenAI       │  模型                          │
│   Kimi         │  ┌──────────────────────┐      │
│   通义千问     │  │ claude-3-5-sonnet  ▼ │      │
│   智谱         │  └──────────────────────┘      │
│   Minimax      │                                │
│   豆包         │  API Key                       │
│   自定义       │  ┌──────────────┐ ┌────┐       │
│                │  │ •••••••••••• │ │显示│       │
│                │  └──────────────┘ └────┘       │
│                │  ✓ 已配置                       │
│                │                                │
│                │  ── 高级设置 ──                 │
│                │                                │
│                │  Base URL                      │
│                │  ┌──────────────────────┐      │
│                │  │ https://api.anthropic│      │
│                │  └──────────────────────┘      │
│                │                                │
│                │  MaxTurns                      │
│                │  ┌────┐                        │
│                │  │ 20 │                        │
│                │  └────┘                        │
│                │                                │
│                │  [测试连接]      [保存并应用]   │
└────────────────┴────────────────────────────────┘
```

### 交互细节

1. **左侧厂商列表**：
   - 圆形品牌色图标 + 厂商名
   - 当前选中的 provider 高亮（背景色 + 左侧竖条）
   - 已配置 API Key 的厂商显示小绿点
   - 点击切换右侧配置区内容

2. **右侧配置区**：
   - 模型下拉：从 `provider.Models` 填充，默认选中 `provider.DefaultModel` 或 `cfg.CurrentModel`
   - API Key 输入框：`type="password"`，旁边有"显示"按钮切换明文
   - 如果已配置 key，显示"✓ 已配置"提示（不显示明文）
   - 高级设置默认折叠，点击展开
   - BaseUrl 输入框：默认填入 provider 的 BaseUrl，可修改
   - MaxTurns 数字输入框：默认 20

3. **底部按钮**：
   - "测试连接"：用当前填入的 API Key + BaseUrl + Model 发一个简单请求，显示成功/失败 + 延迟
   - "保存并应用"：保存配置并立即生效

4. **关闭**：点击右上角 X 或遮罩层关闭

### 厂商图标

使用首字母 + 品牌色圆形图标（避免 AI 风格竖线）：

| 厂商 | 首字母 | 品牌色 |
|---|---|---|
| Anthropic | A | `#D97757` |
| DeepSeek | D | `#4D6BFE` |
| 阶跃星辰 | S | `#7B61FF` |
| OpenAI | O | `#000000` |
| Kimi | K | `#000000` |
| 通义千问 | Q | `#615CED` |
| 智谱 | Z | `#3B82F6` |
| Minimax | M | `#2563EB` |
| 豆包 | D | `#FF1A4B` |
| 自定义 | + | `#6B7280` |

图标用纯 SVG 渲染（无外部图片依赖），样式为 28x28 圆形背景 + 白色首字母文字。

## 通信协议

### 新增消息类型（前端 → C#）

#### 1. `get_model_config` — 获取当前配置

**请求**：
```json
{"type": "get_model_config", "payload": {}}
```

**响应**：
```json
{
  "type": "model_config",
  "payload": {
    "currentProvider": "anthropic",
    "currentModel": "claude-3-5-sonnet-20241022",
    "maxTurns": 20,
    "providers": {
      "anthropic": {
        "displayName": "Claude (Anthropic)",
        "type": "anthropic",
        "baseUrl": "https://api.anthropic.com",
        "models": ["claude-3-5-sonnet-20241022", "claude-3-5-haiku-20241022", "claude-3-opus-20240229"],
        "defaultModel": "claude-3-5-sonnet-20241022",
        "supportsVision": true,
        "hasApiKey": true,
        "apiKeyPreview": "sk-a***...***key"
      },
      "deepseek": { ... },
      ...
    }
  }
}
```

**安全**：
- `hasApiKey`: boolean，是否已配置
- `apiKeyPreview`: 脱敏预览（如 `sk-a***...***key`），仅用于让用户确认是哪个 key
- **绝不返回明文 API Key**

#### 2. `save_model_config` — 保存配置

**请求**：
```json
{
  "type": "save_model_config",
  "payload": {
    "provider": "anthropic",
    "model": "claude-3-5-sonnet-20241022",
    "apiKey": "sk-new-key-here",
    "baseUrl": "https://api.anthropic.com",
    "maxTurns": 20
  }
}
```

**响应**：
```json
{
  "type": "config_saved",
  "payload": {"success": true}
}
```

**C# 处理逻辑**：
1. `SecurityManager.Instance.SaveApiKey(provider, apiKey)` — DPAPI 加密存储
2. `ConfigManager.Instance.Current.Providers[provider].BaseUrl = baseUrl` — 更新 BaseUrl
3. `ConfigManager.Instance.SwitchProvider(provider, model)` — 切换 provider + model
4. `ConfigManager.Instance.Current.General.MaxTurns = maxTurns`
5. `ConfigManager.Instance.Save()` — 持久化
6. `_bridge.RefreshConfigForAllSessions()` — 立即对所有 sidecar 生效

**特殊处理**：
- 如果 `apiKey` 为空字符串，不覆盖现有 key（避免误清空）
- 如果 `apiKey` 为 `"***keep***"` 占位符，跳过保存（用户未修改）
- 如果 `baseUrl` 与默认值相同，仍保存（允许用户覆盖后恢复默认）

#### 3. `test_api_key` — 测试 API 连接

**请求**：
```json
{
  "type": "test_api_key",
  "payload": {
    "provider": "anthropic",
    "apiKey": "sk-test-key",
    "baseUrl": "https://api.anthropic.com",
    "model": "claude-3-5-sonnet-20241022"
  }
}
```

**响应**：
```json
{
  "type": "api_test_result",
  "payload": {
    "success": true,
    "latencyMs": 342,
    "error": null
  }
}
```

**C# 处理逻辑**：
- 用 `HttpClient` 向 `baseUrl` 发一个最小请求（如 Anthropic Messages API 的 1 token 请求）
- 测量响应时间
- 不保存任何数据，纯测试

### 新增消息类型（C# → 前端）

#### `open_model_config` — Ribbon 按钮触发

```json
{"type": "open_model_config"}
```

前端 `App.tsx` 监听后 `setModelConfigOpen(true)`。

## 组件结构

### 前端新增文件

```
src/DeepExcel.UI/src/
├── components/
│   └── ModelConfigPanel.tsx       # 主弹窗组件（~300 行）
├── assets/
│   └── providerIcons.ts           # 厂商图标 SVG 映射
└── types.ts                       # 新增 ModelConfig 类型
```

### ModelConfigPanel.tsx 组件结构

```tsx
interface ProviderInfo {
  displayName: string
  type: string
  baseUrl: string
  models: string[]
  defaultModel: string
  supportsVision: boolean
  hasApiKey: boolean
  apiKeyPreview: string
}

interface ModelConfig {
  currentProvider: string
  currentModel: string
  maxTurns: number
  providers: Record<string, ProviderInfo>
}

interface Props {
  open: boolean
  onClose: () => void
}

function ModelConfigPanel({ open, onClose }: Props) {
  const [config, setConfig] = useState<ModelConfig | null>(null)
  const [selectedProvider, setSelectedProvider] = useState('')
  const [model, setModel] = useState('')
  const [apiKey, setApiKey] = useState('')
  const [baseUrl, setBaseUrl] = useState('')
  const [maxTurns, setMaxTurns] = useState(20)
  const [showAdvanced, setShowAdvanced] = useState(false)
  const [showApiKey, setShowApiKey] = useState(false)
  const [testing, setTesting] = useState(false)
  const [testResult, setTestResult] = useState<{success: boolean, latencyMs?: number, error?: string} | null>(null)
  const [saving, setSaving] = useState(false)

  // 加载配置
  useEffect(() => { if (open) loadConfig() }, [open])

  // 切换 provider 时更新表单
  useEffect(() => { if (config && selectedProvider) syncForm() }, [selectedProvider, config])

  async function loadConfig() { ... }
  function syncForm() { ... }
  async function handleSave() { ... }
  async function handleTest() { ... }

  if (!open) return null

  return (
    <div className="config-overlay" onClick={onClose}>
      <div className="config-panel" onClick={e => e.stopPropagation()}>
        <header>...</header>
        <div className="config-body">
          <aside className="provider-list">...</aside>
          <main className="provider-config">...</main>
        </div>
        <footer>...</footer>
      </div>
    </div>
  )
}
```

## C# 端改动

### 1. AppConfig.cs

- `CreateDefault()`: 新增 5 个国产 provider
- `MigrateConfig()`: 自动补全新 provider（与现有 stepfun 迁移逻辑同模式）

### 2. DeepExcelRibbon.xml

```xml
<button id="btnModelConfig"
        label="模型配置"
        onAction="OnShowModelConfig"
        imageMso="DataSources"
        size="large" />
```

放在 `btnTogglePanel` 和 `btnHelp` 之间。

### 3. IRibbonCallbacks.cs

```csharp
[DispId(4)]
void OnShowModelConfig(object control);
```

### 4. ThisAddIn.cs

```csharp
void IRibbonCallbacks.OnShowModelConfig(object control)
{
    // 找到当前活动窗口的 CTP，通过 WebView 发消息给前端
    var activeWindow = _excelApp.ActiveWindow;
    if (activeWindow == null) return;
    string key = activeWindow.Caption;
    if (_ctpsByWindow.TryGetValue(key, out var pane) && pane != null)
    {
        var control2 = pane.Content as TaskPaneControl;
        if (control2?.WebView != null)
        {
            control2.WebView.CoreWebView2.PostWebMessageAsString(
                JsonSerializer.Serialize(new { type = "open_model_config" }));
        }
    }
}
```

**边界情况**：如果面板未打开，先调用 `OnTogglePanel` 打开面板，再发消息（用 `Task.Delay(500).ContinueWith` 等待 WebView 初始化）。

### 5. MessageBridge.cs HandleMessage

在无 session 消息分支新增：

```csharp
case "get_model_config":
    return HandleGetModelConfig();
case "save_model_config":
    return HandleSaveModelConfig(msg);
case "test_api_key":
    return HandleTestApiKey(msg);
```

### 6. SafeConfig 扩展

现有 `SafeConfig`/`SafeProvider` 需要扩展以包含 `BaseUrl`/`DefaultModel`/`SupportsVision`/`Type`：

```csharp
public class SafeProvider
{
    public string DisplayName { get; set; }
    public string Type { get; set; }          // 新增
    public string BaseUrl { get; set; }       // 新增
    public string DefaultModel { get; set; }  // 新增
    public bool SupportsVision { get; set; }  // 新增
    public string[] Models { get; set; }
    public bool HasApiKey { get; set; }
    public string ApiKeyPreview { get; set; } // 新增（脱敏预览）
}
```

## 安全考虑

1. **API Key 不回传前端**：`get_model_config` 只返回 `hasApiKey` + `apiKeyPreview`（脱敏），绝不返回明文
2. **DPAPI 加密存储**：通过 `SecurityManager.SaveApiKey` 存储
3. **保存时占位符处理**：前端如果用户未修改 API Key 输入框，发送 `"***keep***"` 占位符，C# 跳过保存
4. **测试连接不持久化**：`test_api_key` 纯测试，不保存任何数据
5. **脱敏规则**：`apiKeyPreview = key.Substring(0, 4) + "***..." + key.Substring(key.Length - 3)`，长度不足 8 位时全部 `***`

## 错误处理

- **配置加载失败**：弹窗显示错误信息 + 重试按钮
- **保存失败**：显示错误信息，弹窗不关闭
- **测试连接失败**：显示错误原因（网络错误、401 未授权、模型不存在等）
- **Ribbon 按钮点击时面板未打开**：先打开面板，500ms 后发消息

## 测试要点

1. **厂商切换**：左侧点击不同厂商，右侧表单正确更新
2. **API Key 保存**：保存后重新打开弹窗，显示"已配置"
3. **API Key 脱敏**：预览不暴露完整 key
4. **BaseUrl 修改**：修改后保存，重新打开弹窗显示修改后的值
5. **MaxTurns 修改**：保存后 sidecar 收到新的 max_turns
6. **立即生效**：保存后正在进行的对话下一轮使用新配置
7. **测试连接**：正确 key 返回成功 + 延迟，错误 key 返回 401
8. **Ribbon 按钮入口**：面板未打开时点击按钮，自动打开面板 + 弹窗
9. **迁移**：旧 config.json 升级后自动补全 5 个新 provider

## 实施顺序

1. C# 后端：AppConfig 扩展 + MigrateConfig + MessageBridge 消息处理 + Ribbon 按钮
2. 前端：ModelConfigPanel + providerIcons + App.tsx 接入 + styles.css
3. 集成测试：编译 + 注册 CLSID + Excel 验证
