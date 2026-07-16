# DeepExcel SaaS 客户端改造

**日期：** 2026-07-15
**状态：** 设计已完成，待实现
**关联：** 本 spec 是 [2026-07-14-saas-server-mvp-design.md](2026-07-14-saas-server-mvp-design.md) 的客户端配套。服务端 spec 覆盖服务端 API；本 spec 覆盖 Excel 插件（C# + 前端 + sidecar）的改造。

## 背景与动机

服务端 MVP 已设计完成，提供四类接口：注册、登录、Key 下发、日志上报。客户端需要改造才能接入服务端，实现"用户登录 → 服务端下发 Key → sidecar 用 Key 启动"的流程。

**核心架构决策（服务端 spec 已确认）：**
- Agent loop（Claude Agent SDK）留在客户端本地——工具调用零延迟
- 服务端只做"门卫 + 日志"，不参与 agent loop 编排
- Phase 1：服务端直接下发真实 Key（存内存不落盘）
- Phase 2：短期 token + 服务端代理（客户端只改 base_url，API 不变）

**本 spec 范围：**
- 登录/注册 UI（前端 React 组件）
- C# 端 HTTP 客户端（调服务端 API）
- 会话状态管理（JWT + API Key 内存缓存）
- sidecar 启动流程改造（登录后拉 Key 再启动）
- 现有配置面板的 SaaS 模式适配
- 日志上报机制（异步、脱敏）

**不在本 spec 范围：**
- ❌ 服务端实现（已有 spec）
- ❌ 支付集成、管理后台
- ❌ Phase 2 的代理模式改造（API 不变，后续只改 base_url）
- ❌ JWT 刷新机制（7 天过期后重新登录）

## 架构

```
┌──────────────────────────────────────────────────────────────┐
│ Excel 插件（客户端）                                          │
│                                                               │
│  ┌──────────────┐    ┌──────────────────┐    ┌────────────┐  │
│  │ LoginPanel   │    │ ModelConfigPanel │    │ MessageList│  │
│  │ .tsx         │    │ .tsx (SaaS 适配) │    │ (现有不变) │  │
│  │ 邮箱密码登录 │    │ 隐藏 Key 输入   │    │            │  │
│  └──────┬───────┘    └────────┬─────────┘    └────────────┘  │
│         │ IPC                   │ IPC                          │
│         ▼                       ▼                              │
│  ┌──────────────────────────────────────────────────────────┐ │
│  │ MessageBridge.cs (switch 新增 case)                      │ │
│  │   login / register / logout / get_subscription          │ │
│  └──────┬──────────────────────────────┬────────────────────┘ │
│         │                              │                       │
│         ▼                              ▼                       │
│  ┌──────────────┐              ┌────────────────┐             │
│  │ AuthClient   │              │ SessionManager │             │
│  │ .cs (新增)   │              │ .cs (新增)     │             │
│  │ HTTP 调服务端│              │ 内存缓存 JWT+Key│             │
│  └──────┬───────┘              └────────┬───────┘             │
│         │                               │                      │
│         │ ①登录 ②拉Key                  │ ③提供 Key            │
│         ▼                               ▼                      │
└─────────┼───────────────────────────────┼─────────────────────┘
          │                               │
          ▼                               ▼
┌─────────────────────────┐   ┌─────────────────────────────┐
│ DeepExcel Server        │   │ sidecar.py (几乎不变)       │
│ (已有 spec)             │   │  env_config["ANTHROPIC_API_ │
│  /api/auth/login        │   │   KEY"] = session.Key       │
│  /api/key               │   │  ClaudeSDKClient(options)   │
│  /api/subscription      │   │  Agent loop 留本地          │
└─────────────────────────┘   └─────────────────────────────┘
```

**数据流（SaaS 模式启动）：**
1. Excel 启动 → 插件加载 → 检查 JWT 是否有效（SecurityManager DPAPI 解密读取）
2. JWT 有效 → 调 `/api/key` 拉取 Claude API Key → 存内存（SessionManager）
3. JWT 无效或不存在 → 前端显示 LoginPanel → 用户登录 → 拿 JWT → DPAPI 加密存储 → 调 `/api/key`
4. Key 到手 → Start sidecar → SendConfig（注入 Key）→ sidecar 启动 Claude SDK
5. 用户使用过程中 → 异步批量上报日志到 `/api/logs`

**数据流（本地模式，向后兼容）：**
- 用户未登录且服务端未配置 → 回退到现有本地 Key 模式
- 用户手动填 API Key → DPAPI 存储 → 直接启动 sidecar（现有流程不变）

## 模式切换：本地模式 vs SaaS 模式

**设计原则：** SaaS 化不破坏现有本地模式，两种模式共存，通过配置切换。

### 模式判定

`AppConfig` 新增字段：
```csharp
public class AppConfig
{
    // 现有字段...
    public string AuthMode { get; set; } = "local";  // "local" | "saas"
    public string ServerUrl { get; set; } = "";       // SaaS 模式：服务端地址，如 "https://api.deepexcel.com"
}
```

**判定逻辑：**
- `AuthMode == "saas"` 且 `ServerUrl` 非空 → SaaS 模式
- 其他情况 → 本地模式（现有行为不变）

**用户如何切换：**
- 发布版默认 `AuthMode = "local"`（不破坏现有用户）
- SaaS 版本预配置 `AuthMode = "saas"` + `ServerUrl`
- 后续可在配置面板加切换开关（不在 MVP 范围）

### 两种模式的行为差异

| 行为 | 本地模式（现有） | SaaS 模式（新增） |
|------|----------------|------------------|
| Key 来源 | 用户手动填，DPAPI 落盘 | 服务端下发，内存缓存 |
| 启动流程 | 直接 Start sidecar | 先登录拉 Key，再 Start sidecar |
| 配置面板 | 显示 API Key 输入框 | 隐藏 Key 输入，显示订阅状态 |
| 日志 | 本地文件 | 本地文件 + 异步上报服务端 |
| 未配置时 | 提示填 Key | 显示登录面板 |

## 模块设计

### 1. AuthClient（新增）

**文件：** `src/DeepExcel.AddIn/Auth/AuthClient.cs`

**职责：** HTTP 客户端，封装服务端 API 调用。纯 HTTP，不维护状态。

```csharp
namespace DeepExcel.AddIn.Auth
{
    public class AuthClient
    {
        private readonly string _serverUrl;
        private readonly HttpClient _http;

        public AuthClient(string serverUrl)
        {
            _serverUrl = serverUrl.TrimEnd('/');
            _http = new HttpClient { BaseAddress = new Uri(_serverUrl), Timeout = TimeSpan.FromSeconds(15) };
        }

        // POST /api/auth/register
        public async Task<AuthResponse> RegisterAsync(string email, string password);

        // POST /api/auth/login
        public async Task<AuthResponse> LoginAsync(string email, string password);

        // GET /api/key (需 JWT)
        public async Task<KeyResponse> FetchKeyAsync(string jwt);

        // GET /api/subscription (需 JWT)
        public async Task<SubscriptionResponse> GetSubscriptionAsync(string jwt);

        // POST /api/logs (需 JWT, 批量)
        public async Task<bool> UploadLogsAsync(string jwt, List<LogEntry> logs);
    }

    public class AuthResponse { public string user_id; public string email; public string token; public DateTime expires_at; }
    public class KeyResponse { public string api_key; public string provider; public string base_url; }
    public class SubscriptionResponse { public string tier; public string status; public DateTime expires_at; public int days_remaining; }
    public class LogEntry { public DateTime timestamp; public string level; public string event_name; public string message; public string context_json; }
}
```

**错误处理：**
- 网络异常 → 返回 null + 日志记录，不崩溃 Excel
- 401（JWT 过期）→ 清除 SessionManager 的 JWT，触发前端重新登录
- 403（订阅过期）→ 前端显示订阅过期提示
- 超时（15s）→ 返回 null，提示网络问题

### 2. SessionManager（新增）

**文件：** `src/DeepExcel.AddIn/Auth/SessionManager.cs`

**职责：** 内存缓存 JWT + API Key，管理登录状态。sidecar 退出即销毁 Key。

```csharp
namespace DeepExcel.AddIn.Auth
{
    public class SessionManager
    {
        private string _jwt;          // 内存，不落盘
        private string _apiKey;       // 内存，不落盘
        private string _baseUrl;      // 服务端下发的 base_url
        private DateTime _jwtExpiry;
        private SubscriptionResponse _subscription;

        public bool IsLoggedIn => !string.IsNullOrEmpty(_jwt) && _jwtExpiry > DateTime.UtcNow;
        public bool HasKey => !string.IsNullOrEmpty(_apiKey);
        public string ApiKey => _apiKey;
        public string BaseUrl => _baseUrl;
        public SubscriptionResponse Subscription => _subscription;

        // 登录成功后调用
        public void SetAuth(string jwt, DateTime expiresAt);

        // 拉取 Key 成功后调用
        public void SetKey(string apiKey, string baseUrl);

        // 持久化 JWT（DPAPI 加密），用于下次启动免登录
        public void PersistJwt();
        public bool TryRestoreJwt();  // 从 DPAPI 读取，返回是否有效

        // 清除所有内存状态（登出）
        public void Clear();
    }
}
```

**JWT 持久化策略：**
- JWT 用 SecurityManager 的 DPAPI 机制加密存储到 `%APPDATA%\DeepExcel\credentials\jwt.crypt`
- API Key **不持久化**——每次启动重新拉取
- 下次启动：TryRestoreJwt() → 有效则直接 FetchKeyAsync，无效则前端显示登录

### 3. 登录流程集成到 MessageBridge

**文件：** `src/DeepExcel.AddIn/Bridge/MessageBridge.cs`

**改造点：** 在 `HandleMessage` 的 switch 中新增 case，调整 sidecar 启动顺序。

**新增消息类型：**

| 前端 → C# 消息 | C# 处理 | C# → 前端 响应 |
|---------------|---------|---------------|
| `{type:"login", email, password}` | `AuthClient.LoginAsync` → `SessionManager.SetAuth` → `FetchKeyAsync` → `SessionManager.SetKey` | `{type:"login_result", success, error?}` |
| `{type:"register", email, password}` | `AuthClient.RegisterAsync` → 同 login 后续 | `{type:"register_result", success, error?}` |
| `{type:"logout"}` | `SessionManager.Clear` + 清除 DPAPI JWT | `{type:"logout_result", success}` |
| `{type:"get_auth_status"}` | 检查 `SessionManager.IsLoggedIn` | `{type:"auth_status", loggedIn, subscription?}` |
| `{type:"get_subscription"}` | `AuthClient.GetSubscriptionAsync` | `{type:"subscription_info", tier, status, days_remaining}` |

**sidecar 启动顺序改造（关键）：**

当前流程（本地模式）：
```
用户打开面板 → Start sidecar → SendConfig(本地 Key) → 就绪
```

SaaS 模式新流程：
```
用户打开面板 → 检查登录状态
  ├─ 已登录且有 Key → Start sidecar → SendConfig(内存 Key) → 就绪
  ├─ 已登录但无 Key → FetchKeyAsync → Start sidecar → SendConfig → 就绪
  └─ 未登录 → 前端显示 LoginPanel → 登录成功后 → FetchKey → Start sidecar → SendConfig → 就绪
```

**改造位置：** `MessageBridge.cs` 第 134-135 行（`Start` + `SendConfigToSession`）需要条件化：
- SaaS 模式：先检查 `SessionManager.HasKey`，无 Key 则不立即启动，等登录流程完成
- 本地模式：保持现有行为

### 4. Key 注入点改造

**文件：** `src/DeepExcel.AddIn/Bridge/MessageBridge.cs` 第 213 行

当前：
```csharp
var apiKey = SecurityManager.Instance.GetApiKey(providerKey);
```

改为：
```csharp
string apiKey;
if (IsSaasMode)
{
    apiKey = _sessionManager.ApiKey;  // 内存中服务端下发的 Key
    if (string.IsNullOrEmpty(apiKey))
    {
        Log("SaaS mode: no API key in memory, sidecar will fail to start");
        return;
    }
}
else
{
    apiKey = SecurityManager.Instance.GetApiKey(providerKey);  // 本地模式：现有逻辑
}
```

### 5. 前端 LoginPanel（新增）

**文件：** `src/DeepExcel.UI/src/components/LoginPanel.tsx`

**职责：** 登录/注册 UI，仿 ModelConfigPanel 的弹窗模式。

**UI 结构（遵循用户偏好：CodeX 简洁风格）：**
```
┌─────────────────────────────────────┐
│  DeepExcel 登录                [X]  │
├─────────────────────────────────────┤
│                                     │
│  [登录] [注册]    ← Tab 切换        │
│                                     │
│  邮箱                               │
│  ┌─────────────────────────────┐    │
│  │ user@example.com            │    │
│  └─────────────────────────────┘    │
│                                     │
│  密码                               │
│  ┌─────────────────────────────┐    │
│  │ ••••••••                    │    │
│  └─────────────────────────────┘    │
│                                     │
│  ┌─────────────────────────────┐    │
│  │         登录                │    │
│  └─────────────────────────────┘    │
│                                     │
│  状态消息（错误/成功）              │
│                                     │
└─────────────────────────────────────┘
```

**组件接口：**
```tsx
interface LoginPanelProps {
  open: boolean;
  onClose: () => void;
  onLoginSuccess: () => void;  // 登录成功后回调（关闭面板、刷新状态）
}
```

**交互流程：**
1. 用户输入邮箱密码 → 点击登录
2. 发送 `{type:"login", email, password}` → 等待 `login_result`
3. 成功 → `onLoginSuccess()` → 关闭 LoginPanel → 触发 sidecar 启动
4. 失败 → 显示错误消息（"邮箱或密码错误" / "网络错误"）
5. 注册 Tab 同理，发送 `{type:"register", ...}`

**样式要求（用户偏好）：**
- 深灰 + 靛蓝配色（暗色主题）/ 天蓝色（亮色主题）
- 6-10px 圆角
- SVG 图标（非 emoji）
- 紧凑间距

### 6. ModelConfigPanel SaaS 适配

**文件：** `src/DeepExcel.UI/src/components/ModelConfigPanel.tsx`

**改造点：** SaaS 模式下隐藏 API Key 输入框，显示订阅状态。

```tsx
// 在组件中根据 authMode 条件渲染
{authMode === 'saas' ? (
  // SaaS 模式：显示订阅状态
  <div className="subscription-status">
    <span>订阅状态：{subscription.tier} ({subscription.status})</span>
    <span>剩余天数：{subscription.days_remaining}</span>
    <button onClick={handleLogout}>退出登录</button>
  </div>
) : (
  // 本地模式：现有 API Key 输入框（不变）
  <div className="api-key-input">
    <input type="password" ... />
  </div>
)}
```

**新增 props：**
```tsx
interface ModelConfigPanelProps {
  // 现有 props...
  authMode?: 'local' | 'saas';           // 默认 'local'
  subscription?: SubscriptionInfo;       // SaaS 模式的订阅状态
  onLogout?: () => void;                 // SaaS 模式的退出登录回调
}
```

### 7. 日志上报（异步 + 脱敏）

**文件：** `src/DeepExcel.AddIn/Auth/LogUploader.cs`（新增）

**职责：** 异步批量上报客户端日志到服务端，不阻塞 AI 调用。

```csharp
namespace DeepExcel.AddIn.Auth
{
    public class LogUploader
    {
        private readonly AuthClient _client;
        private readonly SessionManager _session;
        private readonly ConcurrentQueue<LogEntry> _queue = new();
        private readonly Timer _flushTimer;

        public LogUploader(AuthClient client, SessionManager session)
        {
            _client = client;
            _session = session;
            // 每 30 秒批量上报一次
            _flushTimer = new Timer(Flush, null, TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(30));
        }

        public void Enqueue(string level, string eventName, string message, object context = null)
        {
            // 脱敏：移除 API Key、JWT、密码
            var entry = new LogEntry
            {
                timestamp = DateTime.UtcNow,
                level = level,
                event_name = eventName,
                message = Sanitize(message),
                context_json = SanitizeJson(context)
            };
            _queue.Enqueue(entry);
        }

        private async void Flush(object state)
        {
            if (!_session.IsLoggedIn || _queue.IsEmpty) return;
            var batch = new List<LogEntry>();
            while (_queue.TryDequeue(out var entry)) batch.Add(entry);
            if (batch.Count == 0) return;
            try { await _client.UploadLogsAsync(_session.Jwt, batch); }
            catch { /* 静默失败，不影响用户 */ }
        }

        private string Sanitize(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            return text
                .Replace(_session.ApiKey ?? "", "[REDACTED_KEY]")
                .Replace(_session.Jwt ?? "", "[REDACTED_JWT]");
        }
    }
}
```

**上报的事件类型：**
- `sidecar_start` — sidecar 启动（含 model、session_id）
- `tool_call` — 工具调用（含 tool name、耗时，不含参数和结果）
- `error` — 客户端错误（含异常类型、消息，不含堆栈敏感路径）
- `login` / `logout` — 登录/登出事件

**不上报的内容：**
- Excel 单元格数据
- 用户对话内容
- API Key、JWT、密码
- 文件完整路径（只保留文件名）

## 启动流程详解

### 本地模式（现有，不变）

```
Excel 启动
  → 插件加载 (OnConnection)
  → 用户点击"打开面板"
  → CreateTaskPane → WebView2 加载前端
  → 前端发 get_model_config → C# 返回配置
  → 用户发消息 → Start sidecar → SendConfig(本地 Key)
  → sidecar 启动 Claude SDK → 就绪
```

### SaaS 模式（新增）

```
Excel 启动
  → 插件加载 (OnConnection)
  → SessionManager.TryRestoreJwt()  ← 尝试从 DPAPI 恢复 JWT
  → 用户点击"打开面板"
  → CreateTaskPane → WebView2 加载前端
  → 前端发 get_auth_status → C# 检查 SessionManager
    ├─ IsLoggedIn && HasKey → 直接 Start sidecar → SendConfig → 就绪
    ├─ IsLoggedIn && !HasKey → FetchKeyAsync → Start sidecar → SendConfig → 就绪
    └─ !IsLoggedIn → 前端显示 LoginPanel
       → 用户登录 → AuthClient.LoginAsync → SetAuth
       → FetchKeyAsync → SetKey
       → PersistJwt() (DPAPI 加密存)
       → Start sidecar → SendConfig → 就绪
```

**关键时序约束：**
- sidecar 启动必须在拿到 Key 之后（否则 Claude SDK 无 Key 报错）
- JWT 持久化必须在 Key 拉取成功之后（避免存了 JWT 但 Key 拉不到，下次启动误以为已登录）
- 前端 LoginPanel 是异步的，不阻塞 Excel 主线程

## 文件清单

### 新增文件

| 文件 | 职责 |
|------|------|
| `src/DeepExcel.AddIn/Auth/AuthClient.cs` | HTTP 客户端，调服务端 API |
| `src/DeepExcel.AddIn/Auth/SessionManager.cs` | 内存缓存 JWT + Key |
| `src/DeepExcel.AddIn/Auth/LogUploader.cs` | 异步日志上报 |
| `src/DeepExcel.UI/src/components/LoginPanel.tsx` | 登录/注册 UI |
| `src/DeepExcel.UI/src/components/LoginPanel.css` | 登录面板样式 |

### 修改文件

| 文件 | 改造内容 |
|------|---------|
| `src/DeepExcel.AddIn/Bridge/MessageBridge.cs` | switch 新增 login/register/logout case；Key 来源条件化；启动顺序调整 |
| `src/DeepExcel.AddIn/Config/AppConfig.cs` | 新增 `AuthMode` + `ServerUrl` 字段 |
| `src/DeepExcel.AddIn/Security/SecurityManager.cs` | 新增 `SaveJwt`/`GetJwt` 方法（或复用 `SaveApiKey("jwt", ...)`） |
| `src/DeepExcel.AddIn/ThisAddIn.cs` | 初始化 AuthClient + SessionManager + LogUploader；启动流程集成 |
| `src/DeepExcel.UI/src/components/ModelConfigPanel.tsx` | SaaS 模式下隐藏 Key 输入，显示订阅状态 |
| `src/DeepExcel.UI/src/App.tsx` | 集成 LoginPanel，根据 auth_status 切换显示 |
| `src/DeepExcel.UI/src/types.ts` | 新增 AuthStatus、SubscriptionInfo 类型 |

### 不需要改动

| 文件 | 原因 |
|------|------|
| `src/DeepExcel.Sidecar/sidecar.py` | 已接受任意 api_key 注入，Key 来源变更对它透明 |
| `src/DeepExcel.UI/src/bridge.ts` | 通信机制完全可复用 |
| `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs` | 进程启动逻辑不变 |

## 安全考量

- **API Key 不落盘**：SessionManager 的 `_apiKey` 只存内存，Excel 退出即销毁
- **JWT DPAPI 加密**：JWT 持久化用 DPAPI（CurrentUser 范围），其他用户无法解密
- **HTTPS 强制**：AuthClient 校验 `ServerUrl` 必须是 `https://`（开发环境可配 `http://localhost`）
- **日志脱敏**：LogUploader 上报前移除 Key、JWT、密码
- **内存中的 Key 不可被其他进程直接读取**：DPAPI 保护的是落盘数据，内存中的 Key 仍可被 Process Hacker 等工具提取（Phase 1 风险可接受，Phase 2 用代理消除）
- **不记录完整请求体**：日志只记事件类型和元数据，不记完整请求/响应

## 向后兼容性

- `AuthMode = "local"`（默认）→ 完全保持现有行为，不引入任何 SaaS 逻辑
- 现有用户的 config.json 无 `AuthMode` 字段 → 反序列化时默认 `"local"`（AppConfig 的默认值）
- SaaS 模式失败（服务端不可达）→ 不回退到本地模式（避免 Key 泄露），提示用户检查网络
- 现有 `ModelConfigPanel` 的 API Key 输入在本地模式下完全不变

## 测试策略

1. **单元测试（C#）：**
   - AuthClient 的 HTTP 调用（mock HttpClient）
   - SessionManager 的状态转换（登录→拉Key→登出）
   - LogUploader 的脱敏逻辑
2. **前端测试：**
   - LoginPanel 的表单验证和错误显示
   - ModelConfigPanel 在 SaaS 模式下的条件渲染
3. **集成测试：**
   - 本地模式回归：确保现有流程不受影响
   - SaaS 模式端到端：启动服务端 → 客户端登录 → 拉 Key → sidecar 启动 → 对话
4. **安全测试：**
   - API Key 不出现在 config.json / 日志 / 注册表
   - JWT DPAPI 加密验证（其他用户无法解密）
   - 日志脱敏验证（Key/JWT/密码被移除）

## 实现顺序建议

1. **C# 基础设施**：AppConfig 字段 → AuthClient → SessionManager → LogUploader
2. **MessageBridge 集成**：switch case → Key 来源条件化 → 启动顺序调整
3. **前端 LoginPanel**：组件 → 样式 → IPC 联调
4. **ModelConfigPanel 适配**：SaaS 模式条件渲染
5. **ThisAddIn 装配**：初始化 + 启动流程集成
6. **端到端测试**：启动服务端 → 客户端全流程验证

## 后续迭代方向（不在本 spec）

1. **JWT 刷新**：refresh token 机制，避免 7 天后重新登录
2. **Phase 2 代理模式**：客户端 base_url 指向服务端代理，Key 不再下发
3. **用量统计**：客户端上报 token 用量，服务端展示
4. **多设备登录**：JWT 支持多设备同时在线
5. **SSO/OAuth**：第三方登录（Google/Microsoft/微信）
6. **离线模式**：无网络时允许使用本地缓存的 Key（有限次数）
