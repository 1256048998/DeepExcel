# DeepExcel 对抗式审查报告（2026-07-07）

## 审查范围

- 日期：2026-07-07
- 视角：攻击者/对抗者
- 文件：7个核心文件
  - `MessageBridge.cs` (Bridge层)
  - `AppConfig.cs` (配置管理)
  - `ThisAddIn.cs` (Excel集成)
  - `App.tsx` (React主组件)
  - `ModelConfigPanel.tsx` (UI组件)
  - `types.ts` (类型定义)
  - `styles.css` (样式)

---

## 🔴 Critical 级别（3个）

### C-1: API Key 明文泄露 — config.json 中 ApiKey 字段未清空

**文件**: [AppConfig.cs:129-141](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs) + [MessageBridge.cs:400](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
- `ProviderConfig.ApiKey` 在 `CreateDefault()` 中初始化为空字符串，但 config.json 是可编辑的 JSON 文件
- 用户可通过外部编辑器手动修改 config.json，明文 API Key 会持久化在磁盘上
- `HandleSaveModelConfig` 中，如果 `SecurityManager.SaveApiKey` 抛出异常（DPAPI失败），config.json 中 `ApiKey` 字段为空，但内存中仍可能持有旧值，无回滚机制

**攻击场景**:
1. 用户升级旧版本，ApiKey 字段未被清空
2. 用户手动编辑 config.json 添加 ApiKey
3. DPAPI 加密失败时，明文密钥可能残留在内存中

**风险等级**: Critical  
**影响范围**: 所有用户的 API 密钥  
**修复建议**:
```csharp
public void UpdateApiKey(string providerKey, string apiKey)
{
    if (!_config.Providers.ContainsKey(providerKey))
        _config.Providers[providerKey] = new ProviderConfig();
    
    try
    {
        DeepExcel.AddIn.Security.SecurityManager.Instance.SaveApiKey(providerKey, apiKey);
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine("UpdateApiKey: SecurityManager.Save failed: " + ex.Message);
        // DPAPI 失败则拒绝保存，避免明文泄露
        return; // ← 修复：DPAPI 失败则拒绝保存
    }
    _config.Providers[providerKey].ApiKey = ""; // 确保不存明文
    Save();
}
```

---

### C-2: VBA 宏安全设置 — 强制启用 AccessVBOM + VBAWarnings = 严重降级用户安全

**文件**: [ThisAddIn.cs:140-189](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs)

**问题描述**:
- `EnsureVbaSecuritySettings()` 将 `AccessVBOM=1` 和 `VBAWarnings=1` 写入注册表，且不可逆
- 只要加载项启动就强制设置，用户无法关闭

**攻击场景**:
1. 恶意工作簿打开时，如果 DeepExcel 已安装，VBA 宏自动被信任执行，无需用户确认
2. `AccessVBOM=1` 允许程序化访问 VBA 工程对象模型，即使 AI 执行 VBA 是合法功能，也不应该全局禁用 Office 的宏安全策略
3. 其他加载项或恶意脚本也可以利用这个被放宽的安全策略

**风险等级**: Critical  
**影响范围**: 所有安装 DeepExcel 的用户，全局 Office 安全降级  
**修复建议**:
- 不应写 `VBAWarnings=1`（应保持用户原有设置）
- `AccessVBOM=1` 应在卸载时还原，或改为按需启用（仅在执行 VBA 前临时设置，执行后恢复）

```csharp
// 改为按需设置，而非启动时永久设置
private int _originalAccessVbom;
private int _originalVbaWarnings;

private void TempEnableVbaAccess()
{
    var key = Registry.CurrentUser.OpenSubKey(
        $@"Software\Microsoft\Office\{_excelApp.Version}\Excel\Security", true);
    _originalAccessVbom = (int)(key.GetValue("AccessVBOM") ?? 0);
    _originalVbaWarnings = (int)(key.GetValue("VBAWarnings") ?? 2);
    key.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
}

private void RestoreVbaSecurity()
{
    var key = Registry.CurrentUser.OpenSubKey(
        $@"Software\Microsoft\Office\{_excelApp.Version}\Excel\Security", true);
    key.SetValue("AccessVBOM", _originalAccessVbom, RegistryValueKind.DWord);
    key.SetValue("VBAWarnings", _originalVbaWarnings, RegistryValueKind.DWord);
}
```

---

### C-3: SSRF — HandleTestApiKey 允任意 URL 发起请求

**文件**: [MessageBridge.cs:459-566](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
- `HandleTestApiKey` 接受前端传入的 `baseUrl`，拼接 `/v1/messages` 后直接发起 HTTP 请求，没有任何 URL 白名单校验

**攻击场景**:
1. 用户（或恶意网页）构造 `baseUrl = "http://169.254.169.254/latest/meta-data"`，可探测 AWS 元数据服务
2. `baseUrl = "http://localhost:xxxx/admin"` 可探测内网服务
3. `baseUrl = "file:///etc/passwd"` 在某些 HttpClient 实现中可能触发文件读取
4. AI 生成的 baseUrl 也可能被诱导（LLM prompt injection → AI 调用工具 → 返回内网信息）

**风险等级**: Critical  
**影响范围**: 内网安全、SSRF  
**修复建议**:
```csharp
private bool IsValidTestUrl(string url)
{
    if (string.IsNullOrEmpty(url)) return false;
    Uri.TryCreate(url, UriKind.Absolute, out var uri);
    if (uri == null) return false;
    
    // 禁止非 HTTPS（除了 localhost 测试）
    if (uri.Scheme != "https" && !(uri.Host == "localhost" || uri.Host == "127.0.0.1"))
        return false;
    
    // 禁止内网地址
    var ip = System.Net.Dns.GetHostAddresses(uri.Host);
    foreach (var addr in ip)
    {
        if (addr.IsLoopback) return false;
        if (addr.IsPrivate()) return false; // .NET 5+ 有此方法
    }
    return true;
}
```

---

## 🟠 High 级别（5个）

### H-1: 错误信息泄露内部路径和堆栈

**文件**: [MessageBridge.cs:367](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) + [ThisAddIn.cs:444](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs)

**问题描述**:
```csharp
return MakeError($"Handle error: {ex.Message}");  // ← 暴露内部异常
Log("Stack: " + ex.StackTrace);  // ← 暴露堆栈到日志文件
```
`ex.Message` 可能包含文件路径、COM 对象名称、数据库连接字符串等敏感信息。攻击者可通过构造异常消息，观察返回的 error JSON 来推断内部架构。

**风险等级**: High  
**影响范围**: 信息泄露  
**修复建议**: 统一返回模糊错误消息，详细信息只写日志：
```csharp
return MakeError("内部处理错误，请重试");
// ex.Message 和 StackTrace 只写入 Logger
```

---

### H-2: 日志文件路径可预测，可被利用写入任意内容

**文件**: [ThisAddIn.cs:86-95](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs)

**问题描述**:
```csharp
string logPath = Path.Combine(
    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
    "DeepExcel_Load.log");
File.AppendAllText(logPath, "[" + DateTime.Now + "] " + message + Environment.NewLine);
```
- 日志路径是 DLL 所在目录（通常是 Program Files 或 GAC），`File.AppendAllText` 会创建文件
- 如果攻击者能在该目录写入符号链接（symlink），可将日志写入到任意位置
- 日志内容未做转义，可通过构造包含换行符的组件名注入假日志条目

**风险等级**: High  
**影响范围**: 日志伪造、潜在的文件写入  
**修复建议**: 使用 `%APPDATA%\DeepExcel\logs\` 目录，并对日志内容做基本转义

---

### H-3: 附件上传无大小限制 — 内存耗尽 DoS

**文件**: [MessageBridge.cs:716-733](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) + [App.tsx:189-210](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/App.tsx)

**问题描述**:
- 前端通过 `FileReader.readAsDataURL` 将文件转为 base64 字符串，C# 端直接接收 `file_base64` 字符串，没有任何文件大小检查

**攻击场景**:
1. 用户上传 100MB 文件 → base64 膨胀 33% → ~133MB 内存
2. 多个大文件同时上传 → 内存耗尽 → Excel 崩溃
3. AI 可能通过工具调用自动处理超大附件 → 不可预期的内存使用

**风险等级**: High  
**影响范围**: 内存耗尽、应用崩溃  
**修复建议**:
```csharp
// 前端
if (file.size > 10 * 1024 * 1024) { // 10MB
    alert('文件过大，请选择 10MB 以内的文件');
    return;
}

// C# 端
private const int MaxAttachmentBase64Length = 15 * 1024 * 1024; // ~11MB raw
if (fileBase64.Length > MaxAttachmentBase64Length)
    return MakeError("附件大小超过限制（10MB）");
```

---

### H-4: WorkbookSession 字典无并发保护 — 线程安全问题

**文件**: [MessageBridge.cs:33](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
```csharp
private readonly Dictionary<string, WorkbookSession> _sessions = new();
```
`_sessions` 在多个线程中被访问：
- UI 线程调用 `HandleMessage` → `GetOrCreateActiveSession()`
- Sidecar 回调线程调用 `OnStreamDelta` → `FindSessionBySidecar()`
- `RefreshConfigForAllSessions()` 遍历 `_sessions`
- `OnWorkbookClose` 删除 session

`Dictionary<K,V>` **不是线程安全的**，并发读写可能导致 `InvalidOperationException` 或数据损坏。

**风险等级**: High  
**影响范围**: 随机崩溃、数据损坏  
**修复建议**:
```csharp
// 方案1：使用 ConcurrentDictionary
private readonly ConcurrentDictionary<string, WorkbookSession> _sessions = new();

// 方案2：用 lock 保护关键路径
private readonly object _sessionLock = new object();
lock (_sessionLock) {
    _sessions.TryGetValue(key, out var session);
}
```

---

### H-5: 公式黑名单不完整 — 绕过可能性

**文件**: [MessageBridge.cs:1340-1362](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
黑名单包含 `WEBSERVICE`, `FILTERXML`, `CALL`, `EXEC` 等，但缺少以下危险函数：

1. **`INDIRECT`** — 动态引用任意单元格，可绕过地址校验
2. **`INFO`** — 泄露系统信息（操作系统版本、Excel 版本等）
3. **`GETPIVOTDATA`** — 可读取透视表数据
4. **`OBJECT`** — 创建 OLE 对象
5. **`RUN`** — 执行宏（变体）
6. **嵌套绕过** — `=IF(1=1,WEBSERVICE("url"),"")` 中 `upper.Contains("=WEBSERVICE(")` 会匹配，但 `=IF(1=1,SUBSTITUTE("W","W","E")&"BSERVICE","")` 可能绕过

**风险等级**: High  
**影响范围**: 数据外泄  
**修复建议**: 采用白名单方式，或更严格地用正则匹配函数调用模式

---

## 🟡 Medium 级别（8个）

### M-1: HandleSaveModelConfig — BaseUrl 未做 URL 格式校验

**文件**: [MessageBridge.cs:416-419](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
```csharp
if (!string.IsNullOrEmpty(baseUrl))
{
    cfg.Providers[provider].BaseUrl = baseUrl; // ← 任意字符串
}
```
用户可传入 `baseUrl = "javascript:alert(1)"` 或非 URL 字符串，虽然 WebView2 有安全边界，但后续 `HttpClient` 请求会使用此 URL，可能导致异常行为。

**修复建议**: 用 `Uri.TryCreate(baseUrl, UriKind.Absolute, out _)` 校验

---

### M-2: HandleTestApiKey — 阻塞 UI 线程

**文件**: [MessageBridge.cs:504-566](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
```csharp
var response = System.Threading.Tasks.Task.Run(async () => await client.SendAsync(request)).Result;
```
`.Result` 在 UI 线程上会阻塞，如果网络超时（15秒），UI 完全冻结。注释说"用 Task.Run 避免阻塞 UI 线程"，但 `.Result` 仍然阻塞调用线程。

**修复建议**: 改为真正的异步模式：
```csharp
private async Task<string> HandleTestApiKeyAsync(Message msg) { ... }
// 调用方也需改为 async
```

---

### M-3: EnforceSessionLimit — 所有 session 都 busy 时不回收

**文件**: [MessageBridge.cs:1094-1122](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
当所有 session 都处于 `IsBusy` 状态时，`EnforceSessionLimit` 找不到可回收的 session，新 session 无法创建，用户切换工作簿时会得到 `null`。

**修复建议**: 在极端情况下强制回收最老的 busy session（至少记录警告日志）

---

### M-4: 配置迁移不一致 — CreateDefault 与 MigrateConfig 模型列表不同

**文件**: [AppConfig.cs:30-33](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs) vs [AppConfig.cs:297-309](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs)

**问题描述**:
`CreateDefault()` 和 `LatestModelCatalog` 的模型列表不一致：
- `qwen`: CreateDefault `["qwen3.7-max", ...]` vs MigrateConfig `["qwen3-max", ...]`
- `zhipu`: CreateDefault `["glm-5.2", ...]` vs MigrateConfig `["glm-4.6", ...]`
- `minimax`: CreateDefault `["MiniMax-M2.5", ...]` vs MigrateConfig `["MiniMax-M2.1", ...]`

新安装用 `CreateDefault`，升级用 `MigrateConfig`，两者看到的模型列表不同。

**修复建议**: 统一为一个常量源

---

### M-5: 附件无类型校验 — 可上传可执行文件

**文件**: [MessageBridge.cs:716-733](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
`HandleUploadAttachment` 接受任意 `file_name`，不检查文件扩展名。用户可上传 `.exe`、`.bat`、`.ps1` 等可执行文件，如果 AI 被 prompt injection 诱导去"执行附件"，后果严重。

**修复建议**: 添加文件类型白名单（`.xlsx`, `.csv`, `.pdf`, `.txt`, `.png`, `.jpg` 等）

---

### M-6: App.tsx — useEffect 空依赖数组导致闭包捕获过时状态

**文件**: [App.tsx:57-131](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/App.tsx)

**问题描述**:
```tsx
useEffect(() => {
    const unsubscribe = onHostMessage((data) => {
        // 使用 loading, isClarifying 等状态...
    })
    return unsubscribe
}, []) // ← 空依赖数组
```
`onHostMessage` 的回调闭包捕获了初始化时的 `loading` 等状态，后续更新不会反映在回调中。

**修复建议**: 使用 `useRef` 或在依赖数组中添加相关状态

---

### M-7: ConfigManager 非线程安全单例

**文件**: [AppConfig.cs:176-177](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs)

**问题描述**:
```csharp
private static ConfigManager _instance;
public static ConfigManager Instance => _instance ??= new ConfigManager();
```
`??=` 不是原子操作，在多线程环境下可能创建多个实例。

**修复建议**:
```csharp
private static readonly Lazy<ConfigManager> _instance = new Lazy<ConfigManager>(() => new ConfigManager());
public static ConfigManager Instance => _instance.Value;
```

---

### M-8: maxTurns 上限校验不一致

**文件**: [MessageBridge.cs:432](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**问题描述**:
```csharp
if (maxTurns > 0 && maxTurns <= 200)
```
前端 `ModelConfigPanel.tsx:283` 设 `max={200}`，但用户可通过直接构造 JSON 绕过前端限制。如果 `maxTurns` 为 200，AI 可能进行 200 轮工具调用，消耗大量 API 费用。

**修复建议**: 根据实际业务需求设置合理上限（如 50），并在后端强制校验

---

## 🔵 Low 级别（4个）

### L-1: WebView2 `--disable-features=RendererCodeIntegrity` 降低渲染安全

**文件**: [ThisAddIn.cs:682](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs)

**问题描述**:
```csharp
AdditionalBrowserArguments = "--disable-features=RendererCodeIntegrity"
```
禁用渲染器代码完整性保护，降低了浏览器沙箱的安全性。WebView2 推荐使用此参数来避免某些兼容性问题，但应评估是否真正需要。

---

### L-2: CSS 中 `animation: fadeIn 0.2s ease-in` 无性能优化

**文件**: [styles.css:92-93](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/styles.css)

**问题描述**:
每次消息出现都触发 `fadeIn` 动画，大量消息时可能导致性能问题。不过在 Excel Add-in 的 WebView2 中影响有限。

---

### L-3: 前端 `onHostMessage` 未处理未知消息类型

**文件**: [App.tsx:58-131](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/App.tsx)

**问题描述**:
回调只处理 `stream_delta`, `stream_end`, `tool_call`, `tool_result`, `clarify`, `error`, `connection_ok`，其他消息类型被静默忽略。虽然后端已做了消息类型过滤，但前端应有兜底处理。

---

### L-4: 日志文件无限增长

**文件**: [ThisAddIn.cs:92](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs)

**问题描述**:
```csharp
File.AppendAllText(logPath, "[" + DateTime.Now + "] " + message + Environment.NewLine);
```
`DeepExcel_Load.log` 无限追加，长时间运行后可能增长到 GB 级别。

**修复建议**: 启动时检查文件大小，超过阈值时轮转或截断

---

## 优先级总结

| 优先级 | 编号 | 问题 | 风险 |
|--------|------|------|------|
| **P0** | C-1 | API Key 明文泄露 | 密钥泄露 |
| **P0** | C-2 | VBA 安全全局降级 | 恶意宏执行 |
| **P0** | C-3 | SSRF 任意 URL 请求 | 内网探测 |
| **P1** | H-4 | Session 字典线程不安全 | 随机崩溃 |
| **P1** | H-3 | 附件无大小限制 | 内存耗尽 |
| **P1** | H-1 | 错误信息泄露内部细节 | 信息泄露 |
| **P1** | H-2 | 日志路径可预测 | 日志伪造 |
| **P1** | H-5 | 公式黑名单不完整 | 数据外泄 |
| **P2** | M-1 | BaseUrl 未校验 | 异常请求 |
| **P2** | M-2 | 测试连接阻塞 UI | UI 冻结 |
| **P2** | M-3 | 所有 session busy 不回收 | 功能异常 |
| **P2** | M-4 | 模型列表不一致 | 用户困惑 |
| **P2** | M-5 | 附件无类型校验 | 潜在执行 |
| **P2** | M-6 | React 闭包过时状态 | 竞态条件 |
| **P2** | M-7 | 单例非线程安全 | 多实例 |
| **P2** | M-8 | maxTurns 上限过高 | 费用风险 |
| **P3** | L-1~L-4 | 低风险问题 | 边缘场景 |

---

## 修复建议优先级

| 优先级 | 问题 | 建议 |
|--------|------|------|
| **立即** | C-1 API Key泄露 | DPAPI失败时拒绝保存，config.json永不存明文 |
| **立即** | C-2 VBA安全 | 改为按需启用，执行后恢复原设置 |
| **立即** | C-3 SSRF | 添加URL白名单和内网IP过滤 |
| **本周** | H-4 Session字典 | 改用ConcurrentDictionary或加锁 |
| **本周** | H-3 附件大小 | 前后端同时限制10MB |
| **下周** | M-4 模型列表 | 统一为单一常量源 |

---

## 审查结论

**总计风险点**: 20个  
**建议修复时间**: Critical问题需在发布前修复，High问题需在一周内解决。

**关键发现**:
1. API Key 安全机制存在设计缺陷，DPAPI加密路径可能失败导致明文泄露
2. VBA安全设置过度放宽，影响所有Office用户的安全性
3. SSRF漏洞允许探测内网服务，虽攻击面有限但风险较高
4. 线程安全问题可能导致随机崩溃，影响用户体验

**下一步行动**:
1. 优先修复3个Critical问题（C-1, C-2, C-3）
2. 本周内解决5个High问题
3. 下周规划Medium问题的修复
4. Low问题可根据实际情况安排
