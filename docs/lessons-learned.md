# DeepExcel Bug 修复经验总结（2026-06-26 ~ 2026-06-27）

> 本文档汇总 Claude Agent SDK 迁移完成后两天内调试 DeepExcel Excel 加载项所踩的坑、根因分析与防御性经验，供后续开发与多模型扩展时参考。

---

## 一、Bug 修复时间线（共 19 个问题）

按修复顺序排列，每行标注影响层与修复要点。

| # | 现象 | 根因 | 修复 |
|---|------|------|------|
| 1 | 加载项完全不显示 | DLL 重编译后 CLSID 注册项丢失（ProgId 在但 CLSID 项缺失） | 运行 `scripts\register-user.ps1` 注册 `HKCU\Software\Classes\CLSID\{A1B2C3D4-...}` 与 `{B2C3D4E5-...}` |
| 2 | 加载项被自动禁用 | Excel 因前次加载失败把 `LoadBehavior` 改成 2 | 重置 `HKCU\...\Addins\DeepExcel.AddIn\LoadBehavior=3` |
| 3 | sidecar 子进程冲突 | 残留孤儿 Python 进程占用资源 | `_kill_orphans.ps1` 清理 |
| 4 | 面板输入无响应 | C# 启动 sidecar 后未发送 config | `MessageBridge.SetSendToUi` 中调用 `SendConfigToSidecar()` |
| 5 | API 调用 404 | DeepSeek 用 OpenAI 端点 `/v1`，但 SDK 是 Anthropic 协议 | 改为 `https://api.deepseek.com/anthropic` |
| 6 | sidecar hang 住 | 启动顺序错：client 已 connect 但 config 未到 | sidecar 用 `anyio.fail_after(30)` 等 config 后再 `ClaudeSDKClient` |
| 7 | 中文/emoji 崩溃 | Windows 默认 cp1252 编码写中文抛 UnicodeEncodeError | sidecar 入口 `sys.stdout.reconfigure(encoding='utf-8')` |
| 8 | 第二轮 API 响应崩溃 | DeepSeek 不返回 thinking signature，SDK 硬编码 `block["signature"]` 抛 MessageParseError | `ClaudeAgentOptions(thinking={"type": "disabled"})` |
| 9 | BuildContext 崩溃 | `RangeInfo.Values` 是 `Object[,]`，System.Text.Json 不支持 2D 数组序列化 | BuildContext 只放元信息（address/sheet/rowCount），数据由 sidecar 按需读取 |
| 10 | read_range 数据崩溃 | COM 二维数组是 1-based，用 0-based 索引访问 IndexOutOfRangeException | `JsonConverters.cs` 用 `GetLowerBound/GetUpperBound` 遍历 |
| 11 | 模型无限循环 read_range | 序列化失败 fallback 设 `success=true`，模型看到成功但 data 是 error，困惑重试 | fallback 改为 `success=false` + 明确 error 文本 |
| 12 | 多 Excel 实例打不开 | WebView2 固定 `userDataFolder` 被锁 | 路径加 PID 后缀 `WebView2_{PID}` |
| 13 | BeginInvoke 崩溃 Excel | `new Form()` 不自动创建 HWND，`IsHandleCreated=false` 时 BeginInvoke 抛异常 | `SafeBeginInvoke` 包装，访问 `.Handle` 强制创建 |
| 14 | **前端收不到任何回复** | **`bridge.ts` 用 `window.addEventListener('message')`，但 C# 用 `PostWebMessageAsString`** | **改为 `chrome.webview.addEventListener('message')`** |
| 15 | UI 显示 `❌ Sidecar:` | `OnStderrLine` 把诊断日志当 error 转发给前端 | stderr 只写日志文件，不转发 |
| 16 | 工具调用重复显示 | `OnToolUse` 和 `OnToolCall` 都发 `tool_call` 给 UI | `OnToolCall` 只记日志，`OnToolUse` 去掉 `mcp__excel__` 前缀 |
| 17 | 创建图表崩溃 | `Shapes.AddChart2` 在某些 Excel 版本抛 HRESULT | try-catch 回退到 `Shapes.AddChart` |
| 18 | 无法重新编译 | DLL 被 Excel 进程锁定 | 编译前 `Stop-Process EXCEL -Force` |
| 19 | git push 超时 | GitHub 连接慢 | `git config --global http.proxy http://127.0.0.1:7890` |

---

## 二、最关键的成功修复：WebView2 消息监听器（#14）

这是整个调试过程中耗时最长、最反直觉的一个坑，**也是最后一个被发现的根因**。在此之前的所有后端修复（buffer patch、1-based 数组、SendToolResult fallback、AddChart2 回退）都已生效，后端日志显示 `stream_delta` / `stream_end` 已通过 `PostWebMessageAsString` 发出，但前端始终没有任何响应。

### 根因

WebView2 提供两种 C# → JS 的消息发送 API，对应**不同的前端监听器**：

| C# API | 前端监听器 | 数据格式 |
|--------|-----------|---------|
| `PostWebMessageAsString(json)` | `chrome.webview.addEventListener('message', ...)` | `event.data` 是原始字符串，需手动 `JSON.parse` |
| `PostWebMessageAsJson(obj)` | `window.addEventListener('message', ...)` | `event.data` 已自动反序列化为对象 |

`bridge.ts` 错用了 `window.addEventListener`，导致**所有** C# 发来的消息（包括 `stream_delta`、`stream_end`、`tool_call`、`clarify`）都收不到，但前端 → C# 方向（`chrome.webview.postMessage`）是正常的，所以用户输入能进入 sidecar，工具也能执行，**只是结果无法回到 UI**。

### 教训

1. **后端日志显示消息已发送 ≠ 前端能收到**。必须验证全链路最后一公里。
2. WebView2 的两套消息 API 极易混淆，**首次集成时必须在浏览器 DevTools 中验证 `chrome.webview` 对象上确实有事件触发**。
3. 当"后端工作正常但 UI 没反应"时，**第一时间排查 IPC 监听器匹配**，而不是继续调试后端。

### 正确实现（[bridge.ts](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/bridge.ts#L29-L56)）

```typescript
const webview = (window as any).chrome.webview
webview.addEventListener('message', (e: MessageEvent) => {
  try {
    const raw = e.data
    const data = typeof raw === 'string' ? JSON.parse(raw) : raw
    listeners.forEach(l => l(data))
  } catch (err) {
    console.error('Parse host message error:', err, 'raw:', e.data)
  }
})
```

---

## 三、按类别的经验教训

### A. COM / Excel 加载项

1. **每次 DLL 重编译后必须重新注册 CLSID**（`scripts\register-user.ps1`）。ProgId 注册不等于 CLSID 注册，Excel 通过 CLSID 实例化 COM 类。
2. **Excel 会在加载失败时自动把 `LoadBehavior` 改为 2**。调试时若加载项突然消失，先查注册表 `HKCU\Software\Microsoft\Office\Excel\Addins\DeepExcel.AddIn\LoadBehavior`。
3. **Excel COM 二维数组是 1-based**（`GetLowerBound(0) == 1`）。**永远不要**用 `array[0, 0]` 访问 `Range.Value2` 的返回值，必须用 `GetLowerBound/GetUpperBound`。
4. **Excel 会锁定已加载的 DLL**。重新编译前必须 `Stop-Process EXCEL -Force`，否则编译器写入失败。
5. **`Shapes.AddChart2` 在某些 Excel 版本抛 HRESULT**。必须 try-catch 回退到 `Shapes.AddChart`。
6. **COM 操作必须在 STA 线程**。从 sidecar stdout 回调线程访问 Excel 时，必须 `_uiControl.Invoke(...)` 切回 UI 线程。
7. **`new Form()` 不会自动创建 HWND**。在 `BeginInvoke` 前必须确保 `IsHandleCreated`，或访问 `.Handle` 强制创建。

### B. WebView2 通信

1. **`PostWebMessageAsString` ↔ `chrome.webview.addEventListener`**；**`PostWebMessageAsJson` ↔ `window.addEventListener`**。两套 API 不可混用。
2. **`userDataFolder` 必须加 PID 后缀**（`WebView2_{PID}`），否则多 Excel 实例互锁。
3. **前端 → C# 用 `chrome.webview.postMessage(obj)`**，C# 端用 `WebMessageReceived` 事件接收。
4. 开发环境（Vite dev）无 `chrome.webview` 对象，必须提供 mock 实现，但 mock 路径不能掩盖真实路径的差异。

### C. Python Sidecar

1. **必须在所有 import 之前强制 stdout/stderr 为 UTF-8**：
   ```python
   sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
   sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)
   ```
   Windows 默认 cp1252，写中文/emoji 立即 UnicodeEncodeError 崩溃，C# 端会看到"一直在转"。
2. **SDK 内部用 anyio，必须用 `anyio.run(main)`，不要用 `asyncio.run`**，否则事件循环冲突。
3. **`os.environ` 设置 `ANTHROPIC_BASE_URL` 不生效**，必须通过 `ClaudeAgentOptions(env={...})` 传入。
4. **sidecar stderr 是诊断日志，不是错误**（如 `[ipc] route_message`、`[sidecar] client.query() done`）。C# 端**只写日志文件，不转发前端**。
5. **monkey-patch 库函数时必须保留原函数所有特性**。本次 patch 把 `anyio.create_memory_object_stream` 替换为普通函数，破坏了 SDK 使用的 `create_memory_object_stream[dict](...)` 泛型下标语法，导致 `TypeError: 'function' object is not subscriptable` 被 anyio task group 静默吞掉，client 永久 hang。**教训：能不 patch 就不 patch，patch 后必须用最小用例验证泛型/协程/上下文管理器等特性仍可用。**

### D. System.Text.Json 序列化

1. **System.Text.Json 不支持二维数组（`T[,]`）序列化**，会抛 `NotSupportedException`。必须注册自定义 `JsonConverter<T[,]>`。
2. **自定义 Converter 必须用 `GetLowerBound/GetUpperBound`** 遍历，因为 COM 数组是 1-based。
3. **序列化失败 fallback 必须设 `success=false`** + 明确 error 文本，否则模型看到 `success=true` 但 data 是错误信息，会陷入无限重试（实测 DeepSeek 会一直循环 `read_range`）。
4. **`BuildExcelSnapshot` / `BuildContext` 不应包含 2D 数组字段**（`Values`/`Formulas`/`NumberFormats`），只用元信息（address/sheet/rowCount/columnCount），实际数据由 sidecar 按需通过 `read_selection` / `read_range` 工具读取。

### E. Claude Agent SDK 集成

1. **DeepSeek 必须用 `/anthropic` 端点**（`https://api.deepseek.com/anthropic`），不是 `/v1`（OpenAI 格式）。
2. **`thinking` 必须禁用**：`ClaudeAgentOptions(thinking={"type": "disabled"})`。DeepSeek anthropic 兼容端点不返回 thinking block 的 `signature` 字段，SDK 在 `message_parser.py:104` 硬编码 `block["signature"]` 会抛 `MessageParseError`，导致工具调用后第二轮 API 响应解析崩溃，文本永不返回。参考 [anthropics/claude-agent-sdk-python#949](https://github.com/anthropics/claude-agent-sdk-python/issues/949)。
3. **`env` 必须在创建 `ClaudeAgentOptions` 时传入**，`os.environ` 在 SDK 子进程中不生效。
4. **用 `async with ClaudeSDKClient(options=options) as client`** 上下文管理器，不要手动 `connect/close`。
5. **sidecar 错误消息不能用 emoji**（编码风险），纯文本最安全。

### F. UI / UX

1. **stderr ≠ error**。Python sidecar 把诊断日志写到 stderr 是惯例（避免污染 stdout 的 JSON Lines 协议），C# 端必须区分对待：stderr 写日志文件，stdout 走协议解析。
2. **工具调用通知不要重复发送**。SDK 有 `tool_use`（即将调用）和 `tool_call`（实际请求 C# 执行）两个事件，只在 `tool_use` 时通知 UI，`tool_call` 仅记日志。
3. **工具名显示去掉 `mcp__excel__` 前缀**，用户看到 `read_workbook` 比 `mcp__excel__read_workbook` 友好得多。
4. **前端必须处理 `stream_end`** 才能停止 loading 状态和移除流式光标，否则面板永远转圈。

---

## 四、调试方法论

### 1. 用最小用例隔离问题

不要在完整链路中调试。本次调试创建了多个最小测试脚本：

- `_test_sdk_minimal.py` — 无 MCP，验证 SDK + DeepSeek 基础连通
- `_test_sdk_mcp.py` — SDK + MCP server
- `_test_sdk_buffer_patch.py` — 验证 buffer patch 是否导致 hang
- `_test_sidecar_thinking_disabled.py` — sidecar 端到端简单对话
- `_test_sidecar_tool_call.py` — sidecar 工具调用（echo）

**每修一个 bug，先用最小用例验证修复有效，再回到完整链路。**

### 2. 全链路日志分段验证

DeepExcel 的消息链路是：`用户输入 → bridge.ts → C# MessageBridge → C# PythonSidecar → sidecar.py → ClaudeSDKClient → DeepSeek API`，返回路径反向。**任何一段断了，现象都一样：UI 转圈无响应**。

排查时必须分段验证：
- sidecar 直连测试（绕过 C# 和前端）→ 验证 SDK + API
- C# 日志确认 `OnStdoutLine` 收到 `stream_delta` → 验证 sidecar → C# 通道
- C# 日志确认 `PostWebMessageAsString` 已调用 → 验证 C# → 前端发送
- **浏览器 DevTools 确认 `chrome.webview` 上有 `message` 事件触发** → 验证前端接收（**这一步最容易被跳过**）

### 3. "后端正常但 UI 没反应" 的排查清单

按优先级排序：

1. 浏览器 DevTools Console 是否有 JS 错误？
2. `chrome.webview.addEventListener` 是否正确注册？（不是 `window.addEventListener`）
3. C# `PostWebMessageAsString` 是否真的调用了？（查 Logger）
4. `SafeBeginInvoke` 是否因 `IsHandleCreated=false` 丢弃了事件？（查 Logger 的 `DROPPED` 警告）
5. sidecar stdout 是否真的输出了 `stream_delta`？（查 Logger 的 `OnStdoutLine`）
6. sidecar stderr 是否有隐藏异常？（查 `deepexcel-YYYYMMDD.log`）

### 4. monkey-patch 的安全规则

- **能不 patch 就不 patch**。优先升级库版本或换 API。
- 必须patch 时，**保留原函数的所有特性**：泛型下标 `[]`、协程、上下文管理器、属性访问。
- patch 后用**最小用例**验证被 patch 的函数仍能被库的所有调用方式使用。
- anyio 的 `create_memory_object_stream` 支持 `func[T](buffer)` 泛型下标，普通函数不支持——这是本次最隐蔽的坑。

### 5. Excel 加载项"突然消失"的排查清单

1. `HKCU\Software\Microsoft\Office\Excel\Addins\DeepExcel.AddIn\LoadBehavior` 是否为 3？（被自动改成 2 是 Excel 的自我保护）
2. `HKCU\Software\Classes\CLSID\{A1B2C3D4-...}` 和 `{B2C3D4E5-...}` 是否存在？（重编译 DLL 后会丢失）
3. 是否有残留的 Python sidecar 进程占用资源？
4. `%APPDATA%\DeepExcel\logs\deepexcel-YYYYMMDD.log` 是否有异常？
5. `bin\Release\DeepExcel_Load.log` 是否记录了加载失败原因？

---

## 五、防御性编程清单（供 PR review 用）

新增/修改代码时检查以下项：

### C# 侧
- [ ] 所有 `Range.Value2` / `Range.Formula` 等返回的二维数组，遍历用 `GetLowerBound/GetUpperBound`
- [ ] `JsonSerializerOptions` 注册了 `Object2DArrayConverter` 和 `String2DArrayConverter`
- [ ] 从非 UI 线程访问 Excel COM 时，用 `_uiControl.Invoke(...)` 切回 STA
- [ ] `BeginInvoke` 前检查 `IsHandleCreated && !IsDisposed`（用 `SafeBeginInvoke`）
- [ ] 序列化失败 fallback 设 `success=false` + 明确 error
- [ ] stderr 只写日志，不转发前端
- [ ] `PostWebMessageAsString` 发送的是 JSON 字符串，前端用 `chrome.webview.addEventListener`

### Python 侧
- [ ] 入口处 `sys.stdout.reconfigure(encoding='utf-8')`
- [ ] 用 `anyio.run(main)`，不要 `asyncio.run`
- [ ] DeepSeek 等 anthropic 兼容端点：`thinking={"type": "disabled"}`
- [ ] API 配置通过 `ClaudeAgentOptions(env={...})` 传入，不依赖 `os.environ`
- [ ] 错误消息纯文本，不用 emoji
- [ ] monkey-patch 后用最小用例验证泛型/协程特性

### 前端侧
- [ ] `bridge.ts` 用 `chrome.webview.addEventListener('message', ...)` 监听 C# 消息
- [ ] `event.data` 是字符串时手动 `JSON.parse`
- [ ] 处理 `stream_end` 停止 loading 状态
- [ ] 开发环境 mock 不掩盖 `chrome.webview` 真实路径的差异

### 部署侧
- [ ] DLL 重编译后运行 `scripts\register-user.ps1` 重新注册 CLSID
- [ ] 编译前 `Stop-Process EXCEL -Force` 释放 DLL 锁
- [ ] `userDataFolder` 加 PID 后缀
- [ ] `config.json` 和 `*.crypt` 不入 git（已在 `.gitignore`）

---

## 六、后续待办（多模型扩展注意事项）

当后续支持 Kimi / Qwen / ChatGPT / Claude 时，重点关注：

1. **API 端点格式**：每个 provider 的 anthropic 兼容端点路径不同（DeepSeek 是 `/anthropic`，Kimi/Qwen 可能不同），需在 `AppConfig.cs` 中按 provider 配置。
2. **thinking 字段兼容性**：每个 provider 对 thinking block 的支持程度不同，建议**默认禁用** thinking，仅对原生 Claude 模型启用。
3. **error signature 字段**：anthropic 兼容端点普遍不返回 thinking signature，SDK 的硬编码解析是潜在地雷，多模型时这是 P0 检查项。
4. **流式协议差异**：不同 provider 的 stream chunk 格式可能略有差异，SDK 已封装但需验证。
5. **token 限制**：不同模型 `max_turns` / 上下文窗口不同，`ClaudeAgentOptions.max_turns=20` 可能需要按模型调整。

---

## 七、关键文件索引

| 文件 | 关键修复点 |
|------|-----------|
| [sidecar.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) | UTF-8 编码、anyio.run、thinking 禁用、env 传入、移除 buffer patch |
| [bridge.ts](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/bridge.ts) | `chrome.webview.addEventListener`（最关键修复） |
| [PythonSidecar.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) | SafeBeginInvoke、stderr 不转发、SendToolResult fallback success=false |
| [JsonConverters.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/JsonConverters.cs) | GetLowerBound/GetUpperBound 处理 1-based COM 数组 |
| [MessageBridge.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) | BuildContext 只放元信息、SendConfigToSidecar、工具调用去重 |
| [ChartTool.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Tools/ChartTool.cs) | AddChart2 → AddChart 回退 |
| [system_prompt.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/system_prompt.py) | 按需读取策略，避免过度工具调用 |
| [scripts/register-user.ps1](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/scripts/register-user.ps1) | CLSID 注册（每次重编译后必跑） |

---

**最后修订**：2026-06-27
**Git tag**：v1.0（https://github.com/1256048998/DeepExcel）
