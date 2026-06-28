# DeepExcel Claude Agent SDK 迁移设计

**日期**: 2026-06-26
**状态**: 设计已确认，待写实施计划
**作者**: DeepExcel Team

## 1. 背景与动机

### 1.1 当前架构问题

DeepExcel 当前是纯手写实现：HTTP + SSE 解析 + 手写 agent loop（`Orchestrator.cs`）。在测试中暴露以下问题：

1. **模糊指令处理差**：用户输入"统计 A 列数据"，DeepSeek 模型常陷入长篇思考或输出 VBA 代码块让用户手动执行，无法稳定自动调用工具
2. **agent loop 维护成本高**：`Orchestrator.cs` 600+ 行手写多轮循环 + 重试 + follow-up，bug 频出（已修复多轮：`NeedsFollowUp` 误返回 false、`FollowUpWithToolResultStreamAsync` 清空 tools、`GetArg<T>` 无法处理 JsonElement 等）
3. **多模型时 prompt 一致性差**：同一 system prompt 在 Claude/DeepSeek/OpenAI 上行为差异大，DeepSeek 严格遵循 tool use 协议的稳定性不如 Claude
4. **无上下文压缩**：长对话会爆 token，无 prompt caching

### 1.2 迁移目标

迁移到 Claude Agent SDK 作为 AI 底座，解决上述问题：

- **SDK 内置 agent loop** — 删除全部手写循环代码
- **SDK 内置 thinking** — 替代独立"语义转换层"，模型自身做意图推断
- **SDK 内置上下文压缩 + prompt caching** — 长对话不爆 token
- **保留多模型支持** — 通过 `ANTHROPIC_BASE_URL` 环境变量切换端点

### 1.3 关键决策

- **不建独立语义转换层**：Claude 的 thinking 能力足以承担"模糊指令→具体操作"的转换，外置语义层是过度工程
- **采用"先尝试默认值，失败再问"策略**：模型按默认规则推断执行，工具返回 `suggestion` 字段时才触发反问
- **Python sidecar 接入**：Claude Agent SDK 无 C# 版本，通过子进程 + stdin/stdout JSON Lines 通信

## 2. 整体架构

```
┌──────────────────────────────────────────────────────────────┐
│ Excel.exe (STA 主线程)                                        │
│  ┌────────────────────────────────────────────────────────┐  │
│  │ DeepExcel.AddIn (C#/.NET 4.8)                          │  │
│  │  ┌──────────┐  ┌────────────┐  ┌──────────────────┐   │  │
│  │  │ WebView2 │↔│ MessageBridge│↔│ PythonSidecar    │   │  │
│  │  │ React UI │  │ (路由+STA调度)│  │ (Process+IPC)    │   │  │
│  │  └──────────┘  └────────────┘  └─────────┬────────┘   │  │
│  │                                          │ JSON Lines │  │
│  │            ┌─────────────────────────────┘            │  │
│  │            │ 工具调用请求/结果                          │  │
│  │            ▼                                            │  │
│  │  ┌─────────────────────────┐                           │  │
│  │  │ ExcelActionsImpl (C#)   │ ← 保留不动                │  │
│  │  │ Executor/Perception     │                           │  │
│  │  └─────────────────────────┘                           │  │
│  └────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘
                           │ stdin/stdout
                           ▼
┌──────────────────────────────────────────────────────────────┐
│ python.exe sidecar.py                                        │
│  ┌────────────────────────────────────────────────────────┐  │
│  │ claude_agent_sdk.ClaudeSDKClient                       │  │
│  │  · 内置 agent loop（替代 Orchestrator）                 │  │
│  │  · thinking（替代语义转换层）                           │  │
│  │  · 上下文压缩 + prompt caching                          │  │
│  │  ┌──────────────────────────────────────────────┐     │  │
│  │  │ create_sdk_mcp_server("excel", tools=[...])  │     │  │
│  │  │  @tool write_formula                          │     │  │
│  │  │  @tool read_range    ← 模型自己决定何时读    │     │  │
│  │  │  @tool execute_vba                            │     │  │
│  │  │  ... (13 + 1 个工具)                          │     │  │
│  │  └──────────────────────────────────────────────┘     │  │
│  └────────────────────────────────────────────────────────┘  │
│                                                              │
│  模型切换：环境变量 ANTHROPIC_BASE_URL                       │
│    · Claude:  https://api.anthropic.com                      │
│    · GLM:     https://open.bigmodel.cn/api/anthropic         │
│    · DeepSeek: http://localhost:8000/v1 (LiteLLM 代理)       │
└──────────────────────────────────────────────────────────────┘
                           │ HTTPS
                           ▼
                    Anthropic / 兼容端点
```

### 2.1 文件层面变化

| 类型 | 文件 |
|---|---|
| **新增** | `src/DeepExcel.Sidecar/sidecar.py`、`src/DeepExcel.Sidecar/excel_tools.py`、`src/DeepExcel.Sidecar/ipc.py` |
| **新增** | `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs`、`src/DeepExcel.AddIn/Sidecar/SidecarProtocol.cs` |
| **修改** | `ThisAddIn.cs`、`Bridge/MessageBridge.cs` — 改为转发 sidecar |
| **废弃** | `Agent/Orchestrator.cs`、`Agent/ToolRegistry.cs`、`Models/ClaudeAdapter.cs`、`Models/DeepSeekAdapter.cs`、`Models/OpenAIAdapter.cs`、`Models/OpenAICompatibleAdapter.cs`、`Models/IModelAdapter.cs`、`Models/ModelAdapterFactory.cs` |
| **保留不动** | `Bridge/IExcelActions.cs`、`Bridge/ExcelActionsImpl.cs`、`Executor/*`、`Perception/*`、`Security/*` |

## 3. IPC 协议

所有消息一行一个 JSON，以 `\n` 分隔（JSON Lines over stdin/stdout）。

### 3.1 C# → Python

```json
{"type":"user_message","text":"统计A列数据","session_id":"sess-xxx","context":{...}}

{"type":"cancel"}

{"type":"tool_result","call_id":"u1","success":true,"data":{...},"error":null,"suggestion":null,"context":{...}}

{"type":"config","base_url":"https://api.anthropic.com","model":"claude-sonnet-4","api_key":"sk-..."}
```

### 3.2 Python → C#

```json
{"type":"stream_delta","text":"正在读取A列..."}

{"type":"tool_call","call_id":"u1","tool":"read_range","args":{"address":"A:A"}}

{"type":"clarify","question":"A列有数字和文本，您想统计什么？","options":["求和","计数","非空计数"]}

{"type":"stream_end","input_tokens":1234,"output_tokens":567}
```

### 3.3 设计要点

- `tool_result` 用 `call_id` 关联请求 — sidecar 可能并行发出多个工具调用，必须能匹配回
- `clarify` 单列一种消息类型 — 让前端能用选项卡片 UI（不只是显示文本）
- `config` 支持热切换模型 — 不重启 sidecar

## 4. 工具返回值结构

### 4.1 统一返回结构

所有工具统一返回这个结构（C# 侧 `ToolResult` 类，Python 侧 dict）：

```csharp
public class ToolResult
{
    public bool Success { get; set; }
    public object Data { get; set; }
    public string Error { get; set; }
    public string Suggestion { get; set; }    // 新增：建议澄清
    public ExcelSnapshot Context { get; set; } // 新增：当前 Excel 上下文快照
}
```

```python
def make_result(success=True, data=None, error=None, suggestion=None, context=None):
    return {
        "success": success,
        "data": data,
        "error": error,
        "suggestion": suggestion,
        "context": context,
    }
```

### 4.2 "失败再问"触发逻辑

`Suggestion` 字段非空时，SDK 看到工具返回会自动让模型反问用户。典型流程（以"统计 A 列"为例）：

```
用户："统计A列数据"
  ↓ SDK agent loop（自动多轮）

轮 1: 模型 thinking "统计=默认SUM"
      调 read_range("A:A")
      → C# 返回 {success:true, data:["苹果","香蕉","橙子"],
                suggestion:"A列是文本，无法求和。请确认是计数(COUNTA)还是求和(SUM)？"}

轮 2: 模型看到 suggestion → 调 clarify_intent 工具
      → 前端显示选项卡片 ["COUNTA计数","SUM求和（可能为0）"]
      → 用户选 "COUNTA计数"

轮 3: 模型调 write_formula("B1", "=COUNTA(A:A)")
      → C# 返回 {success:true, data:{address:"B1", formula:"=COUNTA(A:A)"}}

轮 4: 模型输出 "已在 B1 写入 =COUNTA(A:A)，结果为 3。"
      → stream_end
```

### 4.3 工具列表

保留现有 13 个工具，新增 1 个：

| 工具 | 改动 |
|---|---|
| `read_workbook` | 不变 |
| `read_selection` | 不变 |
| `read_range` | 返回值加 `data_type` 字段（"number"/"text"/"date"/"mixed"） |
| `write_formula` | 返回值加 `suggestion` 字段，写失败时给澄清提示 |
| `fill_formula_down` | 同上 |
| `replace_formula` | 不变 |
| `clean_data` | 不变 |
| `create_chart` | 不变 |
| `create_pivot_table` | 不变 |
| `execute_vba` | 不变 |
| `execute_python` | 不变 |
| `create_snapshot` | 不变 |
| `rollback` | 不变 |
| **新增** `clarify_intent` | 专用工具，让模型主动向用户提问（配合前端选项卡片） |

`clarify_intent` 工具定义：

```python
@tool("clarify_intent", "向用户提问以澄清模糊指令", {
    "question": str,
    "options": list  # 可选，提供选项时前端渲染为卡片
})
async def clarify_intent(args):
    # 通过 IPC 发送 clarify 消息到 C#，C# 转发到前端
    # 阻塞等待用户回答
    user_answer = await call_csharp_clarify(args["question"], args.get("options", []))
    return {"content": [{"type": "text", "text": f"用户回答：{user_answer}"}]}
```

## 5. System Prompt 策略

### 5.1 完整 System Prompt

```python
SYSTEM_PROMPT = """你是 DeepExcel AI Agent，住在 Excel 里，通过调用工具直接操作工作簿。

## 核心行为准则
1. **直接执行**：通过调用工具完成任务，绝不输出代码块让用户手动运行
2. **先读后写**：写入公式前，先调 read_range 确认目标数据类型
3. **失败再问**：工具返回 success=false 时，向用户说明问题并建议方案；工具返回 suggestion 字段时，按 suggestion 提示用户确认
4. **简洁汇报**：工具成功后用一句话总结结果，不要复述工具返回的原始 JSON

## 模糊指令的默认推断规则

当用户指令含糊时，按以下默认规则执行，不要主动反问：

| 用户说 | 默认推断 | 工具调用 |
|---|---|---|
| "统计 X" | 求和 SUM | write_formula |
| "汇总 X" | 求和 SUM | write_formula |
| "算一下 X" | 求和 SUM | write_formula |
| "数一下 X" | 计数 COUNTA | write_formula |
| "有多少 X" | 计数 COUNTA | write_formula |
| "平均 X" / "X 均值" | 平均值 AVERAGE | write_formula |
| "排名 X" | 排名 RANK | write_formula |
| "去重 X" | 删除重复项 | clean_data |
| "格式化 X" | 统一日期格式 + 去空格 | clean_data |
| "画图 X" / "图表 X" | 柱状图 | create_chart |
| "透视 X" | 按第一列分组求和 | create_pivot_table |

**反问时机（仅在以下情况触发 clarify_intent）：**
- read_range 返回 data_type=mixed（同一列既有数字又有文本）
- write_formula 返回 success=false 且 suggestion 非空
- 用户指令完全无法映射到任何工具（如"帮我做个 PPT"）
- 用户指令的"目标位置"完全无法推断（如"统计 A 列"但没说写哪里，默认写到 B1 即可，不反问）

## 工具使用决策树

遇到"写公式"类指令：
1. 先调 read_range 看数据类型 → 数字→SUM/AVERAGE；日期→COUNTA；文本→COUNTA
2. 默认目标单元格：源数据列的右侧相邻列第一个单元格（如 A 列数据→写 B1）
3. 公式格式：以 = 开头，英文函数名，全列引用用 A:A，范围用 A1:A100

遇到"复杂操作"（循环/条件/多步骤）：
1. 优先用 execute_vba，code 参数放完整 Sub
2. VBA 无法完成的（如正则、复杂数学）用 execute_python
3. Python 可用 ctx 字典获取上下文，用 set_cell/write_range 写回 Excel

## 当前 Excel 上下文
{excel_context}
"""
```

### 5.2 上下文注入

每次 `user_message` 时，C# 自动附带当前 Excel 上下文：

```json
{
  "type": "user_message",
  "text": "统计A列数据",
  "session_id": "sess-xxx",
  "context": {
    "workbook": "销售数据.xlsx",
    "active_sheet": "Sheet1",
    "sheets": [
      {"name": "Sheet1", "used_range": "A1:E100",
       "column_types": {"A": "text", "B": "number", "C": "date"}}
    ],
    "selection": "A1:A10"
  }
}
```

`column_types` 由 C# 侧 `Perception` 模块预扫描得出，模型不用调 `read_range` 就能知道每列数据类型。多数情况直接跳过 read_range，降低轮次开销。

### 5.3 模型选择策略

不同任务用不同模型，但都在 SDK 框架内，靠 `ANTHROPIC_MODEL` + `ANTHROPIC_BASE_URL` 切换：

| 任务类型 | 推荐模型 | 切换方式 |
|---|---|---|
| 简单公式/读写 | Claude Haiku（快） | `config` 消息指定 |
| 复杂 agent loop（多工具链） | Claude Sonnet | 默认 |
| 复杂 VBA 生成 | Claude Sonnet | 默认 |
| 用户选了 DeepSeek/Qwen | 配合 LiteLLM 代理 | `config` 指定 base_url |

**不实现自动选模型** — 让用户在前端手动选，避免过度工程。

## 6. C# 侧实现

### 6.1 PythonSidecar.cs 核心逻辑

```csharp
public class PythonSidecar : IDisposable
{
    private Process _process;
    private readonly IExcelActions _excel;
    private readonly Control _uiControl;  // WebView2 引用，用于 Invoke
    private readonly Action<string> _onStreamDelta;
    private readonly Action<string, List<string>> _onClarify;
    private readonly Action<int, int> _onStreamEnd;

    private readonly ConcurrentDictionary<string, TaskCompletionSource<ToolResult>> _pendingToolCalls
        = new ConcurrentDictionary<string, TaskCompletionSource<ToolResult>>();

    public void Start()
    {
        var psi = new ProcessStartInfo
        {
            FileName = GetPythonPath(),  // 内嵌 Python 或系统 python
            Arguments = $"\"{GetSidecarPath()}\"",
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
        };
        _process = Process.Start(psi);
        _process.OutputDataReceived += OnStdoutLine;
        _process.ErrorDataReceived += OnStderrLine;
        _process.BeginOutputReadLine();
        _process.BeginErrorReadLine();
    }

    public void SendUserMessage(string text, string sessionId, ExcelContext ctx)
    {
        var msg = new { type = "user_message", text, session_id = sessionId, context = ctx };
        WriteLine(JsonSerializer.Serialize(msg));
    }

    public void Cancel() => WriteLine(@"{""type"":""cancel""}");

    public void UpdateConfig(string baseUrl, string model, string apiKey) =>
        WriteLine(JsonSerializer.Serialize(new { type = "config", base_url = baseUrl, model, api_key = apiKey }));

    private void OnStdoutLine(object sender, DataReceivedEventArgs e)
    {
        if (string.IsNullOrEmpty(e.Data)) return;

        var msg = JsonDocument.Parse(e.Data).RootElement;
        var type = msg.GetProperty("type").GetString();

        switch (type)
        {
            case "stream_delta":
                _uiControl.BeginInvoke(new Action(() =>
                    _onStreamDelta?.Invoke(msg.GetProperty("text").GetString())));
                break;

            case "tool_call":
                HandleToolCall(msg);  // 异步，不阻塞 reader
                break;

            case "clarify":
                var q = msg.GetProperty("question").GetString();
                var opts = msg.GetProperty("options").EnumerateArray()
                    .Select(x => x.GetString()).ToList();
                _uiControl.BeginInvoke(new Action(() => _onClarify?.Invoke(q, opts)));
                break;

            case "stream_end":
                var inTok = msg.GetProperty("input_tokens").GetInt32();
                var outTok = msg.GetProperty("output_tokens").GetInt32();
                _uiControl.BeginInvoke(new Action(() => _onStreamEnd?.Invoke(inTok, outTok)));
                break;
        }
    }

    private async void HandleToolCall(JsonElement msg)
    {
        var callId = msg.GetProperty("call_id").GetString();
        var toolName = msg.GetProperty("tool").GetString();
        var args = msg.GetProperty("args").Deserialize<Dictionary<string, object>>();

        // ★ 关键：切回 UI 线程执行 Excel COM 操作（STA 要求）
        var result = await _uiControl.InvokeAsync<ToolResult>(async () =>
        {
            try { return await ExecuteToolAsync(toolName, args); }
            catch (Exception ex) { return new ToolResult { Success = false, Error = ex.Message }; }
        });

        var response = new
        {
            type = "tool_result",
            call_id = callId,
            success = result.Success,
            data = result.Data,
            error = result.Error,
            suggestion = result.Suggestion,  // ★ 澄清提示
            context = BuildExcelSnapshot()   // ★ 实时上下文
        };
        WriteLine(JsonSerializer.Serialize(response));
    }

    private async Task<ToolResult> ExecuteToolAsync(string name, Dictionary<string, object> args)
    {
        switch (name)
        {
            case "read_range":
                var data = _excel.ReadRange(GetArg<string>(args, "address"));
                return new ToolResult
                {
                    Success = true,
                    Data = data,
                    Suggestion = GenerateRangeSuggestion(data),
                    Context = BuildExcelSnapshot()
                };

            case "write_formula":
                var result = _excel.WriteFormula(GetArg<string>(args, "address"), GetArg<string>(args, "formula"));
                if (!result.Success)
                {
                    result.Suggestion = GenerateFormulaSuggestion(args, result.Error);
                }
                return result;

            // ... 其余 11 个工具，复用现有 ExcelActionsImpl
        }
    }

    private string GenerateRangeSuggestion(object data)
    {
        // 检测数据类型混合
        // "A 列包含数字和文本，无法直接求和。建议：求和(忽略文本)、计数(全部)、计数(仅数字)"
        // ...
    }

    private string GenerateFormulaSuggestion(Dictionary<string, object> args, string error)
    {
        var formula = GetArg<string>(args, "formula");
        if (formula.Contains("SUM") && error.Contains("类型不匹配"))
            return "目标区域包含文本，无法求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
        if (formula.Contains("AVERAGE") && error.Contains("除数为零"))
            return "目标区域没有数字，无法计算平均值。建议检查数据源。";
        return null;
    }

    private void WriteLine(string json)
    {
        _process.StandardInput.WriteLine(json);
        _process.StandardInput.Flush();
    }
}
```

### 6.2 线程模型

```
┌─────────────────────────────────────────────────┐
│ Excel STA 主线程                                │
│  ┌─────────────────────────────────────────┐    │
│  │ WebView2 UI  ←  _uiControl.Invoke       │    │
│  │ ExcelActionsImpl ← COM 操作必须在此线程 │    │
│  └─────────────────────────────────────────┘    │
│         ▲                                       │
│         │ BeginInvoke / InvokeAsync             │
└─────────┼───────────────────────────────────────┘
          │
┌─────────┴───────────────────────────────────────┐
│ PythonSidecar OutputDataReceived 线程（后台）   │
│  · 解析 JSON 消息                               │
│  · stream_delta/clarify/stream_end → BeginInvoke│
│  · tool_call → InvokeAsync 等待结果             │
└─────────┬───────────────────────────────────────┘
          │ stdin/stdout
          ▼
┌─────────────────────────────────────────────────┐
│ Python sidecar 进程                             │
│  · ClaudeSDKClient.receive_response() 协程      │
│  · @tool 函数通过 IPC 调 C#                     │
│  · 单线程 asyncio，无并发问题                   │
└─────────────────────────────────────────────────┘
```

**核心约束：**
1. Excel COM 必须在 STA 主线程 — 工具调用通过 `_uiControl.InvokeAsync` 切回
2. stream_delta 不能阻塞 tool_call — 用 BeginInvoke（fire-and-forget），tool_call 用 InvokeAsync（等结果）
3. Python sidecar 是单进程 — asyncio 单线程，不存在并发问题
4. stdout 必须逐行读 — `BeginOutputReadLine` 自动处理

### 6.3 MessageBridge 改动

```csharp
// 原 Orchestrator 调用全部替换为 sidecar 转发
public async Task ProcessAsync(string userMessage)
{
    var context = _perception.CollectContext();
    var sessionId = Guid.NewGuid().ToString();

    _sidecar.SendUserMessage(userMessage, sessionId, context);
}

public void Cancel() => _sidecar.Cancel();

public void UpdateModel(string baseUrl, string model, string apiKey) =>
    _sidecar.UpdateConfig(baseUrl, model, apiKey);
```

## 7. Python Sidecar 实现

### 7.1 sidecar.py 主循环

```python
import asyncio
import json
import os
import sys
from claude_agent_sdk import ClaudeAgentOptions, ClaudeSDKClient, create_sdk_mcp_server
from excel_tools import register_all_tools
from ipc import write_message, _message_buffer

async def main():
    # 注册所有工具到 MCP server
    server = create_sdk_mcp_server(name="excel", tools=register_all_tools())

    options = ClaudeAgentOptions(
        mcp_servers={"excel": server},
        allowed_tools=[f"mcp__excel__{t}" for t in [
            "read_workbook", "read_selection", "read_range",
            "write_formula", "fill_formula_down", "replace_formula",
            "clean_data", "create_chart", "create_pivot_table",
            "execute_vba", "execute_python",
            "create_snapshot", "rollback", "clarify_intent"
        ]],
        system_prompt=os.environ.get("DEEPEXCEL_SYSTEM_PROMPT", ""),
        max_turns=20,  # 限制最大轮次，防止无限循环
    )

    client = ClaudeSDKClient(options=options)
    await client.connect()

    # 启动 stdin reader 协程（唯一读 stdin 的地方）
    asyncio.create_task(stdin_reader_loop())

    # 主循环：从 buffer queue 读已分发的消息
    while True:
        # 等待 user_message（user_message 永远是会话触发点）
        msg = await _message_buffer["user_message"].get()
        if msg is None:
            break

        await client.query(msg["text"])
        async for response in client.receive_response():
            await handle_sdk_message(response)

        # 处理期间到达的 config 消息（不阻塞主循环）
        while not _message_buffer["config"].empty():
            cfg = await _message_buffer["config"].get()
            os.environ["ANTHROPIC_BASE_URL"] = cfg["base_url"]
            os.environ["ANTHROPIC_MODEL"] = cfg["model"]
            os.environ["ANTHROPIC_API_KEY"] = cfg["api_key"]

        # cancel 通过 Event 检查
        if _message_buffer["cancel"].is_set():
            await client.interrupt()
            _message_buffer["cancel"].clear()

async def stdin_reader_loop():
    """唯一的 stdin 读取协程，把消息分发到 _message_buffer"""
    from ipc import read_message, route_message
    while True:
        msg = await read_message()
        if msg is None:
            # stdin 关闭，往 user_message queue 发 None 让主循环退出
            await _message_buffer["user_message"].put(None)
            return
        route_message(msg)
        # cancel 消息也通过 Event 触发
        if msg.get("type") == "cancel":
            _message_buffer["cancel"].set()

async def handle_sdk_message(response):
    """处理 ClaudeSDKClient 的流式消息"""
    from claude_agent_sdk import AssistantMessage, ResultMessage
    from claude_agent_sdk.types import TextBlock, ToolUseBlock

    if isinstance(response, AssistantMessage):
        for block in response.content:
            if isinstance(block, TextBlock):
                await write_message({"type": "stream_delta", "text": block.text})
            elif isinstance(block, ToolUseBlock):
                # SDK 内部已调用 @tool 函数，工具函数会通过 IPC 调 C#
                # 这里不需要额外处理
                pass
    elif isinstance(response, ResultMessage):
        await write_message({
            "type": "stream_end",
            "input_tokens": response.usage.input_tokens,
            "output_tokens": response.usage.output_tokens
        })

asyncio.run(main())
```

**关键设计**：只有一个 `stdin_reader_loop` 协程读 stdin，把消息分发到 `_message_buffer` 里的不同 Queue/Event/Dict。工具协程调用 `call_csharp()` 时从 buffer 取结果，主循环从 `user_message` Queue 取会话触发消息。避免了多协程争抢 stdin。

### 7.2 excel_tools.py（13 + 1 个工具）

```python
import asyncio
import json
from claude_agent_sdk import tool
from ipc import call_csharp, call_csharp_clarify

@tool("read_workbook", "读取当前工作簿的结构信息", {})
async def read_workbook(args):
    result = await call_csharp("read_workbook", {})
    return {"content": [{"type": "text", "text": json.dumps(result)}]}

@tool("read_range", "读取指定范围的单元格数据", {"address": str})
async def read_range(args):
    result = await call_csharp("read_range", {"address": args["address"]})
    return {"content": [{"type": "text", "text": json.dumps(result)}]}

@tool("write_formula", "向指定单元格写入Excel公式", {"address": str, "formula": str})
async def write_formula(args):
    result = await call_csharp("write_formula", {
        "address": args["address"],
        "formula": args["formula"]
    })
    return {"content": [{"type": "text", "text": json.dumps(result)}]}

@tool("clarify_intent", "向用户提问以澄清模糊指令", {"question": str, "options": list})
async def clarify_intent(args):
    user_answer = await call_csharp_clarify(args["question"], args.get("options", []))
    return {"content": [{"type": "text", "text": f"用户回答：{user_answer}"}]}

# ... 其余 10 个工具同理，每个 ~10 行

def register_all_tools():
    return [
        read_workbook, read_selection, read_range,
        write_formula, fill_formula_down, replace_formula,
        clean_data, create_chart, create_pivot_table,
        execute_vba, execute_python,
        create_snapshot, rollback, clarify_intent
    ]
```

### 7.3 ipc.py（通信层）

```python
import asyncio
import json
import sys
from typing import Any, Dict

# 缓存已收到但尚未被消费的消息（按 type 分类）
# 关键：C# 可能并行发来多个 tool_result，必须按 call_id 匹配，不能顺序读
_message_buffer = {
    "tool_result": {},  # call_id -> msg
    "clarify_answer": None,  # 最近一次 clarify 的回答
    "user_message": asyncio.Queue(),
    "cancel": asyncio.Event(),
    "config": asyncio.Queue(),
}

async def call_csharp(tool_name: str, args: dict) -> dict:
    """向 C# 发送工具调用请求，阻塞等待结果"""
    call_id = generate_call_id()
    await write_message({
        "type": "tool_call",
        "call_id": call_id,
        "tool": tool_name,
        "args": args
    })
    # 阻塞等待对应 call_id 的结果，不匹配的消息缓存起来
    while True:
        if call_id in _message_buffer["tool_result"]:
            return _message_buffer["tool_result"].pop(call_id)
        # 没有匹配的，读下一条消息并分类
        msg = await read_message()
        if msg is None:
            raise ConnectionError("stdin closed")
        route_message(msg)

async def call_csharp_clarify(question: str, options: list) -> str:
    """向 C# 发送澄清请求，阻塞等待用户回答"""
    await write_message({
        "type": "clarify",
        "question": question,
        "options": options
    })
    # 阻塞等待 clarify_answer
    while _message_buffer["clarify_answer"] is None:
        msg = await read_message()
        if msg is None:
            raise ConnectionError("stdin closed")
        route_message(msg)
    answer = _message_buffer["clarify_answer"]
    _message_buffer["clarify_answer"] = None
    return answer

def route_message(msg: dict):
    """把收到的消息按 type 路由到对应缓冲区"""
    t = msg.get("type")
    if t == "tool_result":
        _message_buffer["tool_result"][msg["call_id"]] = msg
    elif t == "clarify_answer":
        _message_buffer["clarify_answer"] = msg.get("answer", "")
    elif t == "user_message":
        asyncio.create_task(_message_buffer["user_message"].put(msg))
    elif t == "cancel":
        _message_buffer["cancel"].set()
    elif t == "config":
        asyncio.create_task(_message_buffer["config"].put(msg))

async def write_message(msg: dict):
    """写一行 JSON 到 stdout（C# 读）"""
    sys.stdout.write(json.dumps(msg, ensure_ascii=False) + "\n")
    sys.stdout.flush()

async def read_message() -> dict:
    """从 stdin 读一行 JSON（C# 写）"""
    line = await asyncio.get_event_loop().run_in_executor(None, sys.stdin.readline)
    if not line:
        return None
    return json.loads(line)

def generate_call_id() -> str:
    import uuid
    return str(uuid.uuid4())
```

**关键改进**：用 `_message_buffer` 字典缓存所有收到的消息，按 `call_id` 匹配 tool_result。这避免了"顺序读 stdin 时不匹配的消息被丢弃"的并发问题。

## 8. 部署

### 8.1 目录结构

```
DeepExcel/
├─ src/DeepExcel.AddIn/bin/Release/
│  ├─ DeepExcel.AddIn.dll              # 主加载项
│  ├─ WebViewAssets/                   # 前端
│  ├─ python/                          # 新增：python embeddable
│  │  ├─ python.exe
│  │  └─ Lib/site-packages/
│  │     └─ claude_agent_sdk/          # SDK 及依赖
│  └─ sidecar/                         # 新增
│     ├─ sidecar.py
│     ├─ excel_tools.py
│     └─ ipc.py
```

### 8.2 安装包增量

| 组件 | 大小 |
|---|---|
| Python embeddable | ~50MB |
| claude-agent-sdk + 依赖 | ~150-200MB |
| **总计** | **~200-250MB** |

可接受。可进一步用压缩 + 首次运行下载优化。

### 8.3 配置文件

`%LOCALAPPDATA%\DeepExcel\model-config.json`：

```json
{
  "active_model": "claude-sonnet",
  "models": {
    "claude-sonnet": {
      "base_url": "https://api.anthropic.com",
      "model": "claude-sonnet-4",
      "api_key": "sk-ant-..."
    },
    "claude-haiku": {
      "base_url": "https://api.anthropic.com",
      "model": "claude-haiku-4",
      "api_key": "sk-ant-..."
    },
    "glm": {
      "base_url": "https://open.bigmodel.cn/api/anthropic",
      "model": "glm-4.6",
      "api_key": "your-glm-key"
    },
    "deepseek-via-litellm": {
      "base_url": "http://localhost:8000/v1",
      "model": "deepseek-chat",
      "api_key": "anything"
    }
  }
}
```

## 9. 风险与缓解

| 风险 | 严重度 | 缓解措施 |
|---|---|---|
| SDK 仍是 Alpha（0.2.109） | 中 | 锁版本 `claude-agent-sdk==0.2.109`，CI 烟测 |
| 安装包 +250MB | 中 | 用 python embeddable + 压缩；首次运行下载 |
| COM STA 跨进程调度 | 中 | 复用现有 MessageBridge 调度模式，InvokeAsync 切回主线程 |
| stdin/stdout 阻塞 | 低 | asyncio + run_in_executor 读 stdin，不阻塞事件循环 |
| 40 轮 vs 5 轮 token 消耗 | 低 | `max_turns=20` 限制，监控 usage |
| 模型兼容性（GLM/DeepSeek 是否真支持 Anthropic 协议） | 中 | MVP 阶段只用 Claude 验证，多模型作为 P2 扩展 |

## 10. 演进路径

### MVP 阶段（先跑通一个工具）

1. 写 `sidecar.py` 骨架，跑通 `ClaudeSDKClient` + 一个 `@tool("echo")` 工具
2. C# 用 `Process` 跑通 stdin/stdout 回声
3. 把 `write_formula` 工具接通，端到端验证"在 A1 写入 =SUM(B1:B10)"

### 完整迁移阶段

4. 把 13 + 1 个工具的 `@tool` 函数写完
5. 流式输出接通，MessageBridge 把 sidecar 的 stream_delta 转给 WebView2
6. STA 线程调度 + 错误处理 + 取消机制
7. System prompt 迁移、`max_turns` 调优
8. 打包 Python embeddable

### 多模型扩展阶段（P2）

9. 接入 GLM Anthropic 兼容端点
10. 接入 LiteLLM 代理 + DeepSeek
11. 前端模型切换 UI

## 11. 不在本次设计范围

- 前端 clarify 选项卡片 UI 的具体设计（依赖前端框架，单独立项）
- LiteLLM 代理的部署和配置（P2 多模型扩展时做）
- ExcelMaster 安装包逆向（用户已知会使用 Claude SDK，不再深究）
- 自动选模型策略（用户手动选，不过度工程）
