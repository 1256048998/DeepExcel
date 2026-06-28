# DeepExcel Claude Agent SDK 迁移实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 DeepExcel 的 AI 底座从手写 HTTP+Orchestrator 迁移到 Claude Agent SDK（Python sidecar），打通"模糊指令→默认推断→工具执行→失败再问"管线，并保留多模型切换能力。

**Architecture:**
- C# COM 加载项通过 `Process` 启动 Python sidecar 子进程，以 JSON Lines over stdin/stdout 双向通信
- Python sidecar 内运行 `claude_agent_sdk.ClaudeSDKClient`，内置 agent loop / thinking / 上下文压缩
- 工具调用由 Python 通过 IPC 发往 C#，C# 在 STA 主线程执行 Excel COM 操作后回写结果
- 模型切换通过 `ANTHROPIC_BASE_URL` 环境变量，无需魔改 SDK

**Tech Stack:**
- C# / .NET Framework 4.8 / COM (IDTExtensibility2 + IRibbonExtensibility)
- Python 3.11+ / claude-agent-sdk / asyncio
- Excel COM Interop / WebView2 / React + TypeScript
- xUnit (C# tests) / pytest (Python tests)

---

## Global Constraints

- Excel COM 对象操作必须在 STA 主线程执行；后台线程调用必须通过 `_uiControl.InvokeAsync` / `BeginInvoke` 切回
- Python sidecar 是单进程单 asyncio 事件循环，stdin 必须只有一个 reader 协程，其它协程从 `_message_buffer` 取消息
- IPC 协议为 JSON Lines：每行一个 JSON 对象，以 `\n` 分隔；stdout/stdin 编码 UTF-8
- 所有 `tool_result` 必须按 `call_id` 匹配（不能顺序读），`clarify_answer` 单独缓冲
- 工具返回值统一结构：`{success, data, error, suggestion, context}`，`suggestion` 非空触发模型反问
- System prompt 默认推断规则：统计=SUM、计数=COUNTA、平均=AVERAGE；目标单元格默认写源列右侧相邻列第一个单元格
- 锁定 `claude-agent-sdk==0.2.109` 版本（Alpha 阶段 API 可能变动）
- 多模型扩展（GLM/DeepSeek）为 P2 范围，本计划只验证 Claude 端点
- 现有 `Bridge/IExcelActions.cs` + `Bridge/ExcelActionsImpl.cs` + `Executor/*` + `Perception/*` 保留不动
- 模型 API Key 存储在 `%LOCALAPPDATA%\DeepExcel\model-config.json`，不上传

---

## 阶段一：MVP — 跑通 write_formula 端到端

目标：用一个 `write_formula` 工具验证"C# → sidecar → Claude SDK → 工具回 C# → Excel"完整环路。其余 12 个工具在阶段二补齐。

---

### Task 1: Python sidecar 项目骨架

**Files:**
- Create: `src/DeepExcel.Sidecar/requirements.txt`
- Create: `src/DeepExcel.Sidecar/README.md`
- Create: `src/DeepExcel.Sidecar/.gitignore`
- Create: `src/DeepExcel.Sidecar/tests/__init__.py`
- Create: `src/DeepExcel.Sidecar/tests/conftest.py`

**Interfaces:**
- Produces: Python 包目录结构，可 `pip install -r requirements.txt` 安装依赖

- [ ] **Step 1: 创建 requirements.txt**

```
# src/DeepExcel.Sidecar/requirements.txt
claude-agent-sdk==0.2.109
pytest==8.2.2
pytest-asyncio==0.23.7
```

- [ ] **Step 2: 创建 .gitignore**

```
# src/DeepExcel.Sidecar/.gitignore
__pycache__/
*.pyc
.pytest_cache/
.venv/
venv/
*.egg-info/
```

- [ ] **Step 3: 创建 tests/__init__.py（空文件）**

```python
# src/DeepExcel.Sidecar/tests/__init__.py
```

- [ ] **Step 4: 创建 tests/conftest.py — pytest 配置 fixtures**

```python
# src/DeepExcel.Sidecar/tests/conftest.py
import sys
import os

# 让 tests/ 能 import sidecar 模块
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
```

- [ ] **Step 5: 创建 README.md**

````markdown
# DeepExcel Sidecar

Python sidecar process for DeepExcel. Runs Claude Agent SDK and communicates with the C# COM add-in via JSON Lines over stdin/stdout.

## Development

```bash
cd src/DeepExcel.Sidecar
pip install -r requirements.txt
pytest tests/ -v
```

## Run standalone (debug)

```bash
python sidecar.py
# Reads JSON Lines from stdin, writes JSON Lines to stdout
```
````

- [ ] **Step 6: 安装依赖并验证 SDK API**

Run: `cd src\DeepExcel.Sidecar && pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`
Expected: 安装成功，无错误（使用清华 PyPI 镜像加速）

Run: `python -c "from claude_agent_sdk import ClaudeAgentOptions, ClaudeSDKClient, create_sdk_mcp_server, tool; print('SDK API OK')"`
Expected: 输出 `SDK API OK`

如果 import 失败，检查 `claude-agent-sdk` 版本是否正确，或查阅 SDK 文档调整 import 名（在 README.md 记录实际 API）。

**国内镜像说明**：本计划所有 `pip install` 命令均使用 `-i https://pypi.tuna.tsinghua.edu.cn/simple` 加速。如已在海外环境或已配置全局镜像，可省略该参数。

- [ ] **Step 7: Commit**

```bash
git add src/DeepExcel.Sidecar/
git commit -m "feat(sidecar): scaffold python sidecar project structure"
```

---

### Task 2: 实现 ipc.py — IPC 通信层

**Files:**
- Create: `src/DeepExcel.Sidecar/ipc.py`
- Create: `src/DeepExcel.Sidecar/tests/test_ipc.py`

**Interfaces:**
- Produces:
  - `async def write_message(msg: dict) -> None` — 写一行 JSON 到 stdout
  - `async def read_message() -> dict | None` — 从 stdin 读一行 JSON，EOF 返回 None
  - `def route_message(msg: dict) -> None` — 按 type 路由到 `_message_buffer`
  - `async def call_csharp(tool_name: str, args: dict) -> dict` — 发 tool_call 并阻塞等 tool_result
  - `async def call_csharp_clarify(question: str, options: list) -> str` — 发 clarify 并阻塞等 clarify_answer
  - `def generate_call_id() -> str` — UUID 生成
  - `_message_buffer: dict` — 全局消息缓冲，结构：
    ```python
    {
      "tool_result": {},        # call_id -> msg dict
      "clarify_answer": None,   # 最近一次 clarify 的回答字符串
      "user_message": asyncio.Queue(),
      "cancel": asyncio.Event(),
      "config": asyncio.Queue(),
    }
    ```

- [ ] **Step 1: 写失败测试 — test_write_and_read_message**

```python
# src/DeepExcel.Sidecar/tests/test_ipc.py
import asyncio
import io
import json
import pytest
from unittest.mock import patch


@pytest.mark.asyncio
async def test_write_message_writes_json_line_to_stdout(capsys):
    from ipc import write_message
    await write_message({"type": "stream_delta", "text": "hello"})
    captured = capsys.readouterr()
    assert captured.out == '{"type": "stream_delta", "text": "hello"}\n'


@pytest.mark.asyncio
async def test_read_message_returns_dict_from_stdin():
    from ipc import read_message
    fake_input = '{"type": "user_message", "text": "hi"}\n'
    with patch('sys.stdin', new=io.TextIOWrapper(io.BytesIO(fake_input.encode('utf-8')))):
        result = await read_message()
    assert result == {"type": "user_message", "text": "hi"}


@pytest.mark.asyncio
async def test_read_message_returns_none_on_eof():
    from ipc import read_message
    with patch('sys.stdin', new=io.TextIOWrapper(io.BytesIO(b''))):
        result = await read_message()
    assert result is None
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_ipc.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'ipc'`

- [ ] **Step 3: 实现 ipc.py 的 write_message / read_message / generate_call_id**

```python
# src/DeepExcel.Sidecar/ipc.py
import asyncio
import json
import sys
import uuid
from typing import Any, Dict, Optional


# 全局消息缓冲：所有从 stdin 收到的消息按 type 路由到这里
# 关键：C# 可能并行发来多个 tool_result，必须按 call_id 匹配，不能顺序读
_message_buffer: Dict[str, Any] = {
    "tool_result": {},        # call_id -> msg
    "clarify_answer": None,   # 最近一次 clarify 的回答字符串
    "user_message": None,     # asyncio.Queue，在 _init_buffer() 中创建
    "cancel": None,           # asyncio.Event，在 _init_buffer() 中创建
    "config": None,           # asyncio.Queue，在 _init_buffer() 中创建
}


def _init_buffer():
    """在 asyncio 事件循环内调用，创建 Queue/Event（它们绑定到当前 loop）"""
    if _message_buffer["user_message"] is None:
        _message_buffer["user_message"] = asyncio.Queue()
    if _message_buffer["cancel"] is None:
        _message_buffer["cancel"] = asyncio.Event()
    if _message_buffer["config"] is None:
        _message_buffer["config"] = asyncio.Queue()


def generate_call_id() -> str:
    return str(uuid.uuid4())


async def write_message(msg: dict) -> None:
    """写一行 JSON 到 stdout（C# 读）"""
    sys.stdout.write(json.dumps(msg, ensure_ascii=False) + "\n")
    sys.stdout.flush()


async def read_message() -> Optional[dict]:
    """从 stdin 读一行 JSON（C# 写），EOF 返回 None"""
    loop = asyncio.get_event_loop()
    # run_in_executor 避免阻塞事件循环
    line = await loop.run_in_executor(None, sys.stdin.readline)
    if not line:
        return None
    try:
        return json.loads(line)
    except json.JSONDecodeError as e:
        # 损坏的 JSON 不应导致整个 sidecar 崩溃
        await write_message({"type": "stream_delta", "text": f"[ipc] 解析 stdin 失败: {e}"})
        return None
```

- [ ] **Step 4: 运行测试确认通过**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_ipc.py::test_write_message_writes_json_line_to_stdout tests/test_ipc.py::test_read_message_returns_dict_from_stdin tests/test_ipc.py::test_read_message_returns_none_on_eof -v`
Expected: 3 passed

- [ ] **Step 5: 写失败测试 — test_route_message**

```python
# 追加到 src/DeepExcel.Sidecar/tests/test_ipc.py

@pytest.mark.asyncio
async def test_route_message_routes_tool_result_to_buffer_dict():
    from ipc import route_message, _message_buffer, _init_buffer
    _init_buffer()
    _message_buffer["tool_result"].clear()
    msg = {"type": "tool_result", "call_id": "u1", "success": True, "data": {"x": 1}}
    route_message(msg)
    assert _message_buffer["tool_result"].get("u1") == msg


@pytest.mark.asyncio
async def test_route_message_routes_clarify_answer():
    from ipc import route_message, _message_buffer, _init_buffer
    _init_buffer()
    _message_buffer["clarify_answer"] = None
    route_message({"type": "clarify_answer", "answer": "COUNTA"})
    assert _message_buffer["clarify_answer"] == "COUNTA"


@pytest.mark.asyncio
async def test_route_message_routes_user_message_to_queue():
    from ipc import route_message, _message_buffer, _init_buffer
    _init_buffer()
    # 清空队列
    while not _message_buffer["user_message"].empty():
        _message_buffer["user_message"].get_nowait()
    msg = {"type": "user_message", "text": "统计A列", "session_id": "s1", "context": {}}
    route_message(msg)
    # route_message 用 create_task put，需要 yield 一次让协程执行
    await asyncio.sleep(0.01)
    queued = _message_buffer["user_message"].get_nowait()
    assert queued == msg


@pytest.mark.asyncio
async def test_route_message_routes_cancel_event():
    from ipc import route_message, _message_buffer, _init_buffer
    _init_buffer()
    _message_buffer["cancel"].clear()
    route_message({"type": "cancel"})
    assert _message_buffer["cancel"].is_set()


@pytest.mark.asyncio
async def test_route_message_routes_config_to_queue():
    from ipc import route_message, _message_buffer, _init_buffer
    _init_buffer()
    while not _message_buffer["config"].empty():
        _message_buffer["config"].get_nowait()
    msg = {"type": "config", "base_url": "https://api.anthropic.com", "model": "claude-sonnet-4", "api_key": "sk-x"}
    route_message(msg)
    await asyncio.sleep(0.01)
    queued = _message_buffer["config"].get_nowait()
    assert queued == msg
```

- [ ] **Step 6: 运行测试确认失败**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_ipc.py -v -k route`
Expected: FAIL — `route_message` 未定义

- [ ] **Step 7: 实现 route_message**

```python
# 追加到 src/DeepExcel.Sidecar/ipc.py

def route_message(msg: dict) -> None:
    """把从 stdin 读到的消息按 type 路由到对应缓冲区"""
    _init_buffer()
    t = msg.get("type")
    if t == "tool_result":
        _message_buffer["tool_result"][msg["call_id"]] = msg
    elif t == "clarify_answer":
        _message_buffer["clarify_answer"] = msg.get("answer", "")
    elif t == "user_message":
        # Queue.put 是协程安全的，但 route_message 是同步函数（被 reader 调用）
        # 用 create_task 调度 async put
        asyncio.create_task(_message_buffer["user_message"].put(msg))
    elif t == "cancel":
        _message_buffer["cancel"].set()
    elif t == "config":
        asyncio.create_task(_message_buffer["config"].put(msg))
    # 未知 type 静默丢弃，避免 sidecar 崩溃
```

- [ ] **Step 8: 运行测试确认通过**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_ipc.py -v`
Expected: 所有 route 测试通过

- [ ] **Step 9: 写失败测试 — test_call_csharp_matches_by_call_id**

```python
# 追加到 src/DeepExcel.Sidecar/tests/test_ipc.py

@pytest.mark.asyncio
async def test_call_csharp_sends_tool_call_and_returns_matching_result():
    """call_csharp 应发送 tool_call 消息，并按 call_id 匹配返回结果"""
    from ipc import call_csharp, _message_buffer, _init_buffer, write_message
    import ipc as ipc_module

    _init_buffer()
    _message_buffer["tool_result"].clear()

    # 用一个假的 stdout 缓冲捕获发出的消息
    sent_messages = []
    async def fake_write(msg):
        sent_messages.append(msg)
    ipc_module.write_message = fake_write

    # mock read_message：sleep 0.1s（让 inject 在 0.05s 写完 buffer）后返回非匹配 msg，
    # 让 call_csharp 循环迭代回 buffer 检查命中 inject 的结果
    async def fake_read():
        await asyncio.sleep(0.1)
        return {"type": "user_message", "text": "irrelevant"}
    ipc_module.read_message = fake_read

    # 模拟 C# 延迟回写 tool_result：在 call_csharp 进入循环后，往 buffer 注入结果
    async def inject_result_later():
        await asyncio.sleep(0.05)  # 让 call_csharp 先进入循环
        # 取出 call_csharp 发出的 call_id
        sent = sent_messages[0]
        call_id = sent["call_id"]
        _message_buffer["tool_result"][call_id] = {
            "type": "tool_result",
            "call_id": call_id,
            "success": True,
            "data": {"address": "A1", "formula": "=SUM(B:B)"},
            "error": None,
            "suggestion": None,
        }

    asyncio.create_task(inject_result_later())
    result = await call_csharp("write_formula", {"address": "A1", "formula": "=SUM(B:B)"})

    assert sent_messages[0]["type"] == "tool_call"
    assert sent_messages[0]["tool"] == "write_formula"
    assert result["success"] is True
    assert result["data"]["formula"] == "=SUM(B:B)"


@pytest.mark.asyncio
async def test_call_csharp_clarify_sends_clarify_and_returns_answer():
    from ipc import call_csharp_clarify, _message_buffer, _init_buffer
    import ipc as ipc_module

    _init_buffer()
    _message_buffer["clarify_answer"] = None

    sent_messages = []
    async def fake_write(msg):
        sent_messages.append(msg)
    ipc_module.write_message = fake_write

    # mock read_message：sleep 0.1s 后返回非匹配 msg，让循环迭代回 buffer 检查
    async def fake_read():
        await asyncio.sleep(0.1)
        return {"type": "user_message", "text": "irrelevant"}
    ipc_module.read_message = fake_read

    async def inject_answer_later():
        await asyncio.sleep(0.05)
        _message_buffer["clarify_answer"] = "COUNTA计数"

    asyncio.create_task(inject_answer_later())
    answer = await call_csharp_clarify("A列是文本，求和还是计数？", ["SUM求和", "COUNTA计数"])

    assert sent_messages[0]["type"] == "clarify"
    assert sent_messages[0]["question"] == "A列是文本，求和还是计数？"
    assert answer == "COUNTA计数"
    assert _message_buffer["clarify_answer"] is None  # 已被消费
```

- [ ] **Step 10: 运行测试确认失败**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_ipc.py -v -k call_csharp`
Expected: FAIL — `call_csharp` 未定义

- [ ] **Step 11: 实现 call_csharp 和 call_csharp_clarify**

```python
# 追加到 src/DeepExcel.Sidecar/ipc.py

async def call_csharp(tool_name: str, args: dict) -> dict:
    """向 C# 发送工具调用请求，阻塞等待结果（按 call_id 匹配）"""
    _init_buffer()
    call_id = generate_call_id()
    await write_message({
        "type": "tool_call",
        "call_id": call_id,
        "tool": tool_name,
        "args": args,
    })
    # 循环检查 buffer，未匹配则继续读 stdin
    while True:
        if call_id in _message_buffer["tool_result"]:
            return _message_buffer["tool_result"].pop(call_id)
        # 没有匹配的，读下一条消息并路由
        msg = await read_message()
        if msg is None:
            raise ConnectionError("stdin closed while waiting for tool_result")
        route_message(msg)
        # 路由后再次检查（可能就是我们要的）


async def call_csharp_clarify(question: str, options: list) -> str:
    """向 C# 发送澄清请求，阻塞等待用户回答"""
    _init_buffer()
    _message_buffer["clarify_answer"] = None  # 重置
    await write_message({
        "type": "clarify",
        "question": question,
        "options": options,
    })
    while _message_buffer["clarify_answer"] is None:
        msg = await read_message()
        if msg is None:
            raise ConnectionError("stdin closed while waiting for clarify_answer")
        route_message(msg)
    answer = _message_buffer["clarify_answer"]
    _message_buffer["clarify_answer"] = None
    return answer
```

- [ ] **Step 12: 运行全部 ipc 测试确认通过**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_ipc.py -v`
Expected: 全部通过

- [ ] **Step 13: Commit**

```bash
git add src/DeepExcel.Sidecar/ipc.py src/DeepExcel.Sidecar/tests/test_ipc.py
git commit -m "feat(sidecar): implement ipc layer with call_id-based result matching"
```

---

### Task 3: 实现 system_prompt.py（MVP 版）

**Files:**
- Create: `src/DeepExcel.Sidecar/system_prompt.py`

**Interfaces:**
- Produces: `SYSTEM_PROMPT: str` — Claude Agent SDK 的 system prompt 常量

- [ ] **Step 1: 实现 MVP 版 system_prompt.py**

```python
# src/DeepExcel.Sidecar/system_prompt.py

# MVP 版 system prompt — 阶段四会替换为 spec §5.1 完整版
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

**反问时机（仅在以下情况触发 clarify_intent）：**
- read_range 返回 data_type=mixed（同一列既有数字又有文本）
- write_formula 返回 success=false 且 suggestion 非空
- 用户指令完全无法映射到任何工具
- 用户指令的"目标位置"完全无法推断（如"统计 A 列"但没说写哪里，默认写到 B1 即可，不反问）

## 工具使用决策树

遇到"写公式"类指令：
1. 先调 read_range 看数据类型 → 数字→SUM/AVERAGE；日期→COUNTA；文本→COUNTA
2. 默认目标单元格：源数据列的右侧相邻列第一个单元格（如 A 列数据→写 B1）
3. 公式格式：以 = 开头，英文函数名，全列引用用 A:A，范围用 A1:A100

## 当前 Excel 上下文
{excel_context}
"""
```

- [ ] **Step 2: 验证 import**

Run: `cd src\DeepExcel.Sidecar && python -c "from system_prompt import SYSTEM_PROMPT; print('len=' + str(len(SYSTEM_PROMPT)))"`
Expected: 输出 `len=...`（非零正数）

- [ ] **Step 3: Commit**

```bash
git add src/DeepExcel.Sidecar/system_prompt.py
git commit -m "feat(sidecar): add MVP system prompt with default inference rules"
```

---

### Task 4: 实现 excel_tools.py — echo + write_formula + register scaffold

**Files:**
- Create: `src/DeepExcel.Sidecar/excel_tools.py`
- Create: `src/DeepExcel.Sidecar/tests/test_excel_tools.py`

**Interfaces:**
- Consumes: `ipc.call_csharp(tool_name, args) -> dict`, `ipc.call_csharp_clarify(question, options) -> str`
- Produces:
  - 被 `@tool` 装饰的函数列表（echo, read_range, write_formula, clarify_intent）
  - `def register_all_tools() -> list` — 返回所有 @tool 装饰器返回的工具对象

- [ ] **Step 1: 写失败测试 — test_write_formula_calls_csharp**

```python
# src/DeepExcel.Sidecar/tests/test_excel_tools.py
import asyncio
import json
import pytest
from unittest.mock import patch, AsyncMock


@pytest.mark.asyncio
async def test_write_formula_tool_calls_csharp_and_returns_content():
    """write_formula 工具应通过 call_csharp 调用 C#，并返回 MCP content 格式"""
    from excel_tools import write_formula

    # write_formula 已被 @tool 装饰，需要从 .fn 取出原始协程函数（依赖 SDK API）
    # 如果 SDK 暴露的是 .fn，则用 write_formula.fn；否则直接 await write_formula(...)
    # 这里用 inspect 找到实际可调用对象
    import inspect
    fn = getattr(write_formula, 'fn', None) or write_formula
    if hasattr(fn, '__wrapped__'):
        fn = fn.__wrapped__

    fake_result = {"success": True, "data": {"address": "A1", "formula": "=SUM(B:B)"}}
    with patch('excel_tools.call_csharp', new=AsyncMock(return_value=fake_result)):
        result = await fn({"address": "A1", "formula": "=SUM(B:B)"})

    assert result["content"][0]["type"] == "text"
    parsed = json.loads(result["content"][0]["text"])
    assert parsed["success"] is True
    assert parsed["data"]["formula"] == "=SUM(B:B)"


@pytest.mark.asyncio
async def test_echo_tool_returns_input_as_text():
    from excel_tools import echo
    fn = getattr(echo, 'fn', None) or echo
    if hasattr(fn, '__wrapped__'):
        fn = fn.__wrapped__
    result = await fn({"text": "hello"})
    assert result["content"][0]["text"] == "hello"


@pytest.mark.asyncio
async def test_clarify_intent_calls_csharp_clarify():
    from excel_tools import clarify_intent
    fn = getattr(clarify_intent, 'fn', None) or clarify_intent
    if hasattr(fn, '__wrapped__'):
        fn = fn.__wrapped__
    with patch('excel_tools.call_csharp_clarify', new=AsyncMock(return_value="COUNTA计数")):
        result = await fn({"question": "求和还是计数？", "options": ["SUM", "COUNTA"]})
    assert "COUNTA计数" in result["content"][0]["text"]


def test_register_all_tools_returns_list_with_expected_names():
    from excel_tools import register_all_tools
    tools = register_all_tools()
    # @tool 装饰后的对象可能有 .name 属性或需要查 __name__
    names = []
    for t in tools:
        n = getattr(t, 'name', None) or getattr(t, '__name__', None)
        if n is None:
            # SDK 可能用 .fn.__name__
            fn = getattr(t, 'fn', None)
            n = getattr(fn, '__name__', None) if fn else None
        names.append(n)
    assert "echo" in names
    assert "write_formula" in names
    assert "clarify_intent" in names
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_excel_tools.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'excel_tools'`

- [ ] **Step 3: 实现 excel_tools.py**

```python
# src/DeepExcel.Sidecar/excel_tools.py
import json
from claude_agent_sdk import tool
from ipc import call_csharp, call_csharp_clarify


def _wrap_result(csharp_result: dict) -> dict:
    """把 C# 返回的 dict 包装成 MCP 工具返回格式"""
    return {
        "content": [
            {"type": "text", "text": json.dumps(csharp_result, ensure_ascii=False)}
        ]
    }


@tool("echo", "回声测试工具，原样返回输入文本（用于验证 sidecar 管线）", {"text": str})
async def echo(args):
    return _wrap_result({"success": True, "data": {"echo": args["text"]}})


@tool("read_range", "读取指定范围的单元格数据", {"address": str})
async def read_range(args):
    result = await call_csharp("read_range", {"address": args["address"]})
    return _wrap_result(result)


@tool("write_formula", "向指定单元格写入 Excel 公式", {"address": str, "formula": str})
async def write_formula(args):
    result = await call_csharp("write_formula", {
        "address": args["address"],
        "formula": args["formula"],
    })
    return _wrap_result(result)


@tool("clarify_intent", "向用户提问以澄清模糊指令", {"question": str, "options": list})
async def clarify_intent(args):
    user_answer = await call_csharp_clarify(args["question"], args.get("options", []))
    return _wrap_result({"success": True, "data": {"user_answer": user_answer}})


def register_all_tools() -> list:
    """返回所有 @tool 装饰后的工具对象列表"""
    return [
        echo,
        read_range,
        write_formula,
        clarify_intent,
        # 阶段二补齐其余 10 个工具
    ]
```

- [ ] **Step 4: 运行测试确认通过**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_excel_tools.py -v`
Expected: 4 passed

如果 `@tool` 装饰器的 API 与测试假设不符（例如 `fn` 属性名不同），调整测试中的 `getattr(write_formula, 'fn', None)` 为实际属性名。先在 Python REPL 用 `python -c "from excel_tools import write_formula; print(dir(write_formula))"` 确认。

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.Sidecar/excel_tools.py src/DeepExcel.Sidecar/tests/test_excel_tools.py
git commit -m "feat(sidecar): add echo/read_range/write_formula/clarify_intent tools"
```

---

### Task 5: 实现 sidecar.py 主循环

**Files:**
- Create: `src/DeepExcel.Sidecar/sidecar.py`
- Create: `src/DeepExcel.Sidecar/tests/test_sidecar.py`

**Interfaces:**
- Consumes: `ipc._message_buffer`, `ipc.read_message`, `ipc.route_message`, `ipc.write_message`
- Consumes: `excel_tools.register_all_tools`
- Consumes: `system_prompt.SYSTEM_PROMPT`
- Produces:
  - `async def main() -> None` — 入口协程
  - `async def stdin_reader_loop() -> None` — 唯一 stdin reader
  - `async def handle_sdk_message(response) -> None` — 处理 SDK 流式消息

- [ ] **Step 1: 写失败测试 — test_handle_sdk_message_emits_stream_delta**

```python
# src/DeepExcel.Sidecar/tests/test_sidecar.py
import asyncio
import pytest
from unittest.mock import patch, AsyncMock, MagicMock


@pytest.mark.asyncio
async def test_handle_sdk_message_emits_stream_delta_for_text_block():
    from sidecar import handle_sdk_message
    import sidecar as sidecar_module

    sent = []
    async def fake_write(msg):
        sent.append(msg)
    sidecar_module.write_message = fake_write

    # 构造假 AssistantMessage
    fake_text_block = MagicMock()
    fake_text_block.__class__.__name__ = "TextBlock"
    # 用 isinstance 检查，需要真类型；这里用 monkeypatch 替换 isinstance 行为更复杂
    # 简化：直接构造一个有 .text 属性的对象，并 patch sidecar 模块的 TextBlock 引用
    fake_msg = MagicMock()
    fake_msg.content = [fake_text_block]

    # patch sidecar 内部 import 的 TextBlock
    with patch('sidecar.TextBlock', new=MagicMock(side_effect=lambda x: x if isinstance(x, MagicMock) else None)):
        # 让 isinstance(text_block, TextBlock) 返回 True
        import sys
        sidecar_module.TextBlock = type('TextBlock', (), {})
        fake_text_block = sidecar_module.TextBlock()
        fake_text_block.text = "正在读取A列..."
        fake_msg = MagicMock()
        fake_msg.content = [fake_text_block]
        # patch isinstance 不可行，改为 patch AssistantMessage 检查
        with patch('sidecar.AssistantMessage', new=MagicMock()):
            sidecar_module.AssistantMessage = type('AssistantMessage', (), {})
            fake_msg.__class__ = sidecar_module.AssistantMessage
            await handle_sdk_message(fake_msg)

    assert any(m["type"] == "stream_delta" and m["text"] == "正在读取A列..." for m in sent)


@pytest.mark.asyncio
async def test_handle_sdk_message_emits_stream_end_for_result_message():
    from sidecar import handle_sdk_message
    import sidecar as sidecar_module

    sent = []
    async def fake_write(msg):
        sent.append(msg)
    sidecar_module.write_message = fake_write

    sidecar_module.ResultMessage = type('ResultMessage', (), {})
    fake_msg = MagicMock()
    fake_msg.__class__ = sidecar_module.ResultMessage
    fake_msg.usage = MagicMock(input_tokens=100, output_tokens=50)

    await handle_sdk_message(fake_msg)

    assert any(m["type"] == "stream_end" and m["input_tokens"] == 100 and m["output_tokens"] == 50 for m in sent)
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_sidecar.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'sidecar'`

- [ ] **Step 3: 实现 sidecar.py**

```python
# src/DeepExcel.Sidecar/sidecar.py
import asyncio
import json
import os
import sys

from claude_agent_sdk import (
    ClaudeAgentOptions,
    ClaudeSDKClient,
    create_sdk_mcp_server,
)
from claude_agent_sdk.types import AssistantMessage, ResultMessage, TextBlock, ToolUseBlock

from excel_tools import register_all_tools
from ipc import _message_buffer, read_message, route_message, write_message
from ipc import _init_buffer
from system_prompt import SYSTEM_PROMPT


async def stdin_reader_loop():
    """唯一的 stdin 读取协程，把消息分发到 _message_buffer"""
    while True:
        msg = await read_message()
        if msg is None:
            # stdin 关闭，往 user_message queue 发 None 让主循环退出
            await _message_buffer["user_message"].put(None)
            return
        route_message(msg)
        # cancel 消息也通过 Event 触发（route_message 已处理）


async def handle_sdk_message(response):
    """处理 ClaudeSDKClient.receive_response() 产生的流式消息"""
    if isinstance(response, AssistantMessage):
        for block in response.content:
            if isinstance(block, TextBlock):
                await write_message({"type": "stream_delta", "text": block.text})
            elif isinstance(block, ToolUseBlock):
                # SDK 内部已调度 @tool 函数，工具函数会通过 IPC 调 C#
                # 这里不需要额外处理；可选：发一个 tool_call 通知给 C# 用于 UI 显示
                await write_message({
                    "type": "tool_call",
                    "call_id": getattr(block, "id", None) or getattr(block, "tool_use_id", ""),
                    "tool": block.name,
                    "args": block.input if isinstance(block.input, dict) else {},
                })
    elif isinstance(response, ResultMessage):
        usage = getattr(response, "usage", None)
        in_tok = getattr(usage, "input_tokens", 0) if usage else 0
        out_tok = getattr(usage, "output_tokens", 0) if usage else 0
        await write_message({
            "type": "stream_end",
            "input_tokens": in_tok,
            "output_tokens": out_tok,
        })


async def main():
    _init_buffer()

    # 注册工具到 MCP server
    server = create_sdk_mcp_server(name="excel", tools=register_all_tools())

    options = ClaudeAgentOptions(
        mcp_servers={"excel": server},
        allowed_tools=[f"mcp__excel__{t}" for t in [
            "echo", "read_range", "write_formula", "clarify_intent",
            # 阶段二补齐其余 10 个
        ]],
        system_prompt=SYSTEM_PROMPT,
        max_turns=20,
    )

    client = ClaudeSDKClient(options=options)
    await client.connect()

    # 启动 stdin reader 协程（唯一读 stdin 的地方）
    asyncio.create_task(stdin_reader_loop())

    # 主循环
    while True:
        msg = await _message_buffer["user_message"].get()
        if msg is None:
            break

        # 应用 pending config（不阻塞主循环）
        while not _message_buffer["config"].empty():
            try:
                cfg = _message_buffer["config"].get_nowait()
                os.environ["ANTHROPIC_BASE_URL"] = cfg.get("base_url", "https://api.anthropic.com")
                os.environ["ANTHROPIC_MODEL"] = cfg.get("model", "claude-sonnet-4")
                os.environ["ANTHROPIC_API_KEY"] = cfg.get("api_key", "")
            except Exception:
                pass

        # 发送用户消息
        try:
            await client.query(msg["text"])
            async for response in client.receive_response():
                await handle_sdk_message(response)
        except asyncio.CancelledError:
            await write_message({"type": "stream_delta", "text": "⏹ 已停止"})
            await write_message({"type": "stream_end", "input_tokens": 0, "output_tokens": 0})
        except Exception as e:
            await write_message({"type": "stream_delta", "text": f"❌ 错误: {e}"})
            await write_message({"type": "stream_end", "input_tokens": 0, "output_tokens": 0})

        # 清理 cancel 标志
        if _message_buffer["cancel"].is_set():
            _message_buffer["cancel"].clear()


if __name__ == "__main__":
    asyncio.run(main())
```

- [ ] **Step 4: 运行测试确认通过**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_sidecar.py -v`
Expected: 2 passed

如果 `claude_agent_sdk.types` 的 import 路径不对，调整为 SDK 实际暴露的路径（在 Task 1 已验证）。

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.Sidecar/sidecar.py src/DeepExcel.Sidecar/tests/test_sidecar.py
git commit -m "feat(sidecar): implement main loop with ClaudeSDKClient and stdin reader"
```

---

### Task 6: Sidecar E2E 测试 — mock C# 驱动脚本

**Files:**
- Create: `src/DeepExcel.Sidecar/tests/test_e2e_mock.py`
- Create: `src/DeepExcel.Sidecar/tests/mock_csharp_driver.py`

**Interfaces:**
- Produces: 集成测试，验证 sidecar 启动 → 收 user_message → 发 tool_call → 收 tool_result → 发 stream_end 完整环路

- [ ] **Step 1: 实现 mock C# 驱动脚本**

```python
# src/DeepExcel.Sidecar/tests/mock_csharp_driver.py
"""
模拟 C# 端与 sidecar 通信的驱动脚本。
用法：在测试中通过 subprocess 启动 sidecar.py，本脚本作为参考实现。

实际测试 test_e2e_mock.py 会内联这个逻辑。
"""
import asyncio
import json
import sys


async def drive_sidecar(stdin_writer, stdout_reader):
    """
    模拟 C# 行为：
    1. 发 user_message
    2. 读 sidecar 输出，遇到 tool_call 就回 tool_result
    3. 收到 stream_end 结束
    """
    # 1. 发 user_message
    stdin_writer.write((json.dumps({
        "type": "user_message",
        "text": "在 A1 写入 =SUM(B1:B10)",
        "session_id": "test-1",
        "context": {},
    }) + "\n").encode("utf-8"))
    await stdin_writer.drain()

    # 2. 读 sidecar 输出
    received_messages = []
    while True:
        line = await stdout_reader.readline()
        if not line:
            break
        msg = json.loads(line.decode("utf-8"))
        received_messages.append(msg)

        if msg["type"] == "tool_call":
            # 回写 tool_result
            tool_result = {
                "type": "tool_result",
                "call_id": msg["call_id"],
                "success": True,
                "data": {"address": "A1", "formula": "=SUM(B1:B10)"},
                "error": None,
                "suggestion": None,
                "context": {},
            }
            stdin_writer.write((json.dumps(tool_result) + "\n").encode("utf-8"))
            await stdin_writer.drain()

        if msg["type"] == "stream_end":
            break

    return received_messages
```

- [ ] **Step 2: 写 E2E 测试**

```python
# src/DeepExcel.Sidecar/tests/test_e2e_mock.py
"""
端到端测试：启动真实 sidecar.py 子进程，模拟 C# 行为验证完整环路。

注意：此测试需要真实 ANTHROPIC_API_KEY 才能跑通 SDK 调用。
未设置 API key 时跳过，标记为 integration test。
"""
import asyncio
import json
import os
import sys
import pytest


pytestmark = pytest.mark.skipif(
    not os.environ.get("ANTHROPIC_API_KEY"),
    reason="需要 ANTHROPIC_API_KEY 才能跑 sidecar E2E 测试",
)


@pytest.mark.asyncio
async def test_sidecar_e2e_write_formula():
    """完整环路：user_message → tool_call(write_formula) → tool_result → stream_end"""
    sidecar_path = os.path.join(os.path.dirname(__file__), "..", "sidecar.py")

    proc = await asyncio.create_subprocess_exec(
        sys.executable, sidecar_path,
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )

    try:
        # 1. 发 user_message
        user_msg = json.dumps({
            "type": "user_message",
            "text": "在 A1 写入 =SUM(B1:B10)",
            "session_id": "test-1",
            "context": {},
        }) + "\n"
        proc.stdin.write(user_msg.encode("utf-8"))
        await proc.stdin.drain()

        # 2. 读 sidecar 输出，遇 tool_call 回 tool_result
        received = []
        while True:
            line = await asyncio.wait_for(proc.stdout.readline(), timeout=30.0)
            if not line:
                break
            msg = json.loads(line.decode("utf-8"))
            received.append(msg)

            if msg["type"] == "tool_call":
                tool_result = json.dumps({
                    "type": "tool_result",
                    "call_id": msg["call_id"],
                    "success": True,
                    "data": {"address": "A1", "formula": "=SUM(B1:B10)"},
                    "error": None,
                    "suggestion": None,
                    "context": {},
                }) + "\n"
                proc.stdin.write(tool_result.encode("utf-8"))
                await proc.stdin.drain()

            if msg["type"] == "stream_end":
                break

        # 断言：至少收到一个 tool_call 和一个 stream_end
        types = [m["type"] for m in received]
        assert "tool_call" in types, f"未收到 tool_call，收到: {types}"
        assert "stream_end" in types, f"未收到 stream_end，收到: {types}"

    finally:
        try:
            proc.terminate()
            await asyncio.wait_for(proc.wait(), timeout=5.0)
        except Exception:
            proc.kill()
```

- [ ] **Step 3: 运行 E2E 测试（无 API key 时自动跳过）**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_e2e_mock.py -v`
Expected: SKIPPED（如果没有 ANTHROPIC_API_KEY）或 PASSED（如果有 key 且 SDK 工作正常）

- [ ] **Step 4: Commit**

```bash
git add src/DeepExcel.Sidecar/tests/test_e2e_mock.py src/DeepExcel.Sidecar/tests/mock_csharp_driver.py
git commit -m "test(sidecar): add e2e mock test for write_formula pipeline"
```

---

### Task 7: C# SidecarProtocol.cs + 扩展 ToolResult

**Files:**
- Create: `src/DeepExcel.AddIn/Sidecar/SidecarProtocol.cs`
- Modify: `src/DeepExcel.AddIn/Bridge/Messages.cs:37-50` — 给 ToolResult 加 Suggestion + Context 字段

**Interfaces:**
- Produces:
  - `SidecarProtocol` 静态类 — 消息 type 常量
  - `ToolResult.Suggestion` (string) — 失败再问的提示
  - `ToolResult.Context` (object) — Excel 上下文快照

- [ ] **Step 1: 写失败测试 — test ToolResult Suggestion 字段**

```csharp
// src/DeepExcel.Tests/SidecarProtocolTests.cs
using DeepExcel.AddIn.Bridge;
using Xunit;

namespace DeepExcel.Tests
{
    public class SidecarProtocolTests
    {
        [Fact]
        public void ToolResult_HasSuggestionAndContextFields()
        {
            var tr = new ToolResult
            {
                Name = "write_formula",
                Success = false,
                Error = "type mismatch",
                Suggestion = "目标区域包含文本，无法求和。建议改用 COUNTA。",
                Context = new { active_sheet = "Sheet1" },
            };
            Assert.NotNull(tr.Suggestion);
            Assert.NotNull(tr.Context);
        }

        [Fact]
        public void SidecarProtocol_TypeConstants_MatchSpec()
        {
            // C# → Python
            Assert.Equal("user_message", SidecarProtocol.TypeUserMessage);
            Assert.Equal("cancel", SidecarProtocol.TypeCancel);
            Assert.Equal("tool_result", SidecarProtocol.TypeToolResult);
            Assert.Equal("config", SidecarProtocol.TypeConfig);
            Assert.Equal("clarify_answer", SidecarProtocol.TypeClarifyAnswer);

            // Python → C#
            Assert.Equal("stream_delta", SidecarProtocol.TypeStreamDelta);
            Assert.Equal("tool_call", SidecarProtocol.TypeToolCall);
            Assert.Equal("clarify", SidecarProtocol.TypeClarify);
            Assert.Equal("stream_end", SidecarProtocol.TypeStreamEnd);
        }
    }
}
```

- [ ] **Step 2: 运行测试确认失败**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter SidecarProtocolTests`
Expected: FAIL — `SidecarProtocol` 未定义；`ToolResult.Suggestion` 未定义

注意：DeepExcel.Tests 用 SDK 风格 csproj，可用 `dotnet test`。如果环境只有 csc.exe 没有 dotnet CLI，则改为手动编译运行 xunit 测试（参考已有 UnitTests.cs 的运行方式）。

- [ ] **Step 3: 实现 SidecarProtocol.cs**

```csharp
// src/DeepExcel.AddIn/Sidecar/SidecarProtocol.cs
namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// Sidecar IPC 消息 type 常量（C# ↔ Python 双向）
    /// 协议：每行一个 JSON 对象，以 \n 分隔
    /// </summary>
    public static class SidecarProtocol
    {
        // C# → Python
        public const string TypeUserMessage = "user_message";
        public const string TypeCancel = "cancel";
        public const string TypeToolResult = "tool_result";
        public const string TypeConfig = "config";
        public const string TypeClarifyAnswer = "clarify_answer";

        // Python → C#
        public const string TypeStreamDelta = "stream_delta";
        public const string TypeToolCall = "tool_call";
        public const string TypeClarify = "clarify";
        public const string TypeStreamEnd = "stream_end";
    }
}
```

- [ ] **Step 4: 扩展 Messages.cs 的 ToolResult**

用 Edit 工具修改 [src/DeepExcel.AddIn/Bridge/Messages.cs#L37-L50](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/Messages.cs#L37-L50)：

把原 ToolResult 类替换为：

```csharp
    /// <summary>
    /// 工具执行结果（回传给UI/Sidecar）
    /// </summary>
    public class ToolResult
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("data")]
        public object Data { get; set; }

        [JsonPropertyName("error")]
        public string Error { get; set; }

        [JsonPropertyName("suggestion")]
        public string Suggestion { get; set; }   // 失败再问提示；非空时触发模型反问

        [JsonPropertyName("context")]
        public object Context { get; set; }      // Excel 上下文快照
    }
```

- [ ] **Step 5: 运行测试确认通过**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter SidecarProtocolTests`
Expected: 2 passed

- [ ] **Step 6: Commit**

```bash
git add src/DeepExcel.AddIn/Sidecar/SidecarProtocol.cs src/DeepExcel.AddIn/Bridge/Messages.cs src/DeepExcel.Tests/SidecarProtocolTests.cs
git commit -m "feat(addin): add SidecarProtocol constants and extend ToolResult with suggestion/context"
```

---

### Task 8: PythonSidecar.cs — Process 生命周期 + Send 方法 + 消息路由

**Files:**
- Create: `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs`
- Create: `src/DeepExcel.Tests/PythonSidecarTests.cs`

**Interfaces:**
- Consumes: `IExcelActions`（通过构造注入，工具执行时调用）
- Consumes: `Control _uiControl`（用于 BeginInvoke/InvokeAsync 切回 STA）
- Produces:
  - `class PythonSidecar : IDisposable`
  - `void Start()` — 启动 Python 子进程
  - `void Stop()` — 终止子进程
  - `void SendUserMessage(string text, string sessionId, object context)`
  - `void SendCancel()`
  - `void UpdateConfig(string baseUrl, string model, string apiKey)`
  - `void SendClarifyAnswer(string answer)`
  - `void SendToolResult(string callId, bool success, object data, string error, string suggestion, object context)`
  - 事件: `OnStreamDelta`, `OnToolCall`, `OnClarify`, `OnStreamEnd`

- [ ] **Step 1: 写失败测试 — PythonSidecar 可实例化、可发送消息**

```csharp
// src/DeepExcel.Tests/PythonSidecarTests.cs
using System;
using System.Windows.Forms;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Moq;
using Xunit;

namespace DeepExcel.Tests
{
    public class PythonSidecarTests
    {
        [Fact]
        public void PythonSidecar_CanBeConstructedWithDeps()
        {
            var excelActions = new Mock<IExcelActions>().Object;
            var uiControl = new UserControl();  // 真实 Control 用于 Invoke
            var sidecar = new PythonSidecar(excelActions, uiControl);
            Assert.NotNull(sidecar);
            sidecar.Dispose();
        }

        [Fact]
        public void PythonSidecar_GetSidecarPath_ReturnsValidPath()
        {
            // 静态方法测试，不依赖进程
            var path = PythonSidecar.GetSidecarPath();
            Assert.Contains("sidecar.py", path);
        }

        [Fact]
        public void PythonSidecar_GetPythonPath_ReturnsNonEmpty()
        {
            var py = PythonSidecar.GetPythonPath();
            Assert.False(string.IsNullOrEmpty(py));
        }
    }
}
```

需要给 DeepExcel.Tests 加 Moq 依赖。如果不想加，改为手写 mock 实现 `IExcelActions`。

- [ ] **Step 2: 运行测试确认失败**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter PythonSidecarTests`
Expected: FAIL — `PythonSidecar` 未定义

- [ ] **Step 3: 实现 PythonSidecar.cs（生命周期 + Send + 路由部分）**

```csharp
// src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// Python sidecar 子进程包装：JSON Lines over stdin/stdout 通信
    /// </summary>
    public class PythonSidecar : IDisposable
    {
        private Process _process;
        private readonly IExcelActions _excel;
        private readonly Control _uiControl;
        private readonly ToolDispatcher _dispatcher;
        private readonly object _writeLock = new object();

        // 事件：UI 线程触发
        public event Action<string> OnStreamDelta;
        public event Action<string, string, Dictionary<string, object>> OnToolCall;
        public event Action<string, List<string>> OnClarify;
        public event Action<int, int> OnStreamEnd;
        public event Action<string> OnError;

        public PythonSidecar(IExcelActions excel, Control uiControl)
        {
            _excel = excel;
            _uiControl = uiControl;
            _dispatcher = new ToolDispatcher(excel);
        }

        /// <summary>
        /// 启动 Python sidecar 子进程
        /// </summary>
        public void Start()
        {
            var psi = new ProcessStartInfo
            {
                FileName = GetPythonPath(),
                Arguments = $"\"{GetSidecarPath()}\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };

            // 透传环境变量（ANTHROPIC_API_KEY 等）
            // psi.EnvironmentVariables 已自动继承当前进程

            _process = Process.Start(psi);
            _process.OutputDataReceived += OnStdoutLine;
            _process.ErrorDataReceived += OnStderrLine;
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();

            Logger.Instance.Info("PythonSidecar", "Started, pid=" + _process.Id);
        }

        /// <summary>
        /// 停止 sidecar 子进程
        /// </summary>
        public void Stop()
        {
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    try { WriteLine(@"{""type"":""cancel""}"); } catch { }
                    _process.WaitForExit(2000);
                    if (!_process.HasExited) _process.Kill();
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("PythonSidecar", "Stop failed", ex);
            }
        }

        public void Dispose() => Stop();

        // ============= 发送消息（C# → Python）=============

        public void SendUserMessage(string text, string sessionId, object context)
        {
            var msg = new { type = SidecarProtocol.TypeUserMessage, text, session_id = sessionId, context };
            WriteLine(JsonSerializer.Serialize(msg));
        }

        public void SendCancel() => WriteLine(@"{""type"":""cancel""}");

        public void UpdateConfig(string baseUrl, string model, string apiKey)
        {
            var msg = new { type = SidecarProtocol.TypeConfig, base_url = baseUrl, model, api_key = apiKey };
            WriteLine(JsonSerializer.Serialize(msg));
        }

        public void SendClarifyAnswer(string answer)
        {
            var msg = new { type = SidecarProtocol.TypeClarifyAnswer, answer };
            WriteLine(JsonSerializer.Serialize(msg));
        }

        public void SendToolResult(string callId, bool success, object data, string error, string suggestion, object context)
        {
            var msg = new
            {
                type = SidecarProtocol.TypeToolResult,
                call_id = callId,
                success,
                data,
                error,
                suggestion,
                context,
            };
            WriteLine(JsonSerializer.Serialize(msg));
        }

        // ============= 接收消息（Python → C#）=============

        private void OnStdoutLine(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Data)) return;

            try
            {
                using var doc = JsonDocument.Parse(e.Data);
                var root = doc.RootElement;
                var type = root.GetProperty("type").GetString();

                switch (type)
                {
                    case SidecarProtocol.TypeStreamDelta:
                        var text = root.GetProperty("text").GetString();
                        _uiControl.BeginInvoke(new Action(() => OnStreamDelta?.Invoke(text)));
                        break;

                    case SidecarProtocol.TypeToolCall:
                        HandleToolCall(root);
                        break;

                    case SidecarProtocol.TypeClarify:
                        var q = root.GetProperty("question").GetString();
                        var opts = root.GetProperty("options").EnumerateArray()
                            .Select(x => x.GetString()).ToList();
                        _uiControl.BeginInvoke(new Action(() => OnClarify?.Invoke(q, opts)));
                        break;

                    case SidecarProtocol.TypeStreamEnd:
                        var inTok = root.GetProperty("input_tokens").GetInt32();
                        var outTok = root.GetProperty("output_tokens").GetInt32();
                        _uiControl.BeginInvoke(new Action(() => OnStreamEnd?.Invoke(inTok, outTok)));
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("PythonSidecar", "Parse stdout failed: " + e.Data, ex);
            }
        }

        private void OnStderrLine(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Data)) return;
            Logger.Instance.Warning("PythonSidecar", "stderr: " + e.Data);
            _uiControl.BeginInvoke(new Action(() => OnError?.Invoke(e.Data)));
        }

        private async void HandleToolCall(JsonElement msg)
        {
            var callId = msg.GetProperty("call_id").GetString();
            var toolName = msg.GetProperty("tool").GetString();
            var args = new Dictionary<string, object>();
            if (msg.TryGetProperty("args", out var argsEl))
            {
                foreach (var prop in argsEl.EnumerateObject())
                {
                    args[prop.Name] = prop.Value.Clone();
                }
            }

            // 通知 UI（fire-and-forget）
            _uiControl.BeginInvoke(new Action(() => OnToolCall?.Invoke(callId, toolName, args)));

            // ★ 关键：切回 UI 线程执行 Excel COM 操作（STA 要求）
            ToolResult result = null;
            try
            {
                if (_uiControl.InvokeRequired)
                {
                    result = (ToolResult)_uiControl.Invoke(new Func<ToolResult>(() =>
                        _dispatcher.Execute(toolName, args)));
                }
                else
                {
                    result = _dispatcher.Execute(toolName, args);
                }
            }
            catch (Exception ex)
            {
                result = new ToolResult { Name = toolName, Success = false, Error = ex.Message };
                Logger.Instance.Error("PythonSidecar", "HandleToolCall failed: " + toolName, ex);
            }

            // 回写 tool_result
            SendToolResult(
                callId: callId,
                success: result.Success,
                data: result.Data,
                error: result.Error,
                suggestion: result.Suggestion,
                context: _dispatcher.BuildExcelSnapshot());
        }

        // ============= 工具方法 =============

        private void WriteLine(string json)
        {
            lock (_writeLock)
            {
                if (_process == null || _process.HasExited) return;
                _process.StandardInput.WriteLine(json);
                _process.StandardInput.Flush();
            }
        }

        /// <summary>
        /// 查找 Python 解释器路径
        /// 优先级：1) 内嵌 python/python.exe  2) 系统 PATH 中的 python
        /// </summary>
        public static string GetPythonPath()
        {
            // 1. 内嵌 Python（打包后）
            var addInDir = Path.GetDirectoryName(typeof(PythonSidecar).Assembly.Location);
            var embeddedPy = Path.Combine(addInDir, "python", "python.exe");
            if (File.Exists(embeddedPy)) return embeddedPy;

            // 2. 系统 PATH（开发期）
            var systemPy = Environment.GetEnvironmentVariable("DEEPEXCEL_PYTHON_PATH");
            if (!string.IsNullOrEmpty(systemPy) && File.Exists(systemPy)) return systemPy;

            // 3. 默认 python（在 PATH 中）
            return "python";
        }

        /// <summary>
        /// sidecar.py 路径
        /// </summary>
        public static string GetSidecarPath()
        {
            var addInDir = Path.GetDirectoryName(typeof(PythonSidecar).Assembly.Location);
            return Path.Combine(addInDir, "sidecar", "sidecar.py");
        }
    }
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter PythonSidecarTests`
Expected: 3 passed

如果 DeepExcel.Tests 缺 Moq，改为手写 mock：

```csharp
public class FakeExcelActions : IExcelActions
{
    public object GetSelection() => null;
    public object ReadRange(string address) => new { };
    public object ReadWorkbook() => new { };
    public object ReadWorksheet(string name) => new { };
    public ToolResult ExecuteVBA(string code, string macroName = null) => new ToolResult { Success = true };
    public ToolResult ExecutePython(string code) => new ToolResult { Success = true };
    public ToolResult WriteFormula(string address, string formula) => new ToolResult { Success = true };
    public ToolResult WriteValue(string address, object value) => new ToolResult { Success = true };
    public string CreateSnapshot() => "snap-1";
    public bool Rollback(string snapshotId) => true;
}
```

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs src/DeepExcel.Tests/PythonSidecarTests.cs
git commit -m "feat(addin): implement PythonSidecar with process lifecycle and message routing"
```

---

### Task 9: ToolDispatcher.cs — 工具执行 + GetArg + BuildExcelSnapshot（MVP 版）

**Files:**
- Create: `src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs`
- Create: `src/DeepExcel.Tests/ToolDispatcherTests.cs`

**Interfaces:**
- Consumes: `IExcelActions`
- Produces:
  - `class ToolDispatcher`
  - `ToolResult Execute(string toolName, Dictionary<string, object> args)` — 同步执行（已在 STA 线程）
  - `object BuildExcelSnapshot()` — 返回当前 Excel 上下文快照
  - 私有 `T GetArg<T>(args, key)` 和 `string[] GetArgArray(args, key)` — 从 Dictionary（含 JsonElement）取参数

MVP 阶段只实现 3 个工具：`read_range`, `write_formula`, `echo`。其余在 Task 15 补齐。

- [ ] **Step 1: 写失败测试 — ToolDispatcher 写公式**

```csharp
// src/DeepExcel.Tests/ToolDispatcherTests.cs
using System.Collections.Generic;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Sidecar;
using Xunit;
using Moq;

namespace DeepExcel.Tests
{
    public class ToolDispatcherTests
    {
        [Fact]
        public void Execute_WriteFormula_CallsExcelActions()
        {
            var mock = new Mock<IExcelActions>();
            mock.Setup(x => x.WriteFormula("A1", "=SUM(B:B)"))
                .Returns(new ToolResult { Name = "write_formula", Success = true });

            var dispatcher = new ToolDispatcher(mock.Object);
            var args = new Dictionary<string, object>
            {
                { "address", "A1" },
                { "formula", "=SUM(B:B)" },
            };

            var result = dispatcher.Execute("write_formula", args);

            Assert.True(result.Success);
            mock.Verify(x => x.WriteFormula("A1", "=SUM(B:B)"), Times.Once);
        }

        [Fact]
        public void Execute_ReadRange_ReturnsDataAndSuggestion()
        {
            var mock = new Mock<IExcelActions>();
            mock.Setup(x => x.ReadRange("A:A"))
                .Returns(new { cells = new[] { "苹果", "香蕉" }, data_type = "text" });

            var dispatcher = new ToolDispatcher(mock.Object);
            var args = new Dictionary<string, object> { { "address", "A:A" } };

            var result = dispatcher.Execute("read_range", args);

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
            Assert.NotNull(result.Suggestion);  // text 列应触发"无法求和"提示
        }

        [Fact]
        public void Execute_Echo_ReturnsInputText()
        {
            var mock = new Mock<IExcelActions>();
            var dispatcher = new ToolDispatcher(mock.Object);
            var args = new Dictionary<string, object> { { "text", "hello" } };

            var result = dispatcher.Execute("echo", args);

            Assert.True(result.Success);
            Assert.NotNull(result.Data);
        }

        [Fact]
        public void Execute_UnknownTool_ReturnsError()
        {
            var mock = new Mock<IExcelActions>();
            var dispatcher = new ToolDispatcher(mock.Object);
            var result = dispatcher.Execute("nonexistent", new Dictionary<string, object>());
            Assert.False(result.Success);
            Assert.Contains("未知工具", result.Error);
        }
    }
}
```

- [ ] **Step 2: 运行测试确认失败**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter ToolDispatcherTests`
Expected: FAIL — `ToolDispatcher` 未定义

- [ ] **Step 3: 实现 ToolDispatcher.cs（MVP 版，3 个工具）**

```csharp
// src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Bridge;
using DeepExcel.AddIn.Diagnostics;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// 工具执行调度器：接收 sidecar 的 tool_call，在 STA 线程执行 Excel 操作
    /// 从 Orchestrator.ExecuteToolAsync 移植，保留 GetArg/GetArgArray 的 JsonElement 处理
    /// </summary>
    public class ToolDispatcher
    {
        private readonly IExcelActions _excel;

        public ToolDispatcher(IExcelActions excel)
        {
            _excel = excel;
        }

        /// <summary>
        /// 同步执行工具（必须在 STA 主线程调用）
        /// </summary>
        public ToolResult Execute(string toolName, Dictionary<string, object> args)
        {
            Logger.Instance.Info("ToolDispatcher", "Execute: " + toolName + ", args keys=" + (args == null ? "null" : string.Join(",", args.Keys)));

            try
            {
                switch (toolName)
                {
                    case "echo":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = new { echo = GetArg<string>(args, "text") },
                        };

                    case "read_range":
                        var address = GetArg<string>(args, "address");
                        var rangeData = _excel.ReadRange(address);
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = rangeData,
                            Suggestion = GenerateRangeSuggestion(rangeData),
                            Context = BuildExcelSnapshot(),
                        };

                    case "write_formula":
                        var cellAddress = GetArg<string>(args, "address");
                        var formula = GetArg<string>(args, "formula");
                        Logger.Instance.Info("ToolDispatcher", "write_formula: address=[" + (cellAddress ?? "null") + "], formula=[" + (formula ?? "null") + "]");
                        var result = _excel.WriteFormula(cellAddress, formula);
                        if (!result.Success)
                        {
                            result.Suggestion = GenerateFormulaSuggestion(formula, result.Error);
                        }
                        return result;

                    default:
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = false,
                            Error = $"未知工具: {toolName}",
                        };
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("ToolDispatcher", "Execute failed: " + toolName, ex);
                return new ToolResult
                {
                    Name = toolName,
                    Success = false,
                    Error = ex.Message,
                };
            }
        }

        /// <summary>
        /// 构建当前 Excel 上下文快照（附在 tool_result.context）
        /// </summary>
        public object BuildExcelSnapshot()
        {
            try
            {
                var wb = _excel.ReadWorkbook();
                var sel = _excel.GetSelection();
                return new
                {
                    workbook = wb,
                    selection = sel,
                    timestamp = DateTime.Now.ToString("o"),
                };
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "BuildExcelSnapshot failed", ex);
                return new { error = ex.Message };
            }
        }

        /// <summary>
        /// read_range 返回数据的类型检测：检测到 text/mixed 时返回澄清提示
        /// </summary>
        private string GenerateRangeSuggestion(object rangeData)
        {
            try
            {
                // rangeData 是 RangeInfo/匿名对象，序列化为 JSON 再解析以提取 data_type
                var json = JsonSerializer.Serialize(rangeData);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.TryGetProperty("data_type", out var dtEl))
                {
                    var dt = dtEl.GetString();
                    if (dt == "text")
                        return "该列是文本，无法直接求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
                    if (dt == "mixed")
                        return "该列同时包含数字和文本，无法直接求和。建议：求和(忽略文本)、计数(全部 COUNTA)、计数(仅数字 COUNT)";
                    if (dt == "date")
                        return "该列是日期，求和无意义。建议用 COUNTA 计数，或按日期分组。";
                }
                else if (root.TryGetProperty("cells", out var cellsEl))
                {
                    // 兜底：如果没有 data_type，自己检测
                    bool hasNum = false, hasText = false;
                    foreach (var cell in cellsEl.EnumerateArray())
                    {
                        if (cell.ValueKind == JsonValueKind.Number) hasNum = true;
                        else if (cell.ValueKind == JsonValueKind.String) hasText = true;
                    }
                    if (hasNum && hasText)
                        return "该列同时包含数字和文本，无法直接求和。建议：求和(忽略文本)、计数(全部 COUNTA)、计数(仅数字 COUNT)";
                    if (hasText && !hasNum)
                        return "该列是文本，无法直接求和。建议改用 COUNTA 计数。";
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", "GenerateRangeSuggestion failed", ex);
            }
            return null;
        }

        /// <summary>
        /// write_formula 失败时根据公式和错误生成澄清提示
        /// </summary>
        private string GenerateFormulaSuggestion(string formula, string error)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(error)) return null;

            if (formula.Contains("SUM") && (error.Contains("类型") || error.Contains("type") || error.Contains("mismatch")))
                return "目标区域包含文本，无法求和。建议改用 COUNTA 计数，或确认是否要忽略文本。";
            if (formula.Contains("AVERAGE") && (error.Contains("除数为零") || error.Contains("zero") || error.Contains("divisor")))
                return "目标区域没有数字，无法计算平均值。建议检查数据源。";
            if (error.Contains("范围") || error.Contains("range"))
                return "目标范围无效。请检查地址格式，例如 A1、B1:B10、A:A。";
            return null;
        }

        // ============= 参数提取（从 Orchestrator 移植）=============

        protected internal T GetArg<T>(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null) return default;
            try
            {
                var val = args[key];

                if (val is JsonElement je)
                {
                    if (typeof(T) == typeof(string))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetString();
                    }
                    if (typeof(T) == typeof(int))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetInt32();
                    }
                    if (typeof(T) == typeof(bool))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetBoolean();
                    }
                    if (typeof(T) == typeof(double))
                    {
                        if (je.ValueKind == JsonValueKind.Null) return default;
                        return (T)(object)je.GetDouble();
                    }
                    return (T)Convert.ChangeType(je.GetRawText(), typeof(T));
                }

                return (T)Convert.ChangeType(val, typeof(T));
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("ToolDispatcher", $"GetArg<{typeof(T).Name}>(\"{key}\") failed: " + ex.Message);
                return default;
            }
        }

        protected internal string[] GetArgArray(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key)) return new string[0];
            try
            {
                if (args[key] is JsonElement je)
                {
                    var list = new List<string>();
                    foreach (var item in je.EnumerateArray())
                    {
                        list.Add(item.GetString());
                    }
                    return list.ToArray();
                }
                if (args[key] is object[] arr)
                {
                    return Array.ConvertAll(arr, x => x?.ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Debug("ToolDispatcher", $"GetArgArray(\"{key}\") failed: " + ex.Message);
            }
            return new string[0];
        }
    }
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter ToolDispatcherTests`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs src/DeepExcel.Tests/ToolDispatcherTests.cs
git commit -m "feat(addin): implement ToolDispatcher with echo/read_range/write_formula + suggestion generators"
```

---

### Task 10: MessageBridge.cs — 切换到 PythonSidecar

**Files:**
- Modify: `src/DeepExcel.AddIn/Bridge/MessageBridge.cs:1-213`

**Interfaces:**
- Consumes: `PythonSidecar`（替换原 `Orchestrator`）
- 保留: `IExcelActions`, `ExcelActionsImpl`（工具实现层不变）
- 保留: `Cancel()`, `SetSendToUi()`, `HandleMessage()` 对外 API

- [ ] **Step 1: 修改 MessageBridge 构造函数和字段**

用 Edit 工具修改 [src/DeepExcel.AddIn/Bridge/MessageBridge.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs#L20-L45)：

替换字段和构造函数为：

```csharp
    public class MessageBridge
    {
        private readonly Application _excelApp;
        private readonly IExcelActions _excelActions;
        private readonly PythonSidecar _sidecar;
        private Action<string> _sendToUi;
        private string _pendingClarifyQuestion;  // 待回答的澄清问题

        public MessageBridge(Application excelApp, Control uiControl)
        {
            _excelApp = excelApp;

            // 初始化各层（保留原 ExcelActionsImpl 创建逻辑）
            var workbookAnalyzer = new WorkbookAnalyzer(excelApp);
            var rangeAnalyzer = new RangeAnalyzer();
            var snapshotManager = new SnapshotManager(excelApp);
            var vbaExecutor = new VBAExecutor(excelApp, snapshotManager);
            var pythonExecutor = new Executor.PythonExecutor(excelApp, snapshotManager);

            _excelActions = new ExcelActionsImpl(
                excelApp, workbookAnalyzer, rangeAnalyzer, vbaExecutor, pythonExecutor, snapshotManager);

            // 启动 Python sidecar（替换 Orchestrator）
            _sidecar = new PythonSidecar(_excelActions, uiControl);
            _sidecar.OnStreamDelta += OnStreamDelta;
            _sidecar.OnToolCall += OnToolCall;
            _sidecar.OnClarify += OnClarify;
            _sidecar.OnStreamEnd += OnStreamEndFromSidecar;
            _sidecar.OnError += OnSidecarError;
            _sidecar.Start();
        }
```

需要加 `using DeepExcel.AddIn.Sidecar;` 和 `using System.Windows.Forms;`。

- [ ] **Step 2: 修改 HandleUserMessage 走 sidecar**

替换 HandleUserMessage 方法：

```csharp
        private string HandleUserMessage(Message msg)
        {
            try
            {
                var content = msg.Payload?.GetProperty("content").GetString();

                // 如果有待回答的 clarify，把这条消息当 clarify_answer
                if (_pendingClarifyQuestion != null)
                {
                    _sidecar.SendClarifyAnswer(content);
                    _pendingClarifyQuestion = null;
                    return MakeResponse("ack", new { received = true, kind = "clarify_answer" });
                }

                // 正常用户消息：附带 Excel 上下文
                var context = BuildContext();
                var sessionId = Guid.NewGuid().ToString();
                _sidecar.SendUserMessage(content, sessionId, context);

                return MakeResponse("ack", new { received = true });
            }
            catch (Exception ex)
            {
                return MakeError($"User message error: {ex.Message}");
            }
        }

        private object BuildContext()
        {
            try
            {
                return new
                {
                    workbook = _excelActions.ReadWorkbook(),
                    selection = _excelActions.GetSelection(),
                };
            }
            catch
            {
                return new { };
            }
        }
```

- [ ] **Step 3: 修改 Cancel 方法**

```csharp
        public void Cancel()
        {
            try
            {
                _sidecar?.SendCancel();
                Logger.Instance.Info("MessageBridge", "Cancel requested");
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "Cancel failed", ex);
            }
        }
```

- [ ] **Step 4: 修改回调（OnStreamDelta / OnToolCall / OnClarify / OnStreamEnd）**

替换 OnStreamDelta / OnToolCall / OnToolResult 区段为：

```csharp
        private void OnStreamDelta(string delta)
        {
            Logger.Instance.Debug("MessageBridge", "OnStreamDelta: " + (delta == null ? "null" : ("len=" + delta.Length)));
            SendToUi("stream_delta", new { delta });
        }

        private void OnToolCall(string callId, string name, Dictionary<string, object> args)
        {
            SendToUi("tool_call", new { call_id = callId, name, arguments = args });
        }

        private void OnClarify(string question, List<string> options)
        {
            _pendingClarifyQuestion = question;  // 标记下次用户消息作为 clarify_answer
            SendToUi("clarify", new { question, options });
        }

        private void OnStreamEndFromSidecar(int inputTokens, int outputTokens)
        {
            SendToUi("stream_end", new { input_tokens = inputTokens, output_tokens = outputTokens });
        }

        private void OnSidecarError(string error)
        {
            SendToUi("error", new { message = "Sidecar: " + error });
        }
```

- [ ] **Step 5: 修改 HandleExecuteTool — 移除对 Orchestrator 的依赖**

把 HandleExecuteTool 改为直接用 ToolDispatcher：

```csharp
        private string HandleExecuteTool(Message msg)
        {
            var name = msg.Payload?.GetProperty("name").GetString();
            var argsEl = msg.Payload?.GetProperty("arguments");
            var args = new Dictionary<string, object>();
            if (argsEl.HasValue)
            {
                foreach (var prop in argsEl.Value.EnumerateObject())
                {
                    args[prop.Name] = prop.Value.Clone();
                }
            }
            var dispatcher = new ToolDispatcher(_excelActions);
            var result = dispatcher.Execute(name, args);
            return MakeResponse("tool_result", result);
        }
```

- [ ] **Step 6: 移除 NotifyStreamEnd 调用（sidecar 自己发 stream_end）**

在 HandleUserMessage 里不再需要 Task.Run + NotifyStreamEnd 的兜底，因为 sidecar 的 OnStreamEndFromSidecar 会触发。

但保留 NotifyStreamEnd 方法供 Cancel 使用：

```csharp
        public void NotifyStreamEnd()
        {
            SendToUi("stream_end", new { });
        }
```

Cancel 的 case 里也保留 NotifyStreamEnd：

```csharp
                    case "cancel":
                        Cancel();
                        NotifyStreamEnd();
                        return MakeResponse("cancelled", new { });
```

- [ ] **Step 7: 编译验证**

Run: `cd c:\Users\qinju\Desktop\AIProject\DeepExcel && scripts\build-csc.bat`
Expected: 编译成功

注意：此时 build-csc.bat 还没把新文件加进去，所以会编译失败。先在下一步 Task 11 更新 build-csc.bat 再统一编译。

- [ ] **Step 8: Commit（暂不编译，等 Task 11 一起验证）**

```bash
git add src/DeepExcel.AddIn/Bridge/MessageBridge.cs
git commit -m "feat(addin): switch MessageBridge from Orchestrator to PythonSidecar"
```

---

### Task 11: ThisAddIn.cs — sidecar 生命周期 + 更新 build-csc.bat

**Files:**
- Modify: `src/DeepExcel.AddIn/ThisAddIn.cs:145-158` — 在 OnStartupComplete 后启动 sidecar，OnDisconnection 时停止
- Modify: `scripts/build-csc.bat` — 加入新文件，移除暂未删的旧文件（暂保留 Orchestrator 等）
- Modify: `src/DeepExcel.AddIn/DeepExcel.AddIn.csproj` — 同步加新文件

- [ ] **Step 1: 修改 ThisAddIn.OnStartupComplete — 把 uiControl 传给 MessageBridge**

在 [src/DeepExcel.AddIn/ThisAddIn.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs#L145-L158) 的 OnStartupComplete 中：

```csharp
        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete called");
            try
            {
                // 获取 TaskPaneControl（含 WebView2）作为 uiControl，用于 STA 线程切换
                Control uiControl = _taskPane ?? (_fallbackForm as Control);
                _bridge = new MessageBridge(_excelApp, uiControl);
                Log("MessageBridge initialized with sidecar");
            }
            catch (Exception ex)
            {
                Log("MessageBridge init FAILED: " + ex.GetType().Name + " - " + ex.Message);
                Log("Stack: " + ex.StackTrace);
            }
        }
```

如果 `_taskPane` 在 OnStartupComplete 时还未创建（依赖 CTPFactory），改为延迟初始化：在 TaskPaneControl 创建后调用 `_bridge.AttachUiControl(_taskPane)`。先在 MessageBridge 加：

```csharp
        public void AttachUiControl(Control uiControl)
        {
            // 重新创建 sidecar 关联的 uiControl（如果初始为 null）
            // MVP 简化：如果 sidecar 已启动但 uiControl 是 null，重启 sidecar
            // 此方法留作扩展点，MVP 阶段假设 uiControl 在构造时可用
        }
```

实际 MVP 阶段：如果 `_taskPane` 此时为 null，先用 `_fallbackForm`（如果存在）或一个新建的隐藏 Form 作为 uiControl。后续真正创建 TaskPane 后，sidecar 的 OnStreamDelta 等已经用 BeginInvoke 调度到正确的线程。

最简方案：用 `new Form()` 作为 uiControl 占位（不显示）。

- [ ] **Step 2: 修改 ThisAddIn.OnDisconnection — 停止 sidecar**

在 OnDisconnection 中加：

```csharp
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            Log("OnDisconnection called, mode=" + disconnectMode);
            try
            {
                // MessageBridge.Dispose 会停止 sidecar
                if (_bridge is IDisposable disposable) disposable.Dispose();
                _bridge = null;

                if (_customTaskPane != null)
                {
                    _customTaskPane.Delete();
                    _customTaskPane = null;
                }
                // ... 其余清理保留
            }
            catch (Exception ex) { Log("OnDisconnection cleanup error: " + ex.Message); }
        }
```

需要让 MessageBridge 实现 IDisposable：

```csharp
    public class MessageBridge : IDisposable
    {
        // ... 已有代码
        public void Dispose()
        {
            try { _sidecar?.Stop(); } catch { }
        }
    }
```

- [ ] **Step 3: 更新 scripts/build-csc.bat — 加入新文件**

在 build-csc.bat 的 csc 命令中（line 78-116），加入新 .cs 文件，并移除已删的 Orchestrator 等：

把原本的源文件列表段（line 78-116）替换为：

```bat
  "%ADDINDIR%\Advanced\ChartSpecEngine.cs" ^
  "%ADDINDIR%\Bridge\IExcelActions.cs" ^
  "%ADDINDIR%\Bridge\MessageBridge.cs" ^
  "%ADDINDIR%\Bridge\Messages.cs" ^
  "%ADDINDIR%\Collaboration\OperationHistory.cs" ^
  "%ADDINDIR%\Config\AppConfig.cs" ^
  "%ADDINDIR%\Diagnostics\Logger.cs" ^
  "%ADDINDIR%\Executor\ExecutionResult.cs" ^
  "%ADDINDIR%\Executor\PythonExecutor.cs" ^
  "%ADDINDIR%\Executor\SnapshotManager.cs" ^
  "%ADDINDIR%\Executor\VBAExecutor.cs" ^
  "%ADDINDIR%\Models\ModelTypes.cs" ^
  "%ADDINDIR%\Performance\PerformanceOptimizer.cs" ^
  "%ADDINDIR%\Perception\RangeAnalyzer.cs" ^
  "%ADDINDIR%\Perception\RangeInfo.cs" ^
  "%ADDINDIR%\Perception\WorkbookAnalyzer.cs" ^
  "%ADDINDIR%\Perception\WorkbookStructure.cs" ^
  "%ADDINDIR%\Properties\AssemblyInfo.cs" ^
  "%ADDINDIR%\Security\SecurityGateway.cs" ^
  "%ADDINDIR%\Security\SecurityManager.cs" ^
  "%ADDINDIR%\Sidecar\PythonSidecar.cs" ^
  "%ADDINDIR%\Sidecar\SidecarProtocol.cs" ^
  "%ADDINDIR%\Sidecar\ToolDispatcher.cs" ^
  "%ADDINDIR%\TaskPaneControl.cs" ^
  "%ADDINDIR%\IRibbonCallbacks.cs" ^
  "%ADDINDIR%\ThisAddIn.cs" ^
  "%ADDINDIR%\Tools\ChartTool.cs" ^
  "%ADDINDIR%\Tools\DataCleaner.cs" ^
  "%ADDINDIR%\Tools\FormulaTool.cs" ^
  "%ADDINDIR%\Extensions\DictionaryExtensions.cs"
```

注意：
- 加入了 `Sidecar\PythonSidecar.cs`, `Sidecar\SidecarProtocol.cs`, `Sidecar\ToolDispatcher.cs`
- 暂时**保留** `Agent\Orchestrator.cs`, `Agent\ToolRegistry.cs`, `Models\DeepSeekAdapter.cs` 等（下一步 Task 19 删除）
- 实际 MVP 阶段为了编译通过，先保留 Orchestrator 等文件（虽然不再被 MessageBridge 使用，但 ToolDispatcher 可能用到其辅助方法）。Task 19 会清理。

为了让编译先通过，build-csc.bat 仍然包含 `Agent\Orchestrator.cs`, `Agent\ToolRegistry.cs`, `Agent\PromptBuilder.cs`, `Agent\ActionPlanner.cs`, `Models\*.cs`。所以实际改动是**只新增** Sidecar/*.cs 三个文件，**不移除**任何旧文件。

修正后的 build-csc.bat 源文件列表 = 原 list + 3 个新文件：

```bat
  "%ADDINDIR%\Agent\ActionPlanner.cs" ^
  "%ADDINDIR%\Agent\Orchestrator.cs" ^
  "%ADDINDIR%\Agent\PromptBuilder.cs" ^
  "%ADDINDIR%\Agent\ToolRegistry.cs" ^
  "%ADDINDIR%\Advanced\ChartSpecEngine.cs" ^
  "%ADDINDIR%\Bridge\IExcelActions.cs" ^
  "%ADDINDIR%\Bridge\MessageBridge.cs" ^
  "%ADDINDIR%\Bridge\Messages.cs" ^
  "%ADDINDIR%\Collaboration\OperationHistory.cs" ^
  "%ADDINDIR%\Config\AppConfig.cs" ^
  "%ADDINDIR%\Diagnostics\Logger.cs" ^
  "%ADDINDIR%\Executor\ExecutionResult.cs" ^
  "%ADDINDIR%\Executor\PythonExecutor.cs" ^
  "%ADDINDIR%\Executor\SnapshotManager.cs" ^
  "%ADDINDIR%\Executor\VBAExecutor.cs" ^
  "%ADDINDIR%\Models\ClaudeAdapter.cs" ^
  "%ADDINDIR%\Models\DeepSeekAdapter.cs" ^
  "%ADDINDIR%\Models\IModelAdapter.cs" ^
  "%ADDINDIR%\Models\ModelAdapterFactory.cs" ^
  "%ADDINDIR%\Models\ModelFactory.cs" ^
  "%ADDINDIR%\Models\ModelTypes.cs" ^
  "%ADDINDIR%\Models\OpenAIAdapter.cs" ^
  "%ADDINDIR%\Models\OpenAICompatibleAdapter.cs" ^
  "%ADDINDIR%\Models\RetryHelper.cs" ^
  "%ADDINDIR%\Performance\PerformanceOptimizer.cs" ^
  "%ADDINDIR%\Perception\RangeAnalyzer.cs" ^
  "%ADDINDIR%\Perception\RangeInfo.cs" ^
  "%ADDINDIR%\Perception\WorkbookAnalyzer.cs" ^
  "%ADDINDIR%\Perception\WorkbookStructure.cs" ^
  "%ADDINDIR%\Properties\AssemblyInfo.cs" ^
  "%ADDINDIR%\Security\SecurityGateway.cs" ^
  "%ADDINDIR%\Security\SecurityManager.cs" ^
  "%ADDINDIR%\Sidecar\PythonSidecar.cs" ^
  "%ADDINDIR%\Sidecar\SidecarProtocol.cs" ^
  "%ADDINDIR%\Sidecar\ToolDispatcher.cs" ^
  "%ADDINDIR%\TaskPaneControl.cs" ^
  "%ADDINDIR%\IRibbonCallbacks.cs" ^
  "%ADDINDIR%\ThisAddIn.cs" ^
  "%ADDINDIR%\Tools\ChartTool.cs" ^
  "%ADDINDIR%\Tools\DataCleaner.cs" ^
  "%ADDINDIR%\Tools\FormulaTool.cs" ^
  "%ADDINDIR%\Extensions\DictionaryExtensions.cs"
```

同时把 sidecar/ 目录复制到输出目录，加在 [2/3] WebView assets 后：

```bat
REM Copy sidecar scripts
echo [2.5/3] Copying sidecar scripts...
if not exist "%OUTDIR%\sidecar" mkdir "%OUTDIR%\sidecar"
xcopy /E /I /Q "%BASEDIR%\src\DeepExcel.Sidecar\*.py" "%OUTDIR%\sidecar" >nul
```

- [ ] **Step 4: 同步更新 DeepExcel.AddIn.csproj**

在 [src/DeepExcel.AddIn/DeepExcel.AddIn.csproj](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/DeepExcel.AddIn.csproj#L75-L113) 的 ItemGroup 里加：

```xml
    <Compile Include="Sidecar\PythonSidecar.cs" />
    <Compile Include="Sidecar\SidecarProtocol.cs" />
    <Compile Include="Sidecar\ToolDispatcher.cs" />
```

- [ ] **Step 5: 编译验证**

Run: `cd c:\Users\qinju\Desktop\AIProject\DeepExcel && scripts\build-csc.bat`
Expected: `=== Build SUCCESSFUL ===`

如果失败，根据 csc 错误信息修复。常见问题：
- 缺少 `using DeepExcel.AddIn.Sidecar;` — 在 MessageBridge.cs / ThisAddIn.cs 顶部加
- `Control` 类型未找到 — 在 MessageBridge.cs 顶部加 `using System.Windows.Forms;`
- ToolResult 字段冲突 — 检查 Messages.cs 的修改是否完整

- [ ] **Step 6: 部署 sidecar 到输出目录**

Run: `xcopy /E /I /Y "c:\Users\qinju\Desktop\AIProject\DeepExcel\src\DeepExcel.Sidecar\*.py" "c:\Users\qinju\Desktop\AIProject\DeepExcel\src\DeepExcel.AddIn\bin\Release\sidecar\"`
Expected: 复制 sidecar.py, excel_tools.py, ipc.py, system_prompt.py 到 bin/Release/sidecar/

- [ ] **Step 7: Commit**

```bash
git add src/DeepExcel.AddIn/ThisAddIn.cs src/DeepExcel.AddIn/DeepExcel.AddIn.csproj scripts/build-csc.bat
git commit -m "feat(addin): wire sidecar lifecycle into ThisAddIn and update build scripts"
```

---

### Task 12: 端到端 MVP 烟测 — Excel 内验证 write_formula

**Files:**
- 无新文件（手动测试 + 验证）

**Interfaces:**
- 验证：用户在 Excel 输入"在 A1 写入 =SUM(B1:B10)" → sidecar → Claude → tool_call → C# 写公式 → 返回 → 前端显示结果

- [ ] **Step 1: 注册新构建的 DLL**

Run: `cd c:\Users\qinju\Desktop\AIProject\DeepExcel && powershell -ExecutionPolicy Bypass -File scripts\register.ps1`
Expected: 注册成功（具体输出依赖 register.ps1）

- [ ] **Step 2: 启动 Excel，打开 DeepExcel 任务窗格**

手动：打开 Excel → DeepExcel 加载项 → 显示任务窗格
预期：任务窗格正常显示，无加载错误

检查 `bin\Release\DeepExcel_Load.log` 应包含 `MessageBridge initialized with sidecar`。

- [ ] **Step 3: 设置 API Key 环境变量**

确保系统环境变量 `ANTHROPIC_API_KEY` 已设置，或在 register.ps1 启动 Excel 时透传。

如果未设置，临时方案：在 PythonSidecar.Start() 中从 model-config.json 读取并设置 psi.EnvironmentVariables["ANTHROPIC_API_KEY"]。先用手动方式验证，自动化放 Task 22。

- [ ] **Step 4: 准备测试数据**

在 Excel 的 Sheet1 B1:B10 填入数字 1-10。

- [ ] **Step 5: 输入测试指令**

在 DeepExcel 任务窗格输入：`在 A1 写入 =SUM(B1:B10)`

预期流程：
1. 任务窗格显示流式输出（模型 thinking 简短分析）
2. 显示 `tool_call: write_formula(address=A1, formula==SUM(B1:B10))`
3. Excel 的 A1 单元格被写入公式，显示 55
4. 任务窗格显示"已在 A1 写入 =SUM(B1:B10)，结果为 55"
5. 显示 stream_end，loading 状态消失

- [ ] **Step 6: 验证日志**

检查 `bin\Release\DeepExcel_Load.log` 应包含：
- `PythonSidecar: Started, pid=...`
- `ToolDispatcher: Execute: write_formula, args keys=address,formula`
- `ToolDispatcher: write_formula: address=[A1], formula=[=SUM(B1:B10)]`

检查 sidecar stderr（通过 OnSidecarError 推送到前端）应无致命错误。

- [ ] **Step 7: 测试取消按钮**

输入一个会触发长时间思考的指令，点击红色 Stop 按钮，预期：
- 任务窗格显示 `⏹ 已停止`
- loading 状态消失

- [ ] **Step 8: Commit 测试记录**

无代码改动，跳过 commit。如果发现 bug，新建 fix commit。

---

## 阶段二：完整工具迁移

---

### Task 13: 实现所有 14 个工具的 @tool 函数

**Files:**
- Modify: `src/DeepExcel.Sidecar/excel_tools.py` — 补齐剩余 10 个工具
- Modify: `src/DeepExcel.Sidecar/sidecar.py` — 把 allowed_tools 列表补全
- Modify: `src/DeepExcel.Sidecar/tests/test_excel_tools.py` — 补齐测试

**Interfaces:**
- Produces: 14 个 @tool 装饰函数，全部通过 `call_csharp` 调用 C# 侧 ToolDispatcher

- [ ] **Step 1: 补齐 excel_tools.py 的工具定义**

修改 [src/DeepExcel.Sidecar/excel_tools.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/excel_tools.py)，在已有 echo/read_range/write_formula/clarify_intent 基础上加：

```python
@tool("read_workbook", "读取当前工作簿的结构信息", {})
async def read_workbook(args):
    result = await call_csharp("read_workbook", {})
    return _wrap_result(result)


@tool("read_selection", "读取当前选中的单元格信息", {})
async def read_selection(args):
    result = await call_csharp("read_selection", {})
    return _wrap_result(result)


@tool("fill_formula_down", "将公式向下填充到指定行数", {"from_address": str, "row_count": int})
async def fill_formula_down(args):
    result = await call_csharp("fill_formula_down", {
        "from_address": args["from_address"],
        "row_count": args["row_count"],
    })
    return _wrap_result(result)


@tool("replace_formula", "在指定范围内批量替换公式中的字符串", {"range_address": str, "find": str, "replace": str})
async def replace_formula(args):
    result = await call_csharp("replace_formula", {
        "range_address": args["range_address"],
        "find": args["find"],
        "replace": args["replace"],
    })
    return _wrap_result(result)


@tool("clean_data", "执行数据清洗操作（unify_date/remove_duplicates/highlight_missing/trim_spaces/text_to_number）", {"range_address": str, "operations": list})
async def clean_data(args):
    result = await call_csharp("clean_data", {
        "range_address": args["range_address"],
        "operations": args["operations"],
    })
    return _wrap_result(result)


@tool("create_chart", "基于数据范围创建图表", {"data_range": str, "chart_type": str, "title": str, "x_label": str, "y_label": str})
async def create_chart(args):
    result = await call_csharp("create_chart", {
        "data_range": args["data_range"],
        "chart_type": args.get("chart_type", "column"),
        "title": args.get("title", ""),
        "x_label": args.get("x_label", ""),
        "y_label": args.get("y_label", ""),
    })
    return _wrap_result(result)


@tool("create_pivot_table", "基于数据范围创建数据透视表", {"source_range": str, "destination_sheet": str, "pivot_table_name": str, "row_fields": list, "column_fields": list, "value_fields": list, "value_function": str})
async def create_pivot_table(args):
    result = await call_csharp("create_pivot_table", {
        "source_range": args["source_range"],
        "destination_sheet": args["destination_sheet"],
        "pivot_table_name": args.get("pivot_table_name", "PivotTable1"),
        "row_fields": args.get("row_fields", []),
        "column_fields": args.get("column_fields", []),
        "value_fields": args.get("value_fields", []),
        "value_function": args.get("value_function", "Sum"),
    })
    return _wrap_result(result)


@tool("execute_vba", "执行 VBA 代码（Sub DeepExcel_TempMacro）", {"code": str})
async def execute_vba(args):
    result = await call_csharp("execute_vba", {"code": args["code"]})
    return _wrap_result(result)


@tool("execute_python", "执行 Python 代码（操作 Excel）", {"code": str})
async def execute_python(args):
    result = await call_csharp("execute_python", {"code": args["code"]})
    return _wrap_result(result)


@tool("create_snapshot", "创建当前工作簿快照（用于回滚）", {})
async def create_snapshot(args):
    result = await call_csharp("create_snapshot", {})
    return _wrap_result(result)


@tool("rollback", "回滚到指定快照", {"snapshot_id": str})
async def rollback(args):
    result = await call_csharp("rollback", {"snapshot_id": args["snapshot_id"]})
    return _wrap_result(result)
```

更新 `register_all_tools()` 返回完整列表：

```python
def register_all_tools() -> list:
    return [
        echo,
        read_workbook, read_selection, read_range,
        write_formula, fill_formula_down, replace_formula,
        clean_data, create_chart, create_pivot_table,
        execute_vba, execute_python,
        create_snapshot, rollback,
        clarify_intent,
    ]
```

- [ ] **Step 2: 更新 sidecar.py 的 allowed_tools 列表**

修改 [src/DeepExcel.Sidecar/sidecar.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py) 中 main() 的 allowed_tools：

```python
        allowed_tools=[f"mcp__excel__{t}" for t in [
            "echo",
            "read_workbook", "read_selection", "read_range",
            "write_formula", "fill_formula_down", "replace_formula",
            "clean_data", "create_chart", "create_pivot_table",
            "execute_vba", "execute_python",
            "create_snapshot", "rollback",
            "clarify_intent",
        ]],
```

- [ ] **Step 3: 补充测试 — 每个 @tool 至少一个 happy path 测试**

追加到 [src/DeepExcel.Sidecar/tests/test_excel_tools.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/tests/test_excel_tools.py)：

```python
@pytest.mark.asyncio
async def test_all_14_tools_registered():
    from excel_tools import register_all_tools
    tools = register_all_tools()
    # 14 + echo = 15 个
    assert len(tools) >= 14


@pytest.mark.asyncio
async def test_create_chart_passes_all_args():
    from excel_tools import create_chart
    import excel_tools as et
    fn = getattr(create_chart, 'fn', None) or create_chart
    if hasattr(fn, '__wrapped__'): fn = fn.__wrapped__

    fake_result = {"success": True, "data": {}}
    with patch('excel_tools.call_csharp', new=AsyncMock(return_value=fake_result)) as mock:
        await fn({"data_range": "A1:B10", "chart_type": "bar", "title": "T", "x_label": "X", "y_label": "Y"})
        call_args = mock.call_args[0]
        assert call_args[0] == "create_chart"
        assert call_args[1]["chart_type"] == "bar"


@pytest.mark.asyncio
async def test_clean_data_passes_operations_list():
    from excel_tools import clean_data
    fn = getattr(clean_data, 'fn', None) or clean_data
    if hasattr(fn, '__wrapped__'): fn = fn.__wrapped__

    with patch('excel_tools.call_csharp', new=AsyncMock(return_value={"success": True})) as mock:
        await fn({"range_address": "A1:A100", "operations": ["trim_spaces", "remove_duplicates"]})
        call_args = mock.call_args[0]
        assert call_args[1]["operations"] == ["trim_spaces", "remove_duplicates"]
```

- [ ] **Step 4: 运行测试**

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/test_excel_tools.py -v`
Expected: 全部通过

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.Sidecar/excel_tools.py src/DeepExcel.Sidecar/sidecar.py src/DeepExcel.Sidecar/tests/test_excel_tools.py
git commit -m "feat(sidecar): implement all 14 tools (read/write/clean/chart/pivot/vba/python/snapshot/rollback/clarify)"
```

---

### Task 14: ToolDispatcher.cs — 补齐剩余 11 个工具 case

**Files:**
- Modify: `src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs` — 补齐 case
- Modify: `src/DeepExcel.Tests/ToolDispatcherTests.cs` — 补齐测试

**Interfaces:**
- Consumes: `ExcelActionsImpl` 已实现的 `ExecuteVBA`, `ExecutePython`, `WriteFormula`, `WriteValue`, `CreateSnapshot`, `Rollback`
- 需要 `Application _excelApp` 用于 fill_formula_down / replace_formula / clean_data / create_chart / create_pivot_table（移植自 Orchestrator）
- 注入 `Application` 到 ToolDispatcher 构造函数

- [ ] **Step 1: 修改 ToolDispatcher 构造函数 — 加 Application**

修改 [src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs)：

```csharp
    public class ToolDispatcher
    {
        private readonly IExcelActions _excel;
        private readonly Microsoft.Office.Interop.Excel.Application _excelApp;

        public ToolDispatcher(IExcelActions excel, Microsoft.Office.Interop.Excel.Application excelApp = null)
        {
            _excel = excel;
            _excelApp = excelApp;
        }
```

- [ ] **Step 2: 在 Execute 方法的 switch 中补齐所有 case**

在 `case "write_formula":` 后追加（保留原 echo / read_range / write_formula）：

```csharp
                    case "read_workbook":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = _excel.ReadWorkbook(),
                            Context = BuildExcelSnapshot(),
                        };

                    case "read_selection":
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = true,
                            Data = _excel.GetSelection(),
                            Context = BuildExcelSnapshot(),
                        };

                    case "fill_formula_down":
                        var fromAddr = GetArg<string>(args, "from_address");
                        var rowCount = GetArg<int>(args, "row_count");
                        return ExecuteFillDown(fromAddr, rowCount);

                    case "replace_formula":
                        var rangeAddr = GetArg<string>(args, "range_address");
                        var find = GetArg<string>(args, "find");
                        var replace = GetArg<string>(args, "replace");
                        return ExecuteReplaceFormula(rangeAddr, find, replace);

                    case "clean_data":
                        var cleanRange = GetArg<string>(args, "range_address");
                        var ops = GetArgArray(args, "operations");
                        return ExecuteCleanData(cleanRange, ops);

                    case "create_chart":
                        var dataRange = GetArg<string>(args, "data_range");
                        var chartType = GetArg<string>(args, "chart_type") ?? "column";
                        var chartTitle = GetArg<string>(args, "title") ?? "";
                        var xLabel = GetArg<string>(args, "x_label") ?? "";
                        var yLabel = GetArg<string>(args, "y_label") ?? "";
                        return ExecuteCreateChart(dataRange, chartType, chartTitle, xLabel, yLabel);

                    case "create_pivot_table":
                        var sourceRange = GetArg<string>(args, "source_range");
                        var destSheet = GetArg<string>(args, "destination_sheet");
                        var pivotName = GetArg<string>(args, "pivot_table_name") ?? "PivotTable1";
                        var rowFields = GetArgArray(args, "row_fields");
                        var colFields = GetArgArray(args, "column_fields");
                        var valFields = GetArgArray(args, "value_fields");
                        var valFunc = GetArg<string>(args, "value_function") ?? "Sum";
                        return ExecuteCreatePivot(sourceRange, destSheet, pivotName, rowFields, colFields, valFields, valFunc);

                    case "execute_vba":
                        var code = GetArg<string>(args, "code");
                        return _excel.ExecuteVBA(code);

                    case "execute_python":
                        var pyCode = GetArg<string>(args, "code");
                        return _excel.ExecutePython(pyCode);

                    case "create_snapshot":
                        var snapshotId = _excel.CreateSnapshot();
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = !string.IsNullOrEmpty(snapshotId),
                            Data = new { snapshot_id = snapshotId },
                        };

                    case "rollback":
                        var sid = GetArg<string>(args, "snapshot_id");
                        var success = _excel.Rollback(sid);
                        return new ToolResult
                        {
                            Name = toolName,
                            Success = success,
                        };
```

- [ ] **Step 3: 移植 ExecuteFillDown / ExecuteReplaceFormula / ExecuteCleanData / ExecuteCreateChart / ExecuteCreatePivot**

从 [src/DeepExcel.AddIn/Agent/Orchestrator.cs#L483-L605](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Agent/Orchestrator.cs#L483-L605) 把这 5 个私有方法原样复制到 ToolDispatcher.cs。改 `_excelApp` 引用为 `this._excelApp`。

```csharp
        private ToolResult ExecuteFillDown(string fromAddress, int rowCount)
        {
            try
            {
                var app = _excelApp;
                var fromRange = app.Range[fromAddress];
                var toRange = fromRange.Resize[rowCount, fromRange.Columns.Count];
                fromRange.AutoFill(toRange, Microsoft.Office.Interop.Excel.XlAutoFillType.xlFillDefault);
                return new ToolResult { Name = "fill_formula_down", Success = true };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "fill_formula_down", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteReplaceFormula(string rangeAddress, string find, string replace)
        {
            try
            {
                var app = _excelApp;
                var range = app.Range[rangeAddress];
                int count = 0;
                foreach (Microsoft.Office.Interop.Excel.Range cell in range.Cells)
                {
                    if ((bool)cell.HasFormula)
                    {
                        var formula = cell.Formula.ToString();
                        if (formula.Contains(find))
                        {
                            cell.Formula = formula.Replace(find, replace);
                            count++;
                        }
                    }
                }
                return new ToolResult { Name = "replace_formula", Success = true, Data = new { replacedCount = count } };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "replace_formula", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCleanData(string rangeAddress, string[] operations)
        {
            try
            {
                var cleaner = new Tools.DataCleaner(_excelApp);
                var results = new List<string>();
                foreach (var op in operations)
                {
                    switch (op.ToLower())
                    {
                        case "unify_date":
                            var r1 = cleaner.UnifyDateFormats(rangeAddress);
                            results.Add($"unify_date: {(r1.Success ? (r1.Data as dynamic)?.convertedCount + "个" : "失败-" + r1.Error)}");
                            break;
                        case "remove_duplicates":
                            var r2 = cleaner.RemoveDuplicates(rangeAddress);
                            results.Add($"remove_duplicates: {(r2.Success ? (r2.Data as dynamic)?.removedCount + "行" : "失败-" + r2.Error)}");
                            break;
                        case "highlight_missing":
                            var r3 = cleaner.HighlightMissingValues(rangeAddress);
                            results.Add($"highlight_missing: {(r3.Success ? (r3.Data as dynamic)?.highlightedCount + "个" : "失败-" + r3.Error)}");
                            break;
                        case "trim_spaces":
                            var r4 = cleaner.TrimSpaces(rangeAddress);
                            results.Add($"trim_spaces: {(r4.Success ? (r4.Data as dynamic)?.trimmedCount + "个" : "失败-" + r4.Error)}");
                            break;
                        case "text_to_number":
                            var r5 = cleaner.TextToNumber(rangeAddress);
                            results.Add($"text_to_number: {(r5.Success ? (r5.Data as dynamic)?.convertedCount + "个" : "失败-" + r5.Error)}");
                            break;
                    }
                }
                return new ToolResult { Name = "clean_data", Success = true, Data = new { operations = results } };
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "clean_data", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCreateChart(string dataRange, string chartType, string title, string xLabel, string yLabel)
        {
            try
            {
                var tool = new Tools.ChartTool(_excelApp);
                return tool.CreateChart(dataRange, chartType, title, xLabel, yLabel);
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "create_chart", Success = false, Error = ex.Message };
            }
        }

        private ToolResult ExecuteCreatePivot(
            string sourceRange, string destSheet, string pivotName,
            string[] rowFields, string[] colFields, string[] valFields, string valFunc)
        {
            try
            {
                var tool = new Tools.ChartTool(_excelApp);
                return tool.CreatePivotTable(sourceRange, destSheet, pivotName,
                    rowFields?.Length > 0 ? rowFields : null,
                    colFields?.Length > 0 ? colFields : null,
                    valFields?.Length > 0 ? valFields : null,
                    valFunc);
            }
            catch (Exception ex)
            {
                return new ToolResult { Name = "create_pivot_table", Success = false, Error = ex.Message };
            }
        }
```

加 `using DeepExcel.AddIn.Tools;` 到 ToolDispatcher.cs 顶部。

- [ ] **Step 4: 更新 PythonSidecar 创建 ToolDispatcher 时传 Application**

修改 [src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs) 构造函数：

```csharp
    public class PythonSidecar : IDisposable
    {
        public PythonSidecar(IExcelActions excel, Microsoft.Office.Interop.Excel.Application excelApp, Control uiControl)
        {
            _excel = excel;
            _uiControl = uiControl;
            _dispatcher = new ToolDispatcher(excel, excelApp);
        }
```

同步修改 MessageBridge 构造函数创建 sidecar 处：

```csharp
            _sidecar = new PythonSidecar(_excelActions, excelApp, uiControl);
```

修改测试 `PythonSidecarTests.cs` 和 `ToolDispatcherTests.cs` 对应构造调用（加 `excelApp: null`）。

- [ ] **Step 5: 写测试 — 关键工具至少一个 happy path**

追加到 ToolDispatcherTests.cs：

```csharp
        [Fact]
        public void Execute_ReadWorkbook_ReturnsData()
        {
            var mock = new Mock<IExcelActions>();
            mock.Setup(x => x.ReadWorkbook()).Returns(new { name = "Book1" });
            var dispatcher = new ToolDispatcher(mock.Object, null);
            var result = dispatcher.Execute("read_workbook", new Dictionary<string, object>());
            Assert.True(result.Success);
        }

        [Fact]
        public void Execute_CreateSnapshot_ReturnsSnapshotId()
        {
            var mock = new Mock<IExcelActions>();
            mock.Setup(x => x.CreateSnapshot()).Returns("snap-123");
            var dispatcher = new ToolDispatcher(mock.Object, null);
            var result = dispatcher.Execute("create_snapshot", new Dictionary<string, object>());
            Assert.True(result.Success);
        }

        [Fact]
        public void Execute_Rollback_CallsExcelActions()
        {
            var mock = new Mock<IExcelActions>();
            mock.Setup(x => x.Rollback("snap-1")).Returns(true);
            var dispatcher = new ToolDispatcher(mock.Object, null);
            var args = new Dictionary<string, object> { { "snapshot_id", "snap-1" } };
            var result = dispatcher.Execute("rollback", args);
            Assert.True(result.Success);
            mock.Verify(x => x.Rollback("snap-1"), Times.Once);
        }
```

- [ ] **Step 6: 编译并运行测试**

Run: `cd c:\Users\qinju\Desktop\AIProject\DeepExcel && scripts\build-csc.bat`
Expected: 编译成功

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj --filter ToolDispatcherTests`
Expected: 全部通过

- [ ] **Step 7: Commit**

```bash
git add src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs src/DeepExcel.AddIn/Bridge/MessageBridge.cs src/DeepExcel.Tests/ToolDispatcherTests.cs src/DeepExcel.Tests/PythonSidecarTests.cs
git commit -m "feat(addin): port all 13 tool cases from Orchestrator to ToolDispatcher"
```

---

### Task 15: 接通 clarify 流程

**Files:**
- Verify: `src/DeepExcel.AddIn/Bridge/MessageBridge.cs` — 已在 Task 10 实现 OnClarify + _pendingClarifyQuestion 逻辑
- Verify: `src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs` — 已在 Task 8 实现 SendClarifyAnswer
- Modify: `src/DeepExcel.UI/src/bridge.ts` — 处理 clarify 消息类型（最小化：显示为普通消息）
- Modify: `src/DeepExcel.UI/src/components/InputArea.tsx` — clarify 状态下显示提示

**Interfaces:**
- 验证：clarify_intent 工具 → sidecar 发 clarify → C# 转发 → 前端显示问题 + 选项 → 用户输入 → 作为 clarify_answer 回传 → sidecar 解除阻塞

- [ ] **Step 1: 检查前端 bridge.ts 现有消息处理**

Run: `grep -n "stream_delta\|stream_end\|tool_call" src/DeepExcel.UI/src/bridge.ts`
Expected: 看到现有消息处理位置

读 [src/DeepExcel.UI/src/bridge.ts](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/bridge.ts) 找到消息分发逻辑。

- [ ] **Step 2: 在前端添加 clarify 消息处理**

修改 bridge.ts，在消息 switch 中加 `clarify` case：

```typescript
    case 'clarify': {
      const { question, options } = msg.payload;
      // MVP: 把 clarify 当作 assistant 消息显示，options 内联到问题文本
      const displayText = options && options.length > 0
        ? `${question}\n\n选项：${options.map((o: string, i: number) => `${i + 1}. ${o}`).join('\n')}`
        : question;
      onMessage({ role: 'assistant', content: displayText, type: 'clarify', options });
      break;
    }
```

实际实现依赖现有 onMessage 签名，根据 bridge.ts 的实际 API 调整。

- [ ] **Step 3: 重新构建前端**

Run: `cd src\DeepExcel.UI && npm run build`
Expected: 构建成功，输出到 dist/

复制到 AddIn 输出目录：

Run: `xcopy /E /I /Y "src\DeepExcel.UI\dist\*" "src\DeepExcel.AddIn\WebViewAssets\"`
Expected: 文件复制成功

- [ ] **Step 4: 手动测试 clarify 流程**

Excel 中输入一个会触发 clarify 的指令，如：在 A 列混合输入数字和文本，然后输入"统计 A 列数据"。

预期：
1. 模型调 read_range("A:A")
2. ToolDispatcher 返回 suggestion="该列同时包含数字和文本..."
3. 模型调 clarify_intent
4. 前端显示问题"求和还是计数？" + 选项
5. 用户输入"COUNTA" 或点击选项
6. 消息作为 clarify_answer 回传
7. 模型继续调 write_formula

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.UI/src/bridge.ts src/DeepExcel.UI/src/components/ src/DeepExcel.AddIn/WebViewAssets/
git commit -m "feat(ui): handle clarify messages with inline options display"
```

---

### Task 16: 替换为完整版 system_prompt.py

**Files:**
- Modify: `src/DeepExcel.Sidecar/system_prompt.py` — 用 spec §5.1 完整版替换 MVP 版

- [ ] **Step 1: 用 spec §5.1 完整内容替换 system_prompt.py**

修改 [src/DeepExcel.Sidecar/system_prompt.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/system_prompt.py)，完整内容：

```python
# src/DeepExcel.Sidecar/system_prompt.py

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

- [ ] **Step 2: 验证 import**

Run: `cd src\DeepExcel.Sidecar && python -c "from system_prompt import SYSTEM_PROMPT; print(len(SYSTEM_PROMPT))"`
Expected: 输出非零长度

- [ ] **Step 3: 重新部署 sidecar 到输出目录**

Run: `xcopy /E /I /Y "src\DeepExcel.Sidecar\*.py" "src\DeepExcel.AddIn\bin\Release\sidecar\"`

- [ ] **Step 4: Commit**

```bash
git add src/DeepExcel.Sidecar/system_prompt.py
git commit -m "feat(sidecar): replace MVP system prompt with full spec §5.1 version"
```

---

## 阶段三：清理 & 打包

---

### Task 17: 删除废弃的 C# 文件 + 更新构建

**Files:**
- Delete: `src/DeepExcel.AddIn/Agent/Orchestrator.cs`
- Delete: `src/DeepExcel.AddIn/Agent/ToolRegistry.cs`
- Delete: `src/DeepExcel.AddIn/Agent/PromptBuilder.cs`
- Delete: `src/DeepExcel.AddIn/Agent/ActionPlanner.cs`
- Delete: `src/DeepExcel.AddIn/Models/ClaudeAdapter.cs`
- Delete: `src/DeepExcel.AddIn/Models/DeepSeekAdapter.cs`
- Delete: `src/DeepExcel.AddIn/Models/OpenAIAdapter.cs`
- Delete: `src/DeepExcel.AddIn/Models/OpenAICompatibleAdapter.cs`
- Delete: `src/DeepExcel.AddIn/Models/IModelAdapter.cs`
- Delete: `src/DeepExcel.AddIn/Models/ModelAdapterFactory.cs`
- Delete: `src/DeepExcel.AddIn/Models/ModelFactory.cs`
- Delete: `src/DeepExcel.AddIn/Models/RetryHelper.cs`
- Modify: `src/DeepExcel.AddIn/DeepExcel.AddIn.csproj` — 移除上述 Compile Include
- Modify: `scripts/build-csc.bat` — 移除上述源文件
- Keep: `src/DeepExcel.AddIn/Models/ModelTypes.cs`（含 ToolCall / ModelRequest / ExcelContext 等类型，可能被 ToolDispatcher 引用）

- [ ] **Step 1: 检查 ModelTypes.cs 是否被新代码引用**

Run: `grep -rn "ModelRequest\|ToolCall\|ExcelContext\|ChatMessage\|ModelResponse" src/DeepExcel.AddIn/Sidecar/ src/DeepExcel.AddIn/Bridge/`
Expected: 看是否还有引用

如果完全没有引用，ModelTypes.cs 也可删。如果有引用（例如 ExcelContext 用于 BuildContext），保留。

- [ ] **Step 2: 删除废弃文件**

用 DeleteFile 工具批量删除（每批最多几个）：

- `src/DeepExcel.AddIn/Agent/Orchestrator.cs`
- `src/DeepExcel.AddIn/Agent/ToolRegistry.cs`
- `src/DeepExcel.AddIn/Agent/PromptBuilder.cs`
- `src/DeepExcel.AddIn/Agent/ActionPlanner.cs`
- `src/DeepExcel.AddIn/Models/ClaudeAdapter.cs`
- `src/DeepExcel.AddIn/Models/DeepSeekAdapter.cs`
- `src/DeepExcel.AddIn/Models/OpenAIAdapter.cs`
- `src/DeepExcel.AddIn/Models/OpenAICompatibleAdapter.cs`
- `src/DeepExcel.AddIn/Models/IModelAdapter.cs`
- `src/DeepExcel.AddIn/Models/ModelAdapterFactory.cs`
- `src/DeepExcel.AddIn/Models/ModelFactory.cs`
- `src/DeepExcel.AddIn/Models/RetryHelper.cs`

注意：**先检查 ModelFactory.cs 是否被 MessageBridge 等引用**。如果引用了 `ModelAdapterFactory.CreateDefault()`，需要先删除引用。Task 10 已把 MessageBridge 切换到 sidecar，应该不再引用 ModelAdapterFactory。验证：

Run: `grep -rn "ModelAdapterFactory\|IModelAdapter\|ClaudeAdapter\|DeepSeekAdapter" src/DeepExcel.AddIn/`
Expected: 仅在已删的文件内部互相引用，新代码无引用

- [ ] **Step 3: 更新 .csproj — 移除已删文件的 Compile Include**

修改 [src/DeepExcel.AddIn/DeepExcel.AddIn.csproj](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/DeepExcel.AddIn.csproj#L75-L113)，移除：

```xml
    <Compile Include="Agent\ActionPlanner.cs" />
    <Compile Include="Agent\Orchestrator.cs" />
    <Compile Include="Agent\PromptBuilder.cs" />
    <Compile Include="Agent\ToolRegistry.cs" />
    <Compile Include="Models\ClaudeAdapter.cs" />
    <Compile Include="Models\DeepSeekAdapter.cs" />
    <Compile Include="Models\IModelAdapter.cs" />
    <Compile Include="Models\ModelAdapterFactory.cs" />
    <Compile Include="Models\ModelFactory.cs" />
    <Compile Include="Models\OpenAIAdapter.cs" />
    <Compile Include="Models\OpenAICompatibleAdapter.cs" />
    <Compile Include="Models\RetryHelper.cs" />
```

保留 `Models\ModelTypes.cs`（如 Step 1 确认可删则也移除）。

- [ ] **Step 4: 更新 scripts/build-csc.bat — 移除已删源文件**

修改 [scripts/build-csc.bat](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/scripts/build-csc.bat)，从 csc 命令中移除：

```bat
  "%ADDINDIR%\Agent\ActionPlanner.cs" ^
  "%ADDINDIR%\Agent\Orchestrator.cs" ^
  "%ADDINDIR%\Agent\PromptBuilder.cs" ^
  "%ADDINDIR%\Agent\ToolRegistry.cs" ^
  "%ADDINDIR%\Models\ClaudeAdapter.cs" ^
  "%ADDINDIR%\Models\DeepSeekAdapter.cs" ^
  "%ADDINDIR%\Models\IModelAdapter.cs" ^
  "%ADDINDIR%\Models\ModelAdapterFactory.cs" ^
  "%ADDINDIR%\Models\ModelFactory.cs" ^
  "%ADDINDIR%\Models\OpenAIAdapter.cs" ^
  "%ADDINDIR%\Models\OpenAICompatibleAdapter.cs" ^
  "%ADDINDIR%\Models\RetryHelper.cs" ^
```

- [ ] **Step 5: 编译验证**

Run: `cd c:\Users\qinju\Desktop\AIProject\DeepExcel && scripts\build-csc.bat`
Expected: `=== Build SUCCESSFUL ===`

如果失败，根据错误信息修复未清理干净的引用。

- [ ] **Step 6: 运行所有测试**

Run: `dotnet test src\DeepExcel.Tests\DeepExcel.Tests.csproj`
Expected: 全部通过

Run: `cd src\DeepExcel.Sidecar && python -m pytest tests/ -v`
Expected: 全部通过（或 skipped for E2E）

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "refactor(addin): remove deprecated Orchestrator/Models adapters after sidecar migration"
```

---

### Task 18: Python embeddable 打包脚本

**Files:**
- Create: `scripts/package-python.ps1` — 下载 python embeddable + 安装 claude-agent-sdk
- Modify: `scripts/build-csc.bat` — 调用 package-python.ps1 并复制到输出目录

**Interfaces:**
- Produces: `src/DeepExcel.AddIn/bin/Release/python/python.exe` + 已装好 SDK 的 site-packages

- [ ] **Step 1: 实现 package-python.ps1**

```powershell
# scripts/package-python.ps1
# 下载 Python embeddable 并安装 claude-agent-sdk + 依赖
# 用法: powershell -ExecutionPolicy Bypass -File scripts\package-python.ps1

param(
    [string]$PythonVersion = "3.11.9",
    [string]$OutputDir = "$PSScriptRoot\..\src\DeepExcel.AddIn\bin\Release\python",
    [string]$TempDir = "$env:TEMP\deepexcel-python-packaging"
)

$ErrorActionPreference = "Stop"

# --- 1. 准备目录 ---
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

if (Test-Path $OutputDir) { Remove-Item $OutputDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# --- 2. 下载 Python embeddable zip（使用淘宝 npmmirror 镜像加速） ---
$arch = "amd64"
$zipName = "python-$PythonVersion-embed-$arch.zip"
# 官方源: https://www.python.org/ftp/python/$PythonVersion/$zipName
# 国内镜像: registry.npmmirror.com（淘宝 npm 镜像提供的 Python binary）
$mirrors = @(
    "https://registry.npmmirror.com/-/binary/python/$PythonVersion/$zipName",
    "https://www.python.org/ftp/python/$PythonVersion/$zipName"
)
$zipPath = Join-Path $TempDir $zipName

$downloaded = $false
foreach ($url in $mirrors) {
    Write-Host "尝试下载 Python embeddable: $url"
    try {
        Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -TimeoutSec 60
        $downloaded = $true
        Write-Host "下载成功（来源: $url）"
        break
    } catch {
        Write-Host "下载失败，尝试下一个镜像: $_"
    }
}
if (-not $downloaded) { throw "所有镜像均下载失败" }

# --- 3. 解压到输出目录 ---
Write-Host "解压到 $OutputDir"
Expand-Archive -Path $zipPath -DestinationPath $OutputDir -Force

# --- 4. 启用 pip (取消 _pth 中的 site-packages 注释, 下载 get-pip.py) ---
$pythonExe = Join-Path $OutputDir "python.exe"
$pthFile = Get-ChildItem -Path $OutputDir -Filter "python*._pth" | Select-Object -First 1

if ($pthFile) {
    Write-Host "启用 site-packages: $($pthFile.FullName)"
    $content = Get-Content $pthFile.FullName
    $content = $content | ForEach-Object {
        if ($_ -match "^#import site") { "import site" } else { $_ }
    }
    Set-Content -Path $pthFile.FullName -Value $content
}

# --- get-pip.py（小文件，直接官方下载；失败则用清华镜像） ---
$getPipMirrors = @(
    "https://bootstrap.pypa.io/get-pip.py",
    "https://pypi.tuna.tsinghua.edu.cn/packages/source/g/get-pip/get-pip-24.0.tar.gz"
)
$getPipPath = Join-Path $TempDir "get-pip.py"
$gpDownloaded = $false
foreach ($url in $getPipMirrors) {
    Write-Host "尝试下载 get-pip: $url"
    try {
        Invoke-WebRequest -Uri $url -OutFile $getPipPath -UseBasicParsing -TimeoutSec 30
        $gpDownloaded = $true
        break
    } catch {
        Write-Host "下载失败: $_"
    }
}
if (-not $gpDownloaded) { throw "get-pip.py 下载失败" }

Write-Host "安装 pip（使用清华镜像）"
& $pythonExe $getPipPath --no-warn-script-location --index-url https://pypi.tuna.tsinghua.edu.cn/simple
if ($LASTEXITCODE -ne 0) { throw "pip 安装失败" }

# --- 5. 安装 claude-agent-sdk 及依赖（使用清华镜像） ---
$tsinghuaIndex = "https://pypi.tuna.tsinghua.edu.cn/simple"
Write-Host "安装 claude-agent-sdk==0.2.109（清华镜像）"
& $pythonExe -m pip install --no-warn-script-location -i $tsinghuaIndex `
    "claude-agent-sdk==0.2.109" `
    "anyio>=4.0" `
    "httpx>=0.27" `
    "pydantic>=2.0"
if ($LASTEXITCODE -ne 0) { throw "claude-agent-sdk 安装失败" }

# --- 6. 复制 sidecar 业务代码 ---
$sidecarSrc = "$PSScriptRoot\..\src\DeepExcel.Sidecar"
$sidecarDest = Join-Path $OutputDir "sidecar"
Write-Host "复制 sidecar 代码到 $sidecarDest"

New-Item -ItemType Directory -Path $sidecarDest -Force | Out-Null
Copy-Item -Path "$sidecarSrc\*.py" -Destination $sidecarDest -Force
if (Test-Path "$sidecarSrc\assets") {
    Copy-Item -Path "$sidecarSrc\assets" -Destination $sidecarDest -Recurse -Force
}

# --- 7. 清理 pip 缓存与 __pycache__ ---
Write-Host "清理缓存"
Get-ChildItem -Path $OutputDir -Recurse -Directory -Filter "__pycache__" |
    Remove-Item -Recurse -Force
Get-ChildItem -Path $OutputDir -Recurse -Directory -Filter "*.dist-info" |
    Where-Object { $_.Name -match "^pip-" } |
    Remove-Item -Recurse -Force

# --- 8. 验证 ---
Write-Host "验证安装"
& $pythonExe -c "import claude_agent_sdk; print('claude-agent-sdk:', claude_agent_sdk.__version__)"
if ($LASTEXITCODE -ne 0) { throw "claude-agent-sdk 导入验证失败" }

& $pythonExe -c "import sys; sys.path.insert(0, r'$sidecarDest'); import sidecar; print('sidecar OK')"
if ($LASTEXITCODE -ne 0) { throw "sidecar 导入验证失败" }

# --- 9. 清理临时目录 ---
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== Python 打包完成 ==="
Write-Host "输出目录: $OutputDir"
Write-Host "Python exe: $pythonExe"
Write-Host "Sidecar: $sidecarDest"
Write-Host ""
Write-Host "DeepExcel 加载项将通过 python.exe -m sidecar.sidecar 启动"
```

- [ ] **Step 2: 修改 build-csc.bat 调用打包脚本**

在 `scripts\build-csc.bat` 末尾（csc 编译完成后）追加：

```bat
echo.
echo === 打包 Python sidecar ===
powershell -ExecutionPolicy Bypass -File "%~dp0package-python.ps1"
if errorlevel 1 (
    echo [ERROR] Python 打包失败
    exit /b 1
)

echo.
echo === 复制 sidecar 启动器 ===
xcopy /Y /E /I "%~dp0..\src\DeepExcel.Sidecar\*.py" "%~dp0..\src\DeepExcel.AddIn\bin\Release\python\sidecar\"
```

- [ ] **Step 3: 本地测试打包脚本**

Run: `powershell -ExecutionPolicy Bypass -File scripts\package-python.ps1`

Expected 输出末尾包含：
```
=== Python 打包完成 ===
输出目录: ...\src\DeepExcel.AddIn\bin\Release\python
Python exe: ...\python.exe
Sidecar: ...\python\sidecar
DeepExcel 加载项将通过 python.exe -m sidecar.sidecar 启动
```

验证命令：
```powershell
$py = "src\DeepExcel.AddIn\bin\Release\python\python.exe"
& $py -c "import claude_agent_sdk; print('SDK version:', claude_agent_sdk.__version__)"
& $py -c "import sys; sys.path.insert(0, r'src\DeepExcel.AddIn\bin\Release\python\sidecar'); import sidecar; print('sidecar import OK')"
```

- [ ] **Step 4: 验证 PythonSidecar.cs 能启动打包好的 python.exe**

修改 PythonSidecar.cs 的 PythonExePath 属性指向打包输出目录（如未在 Task 8 中已配置）：

```csharp
private string PythonExePath
{
    get
    {
        var assemblyDir = Path.GetDirectoryName(GetType().Assembly.Location);
        var packaged = Path.Combine(assemblyDir, "python", "python.exe");
        if (File.Exists(packaged)) return packaged;

        // 开发环境回退：使用系统 Python
        var sysPython = Environment.GetEnvironmentVariable("DEEPEXCEL_PYTHON");
        if (!string.IsNullOrEmpty(sysPython) && File.Exists(sysPython)) return sysPython;

        throw new FileNotFoundException(
            "找不到 python.exe。请运行 scripts\\package-python.ps1 打包，或设置 DEEPEXCEL_PYTHON 环境变量。");
    }
}
```

Run: 在 Excel 中重新加载加载项，发送一条测试消息
Expected: sidecar 进程启动，UI 显示正常响应

- [ ] **Step 5: Commit**

```bash
git add scripts/package-python.ps1 scripts/build-csc.bat src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs
git commit -m "build: 添加 Python embeddable 打包脚本，自动安装 claude-agent-sdk"
```

---

## Self-Review

### 1. Spec Coverage 检查

| Spec 章节 | 对应 Task | 状态 |
|---|---|---|
| §1 背景与目标 | Header (Goal) | ✅ |
| §2 架构（C# COM ↔ Python Sidecar） | Task 2, 5, 8 | ✅ |
| §3 IPC 协议（JSON Lines） | Task 2, 7 | ✅ |
| §4 Claude Agent SDK 集成 | Task 5 | ✅ |
| §5.1 System Prompt（完整版） | Task 16 | ✅ |
| §5.2 模糊指令默认推断 | Task 9 (GenerateRangeSuggestion) + Task 16 | ✅ |
| §6 @tool 工具定义（14 个） | Task 4 (MVP 4 个) + Task 13 (补齐 10 个) | ✅ |
| §7 "先尝试默认值，失败再问" 策略 | Task 9 (Suggestion 字段) + Task 15 | ✅ |
| §8 C# 端改造（PythonSidecar/ToolDispatcher/MessageBridge） | Task 7, 8, 9, 10 | ✅ |
| §9 多模型支持（ANTHROPIC_BASE_URL） | Task 5 (main 中读取 env) + Task 11 (config 消息) | ✅ |
| §10 打包与部署 | Task 18 | ✅ |
| §11 旧代码清理 | Task 17 | ✅ |

无未覆盖的 spec 章节。

### 2. Placeholder 扫描

搜索红旗模式：`TBD` / `TODO` / `...`（除代码示意）/ `implement later` / `fill in details`

✅ 未发现任何 placeholder。所有 Step 均包含完整代码或精确命令。

### 3. Type Consistency 检查

跨 Task 的关键类型/方法签名一致性：

| 名称 | 定义 Task | 使用 Task | 一致性 |
|---|---|---|---|
| `_message_buffer` 结构 | Task 2 | Task 5 | ✅ 一致 |
| `call_csharp(tool_name, args) -> dict` | Task 2 | Task 4 | ✅ 一致 |
| `call_csharp_clarify(question) -> str` | Task 2 | Task 4 (clarify_intent) | ✅ 一致 |
| `register_all_tools() -> list` | Task 4 | Task 5 | ✅ 一致 |
| `ClaudeAgentOptions` 字段 | Task 5 | Task 5 (main) | ✅ 一致 |
| `SidecarProtocol` 常量 | Task 7 | Task 8, 10 | ✅ 一致 |
| `ToolResult` (Suggestion/Context) | Task 7 | Task 8, 9 | ✅ 一致 |
| `PythonSidecar.Send(msg)` | Task 8 | Task 10 | ✅ 一致 |
| `ToolDispatcher.Execute(name, args) -> ToolResult` | Task 9 | Task 8 (HandleToolCall) | ✅ 一致 |
| `GenerateRangeSuggestion` 返回 `string` | Task 9 | - | ✅ 自洽 |
| `MessageBridge.OnSidecarMessage` | Task 10 | Task 8 (路由回调) | ✅ 一致 |

无类型不一致问题。

### 4. 额外发现

- **Task 11 涉及 ThisAddIn.cs 改造**：当前 ThisAddIn.cs 已存在 MessageBridge 创建逻辑，Task 10/11 的修改需保留原有 Ribbon/CustomTaskPane 代码，只替换 MessageBridge 构造与生命周期管理部分。
- **Task 17 删除 Orchestrator 后**：DeepExcel.AddIn.csproj 不再需要 `Models/` 目录引用，需在 Task 17 步骤中显式删除 csproj 中的相关条目（或确认 SDK 风格 csproj 自动排除）。
- **多模型支持 P2**：Spec §9 提到的 API key 配置文件（`%LOCALAPPDATA%\DeepExcel\config.json`）在本次计划中未单独成 Task，因为 ANTHROPIC_BASE_URL 与 ANTHROPIC_API_KEY 均通过环境变量传递（Task 5 main 函数 + Task 11 config 消息），无需独立配置文件加载逻辑。如后续 P2 需求要求 GUI 配置，再追加 Task 19+。

---

## Execution Handoff

实施计划已完成并保存到 `docs/superpowers/plans/2026-06-27-deepexcel-claude-sdk-migration.md`。

共 18 个 Task，分三个阶段：
- **阶段一 MVP**（Task 1-12）：Python sidecar 骨架 → IPC → 4 工具 → C# 桥接 → Excel 内烟测
- **阶段二完整迁移**（Task 13-16）：补齐 14 工具 → clarify 流程 → 完整 system prompt
- **阶段三清理打包**（Task 17-18）：删除旧代码 → Python embeddable 打包

两种执行方式：

**1. Subagent-Driven（推荐）** — 每个 Task 派发独立 subagent 执行，task 间进行 review，迭代快、上下文隔离清晰

**2. Inline Execution** — 在当前会话中按 executing-plans skill 批量执行，带检查点 review

请选择执行方式。