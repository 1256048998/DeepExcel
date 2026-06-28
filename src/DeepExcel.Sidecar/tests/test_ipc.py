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

    # 模拟 C# 延迟回写 tool_result：在 call_csharp 进入循环前，往 buffer 注入结果
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
