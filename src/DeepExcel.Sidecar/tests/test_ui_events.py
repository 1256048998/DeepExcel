# src/DeepExcel.Sidecar/tests/test_ui_events.py
# ui_event 信封：工具开始/结束按 tool_use_id 配对、结构化错误、压缩、终态行。
import json

import pytest
from claude_agent_sdk.types import (
    AssistantMessage,
    ResultMessage,
    SystemMessage,
    TextBlock,
    ToolResultBlock,
    ToolUseBlock,
    UserMessage,
)

import sidecar
import ui_events


@pytest.fixture
def sent(monkeypatch):
    out = []

    async def fake_write(msg):
        out.append(msg)

    monkeypatch.setattr(sidecar, "write_message", fake_write)
    monkeypatch.setattr(sidecar, "_run", ui_events.RunTracker())
    return out


def _events(sent, kind=None):
    evs = [m["event"] for m in sent if m.get("type") == "ui_event"]
    return [e for e in evs if kind is None or e["kind"] == kind]


def _tool_json(**payload):
    return [{"type": "text", "text": json.dumps(payload, ensure_ascii=False)}]


@pytest.mark.asyncio
async def test_tool_start_and_end_are_paired_by_tool_use_id(sent):
    await sidecar.handle_sdk_message(AssistantMessage(
        content=[ToolUseBlock(id="tu_1", name="mcp__excel__read_range", input={"address": "A1:D20"})],
        model="m"))
    await sidecar.handle_sdk_message(UserMessage(content=[ToolResultBlock(
        tool_use_id="tu_1",
        content=_tool_json(success=True, data={"values": [[1, 2, 3, 4]] * 20}))]))

    start, = _events(sent, "tool_start")
    end, = _events(sent, "tool_end")
    assert start["id"] == end["id"] == "tu_1"
    assert start["name"] == end["name"] == "read_range"
    assert start["args"] == {"address": "A1:D20"}
    assert end["ok"] is True and end["summary"] == "20 行 × 4 列"
    assert isinstance(end["duration_ms"], int)
    # 旧的 tool_use 仍然发：C# / WPS 用它记历史
    assert any(m.get("type") == "tool_use" for m in sent)


@pytest.mark.asyncio
async def test_failed_tool_carries_message_and_hint(sent):
    await sidecar.handle_sdk_message(AssistantMessage(
        content=[ToolUseBlock(id="tu_2", name="mcp__excel__write_formula", input={})], model="m"))
    await sidecar.handle_sdk_message(UserMessage(content=[ToolResultBlock(
        tool_use_id="tu_2", is_error=True,
        content=_tool_json(success=False, error="工作表受保护", suggestion="先取消保护再写入"))]))
    end, = _events(sent, "tool_end")
    assert end["ok"] is False
    assert end["error"]["message"] == "工作表受保护"
    assert end["error"]["hint"] == "先取消保护再写入"


@pytest.mark.asyncio
async def test_denied_tool_is_reported_as_denied(sent):
    await sidecar.handle_sdk_message(AssistantMessage(
        content=[ToolUseBlock(id="tu_3", name="mcp__excel__delete_sheet", input={})], model="m"))
    await sidecar.handle_sdk_message(UserMessage(content=[ToolResultBlock(
        tool_use_id="tu_3", is_error=True, content="用户拒绝执行此操作")]))
    end, = _events(sent, "tool_end")
    assert end["error"]["code"] == "denied"


@pytest.mark.asyncio
async def test_result_message_emits_run_summary_before_stream_end(sent):
    await sidecar.handle_sdk_message(AssistantMessage(
        content=[ToolUseBlock(id="tu_4", name="mcp__excel__read_range", input={})], model="m"))
    await sidecar.handle_sdk_message(ResultMessage(
        subtype="error_max_turns", duration_ms=1, duration_api_ms=1, is_error=True,
        num_turns=20, session_id="s", usage={"input_tokens": 10, "output_tokens": 5}))
    types = [m.get("type") for m in sent]
    assert types.index("ui_event") < types.index("stream_end")
    # 没收到结果的调用被补上 tool_end，面板不会一直转圈
    end, = _events(sent, "tool_end")
    assert end["id"] == "tu_4" and end["ok"] is False
    summary, = _events(sent, "run_summary")
    assert summary["outcome"] == "max_turns"
    assert summary["num_turns"] == 20
    assert summary["tool_calls"] == 1 and summary["failed_calls"] == 1


@pytest.mark.asyncio
async def test_compact_boundary_becomes_a_compaction_event(sent):
    await sidecar.handle_sdk_message(SystemMessage(
        subtype="compact_boundary",
        data={"compact_metadata": {"trigger": "auto", "pre_tokens": 150000}}))
    ev, = _events(sent, "compaction")
    assert ev["trigger"] == "auto" and ev["pre_tokens"] == 150000


@pytest.mark.asyncio
async def test_api_error_message_becomes_a_classified_error(sent):
    await sidecar.handle_sdk_message(AssistantMessage(
        content=[TextBlock(text="API Error: 401 invalid x-api-key")], model="m",
        error="authentication_failed"))
    ev, = _events(sent, "error")
    assert ev["code"] == "auth"
    assert "API Key" in ev["hint"]
    # 原始英文报错不当作回答流给用户
    assert not any(m.get("type") == "stream_delta" for m in sent)


def test_interrupted_tool_is_reported_in_chinese():
    out = ui_events.parse_tool_result(
        "The user doesn't want to proceed with this tool use. The tool use was rejected.", True)
    assert out["error"] == {"code": "interrupted", "message": "已中断", "hint": None}


@pytest.mark.parametrize("text,code", [
    ("HTTP 429 Too Many Requests", "rate_limit"),
    ('{"reason":"task_call_limit"}', "task_limit"),
    ("prompt is too long: 250000 tokens", "context_too_long"),
    ("ConnectError: getaddrinfo failed", "network"),
    ("something odd", "unknown"),
])
def test_classify_error(text, code):
    assert ui_events.classify_error(text)["code"] == code


def test_summarize_args_truncates_long_code_and_big_arrays():
    args = {"code": "x" * 10000, "values": [[i, i] for i in range(500)]}
    out = ui_events.summarize_args(args)
    assert len(out["code"]) < 4100
    assert out["values"]["__shape"] == [500, 2]
    assert len(out["values"]["head"]) == 3


def test_envelope_drops_none_fields():
    ev = ui_events.envelope("status", text="x", tool=None)["event"]
    assert ev["v"] == 1 and "tool" not in ev
