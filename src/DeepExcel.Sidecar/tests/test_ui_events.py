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


# ---------------------------------------------------------------------------
# 看门狗与代码边写边显示
# ---------------------------------------------------------------------------

def test_watchdog_is_quiet_when_things_move():
    assert ui_events.watchdog_status(3, None) is None
    assert ui_events.watchdog_status(3, ("read_range", 2)) is None


def test_watchdog_explains_what_it_is_waiting_for():
    assert "等待模型" in ui_events.watchdog_status(25, None)
    assert "停止" in ui_events.watchdog_status(130, None)
    assert "执行中" in ui_events.watchdog_status(1, ("execute_vba", 20))
    # 宿主一直不回结果：多半是 Excel 被模态对话框挡住了
    assert "对话框" in ui_events.watchdog_status(1, ("execute_vba", 125))


@pytest.mark.asyncio
async def test_status_watchdog_does_not_nag_while_waiting_for_the_user(sent, monkeypatch):
    import anyio
    import ipc
    ipc._init_buffer()
    monkeypatch.setattr(ui_events, "WATCHDOG_MODEL_WAIT", 0)
    ipc._message_buffer["awaiting_user"] = 1
    with anyio.move_on_after(0.2):
        await sidecar.status_watchdog(tick=0.02)
    assert _events(sent, "status") == []
    ipc._message_buffer["awaiting_user"] = 0
    with anyio.move_on_after(0.2):
        await sidecar.status_watchdog(tick=0.02)
    statuses = _events(sent, "status")
    assert len(statuses) == 1  # 同一句话不重复发
    assert "等待模型" in statuses[0]["text"]


@pytest.mark.parametrize("buffer,expected", [
    (r'{"code": "Sub A()\n  x = 1', "Sub A()\n  x = 1"),
    (r'{"code": "say \"hi\" and \u4e2d', 'say "hi" and 中'),
    ('{"code": "ends mid escape ' + chr(92), "ends mid escape "),
    ('{"code": "done"}', "done"),
    ('{"other": 1', None),
    ('{"code": 12', None),
])
def test_partial_json_string(buffer, expected):
    assert ui_events.partial_json_string(buffer, "code") == expected


def test_tool_gen_streams_code_preview_throttled():
    gen = ui_events.ToolGenTracker()
    gen.start(1, "tu_9", "mcp__excel__execute_vba")
    first = gen.feed(1, r'{"code": "Sub A()\n', now=10.0)["event"]
    assert first["kind"] == "tool_gen" and first["id"] == "tu_9" and first["name"] == "execute_vba"
    assert first["preview"] == "Sub A()\n" and first["lines"] == 2
    # 250ms 内的增量只累积，不发
    assert gen.feed(1, r"  x = 1\n", now=10.1) is None
    later = gen.feed(1, "End Sub", now=10.4)["event"]
    assert later["preview"].endswith("End Sub") and later["lines"] == 3
    gen.stop(1)
    assert gen.feed(1, "more", now=11.0) is None


def test_tool_gen_for_non_code_tools_reports_size_only():
    gen = ui_events.ToolGenTracker()
    gen.start(0, "tu_1", "mcp__excel__write_range")
    ev = gen.feed(0, '{"address": "A1", "values": [[1,2],[3,4]', now=5.0)["event"]
    assert ev["chars"] > 10 and "preview" not in ev


@pytest.mark.asyncio
async def test_stream_events_produce_tool_gen(sent):
    from claude_agent_sdk.types import StreamEvent
    sidecar._gen = ui_events.ToolGenTracker()

    def se(event):
        return StreamEvent(uuid="u", session_id="s", event=event)

    await sidecar.handle_sdk_message(se({"type": "content_block_start", "index": 2,
                                         "content_block": {"type": "tool_use", "id": "tu_5",
                                                           "name": "mcp__excel__execute_python", "input": {}}}))
    await sidecar.handle_sdk_message(se({"type": "content_block_delta", "index": 2,
                                         "delta": {"type": "input_json_delta",
                                                   "partial_json": '{"code": "print(1)'}}))
    ev, = _events(sent, "tool_gen")
    assert ev["id"] == "tu_5" and ev["preview"] == "print(1)"


def test_write_check_is_passed_to_the_panel():
    content = json.dumps({"success": True, "data": {}, "verification": {
        "ok": False, "summary": "体检发现问题：新增 1 个公式错误：Sheet1!D2 #DIV/0!", "new_errors": ["Sheet1!D2 #DIV/0!"]}},
        ensure_ascii=False)
    out = ui_events.parse_tool_result(content, False)
    assert out["ok"] is True
    assert out["check"] == {"ok": False, "summary": "体检发现问题：新增 1 个公式错误：Sheet1!D2 #DIV/0!"}


def test_no_verification_no_check():
    assert "check" not in ui_events.parse_tool_result(json.dumps({"success": True, "data": {}}), False)


def test_checkpoint_id_is_passed_to_the_panel():
    out = ui_events.parse_tool_result(json.dumps({"success": True, "data": {}, "checkpoint_id": "abc"}), False)
    assert out["checkpoint_id"] == "abc"
    failed = ui_events.parse_tool_result(json.dumps({"success": False, "error": "x", "checkpoint_id": "abc"}), False)
    assert "checkpoint_id" not in failed
