# src/DeepExcel.Sidecar/tests/test_interrupt_steer.py
# F9：停止 = 让 CLI 停下（interrupt + 排空），而不只是不再读它的输出；
# 任务进行中发的消息在下一个工具结果后交给模型，不会丢。
import asyncio

import anyio
import pytest
from claude_agent_sdk.types import AssistantMessage, ResultMessage, TextBlock

import ipc
import sidecar
import ui_events


def _result(subtype="success"):
    return ResultMessage(subtype=subtype, duration_ms=1, duration_api_ms=1, is_error=False,
                         num_turns=1, session_id="s", usage={"input_tokens": 1, "output_tokens": 1})


class FakeClient:
    """按轮回放预设的消息。interrupt() 之后当前轮只再吐出 ResultMessage。"""

    def __init__(self, turns, respond_to_interrupt=True, hang=False):
        self.turns = list(turns)
        self.respond_to_interrupt = respond_to_interrupt
        self.hang = hang
        self.queries = []
        self.interrupted = False

    async def query(self, prompt):
        self.queries.append(prompt)

    async def interrupt(self):
        self.interrupted = True

    async def receive_response(self):
        turn = self.turns.pop(0) if self.turns else [_result()]
        for item in turn:
            if self.interrupted and self.respond_to_interrupt:
                self.interrupted = False
                yield _result("error_during_execution")
                return
            if item == "wait":
                # 模拟一个很慢的工具：直到被 interrupt 或永远
                while not self.interrupted:
                    await anyio.sleep(0.01)
                if not self.respond_to_interrupt:
                    await anyio.sleep(3600)
                self.interrupted = False
                yield _result("error_during_execution")
                return
            yield item

    async def get_context_usage(self):
        return {"percentage": 10}


@pytest.fixture
def sent(monkeypatch):
    out = []

    async def fake_write(msg):
        out.append(msg)

    monkeypatch.setattr(sidecar, "write_message", fake_write)
    monkeypatch.setattr(sidecar, "_run", ui_events.RunTracker())
    monkeypatch.setattr(sidecar, "_needs_drain", False)
    ipc._init_buffer()
    return out


def _kinds(sent):
    return [m["event"]["kind"] if m.get("type") == "ui_event" else m.get("type") for m in sent]


async def _run_one_turn(client, text="hi", cancel_after=None):
    ipc._message_buffer["user_message"].put_nowait({"type": "user_message", "text": text, "context": {}})
    ipc._message_buffer["user_message"].put_nowait(None)

    async def press_stop():
        await anyio.sleep(cancel_after)
        ipc._message_buffer["cancel"].set()

    async with anyio.create_task_group() as tg:
        if cancel_after is not None:
            tg.start_soon(press_stop)
        await sidecar.run_agent_loop(client)


@pytest.mark.asyncio
async def test_stop_interrupts_the_cli_and_reports_interrupted(sent, monkeypatch):
    client = FakeClient([[AssistantMessage(content=[TextBlock(text="开始")], model="m"), "wait"]])
    await _run_one_turn(client, cancel_after=0.2)
    kinds = _kinds(sent)
    summary = [m["event"] for m in sent if m.get("type") == "ui_event" and m["event"]["kind"] == "run_summary"]
    assert summary and summary[0]["outcome"] == "interrupted"
    assert kinds.count("stream_end") == 1
    assert "status" in kinds  # 「正在停止…」
    assert sidecar._needs_drain is False


@pytest.mark.asyncio
async def test_hung_cli_is_cut_and_drained_before_the_next_turn(sent, monkeypatch):
    monkeypatch.setattr(sidecar, "INTERRUPT_GRACE_SECONDS", 0.2)
    client = FakeClient([["wait"]], respond_to_interrupt=False)
    await _run_one_turn(client, cancel_after=0.1)
    assert _kinds(sent).count("stream_end") == 1
    assert sidecar._needs_drain is True

    # 下一轮开始前先读掉残留，直到旧的 ResultMessage；新问题的回答不会被它截断
    stale = [AssistantMessage(content=[TextBlock(text="上一轮的残留")], model="m"), _result()]
    fresh = [AssistantMessage(content=[TextBlock(text="新回答")], model="m"), _result()]
    client.turns = [stale, fresh]
    sent.clear()
    await _run_one_turn(client, text="新问题")
    texts = [m["text"] for m in sent if m.get("type") == "stream_delta"]
    assert texts == ["新回答"]
    assert sidecar._needs_drain is False


@pytest.mark.asyncio
async def test_pending_tool_call_returns_immediately_on_stop(sent):
    async def stop_soon():
        await asyncio.sleep(0.1)
        ipc._message_buffer["cancel"].set()

    asyncio.get_event_loop().create_task(stop_soon())
    t0 = asyncio.get_event_loop().time()
    result = await ipc.call_csharp("read_range", {"address": "A1"}, timeout=30)
    assert result["error_code"] == "interrupted"
    assert asyncio.get_event_loop().time() - t0 < 2


def test_steer_is_held_only_while_a_turn_is_running():
    ipc._init_buffer()
    ipc._message_buffer["turn_active"] = True
    ipc.route_message({"type": "user_message", "text": "改成按月汇总", "steer": True})
    assert [m["text"] for m in ipc._message_buffer["steer"]] == ["改成按月汇总"]

    ipc._message_buffer["turn_active"] = False
    ipc.route_message({"type": "user_message", "text": "空闲时发的", "steer": True})
    assert ipc._message_buffer["user_message"].get_nowait()["text"] == "空闲时发的"


@pytest.mark.asyncio
async def test_post_tool_use_injects_steer_as_additional_context(sent):
    ipc._message_buffer["steer"] = [{"text": "改成按月汇总"}, {"text": "别动 Sheet2"}]
    out = await sidecar._post_tool_use_hook({}, "tu_1", None)
    ctx = out["hookSpecificOutput"]["additionalContext"]
    assert "改成按月汇总" in ctx and "别动 Sheet2" in ctx
    assert out["hookSpecificOutput"]["hookEventName"] == "PostToolUse"
    assert ipc._message_buffer["steer"] == []
    assert any(m.get("type") == "ui_event" and m["event"]["kind"] == "steer_delivered" for m in sent)
    # 没有插话时不表态
    assert await sidecar._post_tool_use_hook({}, "tu_2", None) == {}


@pytest.mark.asyncio
async def test_leftover_steer_becomes_the_next_message(sent):
    # 这一轮没有任何工具调用，插话来不及注入 → 下一轮作为普通消息处理
    client = FakeClient([[AssistantMessage(content=[TextBlock(text="好的")], model="m"), _result()],
                         [AssistantMessage(content=[TextBlock(text="已按月汇总")], model="m"), _result()]])

    original_query = client.query

    async def query_and_steer(prompt):
        await original_query(prompt)
        if len(client.queries) == 1:
            ipc.route_message({"type": "user_message", "text": "改成按月汇总", "steer": True})

    client.query = query_and_steer
    ipc._message_buffer["user_message"].put_nowait({"type": "user_message", "text": "汇总", "context": {}})

    async def stop_loop_later():
        await anyio.sleep(0.3)
        ipc._message_buffer["user_message"].put_nowait(None)

    async with anyio.create_task_group() as tg:
        tg.start_soon(stop_loop_later)
        await sidecar.run_agent_loop(client)

    # 每轮末尾固定带一句「用中文思考」的提醒，去掉后应当正好以插话结尾
    assert client.queries[-1].removesuffix(sidecar.THINKING_LANGUAGE_NOTE).endswith("改成按月汇总")
    assert "steer_deferred" in _kinds(sent)


@pytest.mark.asyncio
async def test_stop_discards_pending_steer(sent):
    client = FakeClient([["wait"]])
    ipc._message_buffer["turn_active"] = True
    ipc._message_buffer["steer"] = [{"text": "再加一列"}]
    await _run_one_turn(client, cancel_after=0.1)
    assert ipc._message_buffer["steer"] == []
    assert ipc._message_buffer["user_message"].empty()
