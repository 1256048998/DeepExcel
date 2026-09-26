"""思考过程：SDK 缺 signature 的兼容补丁、思考事件节流、开关。"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import sdk_compat
import ui_events


def _assistant(blocks):
    return {"type": "assistant", "message": {"model": "deepseek-v4-flash", "content": blocks}}


def test_missing_signature_is_filled_before_parsing():
    data = _assistant([{"type": "thinking", "thinking": "先看表头"}, {"type": "text", "text": "好"}])
    sdk_compat.fill_missing_thinking_signature(data)
    assert data["message"]["content"][0]["signature"] == ""
    assert "signature" not in data["message"]["content"][1]


def test_existing_signature_is_kept():
    data = _assistant([{"type": "thinking", "thinking": "x", "signature": "abc"}])
    sdk_compat.fill_missing_thinking_signature(data)
    assert data["message"]["content"][0]["signature"] == "abc"


def test_installed_parser_accepts_thinking_without_signature():
    sdk_compat.install()
    sdk_compat.install()  # 装两次不会套两层
    from claude_agent_sdk._internal import message_parser
    from claude_agent_sdk.types import ThinkingBlock
    message = message_parser.parse_message(_assistant([{"type": "thinking", "thinking": "先看表头"}]))
    assert isinstance(message.content[0], ThinkingBlock)
    assert message.content[0].thinking == "先看表头"


def test_thinking_events_are_batched_and_closed():
    tracker = ui_events.ThinkingTracker()
    start = tracker.start(0, now=0.0)["event"]
    assert start["kind"] == "thinking_start" and start["id"] == "th-1"
    assert tracker.feed(0, "先", now=0.01) is None           # 太少、太快：先攒着
    assert tracker.feed(0, "看表头", now=0.05) is None
    delta = tracker.feed(0, "，", now=0.3)["event"]           # 过了节流间隔：一次发出攒下的
    assert delta == {**delta, "kind": "thinking_delta", "id": "th-1", "text": "先看表头，"}
    assert tracker.feed(0, "再写公式", now=0.35) is None
    tail, end = [e["event"] for e in tracker.stop(0, now=1.5)]
    assert tail["kind"] == "thinking_delta" and tail["text"] == "再写公式"
    assert end["kind"] == "thinking_end" and end["chars"] == 9 and end["duration_ms"] == 1500
    assert tracker.stop(0) == []
    assert tracker.start(3, now=2.0)["event"]["id"] == "th-2"


def test_long_thinking_flushes_without_waiting():
    tracker = ui_events.ThinkingTracker()
    tracker.start(0, now=0.0)
    event = tracker.feed(0, "字" * ui_events.THINKING_FLUSH_CHARS, now=0.01)
    assert event is not None and len(event["event"]["text"]) == ui_events.THINKING_FLUSH_CHARS


def test_thinking_switch(monkeypatch):
    import sidecar
    monkeypatch.delenv("DEEPEXCEL_THINKING", raising=False)
    assert sidecar.thinking_config() == {"type": "enabled", "budget_tokens": sidecar.THINKING_BUDGET_TOKENS}
    monkeypatch.setenv("DEEPEXCEL_THINKING", "off")
    assert sidecar.thinking_config() == {"type": "disabled"}
