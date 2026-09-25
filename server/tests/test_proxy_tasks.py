"""What counts as one task.

The quota is sold in tasks. A user request is an agent loop of a dozen or more
model calls, and the proxy used to count every one of them: a free tier of 50
"tasks" ran out after three to five real ones.
"""

from __future__ import annotations

import json

import pytest
from sqlalchemy import select

from app.db import get_session_factory
from app.models import Entitlement, UsageRecord
from app.proxy.tasks import clean_trace_id, starts_new_user_turn, strip_window_suffix
from tests.test_proxy_streaming import STREAM_BODY, FakeUpstream, _install, hosted  # noqa: F401

TRACE = "task_0123456789abcdef"


def _user_turn(text="按地区汇总销售额"):
    return {"role": "user", "content": text}


def _tool_result_turn():
    return {"role": "user", "content": [
        {"type": "tool_result", "tool_use_id": "toolu_1", "content": "{\"success\": true}"},
    ]}


def _post(client, token, messages, trace=None, model="claude-opus-5"):
    headers = {"Authorization": f"Bearer {token}"}
    if trace is not None:
        headers["x-trace-id"] = trace
    return client.post(
        "/v1/messages",
        json={"model": model, "stream": True, "messages": messages},
        headers=headers,
    )


def _tasks_used(user_id):
    with get_session_factory()() as db:
        return db.scalar(select(Entitlement).where(Entitlement.user_id == user_id)).tasks_used


# ---------------------------------------------------------------------------
# Pure rules
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("messages, expected", [
    ([_user_turn()], True),
    ([_user_turn(), {"role": "assistant", "content": "…"}, _user_turn("再按月份拆开")], True),
    ([_user_turn(), {"role": "assistant", "content": [{"type": "tool_use"}]}, _tool_result_turn()], False),
    ([{"role": "user", "content": [{"type": "text", "text": "hi"}, {"type": "image"}]}], True),
    ([], True),
    (None, True),
    ([{"role": "assistant", "content": "prefill"}], True),
])
def test_new_user_turn_is_recognised_by_shape(messages, expected):
    payload = {} if messages is None else {"messages": messages}
    assert starts_new_user_turn(payload) is expected


@pytest.mark.parametrize("value, expected", [
    ("task_0123456789abcdef", "task_0123456789abcdef"),
    ("  4f9c2d1e-77aa-4b0e-9e1f-0c2b6e8d9a11 ", "4f9c2d1e-77aa-4b0e-9e1f-0c2b6e8d9a11"),
    ("short", None),
    ("x" * 65, None),
    ("has space inside", None),
    ("用户的内容", None),
    ("a/b/c/d/e/f/g/h", None),
    (None, None),
    (12345678, None),
])
def test_trace_ids_must_be_plain_identifiers(value, expected):
    # The column must never become a channel for user content.
    assert clean_trace_id(value) == expected


@pytest.mark.parametrize("model, expected", [
    ("deepseek-v4-pro[1m]", "deepseek-v4-pro"),
    ("deepseek-v4-pro[1M]", "deepseek-v4-pro"),
    ("glm-5.2[200k]", "glm-5.2"),
    ("claude-opus-5", "claude-opus-5"),
    ("weird[model]name", "weird[model]name"),
])
def test_context_window_suffix_is_stripped(model, expected):
    assert strip_window_suffix(model) == expected


# ---------------------------------------------------------------------------
# Through the proxy
# ---------------------------------------------------------------------------

def test_one_task_with_a_trace_id_is_counted_once(hosted, monkeypatch):
    client, token, user_id = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))

    _post(client, token, [_user_turn()], trace=TRACE)
    for _ in range(11):  # the agent loop carrying tool results back
        _post(client, token, [_user_turn(), _tool_result_turn()], trace=TRACE)

    assert _tasks_used(user_id) == 1
    with get_session_factory()() as db:
        records = db.scalars(select(UsageRecord)).all()
        assert len(records) == 12
        assert sum(r.counted_as_task for r in records) == 1
        assert {r.trace_id for r in records} == {TRACE}


def test_a_second_request_in_the_same_conversation_is_a_new_task(hosted, monkeypatch):
    client, token, user_id = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))

    _post(client, token, [_user_turn()], trace="task_aaaaaaaaaaaaaaaa")
    _post(client, token, [_user_turn(), _user_turn("再做个图表")], trace="task_bbbbbbbbbbbbbbbb")

    assert _tasks_used(user_id) == 2


def test_without_a_trace_id_tool_result_calls_are_not_new_tasks(hosted, monkeypatch):
    # Older clients send no x-trace-id; the shape of the request decides.
    client, token, user_id = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))

    _post(client, token, [_user_turn()])
    for _ in range(5):
        _post(client, token, [_user_turn(), _tool_result_turn()])

    assert _tasks_used(user_id) == 1


def test_a_failed_first_call_does_not_use_up_the_task(hosted, monkeypatch):
    client, token, user_id = hosted
    upstream = FakeUpstream([
        (400, b'{"error":"bad"}', {"content-type": "application/json"}),
        (200, STREAM_BODY, None),
    ])
    _install(monkeypatch, upstream)

    _post(client, token, [_user_turn()], trace=TRACE)
    assert _tasks_used(user_id) == 0
    _post(client, token, [_user_turn()], trace=TRACE)
    assert _tasks_used(user_id) == 1


def test_one_trace_id_cannot_carry_unlimited_calls(hosted, monkeypatch):
    # Reusing one id forever would make all traffic "one task".
    import os

    from app import config

    os.environ["MAX_CALLS_PER_TASK"] = "3"
    config.reset_settings_for_tests()
    try:
        client, token, _ = hosted
        upstream = FakeUpstream([(200, STREAM_BODY, None)])
        _install(monkeypatch, upstream)

        for _ in range(3):
            assert _post(client, token, [_user_turn(), _tool_result_turn()], trace=TRACE).status_code == 200
        refused = _post(client, token, [_user_turn(), _tool_result_turn()], trace=TRACE)

        assert refused.status_code == 429
        assert refused.json()["detail"]["reason"] == "task_call_limit"
        # Refused before forwarding: the fourth call cost nothing.
        assert len(upstream.calls) == 3
    finally:
        os.environ.pop("MAX_CALLS_PER_TASK", None)
        config.reset_settings_for_tests()


def test_trace_id_and_window_suffix_never_reach_the_provider(hosted, monkeypatch):
    client, token, _ = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)

    _post(client, token, [_user_turn()], trace=TRACE, model="deepseek-v4-pro[1m]")

    sent = upstream.calls[0]
    assert "x-trace-id" not in {k.lower() for k in sent["headers"]}
    assert json.loads(sent["content"])["model"] == "deepseek-v4-pro"
    with get_session_factory()() as db:
        assert db.scalar(select(UsageRecord)).model == "deepseek-v4-pro"


def test_an_unchanged_body_is_passed_through_byte_for_byte(hosted, monkeypatch):
    client, token, _ = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)

    raw = b'{"model": "claude-opus-5", "stream": true, "messages": [{"role": "user", "content": "hi"}]}'
    client.post("/v1/messages", content=raw, headers={
        "Authorization": f"Bearer {token}", "content-type": "application/json",
    })
    assert upstream.calls[0]["content"] == raw


def test_usage_can_be_read_for_one_task(hosted, monkeypatch):
    client, token, _ = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))

    _post(client, token, [_user_turn()], trace=TRACE)
    _post(client, token, [_user_turn(), _tool_result_turn()], trace=TRACE)
    _post(client, token, [_user_turn()], trace="task_other0000000000")

    usage = client.get(f"/v1/usage?trace_id={TRACE}", headers={"Authorization": f"Bearer {token}"}).json()
    assert usage["trace_id"] == TRACE
    assert usage["calls"] == 2
    assert usage["input_tokens"] == 2400

    bad = client.get("/v1/usage?trace_id=not%20an%20id", headers={"Authorization": f"Bearer {token}"})
    assert bad.status_code == 400
