"""The quota in points: a task costs its model's weight.

One point is one task on the baseline model. A task on a model whose typical
task costs twelve times as much takes twelve points, charged when the task
starts, so the user knows the price before sending and a running task is never
cut off halfway.
"""

from __future__ import annotations

import os

import pytest
from sqlalchemy import select

from app import config
from app.db import get_session_factory
from app.models import Entitlement, UsageRecord
from app.proxy.metering import DEFAULT_PRICES, model_weight, model_weights
from tests.conftest import admin_token, auth_headers
from tests.test_proxy_streaming import STREAM_BODY, FakeUpstream, _install, hosted  # noqa: F401


def _user_turn(text="按地区汇总销售额"):
    return {"role": "user", "content": text}


def _tool_result_turn():
    return {"role": "user", "content": [
        {"type": "tool_result", "tool_use_id": "toolu_1", "content": "{\"success\": true}"},
    ]}


def _post(client, token, messages, model):
    return client.post(
        "/v1/messages",
        json={"model": model, "stream": True, "messages": messages},
        headers={"Authorization": f"Bearer {token}"},
    )


def _entitlement(user_id):
    with get_session_factory()() as db:
        return db.scalar(select(Entitlement).where(Entitlement.user_id == user_id))


def _set_limit(client, user_id, limit):
    response = client.post(
        f"/admin/api/users/{user_id}/entitlement",
        json={"task_limit": limit},
        headers=admin_token(client),
    )
    assert response.status_code == 200, response.text


# ---------------------------------------------------------------------------
# Weights
# ---------------------------------------------------------------------------

def test_weights_follow_the_cost_of_a_typical_task():
    weights = model_weights()
    assert weights["deepseek"] == 1
    assert weights["claude-sonnet"] == 12
    assert weights["claude-opus"] == 58
    assert weights["claude-haiku"] == 3
    # Every priced model has a weight, and none is below one point.
    assert set(weights) == set(DEFAULT_PRICES)
    assert min(weights.values()) == 1


def test_weights_can_be_overridden_per_prefix_and_the_longest_prefix_wins():
    overrides = {"claude-sonnet-5": 8}
    assert model_weight("claude-sonnet-5", overrides) == 8
    assert model_weight("claude-sonnet-4-5", overrides) == 12
    assert model_weight("DeepSeek-V4-Pro") == 1


def test_an_unpriced_model_is_charged_as_the_most_expensive():
    # One point would let an expensive model run at the baseline rate unnoticed.
    assert model_weight("some-new-model") == max(model_weights().values())


@pytest.mark.parametrize("raw", ['{"claude-opus": 0}', '{"claude-opus": "40"}', "[1]", "not json", '{"x": true}'])
def test_a_malformed_override_stops_startup(raw):
    os.environ["MODEL_WEIGHTS"] = raw
    try:
        with pytest.raises(RuntimeError):
            config.Settings()
    finally:
        os.environ.pop("MODEL_WEIGHTS", None)


# ---------------------------------------------------------------------------
# Charging
# ---------------------------------------------------------------------------

def test_a_task_takes_its_models_points_once(hosted, monkeypatch):
    client, token, user_id = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))

    _post(client, token, [_user_turn()], "claude-sonnet-5")
    for _ in range(3):
        _post(client, token, [_user_turn(), _tool_result_turn()], "claude-sonnet-5")
    assert _entitlement(user_id).tasks_used == 12

    _post(client, token, [_user_turn("再按月份拆开")], "deepseek-v4-flash")
    assert _entitlement(user_id).tasks_used == 13

    with get_session_factory()() as db:
        charged = [r.points for r in db.scalars(select(UsageRecord).order_by(UsageRecord.id))]
    assert charged == [12, 0, 0, 0, 1]


def test_an_override_changes_what_a_task_costs(hosted, monkeypatch):
    os.environ["MODEL_WEIGHTS"] = '{"claude-sonnet": 5}'
    config.reset_settings_for_tests()
    try:
        client, token, user_id = hosted
        _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))
        _post(client, token, [_user_turn()], "claude-sonnet-5")
        assert _entitlement(user_id).tasks_used == 5
    finally:
        os.environ.pop("MODEL_WEIGHTS", None)
        config.reset_settings_for_tests()


def test_a_task_the_balance_cannot_pay_for_is_refused_before_forwarding(hosted, monkeypatch):
    client, token, user_id = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)
    _set_limit(client, user_id, 10)

    refused = _post(client, token, [_user_turn()], "claude-sonnet-5")
    assert refused.status_code == 402
    detail = refused.json()["detail"]
    assert detail["reason"] == "quota_insufficient"
    assert detail["points_needed"] == 12
    assert detail["points_remaining"] == 10
    assert "12 点" in detail["message"] and "10 点" in detail["message"]
    assert upstream.calls == []

    # A cheaper model still works with the same balance.
    assert _post(client, token, [_user_turn()], "deepseek-v4-pro").status_code == 200
    assert _entitlement(user_id).tasks_used == 1


def test_a_running_task_is_not_cut_off_when_the_balance_reaches_zero(hosted, monkeypatch):
    client, token, user_id = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))
    _set_limit(client, user_id, 12)

    assert _post(client, token, [_user_turn()], "claude-sonnet-5").status_code == 200
    assert _entitlement(user_id).tasks_remaining == 0
    # The task already paid; its remaining steps go through.
    for _ in range(3):
        assert _post(client, token, [_user_turn(), _tool_result_turn()], "claude-sonnet-5").status_code == 200

    refused = _post(client, token, [_user_turn("下一个任务")], "deepseek-v4-pro")
    assert refused.status_code == 402
    assert refused.json()["detail"]["reason"] == "quota_exhausted"


def test_the_client_is_told_the_weights_only_when_hosted(hosted, make_client):
    client, _, user_id = hosted
    from tests.conftest import register, set_routing_mode

    tokens = register(client, email="points-byok@example.com")
    byok = client.get("/api/v1/session/endpoint", headers=auth_headers(tokens)).json()
    assert byok["entitlement"]["model_weights"] is None

    hosted_user = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    set_routing_mode(client, admin_token(client), hosted_user, "hosted")
    endpoint = client.get("/api/v1/session/endpoint", headers=auth_headers(tokens)).json()
    assert endpoint["entitlement"]["model_weights"]["claude-sonnet"] == 12
    assert endpoint["entitlement"]["model_weights"]["deepseek"] == 1
    assert endpoint["entitlement"]["model_weight_default"] == model_weight("some-new-model")
    assert byok["entitlement"]["model_weight_default"] is None
