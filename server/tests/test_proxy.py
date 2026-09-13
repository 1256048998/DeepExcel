"""Tests for the model proxy.

The proxy is the only point a client cannot bypass, which is what makes hosted
billing possible at all. Three properties matter:

  * It meters what actually went through, not what the client claimed.
  * It refuses when the entitlement says no, before any upstream cost is
    incurred.
  * It streams rather than buffers -- buffering to read the usage totals would
    turn a streaming API into a blocking one.
"""

from __future__ import annotations

import json

import pytest
from sqlalchemy import select

from app.db import get_session_factory
from app.models import Entitlement, RoutingMode, UsageRecord
from app.proxy.metering import StreamUsageCollector, Usage, estimate_cost_usd
from app.proxy.upstream import Upstream, select as select_upstreams
from tests.conftest import admin_token, auth_headers, register, set_routing_mode

HOSTED_URL = "https://proxy.example/v1"


# ---------------------------------------------------------------------------
# Usage extraction
# ---------------------------------------------------------------------------

def _sse(*events: dict) -> bytes:
    return b"".join(
        b"event: " + event["type"].encode() + b"\ndata: " + json.dumps(event).encode() + b"\n\n"
        for event in events
    )


def test_usage_is_extracted_from_a_stream():
    collector = StreamUsageCollector()
    collector.feed(
        _sse(
            {"type": "message_start",
             "message": {"usage": {"input_tokens": 1500, "cache_read_input_tokens": 200}}},
            {"type": "content_block_delta", "delta": {"text": "hello"}},
            {"type": "message_delta", "usage": {"output_tokens": 320},
             "delta": {"stop_reason": "end_turn"}},
        )
    )

    assert collector.usage.input_tokens == 1500
    assert collector.usage.output_tokens == 320
    assert collector.usage.cache_read_tokens == 200
    assert collector.stop_reason == "end_turn"


def test_usage_survives_chunk_boundaries_mid_event():
    """Real streams split wherever the network decides, not on event boundaries."""
    payload = _sse(
        {"type": "message_start", "message": {"usage": {"input_tokens": 100}}},
        {"type": "message_delta", "usage": {"output_tokens": 42}},
    )

    collector = StreamUsageCollector()
    for index in range(0, len(payload), 7):
        collector.feed(payload[index:index + 7])

    assert collector.usage.input_tokens == 100
    assert collector.usage.output_tokens == 42


def test_running_output_totals_are_not_double_counted():
    # message_delta repeats a cumulative output count; summing would inflate the
    # bill on every long response.
    collector = StreamUsageCollector()
    collector.feed(
        _sse(
            {"type": "message_delta", "usage": {"output_tokens": 10}},
            {"type": "message_delta", "usage": {"output_tokens": 25}},
            {"type": "message_delta", "usage": {"output_tokens": 40}},
        )
    )
    assert collector.usage.output_tokens == 40


def test_malformed_stream_does_not_crash_or_grow_unbounded():
    collector = StreamUsageCollector()
    collector.feed(b"data: not json at all\n\n")
    collector.feed(b"garbage without separators " * 50_000)
    collector.feed(_sse({"type": "message_delta", "usage": {"output_tokens": 5}}))

    assert collector.usage.output_tokens == 5


def test_unpriced_model_records_zero_rather_than_guessing():
    usage = Usage(input_tokens=1_000_000, output_tokens=1_000_000)
    # A wrong number on an invoice is worse than a missing one.
    assert estimate_cost_usd("some-unknown-model", usage) == 0.0
    assert estimate_cost_usd("claude-sonnet-5", usage) == pytest.approx(18.0)


# ---------------------------------------------------------------------------
# Upstream selection
# ---------------------------------------------------------------------------

def test_upstreams_are_selected_by_model_and_priority():
    upstreams = [
        Upstream("anthropic", "https://a", "k", ("claude-",), priority=10),
        Upstream("deepseek", "https://d", "k", ("deepseek-",), priority=20),
        Upstream("gateway", "https://g", "k", (), priority=90),
    ]

    candidates = select_upstreams(upstreams, "claude-opus-5")
    assert [u.name for u in candidates] == ["anthropic", "gateway"]

    assert [u.name for u in select_upstreams(upstreams, "deepseek-v4")] == ["deepseek", "gateway"]
    # A catch-all gateway still serves an unknown model, so a new model does not
    # need a deploy.
    assert [u.name for u in select_upstreams(upstreams, "brand-new")] == ["gateway"]


# ---------------------------------------------------------------------------
# Gatekeeping
# ---------------------------------------------------------------------------

def _hosted_client(make_client):
    import os

    from app import config
    from tests.conftest import ADMIN_EMAIL, ADMIN_PASSWORD

    client = make_client(
        BOOTSTRAP_ADMIN_EMAIL=ADMIN_EMAIL,
        BOOTSTRAP_ADMIN_PASSWORD=ADMIN_PASSWORD,
        HOSTED_PROXY_BASE_URL=HOSTED_URL,
    )
    os.environ["HOSTED_PROXY_BASE_URL"] = HOSTED_URL
    config.reset_settings_for_tests()
    return client


def _proxy_token(client, tokens) -> str:
    config_body = client.get("/api/v1/session/endpoint", headers=auth_headers(tokens)).json()
    return config_body["auth_header"].split(" ", 1)[1]


def test_proxy_requires_its_own_token(client):
    assert client.post("/v1/messages", json={"model": "claude-opus-5"}).status_code == 401
    assert client.post(
        "/v1/messages", json={"model": "x"}, headers={"Authorization": "Bearer nonsense"}
    ).status_code == 401


def test_account_token_is_not_accepted_by_the_proxy(make_client):
    # A leak at the proxy must not be an account takeover, and vice versa.
    client = _hosted_client(make_client)
    tokens = register(client)

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"}, headers=auth_headers(tokens)
    )
    assert response.status_code == 401


def test_byok_users_are_refused_by_the_proxy(make_client):
    # Their traffic should never arrive here; if it does, serving it would bill
    # us for a user who is paying their provider directly.
    client = _hosted_client(make_client)
    tokens = register(client)
    user_id = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    set_routing_mode(client, admin_token(client), user_id, "hosted")
    token = _proxy_token(client, tokens)

    # Flip back to BYOK after the token was issued.
    set_routing_mode(client, admin_token(client), user_id, "byok")

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )
    assert response.status_code == 403
    assert response.json()["detail"]["reason"] == "not_hosted"


def test_exhausted_quota_is_refused_before_any_upstream_call(make_client):
    client = _hosted_client(make_client)
    tokens = register(client)
    user_id = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    ops = admin_token(client)
    set_routing_mode(client, ops, user_id, "hosted")
    token = _proxy_token(client, tokens)

    client.post(
        f"/admin/api/users/{user_id}/entitlement",
        json={"task_limit": 0},
        headers=ops,
    )

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )
    assert response.status_code == 402
    assert response.json()["detail"]["reason"] == "quota_exhausted"
    # Nothing was forwarded, so nothing was metered.
    with get_session_factory()() as db:
        assert db.scalar(select(UsageRecord)) is None


def test_disabled_account_is_refused(make_client):
    client = _hosted_client(make_client)
    tokens = register(client)
    user_id = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    ops = admin_token(client)
    set_routing_mode(client, ops, user_id, "hosted")
    token = _proxy_token(client, tokens)

    client.post(f"/admin/api/users/{user_id}/status", json={"status": "disabled"}, headers=ops)

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )
    assert response.status_code == 403


def test_missing_model_is_rejected(make_client):
    client = _hosted_client(make_client)
    tokens = register(client)
    user_id = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    set_routing_mode(client, admin_token(client), user_id, "hosted")
    token = _proxy_token(client, tokens)

    response = client.post(
        "/v1/messages", json={}, headers={"Authorization": f"Bearer {token}"}
    )
    assert response.status_code == 400


def test_no_configured_upstream_is_a_clear_gateway_error(make_client):
    # No UPSTREAM_* variables are set in the test environment.
    client = _hosted_client(make_client)
    tokens = register(client)
    user_id = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    set_routing_mode(client, admin_token(client), user_id, "hosted")
    token = _proxy_token(client, tokens)

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )
    assert response.status_code == 502
    assert response.json()["detail"]["reason"] == "no_upstream"
