"""End-to-end proxy behaviour against a fake upstream.

Covers the parts that only show up when bytes actually move: that the response
is streamed rather than buffered, that usage is metered from what the upstream
really returned, and that a failing upstream fails over instead of becoming the
user's problem.
"""

from __future__ import annotations

import json

import httpx
import pytest
from sqlalchemy import select

from app.db import get_session_factory
from app.models import Entitlement, UsageRecord
from tests.conftest import ADMIN_EMAIL, ADMIN_PASSWORD, admin_token, auth_headers, register, set_routing_mode

HOSTED_URL = "https://proxy.example/v1"


def _sse(*events: dict) -> bytes:
    return b"".join(
        b"event: " + e["type"].encode() + b"\ndata: " + json.dumps(e).encode() + b"\n\n"
        for e in events
    )


STREAM_BODY = _sse(
    {"type": "message_start", "message": {"usage": {"input_tokens": 1200}}},
    {"type": "content_block_delta", "delta": {"text": "你好"}},
    {"type": "message_delta", "usage": {"output_tokens": 350}, "delta": {"stop_reason": "end_turn"}},
)


class FakeUpstream:
    """Stands in for a provider.

    Records what it received so the test can assert the client's credential was
    replaced by ours and the body was passed through untouched.
    """

    def __init__(self, script):
        self.script = script
        self.calls: list[dict] = []
        self.index = 0

    def __call__(self, *args, **kwargs):
        return _FakeClient(self)


class _FakeClient:
    def __init__(self, upstream: FakeUpstream):
        self.upstream = upstream

    def build_request(self, method, url, headers=None, content=None, **kwargs):
        self.upstream.calls.append(
            {"method": method, "url": url, "headers": dict(headers or {}), "content": content}
        )
        return httpx.Request(method, url, headers=headers, content=content)

    async def send(self, request, stream=False):
        step = self.upstream.script[min(self.upstream.index, len(self.upstream.script) - 1)]
        self.upstream.index += 1
        if isinstance(step, Exception):
            raise step
        status_code, body, headers = step
        return httpx.Response(
            status_code,
            content=body,
            headers=headers or {"content-type": "text/event-stream"},
            request=request,
        )

    async def aclose(self):
        return None


@pytest.fixture
def hosted(make_client, monkeypatch):
    import os

    from app import config

    client = make_client(
        BOOTSTRAP_ADMIN_EMAIL=ADMIN_EMAIL,
        BOOTSTRAP_ADMIN_PASSWORD=ADMIN_PASSWORD,
        HOSTED_PROXY_BASE_URL=HOSTED_URL,
    )
    os.environ["HOSTED_PROXY_BASE_URL"] = HOSTED_URL
    os.environ["UPSTREAM_1_NAME"] = "primary"
    os.environ["UPSTREAM_1_BASE_URL"] = "https://upstream-primary.example"
    os.environ["UPSTREAM_1_API_KEY"] = "sk-real-provider-key"
    os.environ["UPSTREAM_1_PRIORITY"] = "10"
    config.reset_settings_for_tests()

    tokens = register(client)
    user_id = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    set_routing_mode(client, admin_token(client), user_id, "hosted")
    endpoint = client.get("/api/v1/session/endpoint", headers=auth_headers(tokens)).json()
    proxy_token = endpoint["auth_header"].split(" ", 1)[1]

    yield client, proxy_token, user_id

    for key in ("UPSTREAM_1_NAME", "UPSTREAM_1_BASE_URL", "UPSTREAM_1_API_KEY",
                "UPSTREAM_1_PRIORITY", "UPSTREAM_2_NAME", "UPSTREAM_2_BASE_URL",
                "UPSTREAM_2_API_KEY", "UPSTREAM_2_PRIORITY", "HOSTED_PROXY_BASE_URL"):
        os.environ.pop(key, None)
    config.reset_settings_for_tests()


def _install(monkeypatch, upstream: FakeUpstream):
    from app.proxy import router as proxy_router

    monkeypatch.setattr(proxy_router.httpx, "AsyncClient", upstream)


def test_streamed_response_is_passed_through_and_metered(hosted, monkeypatch):
    client, token, user_id = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, {"content-type": "text/event-stream"})])
    _install(monkeypatch, upstream)

    response = client.post(
        "/v1/messages",
        json={"model": "claude-opus-5", "stream": True, "messages": [{"role": "user", "content": "hi"}]},
        headers={"Authorization": f"Bearer {token}"},
    )

    assert response.status_code == 200
    # Byte-for-byte passthrough: the client must see exactly what the provider
    # sent, or streaming UIs break in ways that are hard to trace.
    assert response.content == STREAM_BODY

    with get_session_factory()() as db:
        record = db.scalar(select(UsageRecord))
        assert record is not None
        assert record.input_tokens == 1200
        assert record.output_tokens == 350
        assert record.upstream == "primary"
        assert record.cost_usd > 0


def test_the_users_key_is_never_forwarded_and_ours_is(hosted, monkeypatch):
    client, token, _ = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)

    client.post(
        "/v1/messages",
        json={"model": "claude-opus-5", "stream": True},
        headers={"Authorization": f"Bearer {token}", "x-api-key": "sk-user-supplied"},
    )

    sent = upstream.calls[0]["headers"]
    # The proxy authenticates upstream with its own credential; anything the
    # client sent must be discarded rather than relayed.
    assert sent["x-api-key"] == "sk-real-provider-key"
    assert "authorization" not in {k.lower() for k in sent}


def test_successful_call_consumes_one_task(hosted, monkeypatch):
    client, token, user_id = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)

    client.post(
        "/v1/messages", json={"model": "claude-opus-5", "stream": True},
        headers={"Authorization": f"Bearer {token}"},
    )

    with get_session_factory()() as db:
        entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id))
        assert entitlement.tasks_used == 1


def test_upstream_5xx_fails_over_to_the_next(hosted, monkeypatch):
    import os

    from app import config

    os.environ["UPSTREAM_2_NAME"] = "backup"
    os.environ["UPSTREAM_2_BASE_URL"] = "https://upstream-backup.example"
    os.environ["UPSTREAM_2_API_KEY"] = "sk-backup"
    os.environ["UPSTREAM_2_PRIORITY"] = "20"
    config.reset_settings_for_tests()

    client, token, _ = hosted
    upstream = FakeUpstream([
        (503, b"upstream down", {"content-type": "text/plain"}),
        (200, STREAM_BODY, None),
    ])
    _install(monkeypatch, upstream)

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5", "stream": True},
        headers={"Authorization": f"Bearer {token}"},
    )

    assert response.status_code == 200
    assert len(upstream.calls) == 2
    with get_session_factory()() as db:
        assert db.scalar(select(UsageRecord)).upstream == "backup"


def test_client_errors_are_not_retried(hosted, monkeypatch):
    import os

    from app import config

    os.environ["UPSTREAM_2_NAME"] = "backup"
    os.environ["UPSTREAM_2_BASE_URL"] = "https://upstream-backup.example"
    os.environ["UPSTREAM_2_API_KEY"] = "sk-backup"
    config.reset_settings_for_tests()

    client, token, _ = hosted
    upstream = FakeUpstream([(400, b'{"error":"bad model"}', {"content-type": "application/json"})])
    _install(monkeypatch, upstream)

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )

    assert response.status_code == 400
    # Retrying a malformed request against another provider produces the same
    # error twice and doubles the latency for nothing.
    assert len(upstream.calls) == 1


def test_failed_call_is_recorded_but_does_not_consume_quota(hosted, monkeypatch):
    client, token, user_id = hosted
    upstream = FakeUpstream([(400, b'{"error":"bad"}', {"content-type": "application/json"})])
    _install(monkeypatch, upstream)

    client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )

    with get_session_factory()() as db:
        record = db.scalar(select(UsageRecord))
        assert record is not None and record.status_code == 400
        entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id))
        # Charging for a request the provider rejected would be charging for
        # nothing.
        assert entitlement.tasks_used == 0


def test_all_upstreams_failing_is_a_gateway_error(hosted, monkeypatch):
    client, token, _ = hosted
    upstream = FakeUpstream([httpx.ConnectError("refused")])
    _install(monkeypatch, upstream)

    response = client.post(
        "/v1/messages", json={"model": "claude-opus-5"},
        headers={"Authorization": f"Bearer {token}"},
    )
    assert response.status_code == 502
    assert response.json()["detail"]["reason"] == "all_upstreams_failed"


def test_non_streaming_response_is_also_metered(hosted, monkeypatch):
    client, token, _ = hosted
    body = json.dumps({"usage": {"input_tokens": 500, "output_tokens": 90}}).encode()
    upstream = FakeUpstream([(200, body, {"content-type": "application/json"})])
    _install(monkeypatch, upstream)

    response = client.post(
        "/v1/messages", json={"model": "claude-sonnet-5"},
        headers={"Authorization": f"Bearer {token}"},
    )

    assert response.status_code == 200
    with get_session_factory()() as db:
        record = db.scalar(select(UsageRecord))
        assert record.input_tokens == 500
        assert record.output_tokens == 90


def test_usage_endpoint_reports_what_was_metered(hosted, monkeypatch):
    client, token, _ = hosted
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)

    client.post(
        "/v1/messages", json={"model": "claude-opus-5", "stream": True},
        headers={"Authorization": f"Bearer {token}"},
    )

    usage = client.get("/v1/usage", headers={"Authorization": f"Bearer {token}"}).json()
    assert usage["calls"] == 1
    assert usage["input_tokens"] == 1200
    assert usage["output_tokens"] == 350
    assert usage["tasks_used"] == 1
