"""Tests for the endpoint-routing contract.

This is the part of M2 that is expensive to get wrong, because the client is
built against it and turning on hosted routing in M6 depends on it being purely
server-driven. These tests are the specification.
"""

from __future__ import annotations

import jwt
import pytest

from app.security import (
    ACCESS_AUDIENCE_PROXY,
    ACCESS_AUDIENCE_USER,
    decode_access_token,
)
from tests.conftest import (
    admin_token,
    auth_headers,
    register,
    set_routing_mode,
)

HOSTED_URL = "https://api.deepexcel.example/v1"


def test_default_deployment_tells_every_client_byok(client):
    """M2 state: no proxy configured, so nobody is routed through us."""
    tokens = register(client)
    response = client.get("/api/v1/session/endpoint", headers=auth_headers(tokens))

    assert response.status_code == 200, response.text
    body = response.json()
    assert body["mode"] == "byok"
    # A client that received a base_url would start sending model traffic to a
    # proxy that does not exist.
    assert body["base_url"] is None
    assert body["auth_header"] is None
    assert body["entitlement"]["plan"] == "beta"
    assert body["expires_at"] > 0
    assert body["refresh_after_seconds"] > 0


def test_flipping_to_hosted_requires_no_client_change(admin_client):
    """The M6 switch: set the proxy URL, flip the entitlement, done.

    The client sends the identical request in both cases and simply applies
    whatever comes back.
    """
    tokens = register(admin_client)
    headers = auth_headers(tokens)

    before = admin_client.get("/api/v1/session/endpoint", headers=headers).json()
    assert before["mode"] == "byok"

    # Operator action, no deploy of the client.
    import os

    from app import config

    os.environ["HOSTED_PROXY_BASE_URL"] = HOSTED_URL
    config.reset_settings_for_tests()

    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]
    set_routing_mode(admin_client, admin_token(admin_client), user_id, "hosted")

    after = admin_client.get("/api/v1/session/endpoint", headers=headers).json()
    assert after["mode"] == "hosted"
    assert after["base_url"] == HOSTED_URL
    # A complete header value, so the client never encodes the scheme itself.
    assert after["auth_header"].startswith("Bearer ")


def test_proxy_token_cannot_be_used_against_account_api(admin_client):
    """Audience separation.

    The proxy token goes to a different component than the account token. If it
    were interchangeable, a leak at the proxy would be a full account takeover.
    """
    import os

    from app import config

    os.environ["HOSTED_PROXY_BASE_URL"] = HOSTED_URL
    config.reset_settings_for_tests()

    tokens = register(admin_client)
    headers = auth_headers(tokens)
    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]
    set_routing_mode(admin_client, admin_token(admin_client), user_id, "hosted")

    config_body = admin_client.get("/api/v1/session/endpoint", headers=headers).json()
    proxy_token = config_body["auth_header"].split(" ", 1)[1]

    # It is a valid token for its own audience...
    claims = decode_access_token(proxy_token, audience=ACCESS_AUDIENCE_PROXY)
    assert claims.subject == str(user_id)

    # ...and is rejected for any other.
    with pytest.raises(jwt.InvalidAudienceError):
        decode_access_token(proxy_token, audience=ACCESS_AUDIENCE_USER)

    replayed = admin_client.get(
        "/api/v1/auth/me", headers={"Authorization": f"Bearer {proxy_token}"}
    )
    assert replayed.status_code == 401


def test_hosted_without_configured_proxy_fails_loudly(admin_client):
    """Misconfiguration must not silently downgrade to BYOK.

    A hosted user has no local provider key. Quietly answering "byok" would
    leave them with authentication errors pointing at a provider they never
    configured, and the real cause -- our deployment -- would be invisible.
    """
    tokens = register(admin_client)
    headers = auth_headers(tokens)
    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]
    set_routing_mode(admin_client, admin_token(admin_client), user_id, "hosted")

    response = admin_client.get("/api/v1/session/endpoint", headers=headers)
    assert response.status_code == 503
    assert response.json()["detail"]["reason"] == "hosted_routing_unavailable"


def test_inactive_entitlement_is_payment_required(admin_client):
    tokens = register(admin_client)
    headers = auth_headers(tokens)
    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]

    admin_client.post(
        f"/admin/api/users/{user_id}/entitlement",
        json={"status": "suspended"},
        headers=admin_token(admin_client),
    )

    response = admin_client.get("/api/v1/session/endpoint", headers=headers)
    assert response.status_code == 402
    assert response.json()["detail"]["reason"] == "entitlement_inactive"


def test_quota_applies_to_hosted_but_not_byok(admin_client):
    """A BYOK user pays their provider directly, so metering them would be
    charging for traffic we never carry."""
    import os

    from app import config

    tokens = register(admin_client)
    headers = auth_headers(tokens)
    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]
    ops = admin_token(admin_client)

    # Zero remaining quota.
    admin_client.post(
        f"/admin/api/users/{user_id}/entitlement",
        json={"task_limit": 1},
        headers=ops,
    )
    consumed = admin_client.post("/api/v1/session/consume-task", headers=headers)
    assert consumed.status_code == 200
    assert consumed.json()["tasks_remaining"] == 0

    # BYOK: exhausted quota is irrelevant, the session still resolves.
    byok = admin_client.get("/api/v1/session/endpoint", headers=headers)
    assert byok.status_code == 200
    assert byok.json()["mode"] == "byok"
    assert byok.json()["entitlement"]["tasks_remaining"] == 0

    # Hosted: the same quota now blocks.
    os.environ["HOSTED_PROXY_BASE_URL"] = HOSTED_URL
    config.reset_settings_for_tests()
    set_routing_mode(admin_client, ops, user_id, "hosted")

    hosted = admin_client.get("/api/v1/session/endpoint", headers=headers)
    assert hosted.status_code == 402
    assert hosted.json()["detail"]["reason"] == "quota_exhausted"


def test_disabled_account_loses_access_immediately(admin_client):
    """Disabling must not wait for a token to expire."""
    tokens = register(admin_client)
    headers = auth_headers(tokens)
    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]

    assert admin_client.get("/api/v1/session/endpoint", headers=headers).status_code == 200

    admin_client.post(
        f"/admin/api/users/{user_id}/status",
        json={"status": "disabled"},
        headers=admin_token(admin_client),
    )

    # The access token is still cryptographically valid; the account check is
    # what stops it.
    assert admin_client.get("/api/v1/session/endpoint", headers=headers).status_code == 403

    # And the refresh token cannot mint a new one.
    refreshed = admin_client.post(
        "/api/v1/auth/refresh", json={"refresh_token": tokens["refresh_token"]}
    )
    assert refreshed.status_code in (401, 403)


def test_endpoint_requires_authentication(client):
    assert client.get("/api/v1/session/endpoint").status_code == 401
    assert client.get(
        "/api/v1/session/endpoint", headers={"Authorization": "Bearer nonsense"}
    ).status_code == 401
