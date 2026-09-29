"""The hosted model catalog: the server says which models a hosted user may pick.

Before it, hosted users had no provider configured locally, the panel showed no
model choice, and every request carried whatever model name sat in the local
config -- which the proxy forwarded if any upstream accepted it.
"""

from __future__ import annotations

import json
import os

import pytest

from app import config
from tests.conftest import auth_headers, register
from tests.test_proxy_streaming import STREAM_BODY, FakeUpstream, _install, hosted  # noqa: F401

CATALOG = [
    {"model": "deepseek-v4-pro", "label": "DeepSeek V4 Pro", "default": True},
    {"model": "deepseek-v4-flash", "label": "DeepSeek V4 Flash"},
    {"model": "claude-sonnet-5", "label": "Claude Sonnet 5", "plans": ["pro", "team"]},
]


@pytest.fixture
def catalog_env():
    def apply(entries, upstream_models=None):
        os.environ["HOSTED_MODELS"] = json.dumps(entries)
        if upstream_models is not None:
            os.environ["UPSTREAM_1_MODELS"] = upstream_models
        config.reset_settings_for_tests()
    yield apply
    os.environ.pop("HOSTED_MODELS", None)
    os.environ.pop("UPSTREAM_1_MODELS", None)
    config.reset_settings_for_tests()


def _endpoint(client, email="user@deepexcel-qa.com"):
    # The hosted fixture registered this user and switched them to hosted.
    login = client.post("/api/v1/auth/login", json={"email": email, "password": "correct-horse-battery"})
    return client.get("/api/v1/session/endpoint", headers=auth_headers(login.json())).json()


def _post(client, token, model):
    return client.post(
        "/v1/messages",
        json={"model": model, "stream": True, "messages": [{"role": "user", "content": "按地区汇总"}]},
        headers={"Authorization": f"Bearer {token}"},
    )


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("raw", [
    "not json",
    '{"model": "deepseek-v4-pro"}',
    '[{"label": "no model"}]',
    '[{"model": "a"}, {"model": "A"}]',
    '[{"model": "a", "plans": ["gold"]}]',
    '[{"model": "a", "plans": []}]',
    '[{"model": "a", "default": "yes"}]',
])
def test_a_malformed_catalog_stops_startup(raw):
    os.environ["HOSTED_MODELS"] = raw
    try:
        with pytest.raises(RuntimeError):
            config.Settings()
    finally:
        os.environ.pop("HOSTED_MODELS", None)


# ---------------------------------------------------------------------------
# What the client is told
# ---------------------------------------------------------------------------

def test_without_a_catalog_the_client_keeps_its_local_model(hosted):
    client, _, _ = hosted
    assert _endpoint(client)["models"] is None


def test_the_catalog_comes_with_points_and_a_default_and_respects_the_plan(hosted, catalog_env):
    client, _, _ = hosted
    catalog_env(CATALOG)
    models = _endpoint(client)["models"]
    # The test user is on the beta plan: the pro/team-only entry is not offered.
    assert [m["model"] for m in models] == ["deepseek-v4-pro", "deepseek-v4-flash"]
    assert models[0] == {"model": "deepseek-v4-pro", "label": "DeepSeek V4 Pro", "points": 1, "default": True}
    assert models[1]["default"] is False


def test_a_model_no_upstream_serves_is_not_offered(hosted, catalog_env):
    client, _, _ = hosted
    catalog_env(CATALOG + [{"model": "kimi-k2", "label": "Kimi"}], upstream_models="deepseek")
    assert [m["model"] for m in _endpoint(client)["models"]] == ["deepseek-v4-pro", "deepseek-v4-flash"]


def test_without_a_default_the_cheapest_model_is_the_default(hosted, catalog_env):
    client, _, _ = hosted
    catalog_env([{"model": "claude-sonnet-5", "label": "Sonnet"}, {"model": "deepseek-v4-flash", "label": "Flash"}])
    models = _endpoint(client)["models"]
    assert [(m["model"], m["default"]) for m in models] == [("claude-sonnet-5", False), ("deepseek-v4-flash", True)]
    assert models[0]["points"] == 12


def test_byok_users_get_no_catalog(make_client, catalog_env):
    client = make_client()
    catalog_env(CATALOG)
    tokens = register(client, email="catalog-byok@example.com")
    endpoint = client.get("/api/v1/session/endpoint", headers=auth_headers(tokens)).json()
    assert endpoint["mode"] == "byok"
    assert endpoint["models"] is None


# ---------------------------------------------------------------------------
# What the proxy lets through
# ---------------------------------------------------------------------------

def test_the_proxy_only_forwards_catalog_models(hosted, catalog_env, monkeypatch):
    client, token, _ = hosted
    catalog_env(CATALOG)
    upstream = FakeUpstream([(200, STREAM_BODY, None)])
    _install(monkeypatch, upstream)

    assert _post(client, token, "deepseek-v4-pro").status_code == 200
    # In the catalog, but not for this user's plan.
    refused = _post(client, token, "claude-sonnet-5")
    assert refused.status_code == 400
    detail = refused.json()["detail"]
    assert detail["reason"] == "model_not_offered"
    assert detail["offered"] == ["deepseek-v4-pro", "deepseek-v4-flash"]
    assert "deepseek-v4-pro" in detail["message"]
    # Not in the catalog at all.
    assert _post(client, token, "claude-opus-5").status_code == 400
    assert len(upstream.calls) == 1


def test_the_context_window_suffix_does_not_hide_a_catalog_model(hosted, catalog_env, monkeypatch):
    client, token, _ = hosted
    catalog_env(CATALOG)
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))
    assert _post(client, token, "deepseek-v4-pro[1m]").status_code == 200


def test_without_a_catalog_any_served_model_goes_through(hosted, monkeypatch):
    client, token, _ = hosted
    _install(monkeypatch, FakeUpstream([(200, STREAM_BODY, None)]))
    assert _post(client, token, "claude-opus-5").status_code == 200
