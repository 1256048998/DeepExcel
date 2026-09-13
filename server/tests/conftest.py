"""Test fixtures.

Each test gets its own SQLite file and its own process-level settings, because
the endpoint-routing tests need to change deployment configuration (whether a
hosted proxy exists at all) and observe the effect.
"""

from __future__ import annotations

import os
import sys
import tempfile
from collections.abc import Iterator
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import config, db  # noqa: E402
from app.models import Admin  # noqa: E402
from app.security import hash_password  # noqa: E402

ADMIN_EMAIL = "ops@deepexcel-qa.com"
ADMIN_PASSWORD = "admin-password-123456"


def _reset(**env: str | None) -> None:
    for key in [
        "DEEPEXCEL_ENV", "DATABASE_URL", "JWT_SECRET", "REQUIRE_INVITE_CODE",
        "HOSTED_PROXY_BASE_URL", "BOOTSTRAP_ADMIN_EMAIL", "BOOTSTRAP_ADMIN_PASSWORD",
        "TELEMETRY_ENABLED", "TELEMETRY_MAX_BATCH", "ENDPOINT_CONFIG_TTL_SECONDS",
        "ACCESS_TOKEN_TTL_SECONDS", "REFRESH_TOKEN_TTL_SECONDS", "ADMIN_ORIGIN",
        "UPDATE_MANIFEST_DIR",
    ]:
        os.environ.pop(key, None)
    for key, value in env.items():
        if value is not None:
            os.environ[key] = value
    config.reset_settings_for_tests()
    db.reset_engine_for_tests()
    # The feed caches by mtime, which cannot distinguish two writes inside one
    # filesystem tick.
    from app.routers import updates as updates_router

    updates_router.reset_cache()


@pytest.fixture
def make_client(tmp_path: Path):
    """Builds a client with a chosen deployment configuration."""
    created: list[TestClient] = []

    def _factory(**env: str | None) -> TestClient:
        # DEEPEXCEL_TEST_DATABASE_URL points the suite at a real Postgres.
        # SQLite is convenient but it silently accepts things Postgres rejects
        # -- a redundant unique constraint went unnoticed for exactly that
        # reason -- so the suite has to be runnable against the real thing.
        external = os.environ.get("DEEPEXCEL_TEST_DATABASE_URL")
        if external:
            env.setdefault("DATABASE_URL", external)
        else:
            database = tmp_path / f"test-{len(created)}.db"
            env.setdefault("DATABASE_URL", f"sqlite:///{database.as_posix()}")
        env.setdefault("REQUIRE_INVITE_CODE", "false")
        env.setdefault("JWT_SECRET", "test-secret-" + "x" * 40)
        _reset(**env)

        from app import db as db_module
        from app.main import create_app
        from app.models import Base

        if external:
            # Drop and recreate so every client starts from an empty schema.
            Base.metadata.drop_all(bind=db_module.get_engine())

        client = TestClient(create_app())
        client.__enter__()  # triggers lifespan: create_all + bootstrap admin
        if external:
            Base.metadata.create_all(bind=db_module.get_engine())
        created.append(client)
        return client

    yield _factory

    for client in created:
        client.__exit__(None, None, None)
    _reset()


@pytest.fixture
def client(make_client) -> TestClient:
    return make_client()


@pytest.fixture
def admin_client(make_client) -> TestClient:
    return make_client(
        BOOTSTRAP_ADMIN_EMAIL=ADMIN_EMAIL,
        BOOTSTRAP_ADMIN_PASSWORD=ADMIN_PASSWORD,
    )


def register(client: TestClient, email: str = "user@deepexcel-qa.com",
             password: str = "correct-horse-battery", **extra) -> dict:
    response = client.post(
        "/api/v1/auth/register",
        json={"email": email, "password": password, **extra},
    )
    assert response.status_code == 201, response.text
    return response.json()


def auth_headers(tokens: dict) -> dict:
    return {"Authorization": f"Bearer {tokens['access_token']}"}


def admin_token(client: TestClient) -> dict:
    response = client.post(
        "/admin/api/auth/login",
        json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD},
    )
    assert response.status_code == 200, response.text
    return {"Authorization": f"Bearer {response.json()['access_token']}"}


def set_routing_mode(client: TestClient, headers: dict, user_id: int, mode: str) -> None:
    response = client.post(
        f"/admin/api/users/{user_id}/entitlement",
        json={"routing_mode": mode},
        headers=headers,
    )
    assert response.status_code == 200, response.text


@pytest.fixture(autouse=True)
def _isolate_settings() -> Iterator[None]:
    yield
    config.reset_settings_for_tests()
    db.reset_engine_for_tests()
