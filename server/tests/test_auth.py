"""Registration, login, refresh, invite gating."""

from __future__ import annotations

import time

from app.security import hash_password, verify_password
from tests.conftest import admin_token, auth_headers, register


def test_register_creates_an_entitlement(client):
    tokens = register(client)
    me = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()

    assert me["email"] == "user@deepexcel-qa.com"
    # Code asking "what is this user allowed to do" must never have to handle a
    # missing row, so registration creates one even though beta grants all.
    assert me["entitlement"]["plan"] == "beta"
    assert me["entitlement"]["status"] == "active"
    assert me["entitlement"]["tasks_remaining"] is None


def test_email_is_normalized_and_unique(client):
    register(client, email="Someone@DeepExcel-QA.com")
    duplicate = client.post(
        "/api/v1/auth/register",
        json={"email": "someone@deepexcel-qa.com", "password": "correct-horse-battery"},
    )
    assert duplicate.status_code == 409


def test_password_is_not_stored_in_plaintext(client):
    from sqlalchemy import select

    from app.db import get_session_factory
    from app.models import User

    register(client, password="a-very-real-password")
    with get_session_factory()() as db:
        user = db.scalar(select(User))
        assert "a-very-real-password" not in user.password_hash
        assert user.password_hash.startswith("$scrypt$")
        assert verify_password("a-very-real-password", user.password_hash)
        assert not verify_password("a-very-real-passwore", user.password_hash)


def test_login_rejects_wrong_password(client):
    register(client)
    response = client.post(
        "/api/v1/auth/login",
        json={"email": "user@deepexcel-qa.com", "password": "wrong-password-here"},
    )
    assert response.status_code == 401
    # The same message for an unknown email, so the endpoint cannot be used to
    # discover which addresses are registered.
    assert response.json()["detail"] == "Incorrect email or password"

    unknown = client.post(
        "/api/v1/auth/login",
        json={"email": "nobody@deepexcel-qa.com", "password": "wrong-password-here"},
    )
    assert unknown.status_code == 401
    assert unknown.json()["detail"] == response.json()["detail"]


def test_refresh_token_is_single_use(client):
    tokens = register(client)

    first = client.post("/api/v1/auth/refresh", json={"refresh_token": tokens["refresh_token"]})
    assert first.status_code == 200
    rotated = first.json()

    # Replaying the old one after rotation must fail: that is what limits the
    # damage of a stolen refresh token.
    replay = client.post("/api/v1/auth/refresh", json={"refresh_token": tokens["refresh_token"]})
    assert replay.status_code == 401

    assert client.post(
        "/api/v1/auth/refresh", json={"refresh_token": rotated["refresh_token"]}
    ).status_code == 200


def test_logout_revokes_the_refresh_token(client):
    tokens = register(client)
    assert client.post(
        "/api/v1/auth/logout", json={"refresh_token": tokens["refresh_token"]}
    ).status_code == 204
    assert client.post(
        "/api/v1/auth/refresh", json={"refresh_token": tokens["refresh_token"]}
    ).status_code == 401

    # Logging out an already-invalid token still succeeds; reporting otherwise
    # would confirm whether a token was real.
    assert client.post(
        "/api/v1/auth/logout", json={"refresh_token": "never-existed"}
    ).status_code == 204


def test_weak_passwords_are_rejected(client):
    for password in ["short", "aaaaaaaaaaaa"]:
        response = client.post(
            "/api/v1/auth/register",
            json={"email": "weak@deepexcel-qa.com", "password": password},
        )
        assert response.status_code == 422, password


# ---------------------------------------------------------------------------
# Invite gating -- today's answer to "who is allowed to install this"
# ---------------------------------------------------------------------------

def test_invite_required_blocks_open_registration(make_client):
    client = make_client(REQUIRE_INVITE_CODE="true")
    assert client.get("/api/v1/auth/registration-policy").json()["invite_required"] is True

    response = client.post(
        "/api/v1/auth/register",
        json={"email": "nope@deepexcel-qa.com", "password": "correct-horse-battery"},
    )
    assert response.status_code == 403

    bad = client.post(
        "/api/v1/auth/register",
        json={
            "email": "nope@deepexcel-qa.com",
            "password": "correct-horse-battery",
            "invite_code": "made-up",
        },
    )
    assert bad.status_code == 403
    # Identical message for "no such code" and "already used", so the endpoint
    # cannot be used to enumerate valid codes.
    assert bad.json()["detail"] == "Invite code is not valid"


def test_invite_code_respects_max_uses(make_client):
    from tests.conftest import ADMIN_EMAIL, ADMIN_PASSWORD

    client = make_client(
        REQUIRE_INVITE_CODE="true",
        BOOTSTRAP_ADMIN_EMAIL=ADMIN_EMAIL,
        BOOTSTRAP_ADMIN_PASSWORD=ADMIN_PASSWORD,
    )
    headers = admin_token(client)
    code = client.post(
        "/admin/api/invites", json={"max_uses": 2, "note": "first testers"}, headers=headers
    ).json()["code"]

    for index in range(2):
        response = client.post(
            "/api/v1/auth/register",
            json={
                "email": f"tester{index}@deepexcel-qa.com",
                "password": "correct-horse-battery",
                "invite_code": code,
            },
        )
        assert response.status_code == 201, response.text

    exhausted = client.post(
        "/api/v1/auth/register",
        json={
            "email": "tester2@deepexcel-qa.com",
            "password": "correct-horse-battery",
            "invite_code": code,
        },
    )
    assert exhausted.status_code == 403

    listed = client.get("/admin/api/invites", headers=headers).json()[0]
    assert listed["used_count"] == 2


def test_disabled_invite_cannot_be_used(make_client):
    from tests.conftest import ADMIN_EMAIL, ADMIN_PASSWORD

    client = make_client(
        REQUIRE_INVITE_CODE="true",
        BOOTSTRAP_ADMIN_EMAIL=ADMIN_EMAIL,
        BOOTSTRAP_ADMIN_PASSWORD=ADMIN_PASSWORD,
    )
    headers = admin_token(client)
    invite = client.post("/admin/api/invites", json={"max_uses": 5}, headers=headers).json()
    client.post(f"/admin/api/invites/{invite['id']}/disable", headers=headers)

    response = client.post(
        "/api/v1/auth/register",
        json={
            "email": "late@deepexcel-qa.com",
            "password": "correct-horse-battery",
            "invite_code": invite["code"],
        },
    )
    assert response.status_code == 403


def test_expired_access_token_is_rejected(make_client):
    client = make_client(ACCESS_TOKEN_TTL_SECONDS="1")
    tokens = register(client)
    time.sleep(2)
    assert client.get("/api/v1/auth/me", headers=auth_headers(tokens)).status_code == 401


def test_hashing_is_salted(client):
    first = hash_password("identical-password")
    second = hash_password("identical-password")
    # Equal hashes would mean a missing salt, making the whole table vulnerable
    # to one precomputation.
    assert first != second
    assert verify_password("identical-password", first)
    assert verify_password("identical-password", second)


def test_malformed_hash_does_not_crash_verification():
    for broken in ["", "not-a-hash", "$scrypt$bad", "$bcrypt$1$2$3$4$5"]:
        assert verify_password("anything", broken) is False
