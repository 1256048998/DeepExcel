"""Operator API: separation from user accounts, auditing, dashboard."""

from __future__ import annotations

import jwt
import pytest

from app.security import ACCESS_AUDIENCE_ADMIN, ACCESS_AUDIENCE_USER, decode_access_token
from tests.conftest import ADMIN_EMAIL, ADMIN_PASSWORD, admin_token, auth_headers, register


def test_admin_is_bootstrapped_from_environment_only(admin_client):
    headers = admin_token(admin_client)
    assert admin_client.get("/admin/api/auth/me", headers=headers).json()["email"] == ADMIN_EMAIL

    # There is deliberately no public endpoint that creates the first admin:
    # such a thing is either a race to be first or a permanent hole.
    assert admin_client.post(
        "/admin/api/auth/register", json={"email": "x@y.com", "password": "z"}
    ).status_code == 404


def test_bootstrap_is_idempotent(make_client):
    """Restarting the container must not fail or duplicate the operator."""
    from sqlalchemy import func, select

    from app.db import get_session_factory
    from app.models import Admin

    client = make_client(
        BOOTSTRAP_ADMIN_EMAIL=ADMIN_EMAIL, BOOTSTRAP_ADMIN_PASSWORD=ADMIN_PASSWORD
    )
    from app.main import bootstrap_admin

    bootstrap_admin()
    bootstrap_admin()

    with get_session_factory()() as db:
        assert db.scalar(select(func.count()).select_from(Admin)) == 1


def test_user_token_cannot_reach_the_admin_api(admin_client):
    """Audience separation, the other direction.

    Admins live in their own table for the same reason: one authentication bug
    should not be enough to turn a customer into an operator.
    """
    tokens = register(admin_client)
    response = admin_client.get("/admin/api/users", headers=auth_headers(tokens))
    assert response.status_code == 401

    with pytest.raises(jwt.InvalidAudienceError):
        decode_access_token(tokens["access_token"], audience=ACCESS_AUDIENCE_ADMIN)


def test_admin_token_cannot_reach_the_user_api(admin_client):
    headers = admin_token(admin_client)
    assert admin_client.get("/api/v1/auth/me", headers=headers).status_code == 401

    token = headers["Authorization"].split(" ", 1)[1]
    with pytest.raises(jwt.InvalidAudienceError):
        decode_access_token(token, audience=ACCESS_AUDIENCE_USER)


def test_admin_login_rejects_wrong_password(admin_client):
    response = admin_client.post(
        "/admin/api/auth/login", json={"email": ADMIN_EMAIL, "password": "not-the-password"}
    )
    assert response.status_code == 401


def test_disabling_a_user_revokes_their_refresh_tokens(admin_client):
    """Without this, "disable" means nothing for up to a refresh-token lifetime."""
    tokens = register(admin_client)
    headers = auth_headers(tokens)
    user_id = admin_client.get("/api/v1/auth/me", headers=headers).json()["id"]

    admin_client.post(
        f"/admin/api/users/{user_id}/status",
        json={"status": "disabled"},
        headers=admin_token(admin_client),
    )

    assert admin_client.post(
        "/api/v1/auth/refresh", json={"refresh_token": tokens["refresh_token"]}
    ).status_code in (401, 403)
    assert admin_client.post(
        "/api/v1/auth/login",
        json={"email": "user@deepexcel-qa.com", "password": "correct-horse-battery"},
    ).status_code == 403


def test_every_mutation_is_audited(admin_client):
    tokens = register(admin_client)
    headers = admin_token(admin_client)
    user_id = admin_client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]

    admin_client.post("/admin/api/invites", json={"max_uses": 3}, headers=headers)
    admin_client.post(
        f"/admin/api/users/{user_id}/entitlement", json={"plan": "pro"}, headers=headers
    )
    admin_client.post(
        f"/admin/api/users/{user_id}/status", json={"status": "disabled"}, headers=headers
    )

    actions = [row["action"] for row in admin_client.get("/admin/api/audit", headers=headers).json()]
    # "Who disabled this account" needs an answer.
    assert {"invite.create", "entitlement.update", "user.status"} <= set(actions)


def test_user_search_and_paging(admin_client):
    for index in range(3):
        register(admin_client, email=f"person{index}@deepexcel-qa.com")
    headers = admin_token(admin_client)

    everyone = admin_client.get("/admin/api/users", headers=headers).json()
    assert len(everyone) == 3

    filtered = admin_client.get("/admin/api/users?q=person1", headers=headers).json()
    assert len(filtered) == 1
    assert filtered[0]["email"] == "person1@deepexcel-qa.com"

    paged = admin_client.get("/admin/api/users?limit=2&offset=2", headers=headers).json()
    assert len(paged) == 1


def test_dashboard_reports_task_success_rate(admin_client):
    """The north-star metric. Until M2 ships there is no way to measure it at all."""
    tokens = register(admin_client)
    user_headers = auth_headers(tokens)

    events = [
        {"event_type": "task_complete", "payload": {"outcome": "success"}},
        {"event_type": "task_complete", "payload": {"outcome": "success"}},
        {"event_type": "task_complete", "payload": {"outcome": "success"}},
        {"event_type": "task_complete", "payload": {"outcome": "error"}},
        {"event_type": "tool_error", "payload": {"tool_name": "write_range", "error_code": "oob"}},
        {"event_type": "tool_error", "payload": {"tool_name": "write_range", "error_code": "oob"}},
        {"event_type": "tool_error", "payload": {"tool_name": "create_chart", "error_code": "hresult"}},
    ]
    assert admin_client.post(
        "/api/v1/telemetry", json={"events": events}, headers=user_headers
    ).json()["accepted"] == len(events)

    admin_client.post(
        "/api/v1/session/heartbeat?install_id=inst-1&client_version=0.5.0&host=excel",
        headers=user_headers,
    )

    stats = admin_client.get("/admin/api/stats", headers=admin_token(admin_client)).json()
    assert stats["total_users"] == 1
    assert stats["tasks_7d"] == 4
    assert stats["task_success_rate_7d"] == 0.75
    assert stats["top_tool_errors_7d"][0] == {"tool": "write_range", "count": 2}
    assert stats["installs_seen_7d"] == 1
    assert stats["client_versions"][0]["version"] == "0.5.0"


def test_dashboard_reports_auto_update_health(admin_client):
    """Whether the updater is working at all.

    The client-version histogram cannot answer this: an updater that silently
    stopped working looks exactly like "nobody has upgraded yet" -- same flat
    version distribution, no errors anywhere. These counts are the only signal.

    The counting is plain enough to get wrong without noticing. Miscounting
    `started` as `installed` would inflate adoption; missing `blocked` would
    hide the users who cannot install at all. Either way the dashboard stays
    green and says nothing.
    """
    tokens = register(admin_client)
    user_headers = auth_headers(tokens)

    events = [
        # Two machines finished an upgrade.
        {"event_type": "update_event",
         "payload": {"phase": "apply", "outcome": "installed",
                     "from_version": "0.5.0", "to_version": "0.6.0"}},
        {"event_type": "update_event",
         "payload": {"phase": "apply", "outcome": "installed",
                     "from_version": "0.5.0", "to_version": "0.6.0"}},
        # One started an install and never reported back -- which is normal, the
        # process that would report is the one Excel just closed. It must NOT
        # count as an upgrade.
        {"event_type": "update_event",
         "payload": {"phase": "launch", "outcome": "started", "from_version": "0.5.0"}},
        # One has given up after three failed installs: the case that needs a human.
        {"event_type": "update_event",
         "payload": {"phase": "launch", "outcome": "blocked",
                     "reason_code": "max_attempts", "from_version": "0.5.0"}},
        # Failures, grouped by classification code.
        {"event_type": "update_event",
         "payload": {"phase": "check", "outcome": "failed",
                     "reason_code": "feed_timeout", "from_version": "0.5.0"}},
        {"event_type": "update_event",
         "payload": {"phase": "check", "outcome": "failed",
                     "reason_code": "feed_timeout", "from_version": "0.5.0"}},
        {"event_type": "update_event",
         "payload": {"phase": "download", "outcome": "failed",
                     "reason_code": "package_digest_mismatch", "from_version": "0.5.0"}},
        # Routine "already current" must not be mistaken for a problem.
        {"event_type": "update_event",
         "payload": {"phase": "check", "outcome": "up_to_date", "from_version": "0.6.0"}},
    ]
    assert admin_client.post(
        "/api/v1/telemetry", json={"events": events}, headers=user_headers
    ).json()["accepted"] == len(events)

    stats = admin_client.get("/admin/api/stats", headers=admin_token(admin_client)).json()

    assert stats["update_upgrades_7d"] == 2
    assert stats["update_blocked_7d"] == 1
    assert stats["top_update_failures_7d"] == [
        {"reason": "feed_timeout", "count": 2},
        {"reason": "package_digest_mismatch", "count": 1},
    ]


def test_dashboard_reports_no_update_activity_as_zero(admin_client):
    """Zero upgrades is a real, reportable state -- not missing data.

    Unlike the success rate, which is None until there is something to divide,
    "nobody upgraded this week" is itself the answer.
    """
    stats = admin_client.get("/admin/api/stats", headers=admin_token(admin_client)).json()
    assert stats["update_upgrades_7d"] == 0
    assert stats["update_blocked_7d"] == 0
    assert stats["top_update_failures_7d"] == []


def test_dashboard_reports_no_data_as_none_not_zero(admin_client):
    """A brand-new deployment must not read as "0% success"."""
    stats = admin_client.get("/admin/api/stats", headers=admin_token(admin_client)).json()
    assert stats["tasks_7d"] == 0
    assert stats["task_success_rate_7d"] is None


def test_heartbeat_associates_an_install_with_an_account(admin_client):
    tokens = register(admin_client)
    headers = auth_headers(tokens)

    for _ in range(2):
        response = admin_client.post(
            "/api/v1/session/heartbeat?install_id=inst-42&client_version=0.5.0&host=excel",
            headers=headers,
        )
        assert response.status_code == 204

    # Same install seen twice is one installation, not two.
    stats = admin_client.get("/admin/api/stats", headers=admin_token(admin_client)).json()
    assert stats["installs_seen_7d"] == 1


def test_admin_endpoints_require_a_token(admin_client):
    for path in ["/admin/api/users", "/admin/api/stats", "/admin/api/invites", "/admin/api/audit"]:
        assert admin_client.get(path).status_code == 401
