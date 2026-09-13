"""Telemetry ingestion, with the privacy allowlist as the main subject.

The product tells users no workbook content ever leaves their machine. These
tests are what makes that a property of the system rather than a promise.
"""

from __future__ import annotations

import json

from sqlalchemy import select

from app.db import get_session_factory
from app.models import TelemetryEvent
from app.routers.telemetry import sanitize
from tests.conftest import auth_headers, register


def _stored_payloads() -> list[dict]:
    with get_session_factory()() as db:
        return [json.loads(event.payload) for event in db.scalars(select(TelemetryEvent))]


def test_accepts_an_allowlisted_event(client):
    tokens = register(client)
    response = client.post(
        "/api/v1/telemetry",
        json={
            "events": [
                {
                    "event_type": "task_complete",
                    "install_id": "abc123",
                    "payload": {
                        "session_id": "s-1",
                        "duration_ms": 4200,
                        "tool_sequence": ["read_range", "sort_data", "create_chart"],
                        "tokens_in": 1500,
                        "tokens_out": 300,
                        "provider": "anthropic",
                        "model": "claude-opus-5",
                        "outcome": "success",
                    },
                }
            ]
        },
        headers=auth_headers(tokens),
    )
    assert response.status_code == 200, response.text
    assert response.json() == {"accepted": 1, "rejected": 0, "rejected_reasons": []}

    stored = _stored_payloads()[0]
    assert stored["tool_sequence"] == ["read_range", "sort_data", "create_chart"]
    assert stored["outcome"] == "success"


def test_fields_outside_the_allowlist_never_reach_storage(client):
    """The core privacy guarantee.

    A client bug, or a modified client, must not be able to put workbook content
    into the database just by adding a field.
    """
    tokens = register(client)
    response = client.post(
        "/api/v1/telemetry",
        json={
            "events": [
                {
                    "event_type": "task_complete",
                    "payload": {
                        "outcome": "success",
                        "workbook_path": r"C:\Users\alice\2026 财务预算.xlsx",
                        "cell_values": [["华东贸易", 128000]],
                        "user_prompt": "把 A1:F200 按销售额降序排列",
                        "api_key": "sk-ant-super-secret",
                    },
                }
            ]
        },
        headers=auth_headers(tokens),
    )
    assert response.status_code == 200

    stored = _stored_payloads()[0]
    assert stored == {"outcome": "success"}

    blob = json.dumps(stored, ensure_ascii=False)
    for secret in ["alice", "财务预算", "华东贸易", "128000", "降序", "sk-ant"]:
        assert secret not in blob

    # The drop is reported rather than silent, so a misbehaving client is
    # visible instead of quietly losing data.
    reasons = " ".join(response.json()["rejected_reasons"])
    assert "dropped_fields" in reasons


def test_tool_sequence_cannot_smuggle_free_text(client):
    """tool_sequence is the one list field, so it is the obvious place to hide
    a payload. Only identifier-shaped names survive."""
    kept, dropped = sanitize(
        "task_complete",
        {"tool_sequence": ["read_range", r"C:\Users\alice\secret.xlsx", "写入了 128000 元"]},
    )
    assert kept["tool_sequence"] == ["read_range"]
    assert dropped == []

    # A list with nothing valid in it is dropped entirely rather than stored empty.
    kept, dropped = sanitize("task_complete", {"tool_sequence": ["not a tool name!"]})
    assert "tool_sequence" not in kept
    assert dropped == ["tool_sequence"]


def test_tool_error_keeps_the_code_but_not_the_message(client):
    """Raw exception text routinely quotes cell values and file paths, which is
    why only a classification code is allowed."""
    kept, dropped = sanitize(
        "tool_error",
        {"tool_name": "write_range", "error_code": "range_out_of_bounds",
         "error_message": r"Cannot write 128000 to 'C:\budget.xlsx'!Sheet1!C4"},
    )
    assert kept == {"tool_name": "write_range", "error_code": "range_out_of_bounds"}
    assert dropped == ["error_message"]


def test_startup_error_requires_a_diagnostic_code(client):
    kept, _ = sanitize("startup_error", {"diagnostic_code": "E-REG-002"})
    assert kept["diagnostic_code"] == "E-REG-002"

    # Anything not shaped like a diagnostic code is a free-text field wearing a
    # diagnostic code's name.
    kept, dropped = sanitize("startup_error", {"diagnostic_code": "registry write failed at C:\\x"})
    assert kept == {}
    assert dropped == ["diagnostic_code"]


def test_user_feedback_comment_is_allowed_because_the_user_typed_it(client):
    kept, _ = sanitize(
        "user_feedback",
        {"rating": "down", "reason_code": "wrong_cells", "comment": "它改错了 C 列"},
    )
    assert kept["comment"] == "它改错了 C 列"
    assert kept["rating"] == "down"

    kept, dropped = sanitize("user_feedback", {"rating": "maybe"})
    assert kept == {}
    assert dropped == ["rating"]


def test_unknown_event_type_is_rejected(client):
    tokens = register(client)
    response = client.post(
        "/api/v1/telemetry",
        json={"events": [{"event_type": "workbook_dump", "payload": {"rows": [[1, 2]]}}]},
        headers=auth_headers(tokens),
    )
    assert response.status_code == 200
    assert response.json()["accepted"] == 0
    assert response.json()["rejected"] == 1
    assert _stored_payloads() == []


def test_batch_size_is_capped(make_client):
    client = make_client(TELEMETRY_MAX_BATCH="5")
    tokens = register(client)
    response = client.post(
        "/api/v1/telemetry",
        json={"events": [{"event_type": "session_start", "payload": {}} for _ in range(6)]},
        headers=auth_headers(tokens),
    )
    assert response.status_code == 413


def test_future_timestamps_are_clamped(client):
    """A wrong client clock must not be able to distort every dashboard."""
    import datetime as dt

    tokens = register(client)
    far_future = (dt.datetime.now(dt.timezone.utc) + dt.timedelta(days=3650)).isoformat()
    client.post(
        "/api/v1/telemetry",
        json={"events": [{"event_type": "session_start", "occurred_at": far_future,
                          "payload": {"host": "excel"}}]},
        headers=auth_headers(tokens),
    )

    with get_session_factory()() as db:
        event = db.scalar(select(TelemetryEvent))
        occurred = event.occurred_at
        if occurred.tzinfo is None:
            occurred = occurred.replace(tzinfo=dt.timezone.utc)
        assert occurred <= dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=5)


def test_telemetry_requires_authentication(client):
    assert client.post("/api/v1/telemetry", json={"events": []}).status_code == 401


def test_disabled_telemetry_accepts_and_discards(make_client):
    """A deployment with telemetry off should not make every client log an
    error on a fixed interval."""
    client = make_client(TELEMETRY_ENABLED="false")
    tokens = register(client)
    response = client.post(
        "/api/v1/telemetry",
        json={"events": [{"event_type": "session_start", "payload": {"host": "excel"}}]},
        headers=auth_headers(tokens),
    )
    assert response.status_code == 200
    assert response.json()["accepted"] == 0
    assert _stored_payloads() == []


def test_published_schema_matches_the_enforced_allowlist(client):
    """Users are told what is collected; this keeps the claim checkable."""
    published = client.get("/api/v1/telemetry/schema").json()
    from app.routers.telemetry import ALLOWLIST

    assert set(published["event_types"]) == set(ALLOWLIST)
    for name, fields in published["event_types"].items():
        assert set(fields) == set(ALLOWLIST[name])
    assert "cell contents" in published["never_collected"]
