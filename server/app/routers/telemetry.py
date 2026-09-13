"""Telemetry ingestion.

The privacy rule -- never cell contents, workbook names, file paths, user input
or API keys -- is enforced here, on the server, with an allowlist. The client
also filters, but the client is unsigned and modifiable, so its filtering is a
convenience rather than a guarantee. Anything not named below is dropped before
it reaches storage.

Dropping rather than rejecting is deliberate: a client that starts sending a new
field should not lose the events we do understand, and the drop is reported back
so it shows up rather than disappearing.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from ..config import get_settings
from ..db import get_db
from ..deps import get_current_user
from ..models import TelemetryEvent, User, utcnow
from ..schemas import TelemetryAccepted, TelemetryBatch

router = APIRouter(prefix="/api/v1/telemetry", tags=["telemetry"])

_MAX_STRING = 200
_MAX_LIST = 60
_TOOL_NAME = re.compile(r"^[a-z][a-z0-9_]{0,39}$")
_DIAGNOSTIC_CODE = re.compile(r"^E-[A-Z]+-\d{3}$")
_OUTCOMES = {"success", "error", "cancelled", "clarify"}


def _string(value: Any, max_length: int = _MAX_STRING) -> str | None:
    if not isinstance(value, str):
        return None
    cleaned = value.strip()
    return cleaned[:max_length] if cleaned else None


def _int(value: Any, minimum: int = 0, maximum: int = 10_000_000) -> int | None:
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value if minimum <= value <= maximum else None


def _enum(value: Any, allowed: set[str]) -> str | None:
    text = _string(value, 40)
    return text if text in allowed else None


def _tool_names(value: Any) -> list[str] | None:
    """Tool names only. Rejecting anything that is not an identifier is what
    stops a payload being smuggled through this field."""
    if not isinstance(value, list):
        return None
    names = [item for item in value[:_MAX_LIST] if isinstance(item, str) and _TOOL_NAME.match(item)]
    return names or None


def _diagnostic_code(value: Any) -> str | None:
    text = _string(value, 20)
    return text if text and _DIAGNOSTIC_CODE.match(text) else None


def _free_text(value: Any) -> str | None:
    """Only for text the user deliberately typed into a feedback box."""
    return _string(value, 1000)


# event_type -> {field: coercer}. Adding a field here is a privacy decision and
# should be reviewed as one.
ALLOWLIST: dict[str, dict[str, Any]] = {
    "session_start": {
        "client_version": lambda v: _string(v, 32),
        "os_version": lambda v: _string(v, 64),
        "office_version": lambda v: _string(v, 32),
        "office_bitness": lambda v: _enum(v, {"32", "64"}),
        "host": lambda v: _enum(v, {"excel", "wps"}),
    },
    "task_complete": {
        "session_id": lambda v: _string(v, 64),
        "duration_ms": _int,
        "turn_count": lambda v: _int(v, 0, 1000),
        "tool_sequence": _tool_names,
        "tokens_in": _int,
        "tokens_out": _int,
        "provider": lambda v: _string(v, 40),
        "model": lambda v: _string(v, 80),
        "outcome": lambda v: _enum(v, _OUTCOMES),
    },
    "tool_error": {
        "tool_name": lambda v: (_string(v, 40) if isinstance(v, str) and _TOOL_NAME.match(v) else None),
        # A classification code, never the raw exception message: error text
        # routinely quotes cell values and file paths.
        "error_code": lambda v: _string(v, 40),
    },
    "startup_error": {
        "diagnostic_code": _diagnostic_code,
        "client_version": lambda v: _string(v, 32),
    },
    "user_feedback": {
        "rating": lambda v: _enum(v, {"up", "down"}),
        "reason_code": lambda v: _string(v, 40),
        "comment": _free_text,
        "session_id": lambda v: _string(v, 64),
    },
}


def sanitize(event_type: str, payload: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """Returns (kept_fields, dropped_field_names)."""
    schema = ALLOWLIST.get(event_type)
    if schema is None:
        raise KeyError(event_type)

    kept: dict[str, Any] = {}
    dropped: list[str] = []
    for key, value in payload.items():
        coercer = schema.get(key)
        if coercer is None:
            dropped.append(key)
            continue
        coerced = coercer(value)
        if coerced is None:
            dropped.append(key)
            continue
        kept[key] = coerced
    return kept, dropped


@router.post("", response_model=TelemetryAccepted)
def ingest(
    batch: TelemetryBatch,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> TelemetryAccepted:
    settings = get_settings()
    if not settings.telemetry_enabled:
        # Accept and discard. A deployment that turns telemetry off should not
        # make every client log an error on a fixed interval.
        return TelemetryAccepted(accepted=0, rejected=len(batch.events),
                                 rejected_reasons=["telemetry_disabled"])

    if len(batch.events) > settings.telemetry_max_batch:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"At most {settings.telemetry_max_batch} events per batch",
        )

    accepted = 0
    reasons: list[str] = []
    now = utcnow()

    for event in batch.events:
        try:
            kept, dropped = sanitize(event.event_type, event.payload)
        except KeyError:
            reasons.append(f"unknown_event_type:{event.event_type[:40]}")
            continue

        if dropped:
            # Surfaced rather than silent, so a client sending an unexpected
            # field is visible instead of quietly losing data.
            reasons.append(f"dropped_fields:{event.event_type}:{','.join(sorted(dropped)[:5])}")

        occurred_at = event.occurred_at or now
        if occurred_at.tzinfo is None:
            occurred_at = occurred_at.replace(tzinfo=dt.timezone.utc)
        # Clock skew and replayed buffers are normal; a client clock must not be
        # able to write events into the future and distort every dashboard.
        if occurred_at > now + dt.timedelta(minutes=5):
            occurred_at = now

        db.add(
            TelemetryEvent(
                user_id=user.id,
                install_id=(event.install_id or "")[:64] or None,
                event_type=event.event_type,
                payload=json.dumps(kept, ensure_ascii=False, sort_keys=True),
                occurred_at=occurred_at,
                received_at=now,
            )
        )
        accepted += 1

    user.last_seen_at = now
    return TelemetryAccepted(
        accepted=accepted,
        rejected=len(batch.events) - accepted,
        rejected_reasons=reasons[:20],
    )


@router.get("/schema")
def telemetry_schema() -> dict:
    """The allowlist, published.

    Users are told we collect no workbook content. This endpoint makes that
    claim checkable instead of asking them to take our word for it.
    """
    return {
        "event_types": {name: sorted(fields) for name, fields in ALLOWLIST.items()},
        "never_collected": [
            "cell contents",
            "workbook or file names and paths",
            "user prompt text",
            "model responses",
            "API keys or credentials",
            "raw error messages",
        ],
    }
