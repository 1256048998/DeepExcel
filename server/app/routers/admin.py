"""Operator API.

Mounted under /admin/api so a reverse proxy can restrict it by path without
knowing anything about the application.
"""

from __future__ import annotations

import datetime as dt
import json
import secrets
from collections import Counter

from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy import func, select
from sqlalchemy.orm import Session

from ..db import get_db
from ..deps import get_current_admin
from ..models import (
    AccountStatus,
    Admin,
    AuditLog,
    Entitlement,
    EntitlementStatus,
    Installation,
    InviteCode,
    Plan,
    RefreshToken,
    RoutingMode,
    TelemetryEvent,
    User,
    utcnow,
)
from ..schemas import (
    AdminLoginRequest,
    AdminTokenResponse,
    DashboardStats,
    EntitlementUpdate,
    InviteCodeCreate,
    InviteCodeView,
    UserStatusUpdate,
    UserView,
)
from ..security import (
    ACCESS_AUDIENCE_ADMIN,
    issue_access_token,
    verify_password,
)

router = APIRouter(prefix="/admin/api", tags=["admin"])

_DUMMY_ADMIN_HASH = "$scrypt$32768$8$1$AAAAAAAAAAAAAAAAAAAAAA==$" + "A" * 44


def _audit(db: Session, admin: Admin, action: str, target: str | None, detail: dict | None = None) -> None:
    db.add(
        AuditLog(
            admin_id=admin.id,
            action=action,
            target=target,
            detail=json.dumps(detail, ensure_ascii=False, sort_keys=True) if detail else None,
        )
    )


@router.post("/auth/login", response_model=AdminTokenResponse)
def admin_login(payload: AdminLoginRequest, db: Session = Depends(get_db)) -> AdminTokenResponse:
    email = payload.email.strip().lower()
    admin = db.scalar(select(Admin).where(Admin.email == email))
    password_hash = admin.password_hash if admin else _DUMMY_ADMIN_HASH
    password_ok = verify_password(payload.password, password_hash)

    if admin is None or not password_ok or admin.status is not AccountStatus.ACTIVE:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
        )
    # Admin sessions are short and have no refresh token. An operator
    # re-authenticating is cheap; a long-lived admin credential is not.
    token, expires_at = issue_access_token(
        str(admin.id), audience=ACCESS_AUDIENCE_ADMIN, ttl_seconds=8 * 3600
    )
    return AdminTokenResponse(access_token=token, expires_at=expires_at)


@router.get("/auth/me")
def admin_me(admin: Admin = Depends(get_current_admin)) -> dict:
    return {"id": admin.id, "email": admin.email, "status": admin.status.value}


# ---------------------------------------------------------------------------
# Invite codes -- the current gate on who can install
# ---------------------------------------------------------------------------

@router.post("/invites", response_model=InviteCodeView, status_code=status.HTTP_201_CREATED)
def create_invite(
    payload: InviteCodeCreate,
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> InviteCode:
    # URL-safe and unambiguous when read aloud or retyped from a chat message.
    code = secrets.token_urlsafe(9)
    invite = InviteCode(
        code=code,
        note=payload.note,
        max_uses=payload.max_uses,
        expires_at=payload.expires_at,
    )
    db.add(invite)
    db.flush()
    _audit(db, admin, "invite.create", code, {"max_uses": payload.max_uses})
    return invite


@router.get("/invites", response_model=list[InviteCodeView])
def list_invites(
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> list[InviteCode]:
    return list(db.scalars(select(InviteCode).order_by(InviteCode.created_at.desc()).limit(200)))


@router.post("/invites/{invite_id}/disable", response_model=InviteCodeView)
def disable_invite(
    invite_id: int,
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> InviteCode:
    invite = db.get(InviteCode, invite_id)
    if invite is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Invite code not found")
    invite.disabled = True
    _audit(db, admin, "invite.disable", invite.code)
    return invite


# ---------------------------------------------------------------------------
# Users
# ---------------------------------------------------------------------------

@router.get("/users", response_model=list[UserView])
def list_users(
    q: str | None = Query(default=None, max_length=320),
    limit: int = Query(default=50, ge=1, le=200),
    offset: int = Query(default=0, ge=0),
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> list[User]:
    statement = select(User).order_by(User.created_at.desc())
    if q:
        statement = statement.where(User.email.ilike(f"%{q.strip().lower()}%"))
    return list(db.scalars(statement.limit(limit).offset(offset)))


@router.get("/users/{user_id}", response_model=UserView)
def get_user(
    user_id: int,
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> User:
    user = db.get(User, user_id)
    if user is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="User not found")
    return user


@router.post("/users/{user_id}/status", response_model=UserView)
def set_user_status(
    user_id: int,
    payload: UserStatusUpdate,
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> User:
    user = db.get(User, user_id)
    if user is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="User not found")

    user.status = AccountStatus(payload.status)
    if user.status is AccountStatus.DISABLED:
        # Revoke refresh tokens too. Without this the account stays usable for
        # the lifetime of the last access token plus every refresh after it,
        # which makes "disable" mean nothing for up to a month.
        for token in user.refresh_tokens:
            if token.revoked_at is None:
                token.revoked_at = utcnow()
    _audit(db, admin, "user.status", user.email, {"status": payload.status})
    return user


@router.post("/users/{user_id}/entitlement", response_model=UserView)
def update_entitlement(
    user_id: int,
    payload: EntitlementUpdate,
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> User:
    user = db.get(User, user_id)
    if user is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="User not found")
    entitlement = user.entitlement
    if entitlement is None:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="User has no entitlement")

    changes: dict[str, object] = {}
    if payload.plan is not None:
        entitlement.plan = Plan(payload.plan)
        changes["plan"] = payload.plan
    if payload.status is not None:
        entitlement.status = EntitlementStatus(payload.status)
        changes["status"] = payload.status
    if payload.routing_mode is not None:
        entitlement.routing_mode = RoutingMode(payload.routing_mode)
        changes["routing_mode"] = payload.routing_mode
    if payload.task_limit is not None:
        entitlement.task_limit = payload.task_limit
        changes["task_limit"] = payload.task_limit
    if payload.expires_at is not None:
        entitlement.expires_at = payload.expires_at
        changes["expires_at"] = payload.expires_at.isoformat()

    _audit(db, admin, "entitlement.update", user.email, changes)
    return user


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

@router.get("/stats", response_model=DashboardStats)
def stats(
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> DashboardStats:
    since = utcnow() - dt.timedelta(days=7)

    total_users = db.scalar(select(func.count()).select_from(User)) or 0
    active_users = db.scalar(
        select(func.count()).select_from(User).where(User.status == AccountStatus.ACTIVE)
    ) or 0
    installs = db.scalar(
        select(func.count()).select_from(Installation).where(Installation.last_seen_at >= since)
    ) or 0

    task_events = list(
        db.scalars(
            select(TelemetryEvent).where(
                TelemetryEvent.event_type == "task_complete",
                TelemetryEvent.occurred_at >= since,
            )
        )
    )
    outcomes = Counter()
    for event in task_events:
        try:
            outcomes[json.loads(event.payload).get("outcome", "unknown")] += 1
        except (ValueError, TypeError):
            outcomes["unparseable"] += 1

    total_tasks = sum(outcomes.values())
    # The north-star metric. None rather than 0.0 when there is no data, so a
    # brand-new deployment does not read as "0% success".
    success_rate = (outcomes["success"] / total_tasks) if total_tasks else None

    error_events = db.scalars(
        select(TelemetryEvent).where(
            TelemetryEvent.event_type == "tool_error",
            TelemetryEvent.occurred_at >= since,
        )
    )
    tool_errors = Counter()
    for event in error_events:
        try:
            payload = json.loads(event.payload)
        except (ValueError, TypeError):
            continue
        name = payload.get("tool_name")
        if name:
            tool_errors[name] += 1

    # Auto-update health.
    #
    # The client-version histogram says where clients are; it cannot say whether
    # the updater is what put them there. An updater that silently stopped
    # working looks exactly like "nobody has upgraded yet" -- same flat version
    # distribution, no errors anywhere -- so the counts have to be read directly.
    #
    # blocked is the one worth watching. It means a client downloaded and
    # verified an update, failed to install it three times, and has given up
    # offering it. Antivirus is the usual cause, the user sees only a banner
    # telling them to install by hand, and nothing else in this dashboard
    # would ever surface it.
    update_events = db.scalars(
        select(TelemetryEvent).where(
            TelemetryEvent.event_type == "update_event",
            TelemetryEvent.occurred_at >= since,
        )
    )
    upgrades = 0
    blocked = 0
    update_failures = Counter()
    for event in update_events:
        try:
            payload = json.loads(event.payload)
        except (ValueError, TypeError):
            continue
        outcome = payload.get("outcome")
        if outcome == "installed":
            upgrades += 1
        elif outcome == "blocked":
            blocked += 1
        elif outcome == "failed":
            # reason_code is a fixed token by construction; see the privacy note
            # on update_event in routers/telemetry.py.
            update_failures[payload.get("reason_code") or "unknown"] += 1

    versions = Counter(
        version for version in db.scalars(
            select(Installation.client_version).where(Installation.last_seen_at >= since)
        ) if version
    )

    return DashboardStats(
        total_users=total_users,
        active_users=active_users,
        installs_seen_7d=installs,
        tasks_7d=total_tasks,
        task_success_rate_7d=success_rate,
        top_tool_errors_7d=[{"tool": name, "count": count} for name, count in tool_errors.most_common(10)],
        client_versions=[{"version": name, "count": count} for name, count in versions.most_common(10)],
        update_upgrades_7d=upgrades,
        update_blocked_7d=blocked,
        top_update_failures_7d=[
            {"reason": name, "count": count} for name, count in update_failures.most_common(10)
        ],
    )


@router.get("/audit")
def audit_log(
    limit: int = Query(default=100, ge=1, le=500),
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> list[dict]:
    rows = db.scalars(select(AuditLog).order_by(AuditLog.created_at.desc()).limit(limit))
    return [
        {
            "id": row.id,
            "admin_id": row.admin_id,
            "action": row.action,
            "target": row.target,
            "detail": json.loads(row.detail) if row.detail else None,
            "created_at": row.created_at,
        }
        for row in rows
    ]
