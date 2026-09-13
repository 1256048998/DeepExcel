"""Where does model traffic go?

This is the one decision in M2 that is expensive to change later, so it is made
here and only here. The client asks the server on every session and caches the
answer until it expires. It contains no rule of its own.

Two reasons it has to work this way:

1. The client is unsigned and modifiable, so any entitlement check living in the
   client is decoration, not enforcement.
2. Turning on hosted routing in M6 must not require a client release. With this
   in place it is an environment variable plus an entitlement flip.
"""

from __future__ import annotations

import datetime as dt
import time

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from ..config import get_settings
from ..db import get_db
from ..deps import get_current_user
from ..models import Entitlement, Installation, RoutingMode, User, utcnow
from ..schemas import EndpointConfig, EntitlementView
from ..security import ACCESS_AUDIENCE_PROXY, issue_access_token

router = APIRouter(prefix="/api/v1/session", tags=["session"])


def _entitlement_view(entitlement: Entitlement) -> EntitlementView:
    return EntitlementView(
        plan=entitlement.plan.value,
        status=entitlement.status.value,
        routing_mode=entitlement.routing_mode.value,
        task_limit=entitlement.task_limit,
        tasks_used=entitlement.tasks_used,
        tasks_remaining=entitlement.tasks_remaining,
        expires_at=entitlement.expires_at,
    )


def resolve_endpoint(user: User, entitlement: Entitlement) -> EndpointConfig:
    """Pure resolution logic, kept free of request handling so it can be tested
    directly against every combination of plan, mode and configuration."""
    settings = get_settings()

    if not entitlement.is_active:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail={
                "reason": "entitlement_inactive",
                "status": entitlement.status.value,
                "message": "订阅已过期或被暂停，请联系支持或续订。",
            },
        )

    wants_hosted = entitlement.routing_mode is RoutingMode.HOSTED
    proxy_configured = bool(settings.hosted_proxy_base_url)

    if wants_hosted and not proxy_configured:
        # Do NOT silently fall back to BYOK. A user provisioned for hosted has
        # no provider key configured locally, so BYOK would leave them with
        # errors they cannot diagnose. An explicit 503 points at the real cause,
        # which is our deployment.
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={
                "reason": "hosted_routing_unavailable",
                "message": "服务端未配置托管转发地址（HOSTED_PROXY_BASE_URL）。",
            },
        )

    ttl = settings.endpoint_config_ttl_seconds
    expires_at = int(time.time()) + ttl

    if not wants_hosted:
        return EndpointConfig(
            mode="byok",
            base_url=None,
            auth_header=None,
            expires_at=expires_at,
            entitlement=_entitlement_view(entitlement),
            refresh_after_seconds=ttl,
        )

    # Quota only applies to hosted traffic. A BYOK user pays their provider
    # directly, so counting their tasks would be charging for nothing.
    remaining = entitlement.tasks_remaining
    if remaining is not None and remaining <= 0:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail={
                "reason": "quota_exhausted",
                "task_limit": entitlement.task_limit,
                "tasks_used": entitlement.tasks_used,
                "message": "本期任务额度已用完。",
            },
        )

    # The proxy token is a separate audience with a short life. It is handed to
    # a different component than the account API token and must not be usable
    # against account endpoints even if it leaks.
    proxy_token, proxy_expires_at = issue_access_token(
        str(user.id),
        audience=ACCESS_AUDIENCE_PROXY,
        ttl_seconds=ttl,
        extra={"plan": entitlement.plan.value},
    )
    return EndpointConfig(
        mode="hosted",
        base_url=settings.hosted_proxy_base_url,
        # A complete header value: the client never has to know the scheme.
        auth_header=f"Bearer {proxy_token}",
        expires_at=min(expires_at, proxy_expires_at),
        entitlement=_entitlement_view(entitlement),
        refresh_after_seconds=ttl,
    )


@router.get("/endpoint", response_model=EndpointConfig)
def get_endpoint(
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> EndpointConfig:
    entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user.id))
    if entitlement is None:
        # Registration always creates one; a missing row means data corruption
        # rather than a normal state, so it is not silently repaired here.
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Account has no entitlement record",
        )
    user.last_seen_at = utcnow()
    return resolve_endpoint(user, entitlement)


@router.post("/heartbeat", status_code=status.HTTP_204_NO_CONTENT)
def heartbeat(
    install_id: str,
    client_version: str | None = None,
    host: str | None = None,
    office_version: str | None = None,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> None:
    """Associates an installation with an account.

    This is how "how many people actually use this, on what" becomes answerable.
    install_id is a random value generated by the client; it deliberately carries
    no machine fingerprint.
    """
    install_id = install_id.strip()[:64]
    if not install_id:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="install_id is required")

    record = db.scalar(select(Installation).where(Installation.install_id == install_id))
    now = utcnow()
    if record is None:
        record = Installation(install_id=install_id, first_seen_at=now)
        db.add(record)
    record.user_id = user.id
    record.client_version = (client_version or record.client_version or "")[:32] or None
    record.host = (host or record.host or "")[:16] or None
    record.office_version = (office_version or record.office_version or "")[:32] or None
    record.last_seen_at = now
    user.last_seen_at = now


@router.post("/consume-task", status_code=status.HTTP_200_OK)
def consume_task(
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """Records one completed task against the quota.

    Client-reported and therefore spoofable. That is acceptable while every
    account is unmetered: it exists so the counter, the period reset and the
    exhaustion response are exercised before money depends on them. When M6
    turns on hosted routing the proxy becomes the authoritative counter, because
    it is the only place the client cannot bypass.
    """
    entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user.id))
    if entitlement is None:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                            detail="Account has no entitlement record")

    _roll_period_if_needed(entitlement)
    entitlement.tasks_used += 1
    return {
        "tasks_used": entitlement.tasks_used,
        "tasks_remaining": entitlement.tasks_remaining,
    }


def _roll_period_if_needed(entitlement: Entitlement) -> None:
    """Monthly quota window. Rolls forward lazily on first use after expiry."""
    started = entitlement.period_started_at
    if started.tzinfo is None:
        started = started.replace(tzinfo=dt.timezone.utc)
    if utcnow() - started >= dt.timedelta(days=30):
        entitlement.period_started_at = utcnow()
        entitlement.tasks_used = 0
