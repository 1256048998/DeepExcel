"""Model proxy.

Anthropic-protocol passthrough with metering, quota enforcement and upstream
failover. This is what makes hosted routing real: it is the one point a client
cannot bypass, so it is the only place where "how much did this user use" and
"may this user continue" can actually be answered.

Streaming is forwarded chunk by chunk. Buffering the response to read the usage
totals at the end would turn a streaming API into a blocking one and remove the
reason streaming exists.
"""

from __future__ import annotations

import json
import time
import uuid

import httpx
import jwt
from fastapi import APIRouter, Depends, HTTPException, Request, status
from fastapi.responses import StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from sqlalchemy import select
from sqlalchemy.orm import Session

from ..config import get_settings
from ..db import get_db, get_session_factory
from ..models import (
    AccountStatus,
    Entitlement,
    RoutingMode,
    UsageRecord,
    User,
    utcnow,
)
from ..security import ACCESS_AUDIENCE_PROXY, decode_access_token
from .metering import StreamUsageCollector, Usage, estimate_cost_usd, usage_from_payload
from .upstream import load_upstreams, select as select_upstreams

router = APIRouter(prefix="/v1", tags=["proxy"])

_bearer = HTTPBearer(auto_error=False)

# Long enough for a slow model, short enough that a hung upstream does not pin a
# worker forever.
_CONNECT_TIMEOUT = 10.0
_READ_TIMEOUT = 600.0

# Hop-by-hop and authentication headers must not be forwarded: the client's
# credential is ours, not the upstream's.
_STRIPPED_REQUEST_HEADERS = {
    "host", "authorization", "x-api-key", "content-length",
    "connection", "keep-alive", "transfer-encoding", "upgrade",
    "accept-encoding",
}
_STRIPPED_RESPONSE_HEADERS = {
    "content-length", "content-encoding", "connection",
    "keep-alive", "transfer-encoding", "upgrade",
}


def proxy_user(
    credentials: HTTPAuthorizationCredentials | None = Depends(_bearer),
    db: Session = Depends(get_db),
) -> User:
    """Authenticates a proxy token.

    Deliberately a different audience from the account API token: a leak at the
    proxy must not become an account takeover.
    """
    if credentials is None or not credentials.credentials:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token",
            headers={"WWW-Authenticate": "Bearer"},
        )
    try:
        claims = decode_access_token(credentials.credentials, audience=ACCESS_AUDIENCE_PROXY)
    except jwt.PyJWTError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or expired token",
            headers={"WWW-Authenticate": "Bearer"},
        ) from exc

    user = db.get(User, int(claims.subject))
    if user is None or user.status is not AccountStatus.ACTIVE:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Account disabled")
    return user


def _check_quota(db: Session, user: User) -> Entitlement:
    entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user.id))
    if entitlement is None or not entitlement.is_active:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail={"reason": "entitlement_inactive", "message": "订阅已过期或被暂停。"},
        )
    if entitlement.routing_mode is not RoutingMode.HOSTED:
        # A BYOK user has no business here; their traffic should never reach us.
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail={"reason": "not_hosted", "message": "当前账号未启用托管转发。"},
        )
    remaining = entitlement.tasks_remaining
    if remaining is not None and remaining <= 0:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail={
                "reason": "quota_exhausted",
                "task_limit": entitlement.task_limit,
                "tasks_used": entitlement.tasks_used,
                "message": "本期额度已用完。",
            },
        )
    return entitlement


def _record_usage(
    user_id: int, model: str, upstream_name: str, usage: Usage,
    status_code: int, duration_ms: int,
) -> None:
    """Writes the usage row on its own session.

    Separate from the request session because it runs after the response has
    been streamed, by which time the request's session is gone. Failures here
    must not surface to the user: the model call already succeeded, and losing
    one usage row is cheaper than failing a request that worked.
    """
    try:
        with get_session_factory()() as db:
            db.add(
                UsageRecord(
                    user_id=user_id,
                    model=model,
                    upstream=upstream_name,
                    input_tokens=usage.input_tokens,
                    output_tokens=usage.output_tokens,
                    cache_read_tokens=usage.cache_read_tokens,
                    cache_write_tokens=usage.cache_write_tokens,
                    cost_usd=estimate_cost_usd(model, usage),
                    status_code=status_code,
                    duration_ms=duration_ms,
                )
            )
            if status_code < 400:
                # One successful call counts as one task against the quota.
                entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id))
                if entitlement is not None:
                    entitlement.tasks_used += 1
            db.commit()
    except Exception:  # pragma: no cover - accounting must never break serving
        pass


def _forward_headers(request: Request) -> dict[str, str]:
    return {
        key: value
        for key, value in request.headers.items()
        if key.lower() not in _STRIPPED_REQUEST_HEADERS
    }


@router.post("/messages")
async def messages(
    request: Request,
    user: User = Depends(proxy_user),
    db: Session = Depends(get_db),
):
    settings = get_settings()
    if not settings.hosted_proxy_base_url:
        # The proxy is mounted but hosted routing is not turned on. Serving
        # anyway would bill traffic the deployment has not opted into.
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={"reason": "hosted_routing_disabled", "message": "托管转发未启用。"},
        )

    _check_quota(db, user)

    body = await request.body()
    try:
        payload = json.loads(body) if body else {}
    except ValueError:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid JSON body")

    model = str(payload.get("model") or "")
    if not model:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="model is required")

    candidates = select_upstreams(load_upstreams(), model)
    if not candidates:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail={"reason": "no_upstream", "message": f"没有可服务模型 {model} 的上游。"},
        )

    streaming = bool(payload.get("stream"))
    headers = _forward_headers(request)
    user_id = user.id
    started = time.monotonic()

    last_error: str | None = None
    for upstream in candidates:
        client = httpx.AsyncClient(
            timeout=httpx.Timeout(_READ_TIMEOUT, connect=_CONNECT_TIMEOUT)
        )
        upstream_headers = dict(headers)
        upstream_headers["x-api-key"] = upstream.api_key
        upstream_headers.setdefault("anthropic-version", "2023-06-01")

        try:
            upstream_request = client.build_request(
                "POST", f"{upstream.base_url}/v1/messages",
                headers=upstream_headers, content=body,
            )
            response = await client.send(upstream_request, stream=True)
        except httpx.HTTPError as exc:
            await client.aclose()
            last_error = f"{upstream.name}: {type(exc).__name__}"
            continue

        # Only 5xx is worth failing over. A 4xx is the client's request being
        # wrong, and retrying it against another provider just produces the same
        # error twice while doubling the latency.
        if response.status_code >= 500 and upstream is not candidates[-1]:
            await response.aclose()
            await client.aclose()
            last_error = f"{upstream.name}: HTTP {response.status_code}"
            continue

        response_headers = {
            key: value
            for key, value in response.headers.items()
            if key.lower() not in _STRIPPED_RESPONSE_HEADERS
        }

        if streaming:
            collector = StreamUsageCollector()

            async def body_iterator():
                try:
                    async for chunk in response.aiter_bytes():
                        collector.feed(chunk)
                        yield chunk
                finally:
                    await response.aclose()
                    await client.aclose()
                    _record_usage(
                        user_id, model, upstream.name, collector.usage,
                        response.status_code, int((time.monotonic() - started) * 1000),
                    )

            return StreamingResponse(
                body_iterator(),
                status_code=response.status_code,
                headers=response_headers,
                media_type=response.headers.get("content-type", "text/event-stream"),
            )

        content = await response.aread()
        await response.aclose()
        await client.aclose()

        usage = Usage()
        try:
            usage = usage_from_payload(json.loads(content))
        except (ValueError, TypeError):
            pass
        _record_usage(
            user_id, model, upstream.name, usage,
            response.status_code, int((time.monotonic() - started) * 1000),
        )

        from fastapi.responses import Response

        return Response(
            content=content,
            status_code=response.status_code,
            headers=response_headers,
            media_type=response.headers.get("content-type", "application/json"),
        )

    raise HTTPException(
        status_code=status.HTTP_502_BAD_GATEWAY,
        detail={"reason": "all_upstreams_failed", "message": last_error or "上游全部不可用。"},
    )


@router.get("/usage")
def my_usage(
    user: User = Depends(proxy_user),
    db: Session = Depends(get_db),
) -> dict:
    """Lets a client show its own consumption without the account token."""
    import datetime as dt
    from sqlalchemy import func

    since = utcnow() - dt.timedelta(days=30)
    row = db.execute(
        select(
            func.count(),
            func.coalesce(func.sum(UsageRecord.input_tokens), 0),
            func.coalesce(func.sum(UsageRecord.output_tokens), 0),
            func.coalesce(func.sum(UsageRecord.cost_usd), 0.0),
        ).where(UsageRecord.user_id == user.id, UsageRecord.created_at >= since)
    ).one()

    entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user.id))
    return {
        "period_days": 30,
        "calls": row[0],
        "input_tokens": int(row[1]),
        "output_tokens": int(row[2]),
        "cost_usd": round(float(row[3]), 4),
        "tasks_used": entitlement.tasks_used if entitlement else 0,
        "tasks_remaining": entitlement.tasks_remaining if entitlement else None,
    }
