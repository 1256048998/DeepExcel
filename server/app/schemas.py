"""Wire contracts.

The important type in this file is EndpointConfig. It is the reason accounts are
being built before billing: it moves the decision "where does model traffic go"
from client code into a server response, so turning on hosted routing later is a
configuration change rather than a client release.
"""

from __future__ import annotations

import datetime as dt
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, EmailStr, Field, field_validator


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

class RegisterRequest(BaseModel):
    email: EmailStr
    password: str = Field(min_length=10, max_length=256)
    display_name: str | None = Field(default=None, max_length=100)
    invite_code: str | None = Field(default=None, max_length=64)

    @field_validator("password")
    @classmethod
    def _reject_trivial(cls, value: str) -> str:
        # A length floor alone still admits "aaaaaaaaaa". This is not a password
        # policy, just a guard against the obviously indefensible.
        if len(set(value)) < 5:
            raise ValueError("password is too repetitive")
        return value


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class RefreshRequest(BaseModel):
    refresh_token: str


class TokenPair(BaseModel):
    access_token: str
    refresh_token: str
    token_type: Literal["bearer"] = "bearer"
    expires_at: int


class EntitlementView(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    plan: str
    status: str
    # Exposed so the admin UI can show and change it. Not a leak: the client
    # already learns its own mode from EndpointConfig, and it is the server's
    # decision either way.
    routing_mode: str
    task_limit: int | None
    tasks_used: int
    tasks_remaining: int | None
    expires_at: dt.datetime | None


class UserView(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    email: str
    display_name: str | None
    status: str
    created_at: dt.datetime
    entitlement: EntitlementView | None = None


# ---------------------------------------------------------------------------
# Endpoint routing -- the M6 switch
# ---------------------------------------------------------------------------

class EndpointConfig(BaseModel):
    """Tells the client where to send model traffic.

    Contract with the client, which must be honoured on both sides:

    * The client stores no rule about which mode to use. It applies whatever
      this response says. Any `if plan == ...` on the client defeats the point,
      and since the client is unsigned and modifiable, a client-side rule is not
      an enforcement mechanism anyway.
    * `auth_header` is a complete header value, so the client never needs to
      know the scheme.
    * The client re-fetches at `expires_at`. That is the kill switch: access can
      be withdrawn within one TTL without any client release.

    Today every response is mode="byok" because HOSTED_PROXY_BASE_URL is unset.
    M6 sets it and flips entitlements; the client does not change.
    """

    mode: Literal["byok", "hosted"]

    # null in byok mode: the client keeps using its own provider configuration.
    base_url: str | None = None
    auth_header: str | None = None

    expires_at: int
    entitlement: EntitlementView
    # Advisory. Lets the server slow clients down without a client change.
    refresh_after_seconds: int


# ---------------------------------------------------------------------------
# Telemetry
# ---------------------------------------------------------------------------

class TelemetryEventIn(BaseModel):
    event_type: str = Field(max_length=40)
    occurred_at: dt.datetime | None = None
    install_id: str | None = Field(default=None, max_length=64)
    # Filtered against a server-side allowlist before storage. The client is not
    # trusted to have redacted correctly.
    payload: dict[str, Any] = Field(default_factory=dict)


class TelemetryBatch(BaseModel):
    events: list[TelemetryEventIn]


class TelemetryAccepted(BaseModel):
    accepted: int
    rejected: int
    # Named so an operator reading a client log can tell why something vanished.
    rejected_reasons: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Admin
# ---------------------------------------------------------------------------

class AdminLoginRequest(BaseModel):
    email: EmailStr
    password: str


class AdminTokenResponse(BaseModel):
    access_token: str
    token_type: Literal["bearer"] = "bearer"
    expires_at: int


class InviteCodeCreate(BaseModel):
    note: str | None = Field(default=None, max_length=200)
    max_uses: int = Field(default=1, ge=1, le=1000)
    expires_at: dt.datetime | None = None


class InviteCodeView(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    code: str
    note: str | None
    max_uses: int
    used_count: int
    expires_at: dt.datetime | None
    disabled: bool
    created_at: dt.datetime


class UserStatusUpdate(BaseModel):
    status: Literal["active", "disabled"]


class EntitlementUpdate(BaseModel):
    plan: Literal["beta", "free", "pro", "team", "byok"] | None = None
    status: Literal["active", "expired", "suspended"] | None = None
    routing_mode: Literal["byok", "hosted"] | None = None
    task_limit: int | None = Field(default=None, ge=0)
    expires_at: dt.datetime | None = None


class DashboardStats(BaseModel):
    total_users: int
    active_users: int
    installs_seen_7d: int
    tasks_7d: int
    task_success_rate_7d: float | None
    top_tool_errors_7d: list[dict[str, Any]]
    client_versions: list[dict[str, Any]]
    # Auto-update health. The version histogram above says where clients are;
    # these say whether the mechanism that moves them is working at all.
    update_upgrades_7d: int
    update_blocked_7d: int
    top_update_failures_7d: list[dict[str, Any]]


# ---------------------------------------------------------------------------
# Billing
# ---------------------------------------------------------------------------

class PlanView(BaseModel):
    plan: str
    label: str
    price_cents: int
    currency: str
    task_limit: int | None
    summary: str


class OrderCreate(BaseModel):
    plan: Literal["pro", "team", "byok"]
    months: int = Field(default=1, ge=1, le=24)


class OrderView(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    order_no: str
    plan: str
    months: int
    amount_cents: int
    currency: str
    status: str
    channel: str | None
    created_at: dt.datetime
    paid_at: dt.datetime | None
