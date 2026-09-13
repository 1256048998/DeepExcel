"""Database schema.

The entitlement table exists from day one even though every account is on the
same free "beta" plan. Adding it later would mean revisiting every code path
that needs to ask "is this user allowed to do this", which is exactly the kind
of retrofit that makes accounts expensive to bolt on late.
"""

from __future__ import annotations

import datetime as dt
import enum

from sqlalchemy import (
    Boolean,
    DateTime,
    Enum,
    Float,
    ForeignKey,
    Index,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


def utcnow() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


class Base(DeclarativeBase):
    pass


class AccountStatus(str, enum.Enum):
    ACTIVE = "active"
    DISABLED = "disabled"


class Plan(str, enum.Enum):
    """Only BETA is issued today; the rest are the M6 pricing tiers."""

    BETA = "beta"
    FREE = "free"
    PRO = "pro"
    TEAM = "team"
    BYOK = "byok"


class EntitlementStatus(str, enum.Enum):
    ACTIVE = "active"
    EXPIRED = "expired"
    SUSPENDED = "suspended"


class RoutingMode(str, enum.Enum):
    """Where the client should send model traffic.

    BYOK   -- the client uses its own provider configuration and key. Model
              traffic never touches our infrastructure.
    HOSTED -- the client routes through our proxy, which is what makes metering,
              quotas and revocation possible. Requires HOSTED_PROXY_BASE_URL.
    """

    BYOK = "byok"
    HOSTED = "hosted"


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(primary_key=True)
    email: Mapped[str] = mapped_column(String(320), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(255))
    display_name: Mapped[str | None] = mapped_column(String(100), default=None)
    status: Mapped[AccountStatus] = mapped_column(
        Enum(AccountStatus, native_enum=False), default=AccountStatus.ACTIVE
    )
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    last_seen_at: Mapped[dt.datetime | None] = mapped_column(
        DateTime(timezone=True), default=None
    )
    invite_code_id: Mapped[int | None] = mapped_column(
        ForeignKey("invite_codes.id"), default=None
    )

    entitlement: Mapped["Entitlement"] = relationship(
        back_populates="user", uselist=False, cascade="all, delete-orphan"
    )
    refresh_tokens: Mapped[list["RefreshToken"]] = relationship(
        back_populates="user", cascade="all, delete-orphan"
    )


class InviteCode(Base):
    __tablename__ = "invite_codes"

    id: Mapped[int] = mapped_column(primary_key=True)
    code: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    note: Mapped[str | None] = mapped_column(String(200), default=None)
    max_uses: Mapped[int] = mapped_column(Integer, default=1)
    used_count: Mapped[int] = mapped_column(Integer, default=0)
    expires_at: Mapped[dt.datetime | None] = mapped_column(
        DateTime(timezone=True), default=None
    )
    disabled: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    def is_usable(self, now: dt.datetime | None = None) -> bool:
        now = now or utcnow()
        if self.disabled or self.used_count >= self.max_uses:
            return False
        if self.expires_at is not None and _as_utc(self.expires_at) <= now:
            return False
        return True


class Entitlement(Base):
    __tablename__ = "entitlements"

    id: Mapped[int] = mapped_column(primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), unique=True, index=True)
    plan: Mapped[Plan] = mapped_column(Enum(Plan, native_enum=False), default=Plan.BETA)
    status: Mapped[EntitlementStatus] = mapped_column(
        Enum(EntitlementStatus, native_enum=False), default=EntitlementStatus.ACTIVE
    )
    routing_mode: Mapped[RoutingMode] = mapped_column(
        Enum(RoutingMode, native_enum=False), default=RoutingMode.BYOK
    )
    # NULL means unlimited. Counting is in whole tasks, not tokens, because that
    # is the unit a user can actually predict.
    task_limit: Mapped[int | None] = mapped_column(Integer, default=None)
    tasks_used: Mapped[int] = mapped_column(Integer, default=0)
    period_started_at: Mapped[dt.datetime] = mapped_column(
        DateTime(timezone=True), default=utcnow
    )
    expires_at: Mapped[dt.datetime | None] = mapped_column(
        DateTime(timezone=True), default=None
    )
    updated_at: Mapped[dt.datetime] = mapped_column(
        DateTime(timezone=True), default=utcnow, onupdate=utcnow
    )

    user: Mapped[User] = relationship(back_populates="entitlement")

    @property
    def is_active(self) -> bool:
        if self.status is not EntitlementStatus.ACTIVE:
            return False
        if self.expires_at is not None and _as_utc(self.expires_at) <= utcnow():
            return False
        return True

    @property
    def tasks_remaining(self) -> int | None:
        if self.task_limit is None:
            return None
        return max(0, self.task_limit - self.tasks_used)


class RefreshToken(Base):
    __tablename__ = "refresh_tokens"

    id: Mapped[int] = mapped_column(primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    token_hash: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    expires_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True))
    revoked_at: Mapped[dt.datetime | None] = mapped_column(
        DateTime(timezone=True), default=None
    )

    user: Mapped[User] = relationship(back_populates="refresh_tokens")

    @property
    def is_valid(self) -> bool:
        return self.revoked_at is None and _as_utc(self.expires_at) > utcnow()


class Admin(Base):
    """Deliberately separate from users.

    Sharing one table and a role column means one authentication bug is enough
    to turn a customer account into an operator account.
    """

    __tablename__ = "admins"

    id: Mapped[int] = mapped_column(primary_key=True)
    email: Mapped[str] = mapped_column(String(320), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(255))
    status: Mapped[AccountStatus] = mapped_column(
        Enum(AccountStatus, native_enum=False), default=AccountStatus.ACTIVE
    )
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


class Installation(Base):
    """One row per install, so telemetry can be counted without identifying a machine."""

    __tablename__ = "installations"

    id: Mapped[int] = mapped_column(primary_key=True)
    install_id: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    user_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"), default=None, index=True)
    client_version: Mapped[str | None] = mapped_column(String(32), default=None)
    host: Mapped[str | None] = mapped_column(String(16), default=None)
    office_version: Mapped[str | None] = mapped_column(String(32), default=None)
    first_seen_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    last_seen_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


class TelemetryEvent(Base):
    __tablename__ = "telemetry_events"

    id: Mapped[int] = mapped_column(primary_key=True)
    user_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"), default=None, index=True)
    install_id: Mapped[str | None] = mapped_column(String(64), default=None, index=True)
    event_type: Mapped[str] = mapped_column(String(40), index=True)
    # Already filtered through the server-side allowlist in telemetry.py. The
    # client is not trusted to have filtered correctly.
    payload: Mapped[str] = mapped_column(Text, default="{}")
    occurred_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    received_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    __table_args__ = (
        Index("ix_telemetry_type_time", "event_type", "occurred_at"),
    )


class UsageRecord(Base):
    """One proxied model call.

    Written by the proxy, which is the only place the client cannot bypass.
    Cost is stored per row rather than derived later from a price list: prices
    change, and a bill has to reflect what the call cost when it was made.
    """

    __tablename__ = "usage_records"

    id: Mapped[int] = mapped_column(primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    model: Mapped[str] = mapped_column(String(80), index=True)
    upstream: Mapped[str | None] = mapped_column(String(40), default=None)
    input_tokens: Mapped[int] = mapped_column(Integer, default=0)
    output_tokens: Mapped[int] = mapped_column(Integer, default=0)
    cache_read_tokens: Mapped[int] = mapped_column(Integer, default=0)
    cache_write_tokens: Mapped[int] = mapped_column(Integer, default=0)
    cost_usd: Mapped[float] = mapped_column(Float, default=0.0)
    status_code: Mapped[int] = mapped_column(Integer, default=200)
    duration_ms: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    __table_args__ = (
        Index("ix_usage_user_time", "user_id", "created_at"),
    )


class OrderStatus(str, enum.Enum):
    PENDING = "pending"
    PAID = "paid"
    CANCELLED = "cancelled"
    REFUNDED = "refunded"


class Order(Base):
    """A subscription purchase.

    The payment channel itself is not integrated -- WeChat Pay and Alipay both
    require a verified merchant account, which takes weeks. This table plus the
    admin activation path is what that integration will plug into, and it lets
    the entitlement side be exercised before the money side exists.
    """

    __tablename__ = "orders"

    id: Mapped[int] = mapped_column(primary_key=True)
    order_no: Mapped[str] = mapped_column(String(40), unique=True, index=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    plan: Mapped[Plan] = mapped_column(Enum(Plan, native_enum=False))
    months: Mapped[int] = mapped_column(Integer, default=1)
    amount_cents: Mapped[int] = mapped_column(Integer, default=0)
    currency: Mapped[str] = mapped_column(String(8), default="CNY")
    status: Mapped[OrderStatus] = mapped_column(
        Enum(OrderStatus, native_enum=False), default=OrderStatus.PENDING
    )
    channel: Mapped[str | None] = mapped_column(String(24), default=None)
    external_id: Mapped[str | None] = mapped_column(String(64), default=None)
    note: Mapped[str | None] = mapped_column(String(200), default=None)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    paid_at: Mapped[dt.datetime | None] = mapped_column(DateTime(timezone=True), default=None)


class SharedSkill(Base):
    """A user's synced skill, optionally published under a share code.

    Synced and shared are deliberately separate: a skill is private when it is
    backed up, and only becomes reachable by others when its owner explicitly
    shares it. Conflating the two would publish people's work by default.
    """

    __tablename__ = "shared_skills"

    id: Mapped[int] = mapped_column(primary_key=True)
    owner_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    # Client-side skill id, so a re-upload updates rather than duplicates.
    skill_id: Mapped[str] = mapped_column(String(64), index=True)
    name: Mapped[str] = mapped_column(String(120))
    document: Mapped[str] = mapped_column(Text)
    share_code: Mapped[str | None] = mapped_column(
        String(32), unique=True, index=True, default=None
    )
    import_count: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    updated_at: Mapped[dt.datetime] = mapped_column(
        DateTime(timezone=True), default=utcnow, onupdate=utcnow
    )

    __table_args__ = (
        UniqueConstraint("owner_id", "skill_id", name="uq_owner_skill"),
    )


class AuditLog(Base):
    """Every administrative mutation. Without this, "who disabled this account"
    has no answer."""

    __tablename__ = "audit_logs"

    id: Mapped[int] = mapped_column(primary_key=True)
    admin_id: Mapped[int | None] = mapped_column(ForeignKey("admins.id"), default=None)
    action: Mapped[str] = mapped_column(String(60), index=True)
    target: Mapped[str | None] = mapped_column(String(120), default=None)
    detail: Mapped[str | None] = mapped_column(Text, default=None)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


def _as_utc(value: dt.datetime) -> dt.datetime:
    """SQLite hands back naive datetimes; comparing those to aware ones raises."""
    if value.tzinfo is None:
        return value.replace(tzinfo=dt.timezone.utc)
    return value
