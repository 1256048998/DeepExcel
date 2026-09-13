"""Deployment configuration.

Everything that differs between a laptop, staging and production is here, and
nothing here is a code change. In particular HOSTED_PROXY_BASE_URL is the M6
switch: while it is unset the server tells every client to keep using its own
provider key (BYOK), and setting it is what turns on hosted routing.
"""

from __future__ import annotations

import os
import secrets
from dataclasses import dataclass, field


def _env(name: str, default: str | None = None) -> str | None:
    value = os.environ.get(name)
    if value is None or value == "":
        return default
    return value


def _env_int(name: str, default: int) -> int:
    raw = _env(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise RuntimeError(f"{name} must be an integer, got {raw!r}") from exc


def _env_bool(name: str, default: bool) -> bool:
    raw = _env(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    environment: str = field(default_factory=lambda: _env("DEEPEXCEL_ENV", "development"))

    # sqlite for local development and the test suite; postgresql+psycopg in
    # production. The models are identical either way.
    database_url: str = field(
        default_factory=lambda: _env("DATABASE_URL", "sqlite:///./deepexcel.db")
    )

    jwt_secret: str = field(default_factory=lambda: _env("JWT_SECRET", ""))
    access_token_ttl_seconds: int = field(
        default_factory=lambda: _env_int("ACCESS_TOKEN_TTL_SECONDS", 30 * 60)
    )
    refresh_token_ttl_seconds: int = field(
        default_factory=lambda: _env_int("REFRESH_TOKEN_TTL_SECONDS", 30 * 24 * 3600)
    )

    # Registration is invite-only while the product is in closed testing. This
    # is the whole point of building accounts now: today anyone holding the ZIP
    # can install, and we cannot see or limit that.
    require_invite_code: bool = field(
        default_factory=lambda: _env_bool("REQUIRE_INVITE_CODE", True)
    )

    # ---- M6 switch -------------------------------------------------------
    # Unset  -> every client is told mode="byok" and keeps using its own key.
    # Set    -> clients whose entitlement says hosted are routed through the
    #           proxy at this base URL, with a short-lived token.
    # Flipping this needs no client release: see routers/session.py.
    hosted_proxy_base_url: str | None = field(
        default_factory=lambda: _env("HOSTED_PROXY_BASE_URL")
    )
    endpoint_config_ttl_seconds: int = field(
        default_factory=lambda: _env_int("ENDPOINT_CONFIG_TTL_SECONDS", 15 * 60)
    )

    # First admin is created from the environment on startup, never through a
    # public endpoint.
    bootstrap_admin_email: str | None = field(
        default_factory=lambda: _env("BOOTSTRAP_ADMIN_EMAIL")
    )
    bootstrap_admin_password: str | None = field(
        default_factory=lambda: _env("BOOTSTRAP_ADMIN_PASSWORD")
    )

    admin_origin: str | None = field(default_factory=lambda: _env("ADMIN_ORIGIN"))

    telemetry_enabled: bool = field(
        default_factory=lambda: _env_bool("TELEMETRY_ENABLED", True)
    )
    telemetry_max_batch: int = field(
        default_factory=lambda: _env_int("TELEMETRY_MAX_BATCH", 100)
    )

    def __post_init__(self) -> None:
        if self.is_production:
            if not self.jwt_secret:
                raise RuntimeError(
                    "JWT_SECRET must be set in production. Generate one with: "
                    "python -c \"import secrets; print(secrets.token_urlsafe(48))\""
                )
            if len(self.jwt_secret) < 32:
                raise RuntimeError("JWT_SECRET must be at least 32 characters")
            if self.database_url.startswith("sqlite"):
                raise RuntimeError("Refusing to run production on SQLite; set DATABASE_URL")
        elif not self.jwt_secret:
            # A random per-process secret is correct for development: it makes
            # tokens die with the process instead of silently becoming a shared
            # well-known key that could reach production.
            object.__setattr__(self, "jwt_secret", secrets.token_urlsafe(48))

    @property
    def is_production(self) -> bool:
        return self.environment.lower() in {"production", "prod"}


_settings: Settings | None = None


def get_settings() -> Settings:
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings


def reset_settings_for_tests() -> None:
    """Tests mutate os.environ and need the next read to pick it up."""
    global _settings
    _settings = None
