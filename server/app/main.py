"""Application entry point."""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import select

from .config import get_settings
from .db import create_all, get_session_factory
from .models import Admin
from .proxy import router as proxy_router
from .routers import admin, auth, billing, session, skills, telemetry, updates
from .security import hash_password

logger = logging.getLogger("deepexcel.server")


def bootstrap_admin() -> None:
    """Creates the first operator account from the environment.

    There is deliberately no public "create the first admin" endpoint: such an
    endpoint is either a race to be first or a permanent hole if the disable
    condition is ever wrong.
    """
    settings = get_settings()
    if not settings.bootstrap_admin_email or not settings.bootstrap_admin_password:
        return

    email = settings.bootstrap_admin_email.strip().lower()
    with get_session_factory()() as db:
        if db.scalar(select(Admin).where(Admin.email == email)) is not None:
            return
        if len(settings.bootstrap_admin_password) < 12:
            raise RuntimeError("BOOTSTRAP_ADMIN_PASSWORD must be at least 12 characters")
        db.add(Admin(email=email, password_hash=hash_password(settings.bootstrap_admin_password)))
        db.commit()
        logger.info("Bootstrapped administrator account %s", email)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    settings = get_settings()
    if not settings.is_production:
        # Production schema changes go through Alembic; see server/README.md.
        create_all()
    bootstrap_admin()
    yield


def create_app() -> FastAPI:
    settings = get_settings()
    app = FastAPI(
        title="DeepExcel Server",
        version="0.1.0",
        description=(
            "Accounts, entitlements, endpoint routing and telemetry for DeepExcel. "
            "Model traffic does not pass through this service yet: see "
            "GET /api/v1/session/endpoint."
        ),
        lifespan=lifespan,
        # The interactive docs are useful in development and an unnecessary
        # surface in production.
        docs_url=None if settings.is_production else "/docs",
        redoc_url=None,
    )

    if settings.admin_origin:
        app.add_middleware(
            CORSMiddleware,
            allow_origins=[settings.admin_origin],
            allow_credentials=True,
            allow_methods=["GET", "POST", "PATCH", "DELETE", "OPTIONS"],
            allow_headers=["Authorization", "Content-Type"],
        )

    app.include_router(auth.router)
    app.include_router(session.router)
    app.include_router(telemetry.router)
    app.include_router(admin.router)
    app.include_router(billing.router)
    app.include_router(skills.router)
    # Unauthenticated on purpose: a signed-out client still needs to be able to
    # pick up a fix, and the manifest's integrity comes from its signature.
    app.include_router(updates.router)
    # Mounted unconditionally; the endpoint itself refuses to serve while
    # HOSTED_PROXY_BASE_URL is unset, so a deployment cannot start billing
    # traffic just because the route exists.
    app.include_router(proxy_router.router)

    @app.get("/healthz", tags=["ops"])
    def healthz() -> dict:
        return {"status": "ok", "environment": settings.environment}

    @app.get("/api/v1/meta", tags=["ops"])
    def meta() -> dict:
        """What this deployment expects of a client, without a client release."""
        return {
            "invite_required": settings.require_invite_code,
            "hosted_routing_available": bool(settings.hosted_proxy_base_url),
            "telemetry_enabled": settings.telemetry_enabled,
            "update_feed_available": bool(settings.update_manifest_dir),
            "min_client_version": "0.5.0",
        }

    return app


app = create_app()
