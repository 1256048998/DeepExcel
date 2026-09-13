"""Registration, login, refresh, logout."""

from __future__ import annotations

import datetime as dt

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import func, select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from ..config import get_settings
from ..db import get_db
from ..deps import get_current_user
from ..models import (
    AccountStatus,
    Entitlement,
    EntitlementStatus,
    InviteCode,
    Plan,
    RefreshToken,
    RoutingMode,
    User,
    utcnow,
)
from ..schemas import (
    LoginRequest,
    RefreshRequest,
    RegisterRequest,
    TokenPair,
    UserView,
)
from ..security import (
    hash_password,
    hash_refresh_token,
    issue_access_token,
    new_refresh_token,
    verify_password,
)

router = APIRouter(prefix="/api/v1/auth", tags=["auth"])


def _normalize_email(email: str) -> str:
    return email.strip().lower()


def _issue_pair(db: Session, user: User) -> TokenPair:
    access_token, expires_at = issue_access_token(str(user.id))
    plaintext, token_hash = new_refresh_token()
    settings = get_settings()
    db.add(
        RefreshToken(
            user_id=user.id,
            token_hash=token_hash,
            expires_at=utcnow() + dt.timedelta(seconds=settings.refresh_token_ttl_seconds),
        )
    )
    user.last_seen_at = utcnow()
    db.flush()
    return TokenPair(access_token=access_token, refresh_token=plaintext, expires_at=expires_at)


@router.post("/register", response_model=TokenPair, status_code=status.HTTP_201_CREATED)
def register(payload: RegisterRequest, db: Session = Depends(get_db)) -> TokenPair:
    settings = get_settings()
    email = _normalize_email(payload.email)

    invite: InviteCode | None = None
    if settings.require_invite_code:
        if not payload.invite_code:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="An invite code is required to register",
            )
        invite = db.scalar(
            select(InviteCode).where(InviteCode.code == payload.invite_code.strip())
        )
        # Same message for "does not exist" and "already used" so the endpoint
        # cannot be used to enumerate valid codes.
        if invite is None or not invite.is_usable():
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Invite code is not valid",
            )

    user = User(
        email=email,
        password_hash=hash_password(payload.password),
        display_name=payload.display_name,
        invite_code_id=invite.id if invite else None,
    )
    # Every account gets an entitlement row immediately, even though the beta
    # plan grants everything. Code that asks "what is this user allowed to do"
    # must never have to handle a missing row.
    user.entitlement = Entitlement(
        plan=Plan.BETA,
        status=EntitlementStatus.ACTIVE,
        routing_mode=RoutingMode.BYOK,
        task_limit=None,
    )
    db.add(user)

    try:
        db.flush()
    except IntegrityError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="An account with this email already exists",
        ) from exc

    if invite is not None:
        # Re-read under the same transaction and increment conditionally so two
        # concurrent registrations cannot both consume the last use of a code.
        consumed = db.execute(
            InviteCode.__table__.update()
            .where(InviteCode.id == invite.id, InviteCode.used_count < InviteCode.max_uses)
            .values(used_count=InviteCode.used_count + 1)
        )
        if consumed.rowcount != 1:
            db.rollback()
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Invite code is not valid",
            )

    return _issue_pair(db, user)


@router.post("/login", response_model=TokenPair)
def login(payload: LoginRequest, db: Session = Depends(get_db)) -> TokenPair:
    email = _normalize_email(payload.email)
    user = db.scalar(select(User).where(User.email == email))

    # Hash even when the account does not exist, so response time does not
    # reveal which emails are registered.
    password_hash = user.password_hash if user else _DUMMY_HASH
    password_ok = verify_password(payload.password, password_hash)

    if user is None or not password_ok:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
        )
    if user.status is not AccountStatus.ACTIVE:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Account disabled",
        )
    return _issue_pair(db, user)


@router.post("/refresh", response_model=TokenPair)
def refresh(payload: RefreshRequest, db: Session = Depends(get_db)) -> TokenPair:
    token_hash = hash_refresh_token(payload.refresh_token)
    record = db.scalar(select(RefreshToken).where(RefreshToken.token_hash == token_hash))
    if record is None or not record.is_valid:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Refresh token is not valid",
        )

    user = db.get(User, record.user_id)
    if user is None or user.status is not AccountStatus.ACTIVE:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Account disabled",
        )

    # Rotate: a refresh token is single-use. If a stolen one is replayed after
    # the legitimate client has rotated, it is already revoked.
    record.revoked_at = utcnow()
    return _issue_pair(db, user)


@router.post("/logout", status_code=status.HTTP_204_NO_CONTENT)
def logout(payload: RefreshRequest, db: Session = Depends(get_db)) -> None:
    token_hash = hash_refresh_token(payload.refresh_token)
    record = db.scalar(select(RefreshToken).where(RefreshToken.token_hash == token_hash))
    # Succeed regardless: logging out an already-invalid token is not an error,
    # and reporting otherwise would confirm whether a token was real.
    if record is not None and record.revoked_at is None:
        record.revoked_at = utcnow()


@router.get("/me", response_model=UserView)
def me(user: User = Depends(get_current_user)) -> User:
    return user


@router.get("/registration-policy")
def registration_policy() -> dict:
    """Lets the client render the right form without shipping a build per policy."""
    return {"invite_required": get_settings().require_invite_code}


def _count_users(db: Session) -> int:
    return db.scalar(select(func.count()).select_from(User)) or 0


# Cost-matched to a real record so the timing of a failed login against an
# unknown email matches one against a known email.
_DUMMY_HASH = hash_password("deepexcel-timing-equalizer")
