"""Request dependencies: who is calling, and are they allowed to."""

from __future__ import annotations

import jwt
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from sqlalchemy.orm import Session

from .db import get_db
from .models import AccountStatus, Admin, User
from .security import (
    ACCESS_AUDIENCE_ADMIN,
    ACCESS_AUDIENCE_USER,
    decode_access_token,
)

# auto_error=False so a missing header produces our own 401 with a consistent
# body rather than FastAPI's default 403.
_bearer = HTTPBearer(auto_error=False)


def _unauthorized(detail: str) -> HTTPException:
    return HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail=detail,
        headers={"WWW-Authenticate": "Bearer"},
    )


def _token_subject(
    credentials: HTTPAuthorizationCredentials | None,
    audience: str,
) -> str:
    if credentials is None or not credentials.credentials:
        raise _unauthorized("Missing bearer token")
    try:
        claims = decode_access_token(credentials.credentials, audience=audience)
    except jwt.ExpiredSignatureError as exc:
        raise _unauthorized("Token expired") from exc
    except jwt.PyJWTError as exc:
        # Covers a wrong audience, which is how an admin token is stopped from
        # being replayed against user endpoints.
        raise _unauthorized("Invalid token") from exc
    return claims.subject


def get_current_user(
    credentials: HTTPAuthorizationCredentials | None = Depends(_bearer),
    db: Session = Depends(get_db),
) -> User:
    subject = _token_subject(credentials, ACCESS_AUDIENCE_USER)
    try:
        user_id = int(subject)
    except ValueError as exc:
        raise _unauthorized("Invalid token subject") from exc

    user = db.get(User, user_id)
    if user is None:
        raise _unauthorized("Account no longer exists")
    if user.status is not AccountStatus.ACTIVE:
        # 403, not 401: the credentials were fine, the account is not. Returning
        # 401 would send the client into a pointless refresh loop.
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Account disabled",
        )
    return user


def get_current_admin(
    credentials: HTTPAuthorizationCredentials | None = Depends(_bearer),
    db: Session = Depends(get_db),
) -> Admin:
    subject = _token_subject(credentials, ACCESS_AUDIENCE_ADMIN)
    try:
        admin_id = int(subject)
    except ValueError as exc:
        raise _unauthorized("Invalid token subject") from exc

    admin = db.get(Admin, admin_id)
    if admin is None or admin.status is not AccountStatus.ACTIVE:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Administrator account is not active",
        )
    return admin
