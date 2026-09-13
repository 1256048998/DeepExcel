"""Password hashing and token issuance.

Password hashing uses hashlib.scrypt from the standard library rather than
bcrypt or argon2. scrypt is memory-hard, OpenSSL-backed and needs no native
build step, which keeps both the Docker image and the test run free of a
compiler dependency. Parameters follow RFC 7914's interactive-login guidance.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import secrets
import time
import uuid
from dataclasses import dataclass

import jwt

from .config import get_settings

# n=2**15 costs roughly 32 MB and ~50-100 ms per hash on a modern core. That is
# deliberately slow: it is the only thing standing between a stolen database and
# the users' passwords.
_SCRYPT_N = 2 ** 15
_SCRYPT_R = 8
_SCRYPT_P = 1
_SCRYPT_DKLEN = 32
_SALT_BYTES = 16

_PREFIX = "scrypt"


def hash_password(password: str) -> str:
    if not password:
        raise ValueError("password must not be empty")
    salt = secrets.token_bytes(_SALT_BYTES)
    derived = hashlib.scrypt(
        password.encode("utf-8"),
        salt=salt,
        n=_SCRYPT_N,
        r=_SCRYPT_R,
        p=_SCRYPT_P,
        dklen=_SCRYPT_DKLEN,
        maxmem=64 * 1024 * 1024,
    )
    return "${}${}${}${}${}${}".format(
        _PREFIX,
        _SCRYPT_N,
        _SCRYPT_R,
        _SCRYPT_P,
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(derived).decode("ascii"),
    )


def verify_password(password: str, encoded: str) -> bool:
    """Constant-time verification. Returns False for malformed records."""
    try:
        _, prefix, n, r, p, salt_b64, hash_b64 = encoded.split("$")
        if prefix != _PREFIX:
            return False
        salt = base64.b64decode(salt_b64)
        expected = base64.b64decode(hash_b64)
        derived = hashlib.scrypt(
            password.encode("utf-8"),
            salt=salt,
            n=int(n),
            r=int(r),
            p=int(p),
            dklen=len(expected),
            maxmem=64 * 1024 * 1024,
        )
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(derived, expected)


# ---------------------------------------------------------------------------
# Access tokens (JWT) and refresh tokens (opaque)
# ---------------------------------------------------------------------------

ACCESS_AUDIENCE_USER = "deepexcel-user"
ACCESS_AUDIENCE_ADMIN = "deepexcel-admin"
# Issued to a client so it can talk to the model proxy. Deliberately separate
# from the user access token: it is handed to a different component, has a much
# shorter life, and must never be usable against the account APIs.
ACCESS_AUDIENCE_PROXY = "deepexcel-proxy"

_ALGORITHM = "HS256"
_ISSUER = "deepexcel"


@dataclass(frozen=True)
class TokenClaims:
    subject: str
    audience: str
    expires_at: int
    token_id: str


def issue_access_token(
    subject: str,
    audience: str = ACCESS_AUDIENCE_USER,
    ttl_seconds: int | None = None,
    extra: dict | None = None,
) -> tuple[str, int]:
    """Returns (token, expires_at_unix)."""
    settings = get_settings()
    ttl = ttl_seconds if ttl_seconds is not None else settings.access_token_ttl_seconds
    now = int(time.time())
    expires_at = now + ttl
    payload = {
        "sub": str(subject),
        "aud": audience,
        "iss": _ISSUER,
        "iat": now,
        "nbf": now,
        "exp": expires_at,
        "jti": uuid.uuid4().hex,
    }
    if extra:
        # Never let a caller overwrite a registered claim.
        for key, value in extra.items():
            if key in payload:
                raise ValueError(f"extra claim {key!r} would shadow a registered claim")
            payload[key] = value
    token = jwt.encode(payload, settings.jwt_secret, algorithm=_ALGORITHM)
    return token, expires_at


def decode_access_token(token: str, audience: str) -> TokenClaims:
    """Raises jwt.PyJWTError on any problem, including a wrong audience.

    Checking the audience is what stops an admin token from being replayed
    against user endpoints, or a proxy token against either.
    """
    settings = get_settings()
    payload = jwt.decode(
        token,
        settings.jwt_secret,
        algorithms=[_ALGORITHM],
        audience=audience,
        issuer=_ISSUER,
        options={"require": ["exp", "sub", "aud", "iss"]},
    )
    return TokenClaims(
        subject=payload["sub"],
        audience=payload["aud"],
        expires_at=int(payload["exp"]),
        token_id=payload.get("jti", ""),
    )


def new_refresh_token() -> tuple[str, str]:
    """Returns (plaintext, storage_hash).

    Refresh tokens are opaque and stored only as a SHA-256 digest, so a database
    leak cannot be replayed. They are revocable, which JWTs are not -- that is
    why sessions are not built purely on access tokens.
    """
    plaintext = secrets.token_urlsafe(48)
    return plaintext, hash_refresh_token(plaintext)


def hash_refresh_token(plaintext: str) -> str:
    return hashlib.sha256(plaintext.encode("utf-8")).hexdigest()


def new_opaque_id() -> str:
    return uuid.uuid4().hex
