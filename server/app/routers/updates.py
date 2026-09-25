"""The update feed.

This endpoint relays a manifest that was signed offline. It never signs
anything, and the signing key must never exist on this machine: the client
verifies the signature itself, so taking over this server buys an attacker the
ability to withhold updates or serve an older one -- not to push code.

Publishing a release is therefore a file drop:

    UPDATE_MANIFEST_DIR=/srv/deepexcel/updates
    /srv/deepexcel/updates/stable.json      <- output of scripts/update_signing.py

No restart, no database, no deploy. Manifests are re-read when their mtime
changes.

The knowledge pack (skills the agent reads on demand) is published the same
way, as ``knowledge_pack.json`` in that directory, and served by /knowledge.

Deliberately unauthenticated. A user whose session expired, or who never signed
in, still needs to be able to receive a fix, and the manifest is a public
artifact whose integrity comes from its signature rather than from who asked.
"""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import dataclass

from fastapi import APIRouter, HTTPException, Response, status

from ..config import get_settings

logger = logging.getLogger("deepexcel.server.updates")

router = APIRouter(prefix="/api/v1/updates", tags=["updates"])

# Channel names become a file name, so they are validated rather than escaped.
_CHANNEL_PATTERN = re.compile(r"^[a-z0-9][a-z0-9-]{0,31}$")

# Mirrors UpdateManifest.MaxManifestBytes on the client. A file larger than this
# is a mistake somewhere, and the client would refuse it anyway.
MAX_MANIFEST_BYTES = 64 * 1024

_REQUIRED_FIELDS = ("schema", "key_id", "signature", "payload")

_CACHE_SECONDS = 300


@dataclass(frozen=True)
class _Cached:
    mtime: float
    size: int
    body: str


_cache: dict[str, _Cached] = {}


def reset_cache() -> None:
    """Tests write manifests faster than mtime resolution can distinguish."""
    _cache.clear()


def manifest_path(channel: str) -> str | None:
    directory = get_settings().update_manifest_dir
    if not directory:
        return None
    return os.path.join(directory, f"{channel}.json")


def _validate(body: str, path: str) -> None:
    """Structural check only; the signature is the client's business.

    Verifying here would mean putting the public key on the server and adding a
    crypto dependency, for no security gain -- the client is the party that has
    to be convinced. What this does catch is the operator who copied the wrong
    file, which would otherwise look like a working deployment right up until
    every client silently stopped updating.
    """
    try:
        document = json.loads(body)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={"reason": "manifest_invalid", "message": f"Update manifest is not JSON: {exc}"},
        ) from exc

    if not isinstance(document, dict):
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={"reason": "manifest_invalid", "message": "Update manifest is not an object"},
        )

    missing = [field for field in _REQUIRED_FIELDS if field not in document]
    if missing:
        logger.error("Update manifest %s is missing fields: %s", path, ", ".join(missing))
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={
                "reason": "manifest_invalid",
                "message": f"Update manifest is missing: {', '.join(missing)}",
            },
        )


@router.get("/latest")
def latest(channel: str = "stable") -> Response:
    channel = (channel or "").strip().lower()
    if not _CHANNEL_PATTERN.match(channel):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"reason": "bad_channel", "message": "Channel must be lowercase alphanumeric"},
        )

    path = manifest_path(channel)
    if path is None:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={
                "reason": "updates_not_configured",
                "message": "This deployment publishes no update feed (UPDATE_MANIFEST_DIR unset).",
            },
        )

    try:
        stat = os.stat(path)
    except OSError:
        # A channel nobody publishes is a 404, not an error: a client asking for
        # "beta" on a stable-only deployment is behaving correctly.
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={"reason": "no_release", "message": f"No release published for channel {channel}"},
        ) from None

    if stat.st_size > MAX_MANIFEST_BYTES:
        logger.error("Update manifest %s is %d bytes; refusing to serve", path, stat.st_size)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={"reason": "manifest_invalid", "message": "Update manifest is implausibly large"},
        )

    return _relay(channel, path, stat)


def _relay(cache_key: str, path: str, stat: os.stat_result, max_bytes: int = MAX_MANIFEST_BYTES) -> Response:
    cached = _cache.get(cache_key)
    if cached is None or cached.mtime != stat.st_mtime or cached.size != stat.st_size:
        with open(path, "r", encoding="utf-8") as stream:
            body = stream.read(max_bytes + 1)
        _validate(body, path)
        cached = _Cached(mtime=stat.st_mtime, size=stat.st_size, body=body)
        _cache[cache_key] = cached

    return Response(
        content=cached.body,
        media_type="application/json",
        headers={"Cache-Control": f"public, max-age={_CACHE_SECONDS}"},
    )


# The knowledge pack is signed like a manifest (scripts/knowledge_pack.py) and
# dropped into the same directory. The underscore keeps the name outside
# _CHANNEL_PATTERN, so /latest?channel=... can never hand it out as a channel.
KNOWLEDGE_PACK_FILE = "knowledge_pack.json"

# Mirrors KnowledgePack.MaxPackBytes on the client: a 1 MB payload is about
# 1.34 MB once base64-encoded inside the envelope.
MAX_KNOWLEDGE_PACK_BYTES = 2 * 1024 * 1024


@router.get("/knowledge")
def knowledge() -> Response:
    """Relays the signed knowledge pack.

    Knowledge text ends up in the model's context, so it is treated like code:
    signed offline with the update key and verified by the client. Same
    unauthenticated relay as the manifest, for the same reason.
    """
    directory = get_settings().update_manifest_dir
    if not directory:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={
                "reason": "updates_not_configured",
                "message": "This deployment publishes no update feed (UPDATE_MANIFEST_DIR unset).",
            },
        )
    path = os.path.join(directory, KNOWLEDGE_PACK_FILE)
    try:
        stat = os.stat(path)
    except OSError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={"reason": "no_release", "message": "No knowledge pack published"},
        ) from None
    if stat.st_size > MAX_KNOWLEDGE_PACK_BYTES:
        logger.error("Knowledge pack %s is %d bytes; refusing to serve", path, stat.st_size)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail={"reason": "manifest_invalid", "message": "Knowledge pack is implausibly large"},
        )
    return _relay("\0knowledge", path, stat, MAX_KNOWLEDGE_PACK_BYTES)
