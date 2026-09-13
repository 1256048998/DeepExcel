"""Skill sync and sharing.

A skill is a parameterised sequence of tool calls plus the original request that
produced it. Syncing makes the library survive a reinstall; sharing is what lets
one person's solution become everyone's.

The privacy problem is specific and easy to miss: a skill's parameter defaults
are values captured from a real run, so a File parameter holds a real local path
and the original request text can name a real workbook. Uploading those verbatim
would put "C:\\Users\\alice\\2026财务预算.xlsx" in a database, and sharing would
hand it to a stranger. The client scrubs before upload; this module scrubs again,
for the same reason telemetry is filtered server-side -- the client is unsigned
and its filtering is a convenience, not a guarantee.
"""

from __future__ import annotations

import json
import re
import secrets

from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel, Field
from sqlalchemy import select
from sqlalchemy.orm import Session

from ..db import get_db
from ..deps import get_current_user
from ..models import SharedSkill, User, utcnow

router = APIRouter(prefix="/api/v1/skills", tags=["skills"])

MAX_SKILL_BYTES = 64 * 1024
MAX_SKILLS_PER_USER = 200

# Windows and UNC paths, plus anything that looks like a workbook file name.
_PATH = re.compile(r"(?:[a-zA-Z]:\\|\\\\)[^\"'\s]*", re.IGNORECASE)
_WORKBOOK_NAME = re.compile(r"[^\s\\/\"']+\.(?:xlsx|xlsm|xls|csv)", re.IGNORECASE)

_REDACTED = "[已移除路径]"


def scrub(value: str | None) -> str | None:
    """Removes anything that identifies a file on someone's machine."""
    if not value:
        return value
    cleaned = _PATH.sub(_REDACTED, value)
    cleaned = _WORKBOOK_NAME.sub(_REDACTED, cleaned)
    return cleaned


def scrub_skill(payload: dict) -> dict:
    """Scrubs a skill document in place and returns it.

    Only the fields that can carry a real value are touched. Tool names and
    placeholders are structure, not data.
    """
    payload["name"] = scrub(payload.get("name"))
    payload["description"] = scrub(payload.get("description"))
    payload["original_request"] = scrub(payload.get("original_request"))

    for parameter in payload.get("parameters") or []:
        if not isinstance(parameter, dict):
            continue
        # A File default is always a real local path, so it never survives
        # upload: the recipient's file lives somewhere else anyway.
        if parameter.get("kind") == "File" or parameter.get("kind") == 3:
            parameter["default_value"] = ""
        else:
            parameter["default_value"] = scrub(parameter.get("default_value"))

    for step in payload.get("steps") or []:
        if not isinstance(step, dict):
            continue
        arguments = step.get("arguments")
        if isinstance(arguments, dict):
            step["arguments"] = {key: scrub(str(value)) for key, value in arguments.items()}
    return payload


class SkillUpload(BaseModel):
    skill_id: str = Field(max_length=64)
    name: str = Field(max_length=120)
    document: dict


class SkillSummary(BaseModel):
    skill_id: str
    name: str
    updated_at: str
    share_code: str | None


@router.put("", response_model=SkillSummary)
def upsert(
    payload: SkillUpload,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> SkillSummary:
    document = scrub_skill(dict(payload.document))
    encoded = json.dumps(document, ensure_ascii=False, sort_keys=True)
    if len(encoded.encode("utf-8")) > MAX_SKILL_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail="技能内容过大",
        )

    existing = db.scalar(
        select(SharedSkill).where(
            SharedSkill.owner_id == user.id, SharedSkill.skill_id == payload.skill_id
        )
    )
    if existing is None:
        count = len(list(db.scalars(select(SharedSkill).where(SharedSkill.owner_id == user.id))))
        if count >= MAX_SKILLS_PER_USER:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"最多同步 {MAX_SKILLS_PER_USER} 个技能",
            )
        existing = SharedSkill(owner_id=user.id, skill_id=payload.skill_id)
        db.add(existing)

    existing.name = scrub(payload.name) or "未命名技能"
    existing.document = encoded
    existing.updated_at = utcnow()
    db.flush()

    return SkillSummary(
        skill_id=existing.skill_id,
        name=existing.name,
        updated_at=existing.updated_at.isoformat(),
        share_code=existing.share_code,
    )


@router.get("", response_model=list[SkillSummary])
def list_mine(
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> list[SkillSummary]:
    rows = db.scalars(
        select(SharedSkill).where(SharedSkill.owner_id == user.id).order_by(SharedSkill.updated_at.desc())
    )
    return [
        SkillSummary(
            skill_id=row.skill_id,
            name=row.name,
            updated_at=row.updated_at.isoformat(),
            share_code=row.share_code,
        )
        for row in rows
    ]


@router.get("/{skill_id}")
def fetch(
    skill_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    row = db.scalar(
        select(SharedSkill).where(
            SharedSkill.owner_id == user.id, SharedSkill.skill_id == skill_id
        )
    )
    if row is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="技能不存在")
    return json.loads(row.document)


@router.delete("/{skill_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete(
    skill_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> None:
    row = db.scalar(
        select(SharedSkill).where(
            SharedSkill.owner_id == user.id, SharedSkill.skill_id == skill_id
        )
    )
    if row is not None:
        db.delete(row)


@router.post("/{skill_id}/share")
def share(
    skill_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """Publishes a skill under a short code.

    Sharing is opt-in per skill and reversible. A synced skill is private until
    its owner does this, because "backed up" and "published" must never be the
    same action.
    """
    row = db.scalar(
        select(SharedSkill).where(
            SharedSkill.owner_id == user.id, SharedSkill.skill_id == skill_id
        )
    )
    if row is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="技能不存在")
    if row.share_code is None:
        row.share_code = secrets.token_urlsafe(6)
    return {"share_code": row.share_code}


@router.post("/{skill_id}/unshare", status_code=status.HTTP_204_NO_CONTENT)
def unshare(
    skill_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> None:
    row = db.scalar(
        select(SharedSkill).where(
            SharedSkill.owner_id == user.id, SharedSkill.skill_id == skill_id
        )
    )
    if row is not None:
        row.share_code = None


@router.get("/shared/{share_code}")
def fetch_shared(
    share_code: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """Imports someone else's shared skill.

    Requires an account: an open endpoint would turn share codes into a public
    scraping target, and codes are short enough to enumerate.
    """
    row = db.scalar(select(SharedSkill).where(SharedSkill.share_code == share_code))
    if row is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="分享码无效")
    row.import_count += 1
    document = json.loads(row.document)
    # A fresh id, so importing does not overwrite a skill the recipient already
    # has with the same id.
    document["id"] = secrets.token_hex(16)
    return document
