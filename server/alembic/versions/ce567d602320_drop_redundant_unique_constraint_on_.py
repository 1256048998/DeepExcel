"""drop redundant unique constraint on audit_logs.id

Revision ID: ce567d602320
Revises: 795b199bc72d
Create Date: 2026-09-12 20:31:15.177749
"""
from __future__ import annotations

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = 'ce567d602320'
down_revision: Union[str, None] = '795b199bc72d'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # audit_logs.id is the primary key, so a unique constraint on it adds
    # nothing. SQLite accepted it silently; Postgres reported the models and
    # migrations as diverged, which is how it was found.
    with op.batch_alter_table("audit_logs") as batch:
        batch.drop_constraint("uq_audit_id", type_="unique")


def downgrade() -> None:
    with op.batch_alter_table("audit_logs") as batch:
        batch.create_unique_constraint("uq_audit_id", ["id"])
