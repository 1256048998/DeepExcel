"""usage records: points charged

Revision ID: b7d4e2a19c55
Revises: a3f1c9d27e40
Create Date: 2026-09-29 10:00:00

The quota moved from one-per-task to points: a task costs its model's weight.
points records what the counted call took, so a bill can be explained after
the weights change.
"""
from __future__ import annotations

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = 'b7d4e2a19c55'
down_revision: Union[str, None] = 'a3f1c9d27e40'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    with op.batch_alter_table("usage_records", schema=None) as batch_op:
        batch_op.add_column(
            sa.Column("points", sa.Integer(), nullable=False, server_default=sa.text("0"))
        )


def downgrade() -> None:
    with op.batch_alter_table("usage_records", schema=None) as batch_op:
        batch_op.drop_column("points")
