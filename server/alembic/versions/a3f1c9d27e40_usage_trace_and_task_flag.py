"""usage records: trace_id and counted_as_task

Revision ID: a3f1c9d27e40
Revises: ce567d602320
Create Date: 2026-09-25 10:00:00

The quota is sold in tasks, but the proxy used to count every successful model
call as one. A task is an agent loop of a dozen or more calls; trace_id groups
them and counted_as_task marks the single call that consumed the task.
"""
from __future__ import annotations

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = 'a3f1c9d27e40'
down_revision: Union[str, None] = 'ce567d602320'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    with op.batch_alter_table("usage_records", schema=None) as batch_op:
        batch_op.add_column(sa.Column("trace_id", sa.String(length=64), nullable=True))
        batch_op.add_column(
            sa.Column("counted_as_task", sa.Boolean(), nullable=False, server_default=sa.false())
        )
        batch_op.create_index("ix_usage_user_trace", ["user_id", "trace_id"], unique=False)


def downgrade() -> None:
    with op.batch_alter_table("usage_records", schema=None) as batch_op:
        batch_op.drop_index("ix_usage_user_trace")
        batch_op.drop_column("counted_as_task")
        batch_op.drop_column("trace_id")
