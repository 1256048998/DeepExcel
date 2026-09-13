"""Migrations must stay in step with the models.

Changing a model and forgetting the migration is invisible in development,
because the test suite and local runs use create_all. It only surfaces in
production, as a column that does not exist. This test is the gate.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

SERVER_ROOT = Path(__file__).resolve().parents[1]


def _alembic(args: list[str], database_url: str) -> subprocess.CompletedProcess:
    env = dict(os.environ, DATABASE_URL=database_url)
    return subprocess.run(
        [sys.executable, "-m", "alembic", *args],
        cwd=SERVER_ROOT,
        env=env,
        capture_output=True,
        text=True,
    )


@pytest.fixture
def database_url(tmp_path: Path) -> str:
    # Alembic resolves a relative sqlite path against its own working directory,
    # so the file is created under the server root and cleaned up afterwards.
    name = f"migration-test-{os.getpid()}.db"
    yield f"sqlite:///./{name}"
    (SERVER_ROOT / name).unlink(missing_ok=True)


def test_migrations_apply_from_empty(database_url):
    result = _alembic(["upgrade", "head"], database_url)
    assert result.returncode == 0, result.stderr


def test_models_match_migrations(database_url):
    """`alembic check` fails if autogenerate would produce anything.

    A failure here means a model was changed without a migration. Fix it with:
        DATABASE_URL=... alembic revision --autogenerate -m "describe the change"
    """
    assert _alembic(["upgrade", "head"], database_url).returncode == 0

    result = _alembic(["check"], database_url)
    assert result.returncode == 0, (
        "Models and migrations have diverged. Generate a migration:\n"
        + result.stdout
        + result.stderr
    )


def test_migrations_are_reversible(database_url):
    """A migration that cannot be undone turns a bad deploy into a restore."""
    assert _alembic(["upgrade", "head"], database_url).returncode == 0
    result = _alembic(["downgrade", "base"], database_url)
    assert result.returncode == 0, result.stderr
