"""Local preview of the server and the operator console.

    server/.venv/Scripts/python.exe server/scripts/dev_server.py        # API on 127.0.0.1:8000
    server/.venv/Scripts/python.exe server/scripts/dev_server.py seed   # demo data; API must be running

The console is `npm run dev --prefix server/admin` (port 8080, proxies
/admin/api to this server). Both are in .claude/launch.json.

Everything this creates lives in server/.dev/, which is gitignored: the SQLite
database, and admin.json with the generated operator credentials. Delete the
folder to start over. Nothing here is read by a production deployment.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import random
import secrets
import sys
from pathlib import Path

SERVER_DIR = Path(__file__).resolve().parent.parent
STATE_DIR = SERVER_DIR / ".dev"
CREDENTIALS = STATE_DIR / "admin.json"
HOST, PORT = "127.0.0.1", 8000


def _credentials(create: bool) -> dict:
    if CREDENTIALS.exists():
        return json.loads(CREDENTIALS.read_text(encoding="utf-8"))
    if not create:
        raise SystemExit(f"{CREDENTIALS} not found; start the server with this script first")
    STATE_DIR.mkdir(exist_ok=True)
    credentials = {
        "email": "ops@deepexcel-qa.com",
        "password": secrets.token_urlsafe(18),
        # Shared by every seeded user, so the add-in can also sign in as one of
        # them against this server.
        "demo_user_password": secrets.token_urlsafe(18),
    }
    CREDENTIALS.write_text(json.dumps(credentials, indent=2), encoding="utf-8")
    return credentials


def serve() -> None:
    credentials = _credentials(create=True)
    # Never production: that would skip create_all and refuse SQLite.
    os.environ["DEEPEXCEL_ENV"] = "development"
    os.environ.setdefault("DATABASE_URL", f"sqlite:///{(STATE_DIR / 'deepexcel.db').as_posix()}")
    os.environ["BOOTSTRAP_ADMIN_EMAIL"] = credentials["email"]
    os.environ["BOOTSTRAP_ADMIN_PASSWORD"] = credentials["password"]
    print(f"Operator console credentials: {CREDENTIALS}", flush=True)

    import uvicorn

    uvicorn.run("app.main:app", host=HOST, port=PORT, app_dir=str(SERVER_DIR))


# ---------------------------------------------------------------------------
# Demo data
#
# Goes through the public API rather than the database, so it exercises the
# same validation, allowlist and audit paths real traffic does -- and breaks
# loudly when a contract changes instead of seeding rows the API would reject.
# ---------------------------------------------------------------------------

PEOPLE = [
    # (email local part, display name, invite, host, client version)
    ("zhangwei", "张伟", "beta", "excel", "0.5.0"),
    ("liuyang", "刘洋", "beta", "excel", "0.5.0"),
    ("wangfang", "王芳", "beta", "wps", "0.5.0"),
    ("chenjing", "陈静", "beta", "excel", "0.5.0"),
    ("lina", "李娜", "beta", "excel", "0.5.0"),
    ("zhaolei", "赵磊", "beta", "wps", "0.4.2"),
    ("sunli", "孙丽", "xhs", "excel", "0.4.2"),
    ("zhouqiang", "周强", "xhs", "excel", "0.5.0"),
]

TOOLS = [
    "write_formula", "fill_formula_down", "write_range", "sort_data", "clean_data",
    "create_pivot_table", "refresh_pivot", "split_text_to_columns", "execute_vba",
]

# (tool, error code, weight)
TOOL_ERRORS = [
    ("write_formula", "formula_name", 9),
    ("fill_formula_down", "formula_ref", 5),
    ("write_formula", "formula_ref", 4),
    ("clean_amount", "type_mismatch", 4),
    ("create_pivot_table", "range_invalid", 3),
    ("execute_vba", "vba_disabled", 3),
    ("sort_data", "merged_cells", 2),
    ("write_range", "sheet_protected", 2),
]


def seed() -> None:
    import httpx

    credentials = _credentials(create=False)
    rng = random.Random(20260928)
    now = dt.datetime.now(dt.timezone.utc)

    with httpx.Client(base_url=f"http://{HOST}:{PORT}", timeout=30) as client:

        def call(method: str, path: str, token: str | None = None, **kwargs):
            headers = {"Authorization": f"Bearer {token}"} if token else {}
            response = client.request(method, path, headers=headers, **kwargs)
            if response.is_error:
                raise SystemExit(f"{method} {path} -> {response.status_code} {response.text}")
            return response.json() if response.content else None

        admin = call("POST", "/admin/api/auth/login", json={
            "email": credentials["email"], "password": credentials["password"],
        })["access_token"]
        if call("GET", "/admin/api/users", admin):
            print(f"Already seeded; delete {STATE_DIR} and restart the server to reseed.")
            return

        invites = {
            key: call("POST", "/admin/api/invites", admin, json={"note": note, "max_uses": uses})
            for key, note, uses in [
                ("beta", "内测第一批（财务交流群）", 20),
                ("xhs", "渠道：小红书", 5),
                ("vip", "给王总试用", 1),
                ("old", "旧码，已作废", 10),
            ]
        }
        call("POST", f"/admin/api/invites/{invites['old']['id']}/disable", admin)

        tokens: dict[str, str] = {}
        for name, display_name, invite, host, version in PEOPLE:
            tokens[name] = call("POST", "/api/v1/auth/register", json={
                "email": f"{name}@deepexcel-qa.com",
                "password": credentials["demo_user_password"],
                "display_name": display_name,
                "invite_code": invites[invite]["code"],
            })["access_token"]
            call("POST", "/api/v1/session/heartbeat", tokens[name], params={
                "install_id": secrets.token_hex(16),
                "client_version": version,
                "host": host,
                "office_version": "16.0",
            })
            call("POST", "/api/v1/telemetry", tokens[name], json={
                "events": _usage(rng, now, version),
            })

        ids = {user["email"].split("@")[0]: user["id"] for user in call("GET", "/admin/api/users", admin)}

        def order(name: str, plan: str, months: int) -> str:
            return call("POST", "/api/v1/orders", tokens[name],
                        json={"plan": plan, "months": months})["order_no"]

        call("POST", f"/admin/api/orders/{order('wangfang', 'byok', 12)}/mark-paid", admin)
        order("chenjing", "byok", 1)
        call("POST", f"/admin/api/users/{ids['zhouqiang']}/status", admin, json={"status": "disabled"})
        print(f"Seeded {len(PEOPLE)} users, {len(invites)} invites, 2 orders.")

        # Unless this server has hosting configured, both of these must be
        # refused: either one would leave the user unable to work.
        checks = {
            "order pro": client.post("/api/v1/orders", json={"plan": "pro"},
                                     headers={"Authorization": f"Bearer {tokens['liuyang']}"}),
            "route to hosted": client.post(f"/admin/api/users/{ids['sunli']}/entitlement",
                                           json={"routing_mode": "hosted"},
                                           headers={"Authorization": f"Bearer {admin}"}),
        }
        for name, response in checks.items():
            # Only the reason token: the messages are Chinese, and a Windows
            # console may not be able to encode them.
            detail = response.json().get("detail")
            reason = detail.get("reason") if isinstance(detail, dict) else "accepted"
            print(f"{name} -> {response.status_code} {reason}")


def _usage(rng: random.Random, now: dt.datetime, version: str) -> list[dict]:
    """A week of one user's telemetry, spread over the dashboard's 7-day window."""

    def at() -> str:
        return (now - dt.timedelta(seconds=rng.uniform(600, 6.8 * 86400))).isoformat()

    def tool_error() -> dict:
        tool, code, _ = rng.choices(TOOL_ERRORS, weights=[w for *_, w in TOOL_ERRORS])[0]
        return {"event_type": "tool_error", "occurred_at": at(),
                "payload": {"tool_name": tool, "error_code": code}}

    events = []
    for _ in range(rng.randint(8, 22)):
        outcome = rng.choices(["success", "error", "cancelled", "clarify"], weights=[78, 12, 6, 4])[0]
        events.append({"event_type": "task_complete", "occurred_at": at(), "payload": {
            "duration_ms": rng.randint(4_000, 90_000),
            "turn_count": rng.randint(2, 14),
            "tool_sequence": rng.sample(TOOLS, k=rng.randint(1, 4)),
            "outcome": outcome,
        }})
        # Failed tasks always carry a tool failure; successful ones sometimes
        # do too, when the model recovered from it.
        if outcome == "error" or rng.random() < 0.2:
            events.append(tool_error())

    if version == "0.5.0":
        events.append({"event_type": "update_event", "occurred_at": at(), "payload": {
            "phase": "apply", "outcome": "installed", "from_version": "0.4.2", "to_version": "0.5.0",
        }})
    else:
        # Stuck on the old version: the updater gave up after three failed installs.
        for reason in ["package_timeout", "package_network", "package_timeout"]:
            events.append({"event_type": "update_event", "occurred_at": at(), "payload": {
                "phase": "download", "outcome": "failed", "reason_code": reason,
                "from_version": "0.4.2", "to_version": "0.5.0",
            }})
        events.append({"event_type": "update_event", "occurred_at": at(), "payload": {
            "phase": "apply", "outcome": "blocked", "from_version": "0.4.2", "to_version": "0.5.0",
        }})
    return events


if __name__ == "__main__":
    command = sys.argv[1] if len(sys.argv) > 1 else "serve"
    if command == "serve":
        serve()
    elif command == "seed":
        seed()
    else:
        raise SystemExit(__doc__)
