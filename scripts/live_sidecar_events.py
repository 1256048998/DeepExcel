"""Live end-to-end check of the sidecar's ui_event stream against a real model.

Starts sidecar.py the way WPS does (DEEPEXCEL_HOST=wps: it reads the current
user's DeepExcel config and DPAPI credential itself, so this script never sees
the API key), plays a fake host that answers tool calls, and prints the event
stream. One tool call is made to fail on purpose so the error path is exercised.

Costs a few thousand tokens on whatever model is configured. Not part of CI.

    python scripts/live_sidecar_events.py
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SIDECAR = os.path.join(ROOT, "src", "DeepExcel.Sidecar", "sidecar.py")

PROMPT = (
    "先用 read_range 读取 A1:C4，然后用 write_formula 在 D2 写入 =SUM(A2:C2)。"
    "如果写入失败，按提示换一种方式再试一次（例如写到 E2），最后用一句话说明结果。"
)

FAKE_VALUES = [["一月", "二月", "三月"], [1, 2, 3], [4, 5, 6], [7, 8, 9]]


def fake_result(tool: str, args: dict, attempt: dict) -> dict:
    if tool == "read_range":
        return {"success": True, "data": {"address": args.get("address"), "values": FAKE_VALUES}}
    if tool == "write_formula":
        attempt["write_formula"] = attempt.get("write_formula", 0) + 1
        if attempt["write_formula"] == 1:
            return {"success": False, "error": "D 列已被保护，不能写入",
                    "suggestion": "改写到 E 列"}
        return {"success": True, "data": {"message": f"已写入 {args.get('address')}"}}
    return {"success": True, "data": {}}


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, env=env, cwd=os.path.dirname(SIDECAR),
    )
    stderr_tail: list[str] = []

    def drain_stderr() -> None:
        for raw in proc.stderr:
            stderr_tail.append(raw.decode("utf-8", "replace").rstrip())
            del stderr_tail[:-30]

    threading.Thread(target=drain_stderr, daemon=True).start()

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "user_message", "text": PROMPT, "context": {}})
    attempt: dict = {}
    events: list[dict] = []
    deadline = time.time() + 240
    ok = False
    for raw in proc.stdout:
        if time.time() > deadline:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            result = fake_result(msg["tool"], msg.get("args") or {}, attempt)
            send({"type": "tool_result", "call_id": msg["call_id"], "context": {}, **result})
        elif kind == "permission_request":
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "allow"})
        elif kind == "ui_event":
            ev = msg["event"]
            events.append(ev)
            brief = {k: v for k, v in ev.items() if k not in ("v", "ts")}
            print("ui_event", json.dumps(brief, ensure_ascii=False)[:300])
        elif kind == "stream_delta":
            print("text   ", msg.get("text", "").replace("\n", " ")[:120])
        elif kind == "stream_end":
            print("stream_end", msg.get("input_tokens"), msg.get("output_tokens"))
            ok = True
            break
        else:
            print(kind)
    proc.kill()

    if not ok:
        print("\n".join(stderr_tail[-15:]))
        return 1

    starts = {e["id"] for e in events if e["kind"] == "tool_start"}
    ends = {e["id"] for e in events if e["kind"] == "tool_end"}
    problems = []
    if not starts:
        problems.append("no tool_start at all")
    if starts != ends:
        problems.append(f"unpaired tool ids: {sorted(starts ^ ends)}")
    if not any(e["kind"] == "tool_end" and not e["ok"] for e in events):
        problems.append("the deliberate failure did not surface as tool_end ok=false")
    if not any(e["kind"] == "run_summary" for e in events):
        problems.append("no run_summary")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
