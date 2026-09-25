"""Live end-to-end check of the sidecar's ui_event stream against a real model.

Starts sidecar.py the way WPS does (DEEPEXCEL_HOST=wps: it reads the current
user's DeepExcel config and DPAPI credential itself, so this script never sees
the API key), plays a fake host that answers tool calls, and prints the event
stream. One tool call is made to fail on purpose so the error path is exercised.

Costs a few thousand tokens on whatever model is configured. Not part of CI.

    python scripts/live_sidecar_events.py              # 事件配对、失败上报、终态行
    python scripts/live_sidecar_events.py --interrupt  # 停止后再问新问题，不能读到残留
    python scripts/live_sidecar_events.py --steer      # 工具执行期间插话，模型照做
    python scripts/live_sidecar_events.py --codegen    # 写 VBA 时边写边显示（tool_gen）
    python scripts/live_sidecar_events.py --plan       # 多步任务用 todo_write 列计划并更新
    python scripts/live_sidecar_events.py --inspect    # 陌生的表先 inspect_sheet，能转述异常候选
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


def interrupt_scenario() -> int:
    """按停止之后再问一个新问题：回答必须是新问题的，而不是上一轮的残留。

    第一轮让模型调用一个「很慢」的工具（假宿主故意不回结果），按停止；
    然后问「1+1 等于几，只回答数字」。"""
    sys.stdout.reconfigure(encoding="utf-8")
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "user_message", "text": "用 read_range 读取 A1:Z5000，然后告诉我有多少行。", "context": {}})
    turn, texts, outcomes, t0 = 1, {1: "", 2: ""}, {}, time.time()
    stopped_at = None
    for raw in proc.stdout:
        if time.time() - t0 > 240:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call" and turn == 1 and stopped_at is None:
            # 不回结果，模拟宿主卡在一个很慢的操作上；按停止
            stopped_at = time.time()
            send({"type": "cancel"})
        elif kind == "tool_call":
            send({"type": "tool_result", "call_id": msg["call_id"], "success": True,
                  "data": {"values": [[1]]}, "context": {}})
        elif kind == "ui_event":
            ev = msg["event"]
            if ev["kind"] in ("run_summary", "status", "tool_end"):
                print(f"turn {turn} ui_event", json.dumps({k: v for k, v in ev.items() if k not in ("v", "ts")},
                                                         ensure_ascii=False)[:200])
            if ev["kind"] == "run_summary":
                outcomes[turn] = ev["outcome"]
        elif kind == "stream_delta":
            texts[turn] += msg.get("text", "")
        elif kind == "stream_end":
            if turn == 1:
                print(f"turn 1 ended {time.time() - (stopped_at or t0):.1f}s after stop")
                turn = 2
                send({"type": "user_message", "text": "1+1 等于几？只回答数字，不要调用任何工具。", "context": {}})
            else:
                break
    proc.kill()
    print("turn 2 text:", texts[2].strip()[:200])
    problems = []
    if outcomes.get(1) != "interrupted":
        problems.append(f"turn 1 outcome={outcomes.get(1)}")
    if "2" not in texts[2]:
        problems.append("turn 2 did not answer the new question")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def steer_scenario() -> int:
    """任务进行中插话：第一个工具执行期间用户说「改写到 F2」，模型应当照做。"""
    sys.stdout.reconfigure(encoding="utf-8")
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "user_message", "context": {},
          "text": "先用 read_range 读取 A1:C4，再用 write_formula 在 D2 写入 =SUM(A2:C2)。"})
    writes, kinds, steered, t0 = [], [], False, time.time()
    for raw in proc.stdout:
        if time.time() - t0 > 240:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            if msg["tool"] == "read_range" and not steered:
                steered = True
                # 工具还在执行时用户插话，然后再回结果
                send({"type": "user_message", "text": "等一下，公式不要写到 D2，改写到 F2。",
                      "steer": True, "context": {}})
                time.sleep(0.3)
            if msg["tool"] == "write_formula":
                writes.append((msg.get("args") or {}).get("address"))
            result = fake_result(msg["tool"], msg.get("args") or {}, {"write_formula": 1})
            send({"type": "tool_result", "call_id": msg["call_id"], "context": {}, **result})
        elif kind == "permission_request":
            # 模型可能想撤掉已经写进 D2 的公式（高风险工具需要确认）
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "allow"})
        elif kind == "ui_event":
            kinds.append(msg["event"]["kind"])
        elif kind == "stream_end":
            break
    proc.kill()
    print("ui_event kinds:", kinds)
    print("write_formula addresses:", writes)
    problems = []
    if "steer_delivered" not in kinds:
        problems.append("interjection was not delivered mid-run")
    # 插话在下一个工具结果之后才送达：和它同一批发出的调用（例如并行发出的 D2 写入）
    # 已经在路上，撤不回来——Claude Code 也一样。要求的是送达之后照做。
    if "F2" not in writes:
        problems.append(f"model did not follow the interjection: {writes}")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def codegen_scenario() -> int:
    """模型写一段 VBA：写的过程中应当收到带代码预览的 tool_gen，之后同一 id 的 tool_start。"""
    sys.stdout.reconfigure(encoding="utf-8")
    env = dict(os.environ, DEEPEXCEL_HOST="excel", PYTHONIOENCODING="utf-8")
    # Excel 会话才有 execute_vba；配置照样从本机读（DEEPEXCEL_HOST=excel 时侧车等 config 消息，
    # 这里借 WPS 的读取逻辑在本进程里准备好再发过去，Key 只经过内存）
    sys.path.insert(0, os.path.dirname(SIDECAR))
    import sidecar as sidecar_module  # noqa: E402
    cfg = sidecar_module._load_wps_local_config()
    if not cfg:
        print("local DeepExcel config unavailable")
        return 1
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "config", "routing_mode": "byok", **cfg})
    del cfg
    send({"type": "user_message", "context": {},
          "text": "用 execute_vba 写一个大约 25 行的宏：遍历 Sheet1 的 A2:A200，把负数标红、把空单元格填 0，"
                  "最后在 B1 写上处理了多少个单元格。直接调用工具，不要先解释。"})
    gens, starts, t0 = [], [], time.time()
    for raw in proc.stdout:
        if time.time() - t0 > 240:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            send({"type": "tool_result", "call_id": msg["call_id"], "success": True,
                  "data": {"message": "宏已执行"}, "context": {}})
        elif kind == "permission_request":
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "allow"})
        elif kind == "ui_event":
            ev = msg["event"]
            if ev["kind"] == "tool_gen":
                gens.append(ev)
            elif ev["kind"] == "tool_start":
                starts.append(ev)
        elif kind == "stream_end":
            break
    proc.kill()
    print(f"tool_gen events: {len(gens)}; lines seen: {[g.get('lines') for g in gens][:12]}")
    if gens:
        print("last preview tail:", (gens[-1].get("preview") or "")[-160:].replace("\n", " | "))
    problems = []
    vba_gens = [g for g in gens if g["name"] == "execute_vba" and g.get("preview")]
    if len(vba_gens) < 2:
        problems.append("expected several tool_gen events with a code preview")
    if vba_gens and not any(s["id"] == vba_gens[0]["id"] for s in starts):
        problems.append("tool_start did not reuse the tool_gen id")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def plan_scenario() -> int:
    """多步任务：模型应当先用 todo_write 列计划，并随进度更新状态。"""
    sys.stdout.reconfigure(encoding="utf-8")
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "user_message", "context": {},
          "text": "帮我做这几件事：1) 读取 A1:C4；2) 在 D2 写 =SUM(A2:C2)；3) 把 D2 设成两位小数；"
                  "4) 在 A6 写上「已汇总」。"})
    plans, t0 = [], time.time()
    for raw in proc.stdout:
        if time.time() - t0 > 240:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            result = fake_result(msg["tool"], msg.get("args") or {}, {"write_formula": 1})
            send({"type": "tool_result", "call_id": msg["call_id"], "context": {}, **result})
        elif kind == "permission_request":
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "allow"})
        elif kind == "ui_event" and msg["event"]["kind"] == "plan":
            items = msg["event"]["items"]
            plans.append(items)
            print("plan", " | ".join(f"{i['status'][:4]}:{i['content'][:14]}" for i in items))
        elif kind == "stream_end":
            break
    proc.kill()
    problems = []
    if not plans:
        problems.append("model never used todo_write")
    elif not all(i["status"] == "completed" for i in plans[-1]):
        problems.append("final plan is not all completed")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def inspect_scenario() -> int:
    """「检查这张工资表」：模型应当调用 inspect_sheet，并把埋进去的问题（C10 合计漏行、
    E7 被改成死值）说出来。快照用真 Excel 导出的 fixture，宿主只回 sheet_snapshot。"""
    sys.stdout.reconfigure(encoding="utf-8")
    fixture = os.path.join(ROOT, "src", "DeepExcel.Sidecar", "tests", "fixtures", "snapshot_payroll.json")
    with open(fixture, encoding="utf-8") as stream:
        snapshot = json.load(stream)
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "user_message", "context": {},
          "text": "帮我检查一下 Data 这张工资表的公式和合计有没有问题，先别改，告诉我哪里可疑。"})
    host_calls, text, t0 = [], [], time.time()
    for raw in proc.stdout:
        if time.time() - t0 > 240:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            host_calls.append(msg["tool"])
            print("host   ", msg["tool"], json.dumps(msg.get("args") or {}, ensure_ascii=False)[:120])
            if msg["tool"] == "sheet_snapshot":
                result = {"success": True, "data": snapshot}
            elif msg["tool"] == "list":
                result = {"success": True, "data": {"kind": "sheets", "count": 1,
                                                    "items": [{"name": "Data", "used_range": "A1:F10"}]}}
            else:
                result = {"success": False, "error": "这个场景只提供 inspect_sheet / list",
                          "suggestion": "用 inspect_sheet(sheet=\"Data\")"}
            send({"type": "tool_result", "call_id": msg["call_id"],
                  "context": {}, **result})
        elif kind == "permission_request":
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "deny"})
        elif kind == "stream_delta":
            text.append(msg.get("text", ""))
        elif kind == "stream_end":
            break
    proc.kill()
    answer = "".join(text)
    print("\n--- 回复 ---\n" + answer)
    problems = []
    if "sheet_snapshot" not in host_calls:
        problems.append("model never called inspect_sheet")
    for cell in ("C10", "E7"):
        if cell not in answer:
            problems.append(f"answer does not mention {cell}")
    if any(t in host_calls for t in ("write_value", "write_formula", "write_range")):
        problems.append("model modified the sheet although told not to")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


if __name__ == "__main__":
    if "--inspect" in sys.argv:
        sys.exit(inspect_scenario())
    if "--plan" in sys.argv:
        sys.exit(plan_scenario())
    if "--codegen" in sys.argv:
        sys.exit(codegen_scenario())
    if "--interrupt" in sys.argv:
        sys.exit(interrupt_scenario())
    if "--steer" in sys.argv:
        sys.exit(steer_scenario())
    sys.exit(main())
