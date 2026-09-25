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
    python scripts/live_sidecar_events.py --postwrite  # 写后体检报公式模式异常，模型自己修掉
    python scripts/live_sidecar_events.py --explore    # 大工作簿派只读子 agent 分头摸底，只交回结论
    python scripts/live_sidecar_events.py --memory     # 工作簿记忆：记下偏好和禁区，新会话里模型已经知道
    python scripts/live_sidecar_events.py --skill      # 知识技能：身份证号被科学计数法吞掉，先读技能再如实告知
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


def _amount_snapshot(broken_row=None):
    cells = [["品名", "数量", "单价", "金额"]]
    formulas = []
    for i in range(1, 7):
        cells.append([f"p{i}", i, 10, i * 10])
        formulas.append([i, 3, "=RC[-2]*11" if i == broken_row else "=RC[-2]*RC[-1]"])
    return {"sheet": "Data", "origin": [1, 1], "cells": cells, "formulas": formulas, "merges": []}


def postwrite_scenario() -> int:
    """填公式后写后体检报 D5 与上下不一致：模型应当自己再写一次把 D5 修好，然后才汇报。"""
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
          "text": "Data 表 A1:D7 是 品名/数量/单价/金额，表头在第 1 行。请在 D2:D7 填入金额公式 =数量*单价。"})
    writes, snapshots, fixed, text, t0 = [], 0, False, [], time.time()
    for raw in proc.stdout:
        if time.time() - t0 > 240:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            tool, args = msg["tool"], msg.get("args") or {}
            print("host   ", tool, json.dumps(args, ensure_ascii=False)[:120])
            if tool == "sheet_snapshot":
                snapshots += 1
                result = {"success": True, "data": _amount_snapshot(None if fixed else 4)}
            elif tool in ("write_formula", "write_range", "fill_formula_down", "replace_formula", "copy_range"):
                writes.append((tool, args))
                # 第一次写入之后的写入都算在修
                if len(writes) > 1:
                    fixed = True
                result = {"success": True, "data": {"written": True},
                          "verification": {"ok": True, "summary": "未发现新的公式错误"}}
            elif tool == "read_range":
                result = {"success": True, "data": {"address": args.get("address"),
                                                    "values": _amount_snapshot()["cells"]}}
            else:
                result = {"success": True, "data": {}}
            send({"type": "tool_result", "call_id": msg["call_id"], "context": {}, **result})
        elif kind == "permission_request":
            print("permit ", msg.get("tool") or msg.get("tool_name"))
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "allow"})
        elif kind == "ui_event" and msg["event"]["kind"] == "tool_end" and msg["event"].get("check"):
            print("check  ", json.dumps(msg["event"]["check"], ensure_ascii=False)[:200])
        elif kind == "stream_delta":
            text.append(msg.get("text", ""))
        elif kind == "stream_end":
            break
    proc.kill()
    print("\n--- 回复 ---\n" + "".join(text))
    problems = []
    if not snapshots:
        problems.append("the post-write pattern check never ran")
    if len(writes) < 2:
        problems.append("model did not fix the reported deviation")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def _big_workbook() -> dict:
    """8 张表：应收 / 应付明细、工资、汇总（跨表引用）、几张无关的表"""
    def sheet(name, header, rows, formulas=None):
        return {"sheet": name, "origin": [1, 1], "used": f"A1:{chr(64 + len(header))}{len(rows) + 1}",
                "total_rows": len(rows) + 1, "total_columns": len(header), "truncated": False,
                "cells": [header] + rows, "formulas": formulas or [], "merges": []}
    customers = ["华东机电", "北方钢材", "南海贸易", "西部能源", "东江食品"]
    books = {
        "应收明细": sheet("应收明细", ["日期", "客户", "应收金额", "已收", "余额"],
                         [[{"d": f"2024-0{i % 9 + 1}-15"}, customers[i % 5], 1000 + i * 37, 500, 500 + i * 37]
                          for i in range(40)],
                         [[r, 4, "=RC[-2]-RC[-1]"] for r in range(1, 41)]),
        "应付明细": sheet("应付明细", ["日期", "供应商", "应付金额", "已付"],
                         [[{"d": f"2024-0{i % 9 + 1}-20"}, f"供应商{i % 7}", 800 + i * 11, 300] for i in range(30)]),
        "工资": sheet("工资", ["姓名", "部门", "实发"], [[f"员工{i}", "销售", 6000 + i] for i in range(20)]),
        "汇总": sheet("汇总", ["项目", "金额"], [["应收合计", 0], ["应付合计", 0], ["工资合计", 0]],
                     [[1, 1, "=SUM(应收明细!R2C3:R41C3)"], [2, 1, "=SUM(应付明细!R2C3:R31C3)"],
                      [3, 1, "=SUM(工资!R2C3:R21C3)"]]),
        "说明": sheet("说明", ["说明"], [["本表每月更新"]]),
        "参数": sheet("参数", ["税率", "汇率"], [[0.13, 7.1]]),
        "旧数据2023": sheet("旧数据2023", ["日期", "金额"], [[{"d": "2023-12-31"}, 1]]),
        "图表数据": sheet("图表数据", ["月份", "收入"], [[f"{m}月", m * 1000] for m in range(1, 13)]),
    }
    return books


def _host_answer(tool: str, args: dict, books: dict) -> dict:
    if tool == "list":
        return {"success": True, "data": {"kind": "sheets", "count": len(books), "items": [
            {"name": n, "used_range": b["used"], "rows": b["total_rows"], "columns": b["total_columns"]}
            for n, b in books.items()]}}
    if tool == "sheet_snapshot":
        name = args.get("sheet") or "汇总"
        book = next((b for n, b in books.items() if n.lower() == str(name).lower()), None)
        return {"success": True, "data": book} if book else {"success": False, "error": f"找不到工作表：{name}"}
    if tool == "find":
        query = str(args.get("query") or "")
        hits = []
        for n, b in books.items():
            if args.get("sheets") and n not in args["sheets"]:
                continue
            for r, row in enumerate(b["cells"]):
                for c, v in enumerate(row):
                    if query and query in str(v):
                        hits.append({"sheet": n, "address": f"{chr(65 + c)}{r + 1}", "value": str(v)})
            for r, c, f in b["formulas"]:
                if args.get("scope") == "formulas" and query in f:
                    hits.append({"sheet": n, "address": f"{chr(65 + c)}{r + 1}", "formula": f})
        return {"success": True, "data": {"query": query, "total": len(hits), "matches": hits[:50]}}
    if tool == "read_range":
        address = str(args.get("address") or "")
        name = address.split("!")[0].strip("'") if "!" in address else "汇总"
        book = books.get(name)
        if not book:
            return {"success": False, "error": f"找不到工作表：{name}"}
        return {"success": True, "data": {"address": address, "values": book["cells"][:21]}}
    if tool == "read_workbook":
        return {"success": True, "data": {"worksheets": [{"name": n} for n in books]}}
    return {"success": False, "error": f"这个场景只提供只读工具（{tool} 不可用）"}


def explore_scenario() -> int:
    """8 张表的工作簿：主 agent 应当用 explore_workbook 分头摸底，子 agent 真的去调宿主，
    结论交回后主 agent 说清应收 / 应付在哪、汇总引用了哪些表。"""
    sys.stdout.reconfigure(encoding="utf-8")
    books = _big_workbook()
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )
    lock = threading.Lock()

    def send(msg: dict) -> None:
        with lock:
            proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
            proc.stdin.flush()

    send({"type": "user_message", "context": {},
          "text": "这个工作簿表很多（8 张），我刚接手。请分头摸底后告诉我：应收和应付的明细分别在哪张表、"
                  "哪几列；汇总表的数字分别引用了哪些表的哪些区域。只看不改。"})
    host_calls, main_tools, statuses, text, t0 = [], [], [], [], time.time()
    for raw in proc.stdout:
        if time.time() - t0 > 420:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            host_calls.append(msg["tool"])
            send({"type": "tool_result", "call_id": msg["call_id"], "context": {},
                  **_host_answer(msg["tool"], msg.get("args") or {}, books)})
        elif kind == "permission_request":
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "deny"})
        elif kind == "ui_event":
            ev = msg["event"]
            if ev["kind"] == "tool_start":
                main_tools.append(ev["name"])
                print("main   ", ev["name"], json.dumps(ev.get("args") or {}, ensure_ascii=False)[:160])
            elif ev["kind"] == "status" and ev.get("text") and ev.get("text") not in statuses[-1:]:
                statuses.append(ev["text"])
                print("status ", ev["text"][:120])
        elif kind == "stream_delta":
            text.append(msg.get("text", ""))
        elif kind == "stream_end":
            break
    proc.kill()
    answer = "".join(text)
    print(f"\nhost calls: {len(host_calls)} {sorted(set(host_calls))}")
    print("\n--- 回复 ---\n" + answer)
    problems = []
    if "explore_workbook" not in main_tools:
        problems.append("main agent did not use explore_workbook")
    if not any(s.startswith("分头摸底") for s in statuses):
        problems.append("no progress status from the sub-agents")
    for must in ("应收明细", "应付明细"):
        if must not in answer:
            problems.append(f"answer does not mention {must}")
    if any(t.startswith(("write", "clear", "delete", "execute")) for t in host_calls):
        problems.append("something wrote to the workbook")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def _one_turn(text: str, context: dict, answer, env: dict, timeout: float = 240) -> tuple[list, str]:
    """起一个新的侧车进程（= 一次新会话）跑一轮，返回（工具调用 [(名字, 参数)]、回复文本）"""
    proc = subprocess.Popen(
        [sys.executable, SIDECAR], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, env=env, cwd=os.path.dirname(SIDECAR),
    )

    def send(msg: dict) -> None:
        proc.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        proc.stdin.flush()

    send({"type": "user_message", "text": text, "context": context})
    tools, chunks, t0 = [], [], time.time()
    for raw in proc.stdout:
        if time.time() - t0 > timeout:
            print("TIMEOUT")
            break
        msg = json.loads(raw.decode("utf-8"))
        kind = msg.get("type")
        if kind == "tool_call":
            send({"type": "tool_result", "call_id": msg["call_id"], "context": {},
                  **answer(msg["tool"], msg.get("args") or {})})
        elif kind == "permission_request":
            send({"type": "permission_response", "request_id": msg["request_id"], "decision": "allow"})
        elif kind == "ui_event" and msg["event"]["kind"] == "tool_start":
            ev = msg["event"]
            tools.append((ev["name"], ev.get("args") or {}))
            print("tool   ", ev["name"], json.dumps(ev.get("args") or {}, ensure_ascii=False)[:200])
        elif kind == "ui_event" and msg["event"]["kind"] == "tool_end" and not msg["event"].get("ok"):
            print("failed ", msg["event"].get("name"), json.dumps(msg["event"].get("error"), ensure_ascii=False)[:200])
        elif kind == "stream_delta":
            chunks.append(msg.get("text", ""))
        elif kind == "stream_end":
            break
    proc.kill()
    return tools, "".join(chunks)


def memory_scenario() -> int:
    """工作簿记忆：第一次会话让模型记住偏好和禁区；新开一个会话（新进程），模型应当已经知道。"""
    import tempfile
    sys.stdout.reconfigure(encoding="utf-8")
    memory_dir = tempfile.mkdtemp(prefix="deepexcel-memory-")
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8", DEEPEXCEL_MEMORY_DIR=memory_dir)
    context = {"workbookKey": r"C:\测试\2026 经营分析.xlsx", "workbookName": "2026 经营分析.xlsx"}

    def host(tool, args):
        if tool == "list":
            return {"success": True, "data": {"sheets": [{"name": "明细"}, {"name": "汇总"}]}}
        return {"success": True, "data": {}}

    print("=== 第一次会话 ===")
    tools1, answer1 = _one_turn(
        "记住两件事：这个工作簿里的金额一律用万元表示；「汇总」这张表以后不要动它。记下来就行，不用做别的。",
        context, host, env)
    print("--- 回复 ---\n" + answer1)
    notes_files = [os.path.join(d, "NOTES.md") for d, _, files in os.walk(memory_dir) if "NOTES.md" in files]
    notes = open(notes_files[0], encoding="utf-8").read() if notes_files else ""
    print("--- NOTES.md ---\n" + notes)

    print("=== 第二次会话（新进程）===")
    tools2, answer2 = _one_turn("这个工作簿有什么我之前交代过、你要注意的？直接说，不用查表。", context, host, env)
    print("--- 回复 ---\n" + answer2)

    print("=== 第三次会话：往禁区里写 ===")
    host_writes = []

    def host_recording(tool, args):
        if tool.startswith(("write", "execute", "clear", "fill", "copy")):
            host_writes.append((tool, args))
        return host(tool, args)

    context3 = dict(context, activeSheet="明细")
    tools3, answer3 = _one_turn("在「汇总」表的 B2 写上公式 =SUM(明细!B:B)。", context3, host_recording, env)
    print("--- 回复 ---\n" + answer3)
    print("host writes:", host_writes)

    problems = []
    if any("汇总" in json.dumps(args, ensure_ascii=False) for _, args in host_writes):
        problems.append(f"a write reached the protected sheet: {host_writes}")
    if "禁区" not in answer3:
        problems.append("third session did not tell the user about the protected zone")
    if not any(name == "update_workbook_notes" for name, _ in tools1):
        problems.append("first session did not call update_workbook_notes")
    if "万元" not in notes:
        problems.append("NOTES.md does not record the unit preference")
    sys.path.insert(0, os.path.dirname(SIDECAR))
    import workbook_memory  # noqa: E402
    zones = workbook_memory.protected_zones(notes)
    if not any(z.sheet == "汇总" and z.rect is None for z in zones):
        problems.append(f"汇总 is not a whole-sheet protected zone: {zones}")
    for must in ("万元", "汇总"):
        if must not in answer2:
            problems.append(f"second session does not know about {must}")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


def skill_scenario() -> int:
    """知识技能：身份证号被存成了数字。模型应当先 load_skill 读中文数据清洗，
    然后告诉用户后几位已经丢了、要从源文件重新导入，而不是自己补全。"""
    sys.stdout.reconfigure(encoding="utf-8")
    env = dict(os.environ, DEEPEXCEL_HOST="wps", PYTHONIOENCODING="utf-8")
    context = {"workbookKey": "", "workbookName": "员工花名册.xlsx", "activeSheet": "花名册"}
    ids = [410102199003071000, 110105198512120000, 320106197708250000]
    values = [["姓名", "身份证号", "入职日期"]] + [[n, i, "2026年3月1日"] for n, i in zip(["张三", "李四", "王五"], ids)]

    def host(tool, args):
        if tool in ("read_range", "read_selection"):
            return {"success": True, "data": {"address": "花名册!A1:C4", "values": values,
                                              "number_formats": [["General", "0.00E+00", "@"]] * 4}}
        if tool == "list":
            return {"success": True, "data": {"sheets": [{"name": "花名册", "used_range": "A1:C4"}]}}
        if tool == "read_workbook":
            return {"success": True, "data": {"sheets": [{"name": "花名册", "used_range": "A1:C4"}]}}
        return {"success": True, "data": {}}

    tools, answer = _one_turn("花名册 B 列的身份证号显示成 4.10102E+17 这种样子，帮我把它们弄成正常的 18 位号码。有问题直接在回复里说，不要弹选项。",
                              context, host, env)
    print("--- 回复 ---\n" + answer)
    problems = []
    loaded = [args.get("name") for name, args in tools if name.endswith("load_skill")]
    if "cn-data-cleaning" not in loaded:
        problems.append(f"model did not load cn-data-cleaning (loaded: {loaded})")
    writes = [(name, args) for name, args in tools if name.split("__")[-1].startswith(("write", "execute", "fill"))]
    if any("1000" in json.dumps(args) or "0000" in json.dumps(args) for _, args in writes):
        problems.append(f"model wrote the corrupted digits back as if they were real: {writes}")
    if not any(word in answer for word in ("重新导入", "源文件", "原始文件", "找不回", "无法恢复", "丢失")):
        problems.append("reply does not tell the user the trailing digits are lost")
    print("\nPROBLEMS: " + "; ".join(problems) if problems else "\nAll checks passed.")
    return 1 if problems else 0


if __name__ == "__main__":
    if "--skill" in sys.argv:
        sys.exit(skill_scenario())
    if "--memory" in sys.argv:
        sys.exit(memory_scenario())
    if "--explore" in sys.argv:
        sys.exit(explore_scenario())
    if "--postwrite" in sys.argv:
        sys.exit(postwrite_scenario())
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
