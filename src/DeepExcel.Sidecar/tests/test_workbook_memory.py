import asyncio
import hashlib
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import workbook_memory as wm  # noqa: E402

KEY = r"C:\财务\2026 预算.xlsx"


@pytest.fixture(autouse=True)
def memory_root(tmp_path, monkeypatch):
    monkeypatch.setenv("DEEPEXCEL_MEMORY_DIR", str(tmp_path))
    wm._current = None
    wm.reset_injection()
    return tmp_path


def test_directory_matches_the_conversation_history_hash():
    # 与 C# ConversationHistory / WPS conversation-store 同一口径：Excel 和 WPS 共用一份
    expected = hashlib.sha256(KEY.encode("utf-8")).digest()[:16].hex()
    assert wm.key_hash(KEY) == expected
    assert wm.WorkbookMemory(KEY).dir.name == expected


def test_unsaved_workbooks_have_no_memory():
    assert wm.use_context({"workbookKey": "工作簿1"}) is None
    assert wm.use_context({}) is None
    assert wm.use_context({"workbookKey": KEY, "workbookName": "2026 预算.xlsx"}).name == "2026 预算.xlsx"


def test_notes_round_trip_and_size_cap():
    memory = wm.WorkbookMemory(KEY)
    assert memory.save_notes(wm.TEMPLATE + "表头在第 3 行\n") is None
    assert "表头在第 3 行" in memory.notes()
    assert "上限" in memory.save_notes("x" * (wm.MAX_NOTES_CHARS + 1))


def test_the_agent_cannot_drop_a_protected_zone_but_the_user_can():
    memory = wm.WorkbookMemory(KEY)
    memory.save_notes("## 禁区\n- 汇总（老板要看）\n- 明细!A1:D20\n", by_agent=False)
    refusal = memory.save_notes("## 禁区\n- 明细!A1:D20\n", by_agent=True)
    assert refusal and "汇总" in refusal
    assert memory.save_notes("## 禁区\n- 汇总（老板要看）\n- 明细!A1:D20\n- 新表\n", by_agent=True) is None  # 只加可以
    assert memory.save_notes("## 禁区\n", by_agent=False) is None  # 用户在面板里解除


def test_protected_zones_parse_sheets_ranges_and_quoted_names():
    notes = "\n".join([
        "## 用户偏好", "- 金额用万元",
        "## 禁区",
        "- 汇总（老板要看的）",
        "- 明细!A1:D20",
        "- 'Q1 数据'!B:C",
        "- 参数表!3:5：别改",
        "- 坏的!ZZZZ9",  # 看不懂的区域不当禁区
        "## 做过的改动", "- 汇总表加了合计行",
    ])
    zones = wm.protected_zones(notes)
    assert [(z.sheet, z.rect) for z in zones] == [
        ("汇总", None),
        ("明细", (1, 1, 20, 4)),
        ("Q1 数据", (1, 2, wm.MAX_ROW, 3)),
        ("参数表", (3, 1, 5, wm.MAX_COL)),
    ]
    assert wm.zone_hits(zones, "汇总", (100, 1, 100, 1))           # 整表禁区
    assert wm.zone_hits(zones, "明细", (20, 4, 30, 5))             # 相交
    assert not wm.zone_hits(zones, "明细", (21, 1, 30, 4))         # 不相交
    assert wm.zone_hits(zones, "明细", None)                       # 目标区域不明：保守地算碰到
    assert wm.zone_hits(zones, "q1 数据", (7, 2, 7, 2))            # 表名不区分大小写
    assert not wm.zone_hits(zones, "其他", None)


def test_history_records_writes_not_reads_and_never_content():
    memory = wm.WorkbookMemory(KEY)
    ok = {"ok": True}
    assert wm.history_entry("read_range", {"address": "A1"}, ok) is None
    entry = wm.history_entry("write_range", {"address": "明细!A1:B2", "values": [["机密", 1]]}, ok)
    assert entry["target"] == "明细!A1:B2" and entry["ok"] is True
    assert "机密" not in json.dumps(entry, ensure_ascii=False)
    failed = wm.history_entry("execute_vba", {"code": "Sub A()"}, {"ok": False, "error": {"code": "vba_error"}})
    assert failed["error_code"] == "vba_error" and "target" not in failed
    memory.append_history(entry)
    memory.append_history(failed)
    assert [e["tool"] for e in memory.history()] == ["write_range", "execute_vba"]


def test_history_is_bounded(monkeypatch):
    monkeypatch.setattr(wm, "MAX_HISTORY_LINES", 10)
    monkeypatch.setattr(wm, "KEEP_HISTORY_LINES", 5)
    memory = wm.WorkbookMemory(KEY)
    for i in range(12):
        memory.append_history({"ts": i, "tool": f"t{i}", "ok": True})
    tools = [e["tool"] for e in memory.history()]
    assert len(tools) <= 10 and tools[-1] == "t11"


def test_summary_is_injected_once_then_again_after_user_edits_or_compaction():
    memory = wm.use_context({"workbookKey": KEY})
    assert wm.injection_for_turn(memory) is None  # 什么都没记过：不注入
    memory.save_notes("## 用户偏好\n- 金额用万元\n## 禁区\n- 汇总\n", by_agent=False)
    first = wm.injection_for_turn(memory)
    assert "<workbook-memory>" in first and "金额用万元" in first and "禁区" in first
    assert wm.injection_for_turn(memory) is None                 # 同一会话不重复
    assert wm.injection_for_turn(memory, force=True)             # 压缩过：再注入
    import os, time
    later = time.time() + 5
    os.utime(memory.notes_path, (later, later))                  # 用户在面板里改了
    assert wm.injection_for_turn(memory)


def test_notes_the_agent_just_wrote_are_not_injected_back(monkeypatch):
    import excel_tools
    memory = wm.use_context({"workbookKey": KEY})
    result = asyncio.run(excel_tools.update_workbook_notes.handler({"notes": wm.TEMPLATE + "- 表头在第 3 行\n"}))
    payload = json.loads(result["content"][0]["text"])
    assert payload["success"] is True, payload
    assert wm.injection_for_turn(memory) is None


def test_update_tool_refuses_without_a_saved_workbook():
    import excel_tools
    wm.use_context({"workbookKey": "工作簿1"})
    result = asyncio.run(excel_tools.update_workbook_notes.handler({"notes": "x"}))
    assert json.loads(result["content"][0]["text"])["success"] is False


def test_recent_history_shows_in_the_summary():
    memory = wm.use_context({"workbookKey": KEY})
    memory.append_history({"ts": 0, "tool": "write_formula", "target": "D2", "ok": False, "error_code": "com_error"})
    text = wm.injection_for_turn(memory)
    assert "write_formula D2 失败（com_error）" in text


ZONES = "## 禁区\n- 汇总\n- 明细!A1:D20\n"


def test_writes_into_a_protected_zone_are_refused():
    assert "禁区" in wm.check_write("write_value", {"address": "汇总!B2", "value": 1}, ZONES, "明细")
    assert wm.check_write("write_value", {"address": "明细!E1"}, ZONES, "汇总") is None       # 区域外
    assert wm.check_write("write_formula", {"address": "C5"}, ZONES, "明细")                  # 没带表名：落在活动表
    assert wm.check_write("write_formula", {"address": "C5"}, ZONES, "其他") is None
    assert wm.check_write("fill_formula_down", {"from_address": "明细!B18", "row_count": 100}, ZONES, None)
    assert wm.check_write("copy_range", {"source_address": "汇总!A1:B2", "dest_address": "其他!A1"}, ZONES, None) is None
    assert wm.check_write("delete_sheet", {"name": "汇总"}, ZONES, None)
    assert wm.check_write("rename_sheet", {"old_name": "汇总", "new_name": "x"}, ZONES, None)
    assert wm.check_write("create_chart", {"data_range": "明细!A1:B5"}, ZONES, "汇总")           # 图表落在活动表
    assert wm.check_write("freeze_panes", {"address": "B2"}, ZONES, "汇总") is None               # 只改视图
    assert wm.check_write("read_range", {"address": "汇总!A1"}, ZONES, None) is None


def test_inserting_rows_above_a_zone_counts_as_touching_it():
    assert wm.check_write("insert_rows", {"row": 5, "count": 1}, ZONES, "明细")      # 会把 A5:D20 往下推
    assert wm.check_write("insert_rows", {"row": 30, "count": 1}, ZONES, "明细") is None
    assert wm.check_write("delete_columns", {"column": "C", "count": 1}, ZONES, "明细")
    assert wm.check_write("delete_columns", {"column": "E", "count": 1}, ZONES, "明细") is None


def test_code_that_names_a_protected_sheet_is_refused():
    assert wm.check_write("execute_vba", {"code": 'Worksheets("汇总").Range("A1").Value = 1'}, ZONES, "其他")
    assert wm.check_write("execute_jsa", {"code": "Application.Sheets.Item('汇总').Range('A1').Value2 = 1"}, ZONES, None)
    assert wm.check_write("execute_vba", {"code": "ActiveSheet.Range(\"A1\").Value = 1"}, ZONES, "汇总")
    assert wm.check_write("execute_vba", {"code": 'Worksheets("其他").Range("A1").Value = 1'}, ZONES, "明细") is None


def test_no_zones_means_no_checks_and_unknown_sheets_ask_for_a_name():
    assert wm.check_write("write_value", {"address": "汇总!A1"}, "## 用户偏好\n- 万元\n", None) is None
    refusal = wm.check_write("write_value", {"address": "A1"}, ZONES, None)
    assert refusal and "表名" in refusal


def test_active_sheet_comes_from_either_host():
    wm.use_context({"workbookKey": KEY, "workbook": {"activeSheet": "明细"}})   # Excel
    assert wm.active_sheet() == "明细"
    wm.use_context({"workbookKey": KEY, "activeSheet": "汇总"})                  # WPS
    assert wm.active_sheet() == "汇总"


def test_the_hook_denies_protected_writes(tmp_path):
    import sidecar
    memory = wm.use_context({"workbookKey": KEY, "workbook": {"activeSheet": "明细"}})
    memory.save_notes(ZONES, by_agent=False)
    denied = asyncio.run(sidecar._pre_tool_use_hook(
        {"tool_name": "mcp__excel__write_value", "tool_input": {"address": "汇总!A1", "value": 1}}, "t1", None))
    assert denied["hookSpecificOutput"]["permissionDecision"] == "deny"
    assert "禁区" in denied["hookSpecificOutput"]["permissionDecisionReason"]
    allowed = asyncio.run(sidecar._pre_tool_use_hook(
        {"tool_name": "mcp__excel__write_value", "tool_input": {"address": "其他!A1", "value": 1}}, "t2", None))
    assert allowed["hookSpecificOutput"]["permissionDecision"] == "allow"
