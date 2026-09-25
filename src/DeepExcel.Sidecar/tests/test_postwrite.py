"""写公式之后的模式检查并入写后体检。"""

import json
from unittest.mock import AsyncMock, patch

import pytest

from perception.postwrite import parse_a1, pattern_check, summary_line


def _amount_snapshot(broken_row=None, constant_row=None, n=6):
    """品名 / 数量 / 单价 / 金额，金额列 =RC[-2]*RC[-1]；可在某行放一个不同的公式或死值"""
    cells = [["品名", "数量", "单价", "金额"]]
    formulas = []
    for i in range(1, n + 1):
        cells.append([f"p{i}", i, 10, i * 10])
        if i == constant_row:
            continue
        formulas.append([i, 3, "=RC[-2]*11" if i == broken_row else "=RC[-2]*RC[-1]"])
    return {"sheet": "Data", "origin": [1, 1], "cells": cells, "formulas": formulas, "merges": []}


def _fn(tool):
    fn = getattr(tool, "handler", None) or tool
    return getattr(fn, "__wrapped__", fn)


def test_parse_a1():
    assert parse_a1("D2") == (None, 2, 4, 2, 4)
    assert parse_a1("Sheet1!$B$2:D9") == ("Sheet1", 2, 2, 9, 4)
    assert parse_a1("'Q1 汇总'!C3") == ("Q1 汇总", 3, 3, 3, 3)
    assert parse_a1("D9:B2") == (None, 2, 2, 9, 4)
    assert parse_a1("A:A") is None
    assert parse_a1("税率") is None


def test_a_deviation_inside_the_written_range_is_reported():
    check = pattern_check(_amount_snapshot(broken_row=4), (2, 4, 7, 4))  # 写的是 D2:D7
    assert [c["address"] for c in check["in_written_range"]] == ["D5"]
    assert check["inconsistent_columns"] == ["D 列（金额）"]
    assert "D5" in summary_line(check)


def test_a_problem_outside_the_written_range_is_only_context():
    check = pattern_check(_amount_snapshot(constant_row=5), (2, 4, 3, 4))  # 只写了 D2:D3
    assert "in_written_range" not in check
    assert [c["address"] for c in check["elsewhere_in_column"]] == ["D6"]
    assert summary_line(check) is None


def test_a_clean_write_reports_nothing():
    assert pattern_check(_amount_snapshot(), (2, 4, 7, 4)) is None
    assert pattern_check(_amount_snapshot(broken_row=4), (2, 1, 7, 1)) is None  # 写的是 A 列


def _host(snapshot, write_result=None):
    """按工具名回答的假宿主"""
    calls = []

    async def call(tool, args, **_):
        calls.append((tool, args))
        if tool == "sheet_snapshot":
            return snapshot
        return write_result or {"success": True, "data": {"written": True},
                                "verification": {"ok": True, "summary": "体检通过"}}
    return call, calls


@pytest.mark.asyncio
async def test_fill_down_merges_the_pattern_check_into_verification():
    from excel_tools import fill_formula_down
    call, calls = _host({"success": True, "data": _amount_snapshot(broken_row=4)})
    with patch("excel_tools.call_csharp", new=AsyncMock(side_effect=call)):
        out = await _fn(fill_formula_down)({"from_address": "D2", "row_count": 5})
    parsed = json.loads(out["content"][0]["text"])
    v = parsed["verification"]
    assert v["ok"] is False
    assert v["summary"].startswith("体检通过；公式模式检查：D5")
    assert v["pattern_check"]["in_written_range"][0]["address"] == "D5"
    assert calls[1] == ("sheet_snapshot", {"max_cells": 20000})
    assert out["is_error"] is False  # 写入本身成功，只是体检没过


@pytest.mark.asyncio
async def test_snapshot_failure_leaves_the_result_alone():
    from excel_tools import write_formula
    call, _ = _host({"success": False, "error": "读不到"})
    with patch("excel_tools.call_csharp", new=AsyncMock(side_effect=call)):
        out = await _fn(write_formula)({"address": "Data!D2", "formula": "=B2*C2"})
    parsed = json.loads(out["content"][0]["text"])
    assert parsed["verification"] == {"ok": True, "summary": "体检通过"}


@pytest.mark.asyncio
async def test_failed_writes_and_plain_values_skip_the_check():
    from excel_tools import write_formula, write_range
    call, calls = _host({"success": True, "data": _amount_snapshot(broken_row=4)},
                        write_result={"success": False, "error": "被保护"})
    with patch("excel_tools.call_csharp", new=AsyncMock(side_effect=call)):
        await _fn(write_formula)({"address": "D2", "formula": "=1"})
    assert [t for t, _ in calls] == ["write_formula"]

    call, calls = _host({"success": True, "data": _amount_snapshot(broken_row=4)})
    with patch("excel_tools.call_csharp", new=AsyncMock(side_effect=call)):
        await _fn(write_range)({"address": "A1", "values": [["a", 1], ["b", 2]]})
        assert [t for t, _ in calls] == ["write_range"]
        await _fn(write_range)({"address": "Data!D2", "values": [["=B2*C2"], ["=B3*C3"]]})
    assert calls[-1] == ("sheet_snapshot", {"max_cells": 20000, "sheet": "Data"})


@pytest.mark.asyncio
async def test_copy_range_checks_the_whole_destination():
    from excel_tools import copy_range
    call, calls = _host({"success": True, "data": _amount_snapshot(broken_row=4)})
    with patch("excel_tools.call_csharp", new=AsyncMock(side_effect=call)):
        out = await _fn(copy_range)({"source_address": "D2", "dest_address": "D3:D7"})
    v = json.loads(out["content"][0]["text"])["verification"]
    assert v["ok"] is False
    assert v["pattern_check"]["in_written_range"][0]["address"] == "D5"
