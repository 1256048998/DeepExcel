"""perception 包：区块检测、多级表头、R1C1 公式模式、异常候选、分层 inspect。"""

from perception import inspect, normalize_layers
from perception.anomalies import find_anomalies
from perception.blocks import detect_blocks
from perception.formulas import pattern_id, scan_formulas, to_a1
from perception.grid import Rect, Snapshot
from perception.headers import analyze_header


def snap(rows, formulas=None, merges=None, origin=(1, 1), **extra):
    """rows 是二维值数组（窗口坐标）；formulas 是 {(r, c): R1C1}；merges 是 [(r1, c1, r2, c2)]"""
    width = max((len(r) for r in rows), default=0)
    cells = [list(r) + [None] * (width - len(r)) for r in rows]
    data = {
        "sheet": extra.pop("sheet", "Data"),
        "origin": list(origin),
        "cells": cells,
        "formulas": [[r, c, f] for (r, c), f in (formulas or {}).items()],
        "merges": [list(m) for m in (merges or [])],
    }
    data.update(extra)
    return data


def only_table(data):
    s = Snapshot(data)
    scan = detect_blocks(s)
    tables = [b for b in scan.blocks if b.kind == "table"]
    assert len(tables) == 1, scan.blocks
    return s, tables[0]


# ============ 区块 ============

def test_separate_tables_and_a_title_note_are_different_blocks():
    rows = [
        ["2024 年销售报表"],
        [],
        ["客户", "金额", None, None, "地区", "目标"],
        ["甲", 10, None, None, "华东", 100],
        ["乙", 20, None, None, "华北", 200],
    ]
    s = Snapshot(snap(rows))
    scan = detect_blocks(s)
    ranges = [(s.rect_address(b.rect), b.kind) for b in scan.blocks]
    assert ranges == [("A1", "note"), ("A3:B5", "table"), ("E3:F5", "table")]
    assert scan.complete


def test_a_single_blank_row_inside_a_table_does_not_split_it():
    rows = [["客户", "金额", "日期"], ["甲", 1, 2], [], ["乙", 3, 4], ["丙", 5, 6]]
    s = Snapshot(snap(rows))
    scan = detect_blocks(s)
    assert [s.rect_address(b.rect) for b in scan.blocks] == ["A1:C5"]


def test_two_blank_rows_do_split_and_narrow_overlap_does_not_merge():
    rows = [["a", "b", "c"], [1, 2, 3], [], [], ["x", "y", "z"], [4, 5, 6]]
    s = Snapshot(snap(rows))
    assert len(detect_blocks(s).blocks) == 2

    rows = [["a", "b", "c", "d", "e"], [1, 2, 3, 4, 5], [], [None, None, None, "p", "q"], [None, None, None, 7, 8]]
    s = Snapshot(snap(rows))
    # 下面那块只有 D:E 两列，与上面 5 列重合 2/2=100%：会合并
    assert len(detect_blocks(s).blocks) == 1
    rows = [["a", "b", "c", "d", "e"], [1, 2, 3, 4, 5], [], [None, None, None, None, "q", "r", "s"], [None] * 4 + [7, 8, 9]]
    s = Snapshot(snap(rows))
    # 下面那块 E:G，与上面 A:E 只重合 1/3：不合并
    assert len(detect_blocks(s).blocks) == 2


def test_a_merged_title_connects_to_nothing_but_counts_as_filled():
    rows = [["工资表", None, None], ["姓名", "基本", "奖金"], ["张三", 1, 2]]
    s = Snapshot(snap(rows, merges=[(0, 0, 0, 2)]))
    scan = detect_blocks(s)
    assert [s.rect_address(b.rect) for b in scan.blocks] == ["A1:C3"]


def test_contained_islands_are_absorbed():
    # 外框是一圈数据，中间有一个被空白隔开的格子
    rows = [[1, 1, 1, 1, 1], [1, None, None, None, 1], [1, None, 9, None, 1], [1, None, None, None, 1], [1, 1, 1, 1, 1]]
    s = Snapshot(snap(rows))
    scan = detect_blocks(s)
    assert len(scan.blocks) == 1
    assert scan.blocks[0].cells == 17


def test_too_many_blocks_keeps_the_biggest_and_lists_the_rest():
    rows = []
    for i in range(20):
        size = 2 if i else 5
        rows.append(["x"] * size)
        rows.append([])
        rows.append([])
    s = Snapshot(snap(rows))
    scan = detect_blocks(s)
    assert len(scan.blocks) == 12
    assert scan.others_total == 8
    assert scan.blocks[0].rect.cols == 5


def test_truncated_snapshot_marks_the_scan_partial():
    s = Snapshot(snap([["a", "b"], [1, 2]], truncated=True, total_rows=50000))
    assert detect_blocks(s).complete is False


# ============ 表头 ============

def test_single_header_row_with_types():
    rows = [["日期", "客户", "金额"], [{"d": "2024-01-01"}, "甲", 10], [{"d": "2024-01-02"}, "乙", 20.5]]
    s, block = only_table(snap(rows))
    h = analyze_header(s, block)
    assert h.header_rows == [0]
    assert h.data_start == 1
    assert [(c.label, c.type) for c in h.columns] == [("日期", "date"), ("客户", "text"), ("金额", "number")]
    assert not h.uncertain


def test_title_then_two_level_merged_header():
    #   A        B     C      D      E      F
    # 1 2024 年 9 月工资表（A1:F1 合并）
    # 2 姓名     部门  扣款（C2:E2 合并）   实发
    # 3 （合并）      养老   医疗   失业   （合并）
    rows = [
        ["2024 年 9 月工资表"],
        ["姓名", "部门", "扣款", None, None, "实发"],
        [None, None, "养老", "医疗", "失业", None],
        ["张三", "销售", 100, 50, 10, 5000],
        ["李四", "财务", 120, 60, 12, 6000],
    ]
    merges = [(0, 0, 0, 5), (1, 0, 2, 0), (1, 1, 2, 1), (1, 2, 1, 4), (1, 5, 2, 5)]
    s, block = only_table(snap(rows, merges=merges))
    h = analyze_header(s, block)
    assert h.title_rows == [0]
    assert h.title == "2024 年 9 月工资表"
    assert h.header_rows == [1, 2]
    assert h.data_start == 3
    assert [c.label for c in h.columns] == ["姓名", "部门", "扣款/养老", "扣款/医疗", "扣款/失业", "实发"]
    assert not h.uncertain


def test_three_level_header_builds_full_paths():
    rows = [
        ["姓名", "扣款", None, None, None],
        [None, "养老", None, "医疗", None],
        [None, "公司", "个人", "公司", "个人"],
        ["张三", 1, 2, 3, 4],
    ]
    merges = [(0, 0, 2, 0), (0, 1, 0, 4), (1, 1, 1, 2), (1, 3, 1, 4)]
    s, block = only_table(snap(rows, merges=merges))
    h = analyze_header(s, block)
    assert h.header_rows == [0, 1, 2]
    assert h.columns[1].label == "扣款/养老/公司"
    assert h.columns[4].label == "扣款/医疗/个人"


def test_center_across_selection_without_merges_inherits_the_parent_label():
    rows = [
        ["姓名", "扣款", None, None, "实发"],
        [None, "养老", "医疗", "失业", None],
        ["张三", 1, 2, 3, 4],
    ]
    s, block = only_table(snap(rows))
    h = analyze_header(s, block)
    assert h.header_rows == [0, 1]
    assert [c.label for c in h.columns] == ["姓名", "扣款/养老", "扣款/医疗", "扣款/失业", "实发"]


def test_a_text_first_data_row_is_not_mistaken_for_a_second_header():
    rows = [["姓名", "部门", "备注"], ["张三", "销售", "新人"], ["李四", "财务", "老员工"]]
    s, block = only_table(snap(rows))
    h = analyze_header(s, block)
    assert h.header_rows == [0]


def test_numbers_only_block_says_the_header_is_uncertain():
    rows = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    s, block = only_table(snap(rows))
    h = analyze_header(s, block)
    assert h.header_rows == []
    assert h.uncertain
    assert "没有表头" in h.reasons[0] or "表头" in h.reasons[0]


def test_total_rows_are_found_and_kept_out_of_column_types():
    rows = [["项目", "金额"], ["a", 1], ["b", 2], ["合计", 3]]
    s, block = only_table(snap(rows, formulas={(3, 1): "=SUM(R[-2]C:R[-1]C)"}))
    h = analyze_header(s, block)
    assert h.total_rows == [3]
    assert h.columns[1].type == "number"
    assert not h.columns[1].formula


# ============ R1C1 公式模式 ============

def _amount_table(n=5, overrides=None, constants=None):
    rows = [["品名", "数量", "单价", "金额"]]
    formulas = {}
    for i in range(1, n + 1):
        rows.append([f"p{i}", i, 10, i * 10])
        formulas[(i, 3)] = "=RC[-2]*RC[-1]"
    for (r, c), f in (overrides or {}).items():
        formulas[(r, c)] = f
    for (r, c) in constants or ():
        formulas.pop((r, c), None)
    return rows, formulas


def test_same_r1c1_down_a_column_is_one_pattern_rendered_with_header_names():
    rows, formulas = _amount_table()
    s, block = only_table(snap(rows, formulas=formulas))
    scan = scan_formulas(s, block, analyze_header(s, block))
    assert len(scan.patterns) == 1
    p = scan.patterns[0]
    assert p.rendered == "=[数量]*[单价]"
    assert p.id == pattern_id("=RC[-2]*RC[-1]")
    assert [s.rect_address(r) for r in p.rects] == ["D2:D6"]
    assert p.cells == 5
    assert scan.inconsistent == []
    assert scan.complete


def test_adjacent_columns_with_the_same_formula_become_one_rectangle():
    rows = [["a", "b", "c"], [1, 2, 3], [4, 5, 6], ["合计", 9, 9]]
    formulas = {(3, 1): "=SUM(R[-2]C:R[-1]C)", (3, 2): "=SUM(R[-2]C:R[-1]C)"}
    s, block = only_table(snap(rows, formulas=formulas))
    scan = scan_formulas(s, block, analyze_header(s, block))
    assert [s.rect_address(r) for r in scan.patterns[0].rects] == ["B4:C4"]
    assert scan.patterns[0].rendered == "=SUM(B2:B3)"


def test_mixed_patterns_in_one_column_are_flagged_inconsistent():
    rows, formulas = _amount_table(6, overrides={(4, 3): "=RC[-2]*RC[-1]*1.1"})
    s, block = only_table(snap(rows, formulas=formulas))
    scan = scan_formulas(s, block, analyze_header(s, block))
    assert len(scan.inconsistent) == 1
    assert scan.inconsistent[0].col == 3
    assert scan.inconsistent[0].patterns[0][1] == 5


def test_to_a1_handles_absolute_cross_sheet_strings_and_names():
    assert to_a1("=R1C1+RC[1]", 5, 3) == "=$A$1+D5"
    assert to_a1("=Sheet2!RC+'Q1 汇总'!R2C[-1]", 5, 3) == "=Sheet2!C5+'Q1 汇总'!B$2"
    assert to_a1('=IF(RC[-1]="RC",ROUND(RC[-1],2),0)', 2, 2) == '=IF(A2="RC",ROUND(A2,2),0)'
    assert to_a1("=R[-5]C", 3, 1) == "=#REF!"


def test_partial_snapshot_gives_no_counts():
    rows, formulas = _amount_table()
    out = inspect(snap(rows, formulas=formulas, truncated=True, total_rows=9999), ["formulas"])
    layer = out["layers"]["formulas"]
    assert layer["status"] == "partial"
    assert "统计不完整" in layer["note"]
    assert "cells" not in layer["patterns"][0]
    assert "formula_cells" not in layer
    assert "total" not in layer["anomalies"]


# ============ 异常候选 ============

def _anomalies(data):
    s, block = only_table(data)
    header = analyze_header(s, block)
    return find_anomalies(s, block, header, scan_formulas(s, block, header))


def test_isolated_deviation_is_a_candidate():
    rows, formulas = _amount_table(6, overrides={(3, 3): "=RC[-2]*11"})
    found = _anomalies(snap(rows, formulas=formulas))
    assert [c["address"] for c in found["candidates"] if c["type"] == "isolated"] == ["D4"]
    assert "=[数量]" not in found["candidates"][0]["detail"]  # 细节里是 A1，不是表头名
    assert "D4" in found["candidates"][0]["address"]


def test_a_hardcoded_number_between_formulas_is_a_candidate():
    rows, formulas = _amount_table(6, constants=[(3, 3)])
    found = _anomalies(snap(rows, formulas=formulas))
    hits = [c for c in found["candidates"] if c["type"] == "hardcoded"]
    assert [c["address"] for c in hits] == ["D4"]
    assert "数值 30" in hits[0]["detail"]


def test_total_range_that_misses_the_last_row_is_a_candidate():
    rows = [["项目", "金额"], ["a", 1], ["b", 2], ["c", 3], ["合计", 3]]
    formulas = {(4, 1): "=SUM(R[-3]C:R[-2]C)"}  # 合计在第 5 行，只汇总了第 2–3 行，漏了第 4 行（c）
    found = _anomalies(snap(rows, formulas=formulas))
    hits = [c for c in found["candidates"] if c["type"] == "total_gap"]
    assert len(hits) == 1
    assert hits[0]["address"] == "B5"
    assert "漏了第 4 行" in hits[0]["detail"]
    assert hits[0]["formula"] == "=SUM(B2:B3)"


def test_complete_total_range_is_not_a_candidate():
    rows = [["项目", "金额"], ["a", 1], ["b", 2], ["合计", 3]]
    found = _anomalies(snap(rows, formulas={(3, 1): "=SUM(R[-2]C:R[-1]C)"}))
    assert found["total"] == 0


def test_broken_references_are_candidates():
    rows = [["a", "b"], [1, {"e": "#REF!"}], [2, {"e": "#REF!"}]]
    formulas = {(1, 1): "=#REF!*2", (2, 1): "=[old.xlsx]Sheet1!R1C1"}
    found = _anomalies(snap(rows, formulas=formulas))
    assert sorted(c["address"] for c in found["candidates"] if c["type"] == "broken_ref") == ["B2", "B3"]


def test_merged_and_spilled_cells_are_excluded():
    rows, formulas = _amount_table(6, constants=[(3, 3)])
    found = _anomalies(snap(rows, formulas=formulas, spills=[[3, 3, 3, 3]]))
    assert not [c for c in found["candidates"] if c["type"] == "hardcoded"]
    rows, formulas = _amount_table(6, constants=[(3, 3)])
    found = _anomalies(snap(rows, formulas=formulas, merges=[(3, 3, 3, 3)]))
    assert not [c for c in found["candidates"] if c["type"] == "hardcoded"]


# ============ 分层 inspect ============

def test_layers_default_and_all():
    assert normalize_layers(None) == ["blocks", "formulas"]
    assert normalize_layers("all") == ["blocks", "formulas", "objects", "dependencies"]
    assert normalize_layers("objects, blocks") == ["blocks", "objects"]
    assert normalize_layers(["bogus"]) == ["blocks", "formulas"]


def test_inspect_returns_every_requested_layer_with_status():
    rows, formulas = _amount_table()
    formulas[(1, 0)] = "='Q1 汇总'!R1C1&[book2.xlsx]Rates!R1C1"
    data = snap(rows, formulas=formulas, used="A1:D6", objects={"tables": [{"name": "T1"}]})
    out = inspect(data, "all")
    layers = out["layers"]
    assert out["coverage"]["complete"] is True
    assert {k: v["status"] for k, v in layers.items()} == {
        "blocks": "complete", "formulas": "complete", "objects": "complete", "dependencies": "complete"}
    block = layers["blocks"]["items"][0]
    assert block["range"] == "A1:D6"
    assert block["header_rows"] == [1]
    assert block["data_rows"] == "2-6"
    assert block["columns_detail"][3] == {"column": "D", "label": "金额", "type": "number", "formula": True}
    assert layers["formulas"]["formula_cells"] == 6
    assert layers["objects"]["tables"] == [{"name": "T1"}]
    assert layers["dependencies"]["sheets_referenced"] == ["Q1 汇总"]
    assert layers["dependencies"]["external_workbooks"] == ["book2.xlsx"]


def test_coverage_is_honest_about_truncation():
    out = inspect(snap([["a", "b"], [1, 2]], truncated=True, total_rows=90000, total_columns=2, used="A1:B90000"))
    assert out["coverage"]["complete"] is False
    assert "只读了前 2 行" in out["coverage"]["note"]
    assert out["layers"]["blocks"]["status"] == "partial"


def test_one_broken_layer_does_not_take_down_the_others(monkeypatch):
    import perception.inspect_sheet as mod

    def boom(*_):
        raise ValueError("坏了")

    monkeypatch.setattr(mod, "_formulas_layer", boom)
    rows, formulas = _amount_table()
    out = inspect(snap(rows, formulas=formulas), "all")
    assert out["layers"]["formulas"]["status"] == "error"
    assert "坏了" in out["layers"]["formulas"]["error"]
    assert out["layers"]["blocks"]["status"] == "complete"
    assert out["layers"]["dependencies"]["status"] == "complete"


def test_empty_sheet():
    out = inspect(snap([]))
    assert out["layers"]["blocks"]["items"] == []
    assert out["layers"]["blocks"]["note"] == "这张表是空的"


def test_origin_offsets_every_address():
    rows, formulas = _amount_table()
    out = inspect(snap(rows, formulas=formulas, origin=(10, 3)), ["blocks"])
    assert out["layers"]["blocks"]["items"][0]["range"] == "C10:F15"


def test_rect_helpers():
    a = Rect(0, 0, 2, 2)
    assert a.contains(Rect(1, 1, 1, 1))
    assert a.union(Rect(3, 3, 3, 3)) == Rect(0, 0, 3, 3)


# ============ 真 Excel 导出的快照（C# SheetSnapshotExcelTests 生成） ============

def test_real_excel_payroll_snapshot():
    import json
    from pathlib import Path

    data = json.loads((Path(__file__).parent / "fixtures" / "snapshot_payroll.json").read_text(encoding="utf-8"))
    out = inspect(data, "all")
    block = out["layers"]["blocks"]["items"][0]
    assert block["title"] == "2024年9月工资表"
    assert block["header_rows"] == [2, 3]
    assert block["total_rows"] == [10]
    assert [c["label"] for c in block["columns_detail"]] == ["姓名", "入职日期", "扣款/养老", "扣款/医疗", "扣款/小计", "实发"]
    assert block["columns_detail"][1]["type"] == "date"

    formulas = out["layers"]["formulas"]
    assert formulas["patterns"][0]["formula"] == "=[养老]+[医疗]"
    found = {(c["type"], c["address"]) for c in formulas["anomalies"]["candidates"]}
    assert found == {("total_gap", "C10"), ("hardcoded", "E7"), ("isolated", "F8")}
