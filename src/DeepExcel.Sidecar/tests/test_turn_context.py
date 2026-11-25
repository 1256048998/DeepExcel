"""Per-turn message assembly: context first, request last; the structure summary
and host line are sent once and again only when they change or after compaction."""
import pytest

import turn_context
from turn_context import Turn, compose, host_line, STRUCTURE_UNCHANGED


@pytest.fixture(autouse=True)
def _fresh():
    turn_context.reset()
    yield
    turn_context.reset()


STRUCTURE = "## 工作簿结构\n### 销售 (1..900 行)\n  A 日期 date"
HOST_365 = {"name": "Microsoft Excel", "version": "16.0", "build": "17531", "dynamicArrays": True}


def test_the_request_comes_last_in_its_own_tag():
    text = compose("把 B 列求和", ["[工作簿] a.xlsx", STRUCTURE])
    assert text.index("<context>") < text.index("[工作簿] a.xlsx") < text.index("</context>")
    assert text.rstrip().endswith("<request>\n把 B 列求和\n</request>")


def test_the_mode_note_sits_right_before_the_request():
    note = "<permission-mode>只出方案</permission-mode>"
    text = compose("改表头", ["[工作簿] a.xlsx"], instruction=note)
    assert text.index("</context>") < text.index(note) < text.index("<request>")


def test_no_context_sends_just_the_request():
    assert compose("你好", ["", None], trailer="（请用中文思考。）") == "你好（请用中文思考。）"


def test_the_trailer_stays_at_the_very_end():
    text = compose("求和", ["[工作簿] a.xlsx"], trailer="\n\n（请用中文思考。）")
    assert text.endswith("</request>\n\n（请用中文思考。）")


def test_request_first_is_one_switch_away(monkeypatch):
    monkeypatch.setattr(turn_context, "REQUEST_LAST", False)
    text = compose("求和", ["[工作簿] a.xlsx"])
    assert text.index("<request>") < text.index("<context>")


def test_structure_is_sent_once_then_referenced():
    first = Turn(STRUCTURE, None)
    assert first.structure_text == STRUCTURE and first.structure_sent
    first.commit()
    second = Turn(STRUCTURE, None)
    assert second.structure_text == STRUCTURE_UNCHANGED and not second.structure_sent


def test_a_changed_structure_is_sent_again():
    Turn(STRUCTURE, None).commit()
    changed = STRUCTURE + "\n  B 金额 number"
    assert Turn(changed, None).structure_text == changed


def test_an_unsent_turn_does_not_count():
    """client.query failed: the model never saw it, so the next turn sends it again."""
    Turn(STRUCTURE, None)            # no commit
    assert Turn(STRUCTURE, None).structure_text == STRUCTURE


def test_compaction_sends_everything_again():
    t = Turn(STRUCTURE, HOST_365)
    t.commit()
    turn_context.reset()
    again = Turn(STRUCTURE, HOST_365)
    assert again.structure_text == STRUCTURE and again.host_text


def test_the_host_line_is_sent_once():
    t = Turn(None, HOST_365)
    assert "Microsoft Excel 16.0" in t.host_text and "支持动态数组" in t.host_text
    t.commit()
    assert Turn(None, HOST_365).host_text == ""


def test_no_structure_means_no_structure_line():
    t = Turn(None, None)
    assert t.structure_text == "" and t.host_text == "" and not t.structure_sent


def test_old_excel_is_told_to_avoid_dynamic_array_functions():
    line = host_line({"name": "Microsoft Excel", "version": "16.0", "build": "10827", "dynamicArrays": False})
    assert "不支持动态数组" in line and "FILTER" in line and "INDEX" in line


def test_unknown_dynamic_array_support_is_said_plainly():
    line = host_line({"name": "WPS 表格", "version": "12.0", "dynamicArrays": None})
    assert line.startswith("[宿主] WPS 表格 12.0") and "未知" in line


def test_an_empty_host_says_nothing():
    assert host_line(None) == "" and host_line({}) == ""
