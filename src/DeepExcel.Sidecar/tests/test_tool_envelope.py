"""Every tool result the model reads starts with the same envelope:
ok / status / checkpoint_id / next / warnings (the last three only when they
say something), followed by the host's original fields."""
import json

import excel_tools
import ui_events
from excel_tools import build_envelope, _wrap_result


def test_a_plain_success_is_done_with_only_ok_and_status_added():
    out = build_envelope({"success": True, "data": {"values": [[1]]}})
    assert list(out)[:2] == ["ok", "status"]
    assert out["ok"] is True and out["status"] == "done"
    assert "next" not in out and "warnings" not in out and "checkpoint_id" not in out
    assert out["success"] is True and out["data"] == {"values": [[1]]}


def test_envelope_keys_come_first_and_original_fields_survive():
    raw = {"success": True, "data": {"cells_written": 4}, "checkpoint_id": "cp-1",
           "backup_snapshot_id": "snap-1", "changes": {"changed": 4}}
    out = build_envelope(raw)
    assert list(out)[:3] == ["ok", "status", "checkpoint_id"]
    assert out["checkpoint_id"] == "cp-1"
    for key in ("success", "data", "backup_snapshot_id", "changes"):
        assert out[key] == raw[key]


def test_a_failure_points_at_the_hosts_suggestion():
    out = build_envelope({"success": False, "error": "工作表不存在", "suggestion": "先用 list(kind=sheets) 看有哪些表"})
    assert out["ok"] is False and out["status"] == "failed"
    assert out["next"] == "先用 list(kind=sheets) 看有哪些表"
    assert out["error"] == "工作表不存在"


def test_a_failure_without_suggestion_still_says_not_to_loop():
    out = build_envelope({"success": False, "error": "COM 异常"})
    assert out["status"] == "failed"
    assert "不要原样重试" in out["next"]


def test_an_interrupt_tells_the_model_to_stop():
    out = build_envelope({"success": False, "error": "用户已中断", "error_code": "interrupted"})
    assert "不要重试" in out["next"] and "中断" in out["next"]


def test_a_declined_step_must_not_be_worked_around():
    out = build_envelope({"success": False, "error": "用户取消了高风险操作"})
    assert "绕过" in out["next"]


def test_excel_or_the_executor_refusing_is_not_the_user_declining():
    for error in ("Excel 拒绝了这次操作。常见原因：地址或公式写错", "VBA 代码包含独立的 End 语句，已拒绝执行。"):
        out = build_envelope({"success": False, "error": error})
        assert "绕过" not in out["next"], error


def test_a_failed_post_write_check_is_its_own_status():
    out = build_envelope({"success": True, "checkpoint_id": "cp-9",
                          "verification": {"ok": False, "summary": "新增 3 个 #REF!"}})
    assert out["ok"] is True and out["status"] == "check_failed"
    assert "verification.summary" in out["next"] and "checkpoint_id" in out["next"]
    assert out["verification"]["summary"] == "新增 3 个 #REF!"


def test_a_passing_check_is_just_done():
    out = build_envelope({"success": True, "verification": {"ok": True, "summary": "回读一致"}})
    assert out["status"] == "done" and "next" not in out


def test_a_paged_read_is_partial_and_reuses_the_hosts_hint():
    hint = '共 900 行，本页是第 1–200 行，还有 700 行未读。下一页：read_range(address="A:D", offset=200)'
    out = build_envelope({"success": True, "data": {"paging": {"next_offset": 200}, "hint": hint}})
    assert out["status"] == "partial" and out["next"] == hint


def test_a_paged_read_without_hint_names_the_offset():
    out = build_envelope({"success": True, "data": {"paging": {"next_offset": 400}}})
    assert out["status"] == "partial" and "offset=400" in out["next"]


def test_the_last_page_is_done():
    out = build_envelope({"success": True, "data": {"paging": {"next_offset": None, "columns_truncated": False},
                                                     "hint": None}})
    assert out["status"] == "done" and "next" not in out


def test_a_truncated_search_is_partial():
    out = build_envelope({"success": True, "data": {"total": 300, "truncated": True}})
    assert out["status"] == "partial" and "缩小范围" in out["next"]


def test_warnings_are_collected_once_and_the_singular_field_is_dropped():
    out = build_envelope({"success": True, "warning": "参数已自动纠正", "data": {"warnings": ["参数已自动纠正", "有合并单元格"]}})
    assert out["warnings"] == ["参数已自动纠正", "有合并单元格"]
    assert "warning" not in out


def test_non_dict_results_pass_through():
    assert build_envelope("plain text") == "plain text"


def test_the_wrapped_text_is_the_envelope_and_failures_stay_errors():
    wrapped = _wrap_result({"success": False, "error": "x"})
    assert wrapped["is_error"] is True
    payload = json.loads(wrapped["content"][0]["text"])
    assert payload["ok"] is False and payload["status"] == "failed"


def test_the_panel_still_reads_the_envelope():
    """ui_events.parse_tool_result parses the same text the model reads."""
    ok_text = _wrap_result({"success": True, "checkpoint_id": "cp-2", "data": {"cells_written": 3}})["content"]
    parsed = ui_events.parse_tool_result(ok_text, False)
    assert parsed["ok"] is True and parsed["checkpoint_id"] == "cp-2" and parsed["summary"] == "3 个单元格"
    bad_text = _wrap_result({"success": False, "error": "坏了", "suggestion": "换个地址"})["content"]
    parsed = ui_events.parse_tool_result(bad_text, True)
    assert parsed["ok"] is False and parsed["error"]["message"] == "坏了" and parsed["error"]["hint"] == "换个地址"


def test_envelope_keys_are_documented_in_the_system_prompt():
    from system_prompt import SYSTEM_PROMPT
    for key in excel_tools.ENVELOPE_KEYS:
        assert key in SYSTEM_PROMPT, key
    for status in (excel_tools.STATUS_DONE, excel_tools.STATUS_PARTIAL,
                   excel_tools.STATUS_CHECK_FAILED, excel_tools.STATUS_FAILED):
        assert status in SYSTEM_PROMPT, status
