# src/DeepExcel.Sidecar/tests/test_context_lite.py
# 注入用户消息的轻量 Excel 上下文。
import sidecar


def test_user_edits_between_turns_are_mentioned():
    text = sidecar._build_excel_context_lite({"workbookName": "a.xlsx", "userEdits": ["Sheet1!B3", "Data!A1:A5"]})
    assert "[用户改动]" in text
    assert "Sheet1!B3、Data!A1:A5" in text


def test_no_edits_no_line():
    assert "[用户改动]" not in sidecar._build_excel_context_lite({"workbookName": "a.xlsx", "userEdits": []})
    assert "[用户改动]" not in sidecar._build_excel_context_lite({"workbookName": "a.xlsx", "userEdits": None})


def test_host_notices_are_shown():
    text = sidecar._build_excel_context_lite({"hostNotices": ["用户在面板上把工作簿回退到了 10:02:03 的检查点"]})
    assert "[提示] 用户在面板上把工作簿回退到了 10:02:03 的检查点" in text
