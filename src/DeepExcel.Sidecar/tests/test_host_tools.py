# src/DeepExcel.Sidecar/tests/test_host_tools.py
# 按宿主注册工具（F7）+ 回答缓存已删除（F8）。
import re
from pathlib import Path

import excel_tools
import sidecar
from excel_tools import (
    SIDECAR_CHANNEL_TOOLS,
    WPS_HOST_TOOLS,
    host_supports_tool,
    register_all_tools,
)

ROOT = Path(__file__).resolve().parents[3]


def _names(tools):
    return {t.name for t in tools}


def test_wps_host_tools_match_the_js_dispatcher():
    """WPS_HOST_TOOLS 必须与 tool-dispatcher.js 的 case 分支逐条一致。

    WPS 端新增或删掉一个工具时，这里不同步就会重新出现「模型能看到、WPS 执行不了」。
    """
    js = (ROOT / "src" / "DeepExcel.Wps" / "tool-dispatcher.js").read_text(encoding="utf-8")
    cases = set(re.findall(r"case '(\w+)':", js))
    assert cases, "没解析到任何 case 分支，正则可能过时了"
    assert cases == set(WPS_HOST_TOOLS)


def test_wps_session_only_sees_tools_wps_can_run():
    wps = _names(register_all_tools("wps"))
    for name in wps:
        assert host_supports_tool("wps", name)
    # 图表、透视、快照、VBA 这些 WPS 没实现的工具不能出现
    for name in ("create_chart", "create_pivot_table", "create_snapshot", "rollback",
                 "execute_vba", "execute_python", "screenshot_excel", "send_keys"):
        assert name not in wps
    assert "execute_jsa" in wps
    assert SIDECAR_CHANNEL_TOOLS <= wps


def test_excel_session_does_not_see_wps_only_tools():
    excel = _names(register_all_tools("excel"))
    assert "execute_jsa" not in excel
    assert "execute_vba" in excel


def test_every_wps_host_tool_is_actually_registered():
    """清单里写了、注册表里却没有的名字，说明 JS 端和 Python 端有一边改了名。"""
    wps = _names(register_all_tools("wps"))
    assert set(WPS_HOST_TOOLS) <= wps


def test_answer_cache_is_gone():
    """工作簿一变，缓存的回答就是错的（F8）。"""
    assert not hasattr(sidecar, "_ResponseCache")
    assert not hasattr(sidecar, "_response_cache")
    assert not hasattr(sidecar, "_build_cache_key")
    assert excel_tools  # 模块可导入


def test_wps_prompt_tells_the_model_which_tools_exist():
    names = sorted(t.name for t in register_all_tools("wps"))
    note = excel_tools.host_tool_note("wps", names)
    assert "execute_jsa" in note and "create_chart" not in note
    assert excel_tools.host_tool_note("excel", names) == ""


def test_todo_write_is_available_in_both_hosts():
    assert "todo_write" in {t.name for t in register_all_tools("excel")}
    assert "todo_write" in {t.name for t in register_all_tools("wps")}


def test_normalize_todos_keeps_one_active_step():
    items = excel_tools.normalize_todos([
        {"content": "读取数据", "status": "completed"},
        {"content": "写公式", "status": "in_progress"},
        {"content": "画图", "status": "in_progress"},
        {"content": "  ", "status": "pending"},
        "检查结果",
        {"content": "x", "status": "bogus"},
    ])
    assert [i["status"] for i in items] == ["completed", "in_progress", "pending", "pending", "pending"]
    assert items[3]["content"] == "检查结果"
    assert excel_tools.normalize_todos("not a list") == []
