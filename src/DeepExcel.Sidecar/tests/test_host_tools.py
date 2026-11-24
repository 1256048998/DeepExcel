# src/DeepExcel.Sidecar/tests/test_host_tools.py
# 工具注册（F7）+ 回答缓存已删除（F8）。
# WPS 加载的是和 Excel 同一个 COM 加载项，侧车只有一种宿主、注册完整工具集。
import excel_tools
import sidecar
from excel_tools import register_all_tools


def _names(tools):
    return {t.name for t in tools}


def test_inspect_sheet_is_offered_but_its_host_primitive_is_not():
    """inspect_sheet 在侧车里算，靠宿主的 sheet_snapshot 原语；原语本身不给模型"""
    names = _names(register_all_tools())
    assert "inspect_sheet" in names
    assert "sheet_snapshot" not in names


def test_host_mode_split_is_gone():
    """WPS 宿主模式（DEEPEXCEL_HOST=wps、工具子集、execute_jsa）随 JS 加载项一起删除"""
    names = _names(register_all_tools())
    assert "execute_jsa" not in names and "execute_vba" in names
    for gone in ("WPS_HOST_TOOLS", "WPS_HOST_PRIMITIVES", "host_supports_tool", "host_tool_note"):
        assert not hasattr(excel_tools, gone), gone
    assert not hasattr(sidecar, "_load_wps_local_config")
    assert "DEEPEXCEL_HOST" not in open(sidecar.__file__, encoding="utf-8").read()


def test_answer_cache_is_gone():
    """工作簿一变，缓存的回答就是错的（F8）。"""
    assert not hasattr(sidecar, "_ResponseCache")
    assert not hasattr(sidecar, "_response_cache")
    assert not hasattr(sidecar, "_build_cache_key")
    assert excel_tools  # 模块可导入


def test_todo_write_is_registered():
    assert "todo_write" in _names(register_all_tools())


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
