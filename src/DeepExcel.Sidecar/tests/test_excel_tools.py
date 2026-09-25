# src/DeepExcel.Sidecar/tests/test_excel_tools.py
import asyncio
import json
import pytest
from unittest.mock import patch, AsyncMock


@pytest.mark.asyncio
async def test_write_formula_tool_calls_csharp_and_returns_content():
    """write_formula 工具应通过 call_csharp 调用 C#，并返回 MCP content 格式"""
    from excel_tools import write_formula

    # write_formula 已被 @tool 装饰，需要从 .handler 取出原始协程函数
    # claude-agent-sdk 0.2.109 的 SdkMcpTool 是 dataclass，handler 字段持有原始 async 函数
    fn = getattr(write_formula, 'handler', None) or getattr(write_formula, 'fn', None) or write_formula
    if hasattr(fn, '__wrapped__'):
        fn = fn.__wrapped__

    fake_result = {"success": True, "data": {"address": "A1", "formula": "=SUM(B:B)"}}
    with patch('excel_tools.call_csharp', new=AsyncMock(return_value=fake_result)):
        result = await fn({"address": "A1", "formula": "=SUM(B:B)"})

    assert result["content"][0]["type"] == "text"
    parsed = json.loads(result["content"][0]["text"])
    assert parsed["success"] is True
    assert parsed["data"]["formula"] == "=SUM(B:B)"


@pytest.mark.asyncio
async def test_clarify_intent_calls_csharp_clarify():
    from excel_tools import clarify_intent
    fn = getattr(clarify_intent, 'handler', None) or getattr(clarify_intent, 'fn', None) or clarify_intent
    if hasattr(fn, '__wrapped__'):
        fn = fn.__wrapped__
    with patch('excel_tools.call_csharp_clarify', new=AsyncMock(return_value="COUNTA计数")):
        result = await fn({"question": "求和还是计数？", "options": ["SUM", "COUNTA"]})
    assert "COUNTA计数" in result["content"][0]["text"]


def test_register_all_tools_returns_list_with_expected_names():
    from excel_tools import register_all_tools
    tools = register_all_tools()
    # @tool 装饰后的对象可能有 .name 属性或需要查 __name__
    names = []
    for t in tools:
        n = getattr(t, 'name', None) or getattr(t, '__name__', None)
        if n is None:
            # SDK 可能用 .fn.__name__ 或 .handler.__name__
            fn = getattr(t, 'handler', None) or getattr(t, 'fn', None)
            n = getattr(fn, '__name__', None) if fn else None
        names.append(n)
    assert "read_range" in names
    assert "write_formula" in names
    assert "clarify_intent" in names


@pytest.mark.asyncio
async def test_all_14_tools_registered():
    from excel_tools import register_all_tools
    tools = register_all_tools()
    assert len(tools) >= 14


@pytest.mark.asyncio
async def test_create_chart_passes_all_args():
    from excel_tools import create_chart
    fn = getattr(create_chart, 'handler', None) or getattr(create_chart, 'fn', None) or create_chart
    if hasattr(fn, '__wrapped__'): fn = fn.__wrapped__

    fake_result = {"success": True, "data": {}}
    with patch('excel_tools.call_csharp', new=AsyncMock(return_value=fake_result)) as mock:
        await fn({"data_range": "A1:B10", "chart_type": "bar", "title": "T", "x_label": "X", "y_label": "Y"})
        call_args = mock.call_args[0]
        assert call_args[0] == "create_chart"
        assert call_args[1]["chart_type"] == "bar"


@pytest.mark.asyncio
async def test_clean_data_passes_operations_list():
    from excel_tools import clean_data
    fn = getattr(clean_data, 'handler', None) or getattr(clean_data, 'fn', None) or clean_data
    if hasattr(fn, '__wrapped__'): fn = fn.__wrapped__

    with patch('excel_tools.call_csharp', new=AsyncMock(return_value={"success": True})) as mock:
        await fn({"range_address": "A1:A100", "operations": ["trim_spaces", "remove_duplicates"]})
        call_args = mock.call_args[0]
        assert call_args[1]["operations"] == ["trim_spaces", "remove_duplicates"]


def _registered_tool_names(host="excel"):
    from excel_tools import register_all_tools
    names = []
    for t in register_all_tools(host):
        n = getattr(t, 'name', None) or getattr(t, '__name__', None)
        if n is None:
            fn = getattr(t, 'handler', None) or getattr(t, 'fn', None)
            n = getattr(fn, '__name__', None) if fn else None
        names.append(n)
    return names


# 只在 WPS 宿主实现的工具（Excel 端没有对应分支是预期的）
_WPS_ONLY_TOOLS = {"execute_jsa"}


def test_every_host_call_has_a_csharp_handler():
    """sidecar 转发给宿主的每个工具，ToolDispatcher 里都必须有对应分支。

    auto_analyze 曾经只在这里注册、C# 侧没有实现：模型每次调用都拿到"未知工具"，
    白白浪费一轮，却没有任何测试会因此变红。
    """
    import re
    from pathlib import Path
    root = Path(__file__).resolve().parents[3]
    py = (root / "src" / "DeepExcel.Sidecar" / "excel_tools.py").read_text(encoding="utf-8")
    cs = (root / "src" / "DeepExcel.AddIn" / "Sidecar" / "ToolDispatcher.cs").read_text(encoding="utf-8")
    called = set(re.findall(r'call_csharp\(\s*"(\w+)"', py))
    handled = set(re.findall(r'case "(\w+)":', cs))
    assert called, "没解析到任何 call_csharp 调用，正则可能过时了"
    missing = sorted(called - handled - _WPS_ONLY_TOOLS)
    assert missing == [], f"这些工具在 C# 侧没有实现，调用必然失败：{missing}"


@pytest.mark.parametrize("name", ["auto_analyze", "echo", "quick_summary", "create_plan", "update_plan"])
def test_placeholder_tools_are_not_offered_to_the_model(name):
    """这些工具要么宿主没实现，要么只是在 Python 里回显参数，没有任何作用，
    却每轮占用模型的上下文，还会诱导模型去调用。"""
    import sidecar
    assert name not in _registered_tool_names("excel")
    assert name not in _registered_tool_names("wps")
    assert f'"{name}"' not in open(sidecar.__file__, encoding="utf-8").read()


def test_wps_only_tools_are_registered_only_for_wps():
    assert "execute_jsa" not in _registered_tool_names("excel")
    assert "execute_jsa" in _registered_tool_names("wps")
    assert set(_registered_tool_names("excel")) <= set(_registered_tool_names("wps"))


def test_system_prompt_tool_list_matches_what_is_registered():
    """<available-tools> 必须与实际注册的工具一致。

    曾经清单里写着根本不存在的 remove_duplicates（它是 clean_data 的一个操作），
    又漏掉了 20 个已注册的图表、透视、清洗工具，模型不知道有这些能力。
    """
    import re
    from system_prompt import SYSTEM_PROMPT
    block = SYSTEM_PROMPT.split("<available-tools>")[1].split("</available-tools>")[0]
    # 只看每行冒号后面、括号外的工具名；括号里是说明文字
    listed = set()
    for line in block.strip().splitlines():
        body = re.sub(r"（[^）]*）", "", line.split("：", 1)[-1])
        listed |= set(re.findall(r"\b[a-z]+(?:_[a-z]+)*\b", body))
    registered = set(_registered_tool_names("excel")) | set(_registered_tool_names("wps"))
    assert sorted(listed - registered) == [], "清单里有不存在的工具"
    assert sorted(registered - listed) == [], "已注册的工具没写进清单"
