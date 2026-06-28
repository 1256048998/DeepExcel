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
async def test_echo_tool_returns_input_as_text():
    from excel_tools import echo
    fn = getattr(echo, 'handler', None) or getattr(echo, 'fn', None) or echo
    if hasattr(fn, '__wrapped__'):
        fn = fn.__wrapped__
    result = await fn({"text": "hello"})
    # echo 实现逐字使用 _wrap_result（与 write_formula 一致），返回 JSON 字符串
    parsed = json.loads(result["content"][0]["text"])
    assert parsed["success"] is True
    assert parsed["data"]["echo"] == "hello"


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
    assert "echo" in names
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
