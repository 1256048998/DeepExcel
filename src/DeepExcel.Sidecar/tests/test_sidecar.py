# src/DeepExcel.Sidecar/tests/test_sidecar.py
import pytest


@pytest.mark.asyncio
async def test_handle_sdk_message_emits_stream_delta_for_text_block(monkeypatch):
    """AssistantMessage 内的 TextBlock 应触发 stream_delta IPC 消息"""
    from claude_agent_sdk.types import AssistantMessage, TextBlock
    import sidecar as sidecar_module

    sent = []

    async def fake_write(msg):
        sent.append(msg)

    monkeypatch.setattr(sidecar_module, "write_message", fake_write)

    fake_msg = AssistantMessage(
        content=[TextBlock(text="正在读取A列...")],
        model="claude-sonnet-4",
    )
    await sidecar_module.handle_sdk_message(fake_msg)

    assert any(
        m["type"] == "stream_delta" and m["text"] == "正在读取A列..."
        for m in sent
    )


@pytest.mark.asyncio
async def test_handle_sdk_message_emits_stream_end_for_result_message(monkeypatch):
    """ResultMessage 应触发 stream_end IPC 消息，并携带 input/output tokens"""
    from claude_agent_sdk.types import ResultMessage
    import sidecar as sidecar_module

    sent = []

    async def fake_write(msg):
        sent.append(msg)

    monkeypatch.setattr(sidecar_module, "write_message", fake_write)

    fake_msg = ResultMessage(
        subtype="success",
        duration_ms=100,
        duration_api_ms=50,
        is_error=False,
        num_turns=1,
        session_id="test",
        usage={"input_tokens": 100, "output_tokens": 50},
    )
    await sidecar_module.handle_sdk_message(fake_msg)

    assert any(
        m["type"] == "stream_end"
        and m["input_tokens"] == 100
        and m["output_tokens"] == 50
        for m in sent
    )


# ---------------------------------------------------------------------------
# PreToolUse 钩子：低风险工具必须显式 allow
#
# 返回 continue_ 等于钩子不表态，CLI 会按 allowed_tools 决定；不在名单里的工具
# 在无交互模式下被直接拒掉。2026-09-24 用发布包同款 SDK 0.2.109 + claude.exe 实测：
# continue_ → "Claude requested permissions to use ..., but you haven't granted it yet."
# ---------------------------------------------------------------------------

def _decision(result):
    return (result or {}).get("hookSpecificOutput", {}).get("permissionDecision")


@pytest.mark.asyncio
@pytest.mark.parametrize("tool", ["set_chart_title", "refresh_pivot", "highlight_duplicates", "read_range"])
async def test_low_risk_tools_are_explicitly_allowed(tool):
    import sidecar
    result = await sidecar._pre_tool_use_hook({"tool_name": f"mcp__excel__{tool}", "tool_input": {}}, "id", None)
    assert _decision(result) == "allow"


@pytest.mark.asyncio
async def test_non_excel_tools_are_still_denied():
    import sidecar
    result = await sidecar._pre_tool_use_hook({"tool_name": "Bash", "tool_input": {}}, "id", None)
    assert _decision(result) == "deny"


@pytest.mark.asyncio
async def test_a_broken_confirmation_flow_denies_high_risk_tools(monkeypatch):
    """确认流程出错时，高风险工具不能被当作用户已同意（以前一律放行）"""
    import sidecar

    async def boom(name, tool_input):
        raise RuntimeError("ipc broken")

    monkeypatch.setattr(sidecar, "request_permission", boom)
    sidecar._allowed_tools_session.clear()
    result = await sidecar._pre_tool_use_hook({"tool_name": "mcp__excel__delete_rows", "tool_input": {}}, "id", None)
    assert _decision(result) == "deny"


@pytest.mark.asyncio
async def test_jsa_asks_like_vba(monkeypatch):
    """execute_jsa 在 WPS 里跑任意代码，与 VBA 同级：要确认，且可"允许并记住\""""
    import sidecar
    asked = []

    async def fake_request(name, tool_input):
        asked.append(name)
        return "allow"

    monkeypatch.setattr(sidecar, "request_permission", fake_request)
    sidecar._allowed_tools_session.clear()
    first = await sidecar._pre_tool_use_hook({"tool_name": "mcp__excel__execute_jsa", "tool_input": {"code": "1"}}, "id", None)
    second = await sidecar._pre_tool_use_hook({"tool_name": "mcp__excel__execute_jsa", "tool_input": {"code": "2"}}, "id", None)
    assert _decision(first) == "allow" and _decision(second) == "allow"
    assert asked == ["execute_jsa"]
    sidecar._allowed_tools_session.clear()
