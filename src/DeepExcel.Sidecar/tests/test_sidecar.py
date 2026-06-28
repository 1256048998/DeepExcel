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
