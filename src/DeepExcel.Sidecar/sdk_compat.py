"""claude-agent-sdk 的兼容补丁。

思考块的 signature：SDK 0.2.109 的 message_parser 直接取 block["signature"]，
而 Anthropic 兼容端点（DeepSeek 等）返回的思考块可以不带这个字段——缺了就抛
MessageParseError，整轮回复解析失败。signature 只用于 Anthropic 官方端点回传
校验，面板用不到，缺省补空串即可。
"""

from __future__ import annotations

from typing import Any


def fill_missing_thinking_signature(data: Any) -> Any:
    """给 assistant 消息里缺 signature 的思考块补空串（原地修改并返回 data）。"""
    if not isinstance(data, dict) or data.get("type") != "assistant":
        return data
    message = data.get("message")
    content = message.get("content") if isinstance(message, dict) else None
    if isinstance(content, list):
        for block in content:
            if isinstance(block, dict) and block.get("type") == "thinking" and "signature" not in block:
                block["signature"] = ""
    return data


def install() -> None:
    """包一层 parse_message。ClaudeSDKClient 每次调用时才从模块里取它，替换模块属性即可生效。"""
    from claude_agent_sdk._internal import message_parser

    original = message_parser.parse_message
    if getattr(original, "_deepexcel_patched", False):
        return

    def parse_message(data: Any):
        return original(fill_missing_thinking_signature(data))

    parse_message._deepexcel_patched = True  # type: ignore[attr-defined]
    message_parser.parse_message = parse_message
    try:
        from claude_agent_sdk._internal import client as internal_client
        internal_client.parse_message = parse_message
    except ImportError:
        pass
