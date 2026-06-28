# src/DeepExcel.Sidecar/tests/mock_csharp_driver.py
"""
模拟 C# 端与 sidecar 通信的驱动脚本。
用法：在测试中通过 subprocess 启动 sidecar.py，本脚本作为参考实现。

实际测试 test_e2e_mock.py 会内联这个逻辑。
"""
import asyncio
import json
import sys


async def drive_sidecar(stdin_writer, stdout_reader):
    """
    模拟 C# 行为：
    1. 发 user_message
    2. 读 sidecar 输出，遇到 tool_call 就回 tool_result
    3. 收到 stream_end 结束
    """
    # 1. 发 user_message
    stdin_writer.write((json.dumps({
        "type": "user_message",
        "text": "在 A1 写入 =SUM(B1:B10)",
        "session_id": "test-1",
        "context": {},
    }) + "\n").encode("utf-8"))
    await stdin_writer.drain()

    # 2. 读 sidecar 输出
    received_messages = []
    while True:
        line = await stdout_reader.readline()
        if not line:
            break
        msg = json.loads(line.decode("utf-8"))
        received_messages.append(msg)

        if msg["type"] == "tool_call":
            # 回写 tool_result
            tool_result = {
                "type": "tool_result",
                "call_id": msg["call_id"],
                "success": True,
                "data": {"address": "A1", "formula": "=SUM(B1:B10)"},
                "error": None,
                "suggestion": None,
                "context": {},
            }
            stdin_writer.write((json.dumps(tool_result) + "\n").encode("utf-8"))
            await stdin_writer.drain()

        if msg["type"] == "stream_end":
            break

    return received_messages
