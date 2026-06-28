# src/DeepExcel.Sidecar/tests/test_e2e_mock.py
"""
端到端测试：启动真实 sidecar.py 子进程，模拟 C# 行为验证完整环路。

注意：此测试需要真实 ANTHROPIC_API_KEY 才能跑通 SDK 调用。
未设置 API key 时跳过，标记为 integration test。
"""
import asyncio
import json
import os
import sys
import pytest


pytestmark = pytest.mark.skipif(
    not os.environ.get("ANTHROPIC_API_KEY"),
    reason="需要 ANTHROPIC_API_KEY 才能跑 sidecar E2E 测试",
)


@pytest.mark.asyncio
async def test_sidecar_e2e_write_formula():
    """完整环路：user_message → tool_call(write_formula) → tool_result → stream_end"""
    sidecar_path = os.path.join(os.path.dirname(__file__), "..", "sidecar.py")

    proc = await asyncio.create_subprocess_exec(
        sys.executable, sidecar_path,
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )

    try:
        # 1. 发 user_message
        user_msg = json.dumps({
            "type": "user_message",
            "text": "在 A1 写入 =SUM(B1:B10)",
            "session_id": "test-1",
            "context": {},
        }) + "\n"
        proc.stdin.write(user_msg.encode("utf-8"))
        await proc.stdin.drain()

        # 2. 读 sidecar 输出，遇 tool_call 回 tool_result
        received = []
        while True:
            line = await asyncio.wait_for(proc.stdout.readline(), timeout=30.0)
            if not line:
                break
            msg = json.loads(line.decode("utf-8"))
            received.append(msg)

            if msg["type"] == "tool_call":
                tool_result = json.dumps({
                    "type": "tool_result",
                    "call_id": msg["call_id"],
                    "success": True,
                    "data": {"address": "A1", "formula": "=SUM(B1:B10)"},
                    "error": None,
                    "suggestion": None,
                    "context": {},
                }) + "\n"
                proc.stdin.write(tool_result.encode("utf-8"))
                await proc.stdin.drain()

            if msg["type"] == "stream_end":
                break

        # 断言：至少收到一个 tool_call 和一个 stream_end
        types = [m["type"] for m in received]
        assert "tool_call" in types, f"未收到 tool_call，收到: {types}"
        assert "stream_end" in types, f"未收到 stream_end，收到: {types}"

    finally:
        try:
            proc.terminate()
            await asyncio.wait_for(proc.wait(), timeout=5.0)
        except Exception:
            proc.kill()
