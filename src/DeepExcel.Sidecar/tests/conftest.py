import sys
import os

import pytest

# 让 tests/ 能 import sidecar 模块
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


@pytest.fixture(autouse=True)
def _reset_ipc_state():
    """ipc 的消息缓冲是模块级的：一个测试按下的「停止」、留下的插话，
    不能漏到下一个测试里（等待中的工具调用会因此立刻返回「已中断」）。"""
    import ipc
    yield
    cancel = ipc._message_buffer.get("cancel")
    if cancel is not None:
        cancel.clear()
    ipc._message_buffer["steer"] = []
    ipc._message_buffer["turn_active"] = False
    ipc._message_buffer["awaiting_user"] = 0
