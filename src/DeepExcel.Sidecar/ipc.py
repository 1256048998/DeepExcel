# src/DeepExcel.Sidecar/ipc.py
import asyncio
import json
import sys
import uuid
from typing import Any, Dict, Optional


def _log(msg: str) -> None:
    """写入 stderr 供 C# 端 Logger 捕获（不经过 stdout，避免干扰 IPC）"""
    try:
        sys.stderr.write(f"[ipc] {msg}\n")
        sys.stderr.flush()
    except Exception:
        pass


# 全局消息缓冲：所有从 stdin 收到的消息按 type 路由到这里
# 关键：C# 可能并行发来多个 tool_result，必须按 call_id 匹配，不能顺序读
_message_buffer: Dict[str, Any] = {
    "tool_result": {},        # call_id -> msg
    "clarify_answer": None,   # 最近一次 clarify 的回答字符串
    "user_message": None,     # asyncio.Queue，在 _init_buffer() 中创建
    "cancel": None,           # asyncio.Event，在 _init_buffer() 中创建
    "config": None,           # asyncio.Queue，在 _init_buffer() 中创建
    "restore_history": None,  # ★ 历史对话恢复（list of {role, content}），None 表示无历史
}


def _init_buffer():
    """在 asyncio 事件循环内调用，创建 Queue/Event（它们绑定到当前 loop）"""
    if _message_buffer["user_message"] is None:
        _message_buffer["user_message"] = asyncio.Queue()
    if _message_buffer["cancel"] is None:
        _message_buffer["cancel"] = asyncio.Event()
    if _message_buffer["config"] is None:
        _message_buffer["config"] = asyncio.Queue()


def generate_call_id() -> str:
    return str(uuid.uuid4())


async def write_message(msg: dict) -> None:
    """写一行 JSON 到 stdout（C# 读）"""
    sys.stdout.write(json.dumps(msg, ensure_ascii=False) + "\n")
    sys.stdout.flush()


async def read_message() -> Optional[dict]:
    """从 stdin 读一行 JSON（C# 写），EOF 返回 None"""
    loop = asyncio.get_event_loop()
    # run_in_executor 避免阻塞事件循环
    line = await loop.run_in_executor(None, sys.stdin.readline)
    if not line:
        return None
    try:
        return json.loads(line)
    except json.JSONDecodeError as e:
        # 损坏的 JSON 不应导致整个 sidecar 崩溃
        await write_message({"type": "stream_delta", "text": f"[ipc] 解析 stdin 失败: {e}"})
        return None


# 追加到 src/DeepExcel.Sidecar/ipc.py

def route_message(msg: dict) -> None:
    """把从 stdin 读到的消息按 type 路由到对应缓冲区"""
    _init_buffer()
    t = msg.get("type")
    _log(f"route_message: type={t}")
    if t == "tool_result":
        cid = msg.get("call_id", "?")
        _message_buffer["tool_result"][cid] = msg
        _log(f"route_message: tool_result stored, call_id={cid}, waiting_keys={list(_message_buffer['tool_result'].keys())}")
    elif t == "clarify_answer":
        _message_buffer["clarify_answer"] = msg.get("answer", "")
        _log("route_message: clarify_answer stored")
    elif t == "user_message":
        # 用 put_nowait 避免在同步函数里调 async（anyio 下 create_task 可能有问题）
        _message_buffer["user_message"].put_nowait(msg)
        _log("route_message: user_message enqueued")
    elif t == "cancel":
        _message_buffer["cancel"].set()
        _log("route_message: cancel set")
    elif t == "config":
        _message_buffer["config"].put_nowait(msg)
        _log("route_message: config enqueued")
    elif t == "restore_history":
        # ★ 历史对话恢复：存到全局变量，sidecar 主循环会在首条 user_message 时注入
        _message_buffer["restore_history"] = msg.get("messages", [])
        _log(f"route_message: restore_history stored, count={len(_message_buffer['restore_history'] or [])}")
    else:
        _log(f"route_message: UNKNOWN type={t}, ignored")
    # 未知 type 静默丢弃，避免 sidecar 崩溃


# 追加到 src/DeepExcel.Sidecar/ipc.py

async def call_csharp(tool_name: str, args: dict, timeout: float = 60.0) -> dict:
    """向 C# 发送工具调用请求，阻塞等待结果（按 call_id 匹配）。
    ★ 默认 60 秒超时，避免 C# 卡死时 sidecar 永久 hang。
    超时返回 success=False 的错误结果，让 LLM 能继续推理。"""
    _init_buffer()
    call_id = generate_call_id()
    _log(f"call_csharp: sending tool_call, tool={tool_name}, call_id={call_id}")
    await write_message({
        "type": "tool_call",
        "call_id": call_id,
        "tool": tool_name,
        "args": args,
    })
    # 只检查 buffer，不直接读 stdin（stdin 由 stdin_reader_loop 统一读取）
    poll_count = 0
    deadline = asyncio.get_event_loop().time() + timeout
    while True:
        if call_id in _message_buffer["tool_result"]:
            result = _message_buffer["tool_result"].pop(call_id)
            _log(f"call_csharp: got result, tool={tool_name}, call_id={call_id}, success={result.get('success')}")
            return result
        poll_count += 1
        if poll_count % 20 == 0:  # 每 1 秒打印一次等待日志
            _log(f"call_csharp: waiting... tool={tool_name}, call_id={call_id}, poll_count={poll_count}, buffer_keys={list(_message_buffer['tool_result'].keys())}")
        # ★ 超时检查：C# 卡死时返回明确错误，避免 sidecar 永久阻塞
        if asyncio.get_event_loop().time() > deadline:
            _log(f"call_csharp: TIMEOUT after {timeout}s, tool={tool_name}, call_id={call_id}")
            return {
                "success": False,
                "error": f"工具 {tool_name} 调用超时（{timeout}s），C# 端可能已卡死",
            }
        await asyncio.sleep(0.05)


async def call_csharp_clarify(question: str, options: list) -> str:
    """向 C# 发送澄清请求，阻塞等待用户回答"""
    _init_buffer()
    _message_buffer["clarify_answer"] = None  # 重置
    await write_message({
        "type": "clarify",
        "question": question,
        "options": options,
    })
    while _message_buffer["clarify_answer"] is None:
        await asyncio.sleep(0.05)
    answer = _message_buffer["clarify_answer"]
    _message_buffer["clarify_answer"] = None
    return answer
