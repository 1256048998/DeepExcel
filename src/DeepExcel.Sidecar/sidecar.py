# src/DeepExcel.Sidecar/sidecar.py
# 关键修复：
# 1. 用 ClaudeAgentOptions(env={...}) 传 DeepSeek 配置（os.environ 不生效）
# 2. 用 anyio 而非 asyncio.run（SDK 内部用 anyio，混用 asyncio.run 会事件循环冲突）
# 3. 用 async with ClaudeSDKClient 上下文管理器
# 4. 先等 C# 发来 config，再创建 options + client（env 必须在创建 options 时传入）
# 5. stdin_reader_loop 作为后台任务在整个生命周期内运行
# 6. 强制 stdout/stderr 为 UTF-8（Windows 默认 cp1252，中文/emoji 会 UnicodeEncodeError 崩溃）

import sys

# ★ 必须在所有其他操作之前强制 UTF-8 编码
# Windows 下 Python 子进程的 stdout 默认是 cp1252，写入中文/emoji 会抛 UnicodeEncodeError
# C# 端 PythonSidecar.cs 已设置 StandardOutputEncoding=UTF8，Python 端必须匹配
try:
    sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
    sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)
except Exception:
    # Python < 3.7 没有 reconfigure，用 TextIOWrapper 兜底
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', line_buffering=True)

import anyio
import json
import os

from claude_agent_sdk import (
    ClaudeAgentOptions,
    ClaudeSDKClient,
    create_sdk_mcp_server,
)
from claude_agent_sdk.types import AssistantMessage, ResultMessage, TextBlock, ToolUseBlock

from excel_tools import register_all_tools
from ipc import _message_buffer, read_message, route_message, write_message
from ipc import _init_buffer
from system_prompt import SYSTEM_PROMPT


async def stdin_reader_loop():
    """唯一的 stdin 读取协程，把消息分发到 _message_buffer。整个生命周期内运行。"""
    while True:
        msg = await read_message()
        if msg is None:
            # stdin 关闭，往 user_message queue 发 None 让主循环退出
            await _message_buffer["user_message"].put(None)
            return
        route_message(msg)


async def handle_sdk_message(response):
    """处理 ClaudeSDKClient.receive_response() 产生的流式消息"""
    if isinstance(response, AssistantMessage):
        for block in response.content:
            if isinstance(block, TextBlock):
                await write_message({"type": "stream_delta", "text": block.text})
            elif isinstance(block, ToolUseBlock):
                # SDK 内部已调度 @tool 函数，工具函数会通过 call_csharp() 发 tool_call 给 C#
                # 这里只发一个轻量通知给 C# 用于 UI 显示
                await write_message({
                    "type": "tool_use",
                    "tool": block.name,
                    "args": block.input if isinstance(block.input, dict) else {},
                })
    elif isinstance(response, ResultMessage):
        usage = getattr(response, "usage", None)
        if isinstance(usage, dict):
            in_tok = usage.get("input_tokens", 0)
            out_tok = usage.get("output_tokens", 0)
        elif usage:
            in_tok = getattr(usage, "input_tokens", 0)
            out_tok = getattr(usage, "output_tokens", 0)
        else:
            in_tok = 0
            out_tok = 0
        await write_message({
            "type": "stream_end",
            "input_tokens": in_tok,
            "output_tokens": out_tok,
        })


async def run_agent_loop(client):
    """主循环：从 user_message queue 取消息，发给 SDK 处理"""
    while True:
        msg = await _message_buffer["user_message"].get()
        if msg is None:
            return

        # 重置 cancel 标志（每次新对话开始前）
        if _message_buffer["cancel"].is_set():
            _message_buffer["cancel"].clear()

        user_text = msg.get("text", "")
        sys.stderr.write(f"[sidecar] run_agent_loop: processing user message (len={len(user_text)})\n")
        sys.stderr.flush()

        try:
            await client.query(user_text)
            sys.stderr.write("[sidecar] client.query() done, receiving response...\n")
            sys.stderr.flush()

            # 用 cancel_scope 包装 receive_response，以便在 cancel 时能中断
            async with anyio.create_task_group() as inner_tg:
                # 启动一个 watchdog 协程：检测 cancel 标志，触发 task group 取消
                async def _cancel_watchdog():
                    while True:
                        if _message_buffer["cancel"].is_set():
                            sys.stderr.write("[sidecar] cancel detected by watchdog, cancelling receive_response\n")
                            sys.stderr.flush()
                            inner_tg.cancel_scope.cancel()
                            return
                        await anyio.sleep(0.1)

                inner_tg.start_soon(_cancel_watchdog)

                try:
                    async for response in client.receive_response():
                        await handle_sdk_message(response)
                except Exception as inner_e:
                    sys.stderr.write(f"[sidecar] receive_response exception: {type(inner_e).__name__}: {inner_e}\n")
                    sys.stderr.flush()
                    raise

                # 正常结束，取消 watchdog
                inner_tg.cancel_scope.cancel()

            sys.stderr.write("[sidecar] receive_response completed normally\n")
            sys.stderr.flush()

        except Exception as e:
            sys.stderr.write(f"[sidecar] run_agent_loop exception: {type(e).__name__}: {e}\n")
            sys.stderr.flush()
            # anyio 的取消异常不在这里捕获，让它传播
            if isinstance(e, (KeyboardInterrupt, SystemExit)):
                raise
            # 不用 emoji（避免编码问题），纯文本错误消息
            await write_message({"type": "stream_delta", "text": f"[错误] {type(e).__name__}: {e}"})
            await write_message({"type": "stream_end", "input_tokens": 0, "output_tokens": 0})

        # 如果是 cancel 触发的结束，主动发 stream_end 让 UI 恢复
        if _message_buffer["cancel"].is_set():
            sys.stderr.write("[sidecar] cancel was set, sending stream_end to unblock UI\n")
            sys.stderr.flush()
            await write_message({"type": "stream_end", "input_tokens": 0, "output_tokens": 0})
            _message_buffer["cancel"].clear()


async def main():
    _init_buffer()
    sys.stderr.write("[sidecar] main() started\n")
    sys.stderr.flush()

    # 启动 stdin reader 作为后台任务（整个生命周期内运行）
    # 用 task_group 管理后台 reader
    async with anyio.create_task_group() as tg:
        tg.start_soon(stdin_reader_loop)
        sys.stderr.write("[sidecar] stdin_reader_loop started, waiting for config...\n")
        sys.stderr.flush()

        # 等待 config 消息（最多等 30 秒）
        cfg = None
        try:
            with anyio.fail_after(30.0):
                cfg = await _message_buffer["config"].get()
            sys.stderr.write(f"[sidecar] config received: base_url={cfg.get('base_url')}, model={cfg.get('model')}, hasKey={bool(cfg.get('api_key'))}\n")
            sys.stderr.flush()
        except TimeoutError:
            sys.stderr.write("[sidecar] WARNING: config timeout (30s), falling back to env vars\n")
            sys.stderr.flush()

        # 构建 env 配置（DeepSeek 必须通过 env 传给 SDK，os.environ 不生效）
        if cfg:
            env_config = {
                "ANTHROPIC_BASE_URL": cfg.get("base_url", "https://api.anthropic.com"),
                "ANTHROPIC_API_KEY": cfg.get("api_key", ""),
            }
            model = cfg.get("model", "claude-sonnet-4")
        else:
            env_config = {
                "ANTHROPIC_BASE_URL": os.environ.get("ANTHROPIC_BASE_URL", "https://api.anthropic.com"),
                "ANTHROPIC_API_KEY": os.environ.get("ANTHROPIC_API_KEY", ""),
            }
            model = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4")

        # 注册工具到 MCP server
        server = create_sdk_mcp_server(name="excel", tools=register_all_tools())

        options = ClaudeAgentOptions(
            model=model,
            mcp_servers={"excel": server},
            allowed_tools=[f"mcp__excel__{t}" for t in [
                "echo",
                "read_workbook", "read_selection", "read_range",
                "write_formula", "fill_formula_down", "replace_formula",
                "clean_data", "create_chart", "create_pivot_table",
                "execute_vba", "execute_python",
                "create_snapshot", "rollback",
                "add_sheet", "delete_sheet", "rename_sheet",
                "set_number_format", "set_column_width",
                "sort_data", "filter_data",
                "merge_cells", "unmerge_cells",
                "set_cell_style", "copy_range", "clear_range",
                "insert_rows", "delete_rows", "insert_columns", "delete_columns",
                "freeze_panes",
                "apply_conditional_format", "write_table",
                "clarify_intent",
            ]],
            system_prompt=SYSTEM_PROMPT,
            max_turns=20,
            env=env_config,  # ★ DeepSeek 配置必须在这里传
            # ★ 禁用 thinking：DeepSeek anthropic 兼容端点不返回 thinking block 的 signature 字段，
            # SDK 在 message_parser.py:104 硬编码 block["signature"] 会抛 MessageParseError，
            # 导致 tool_use 后第二轮 API 响应解析崩溃，最终文本永不返回。
            # 参考：https://github.com/anthropics/claude-agent-sdk-python/issues/949
            thinking={"type": "disabled"},
        )

        # 用 async with 上下文管理器创建 client
        sys.stderr.write(f"[sidecar] creating ClaudeSDKClient, model={model}, base_url={env_config.get('ANTHROPIC_BASE_URL')}\n")
        sys.stderr.flush()
        async with ClaudeSDKClient(options=options) as client:
            sys.stderr.write("[sidecar] ClaudeSDKClient connected, entering agent loop\n")
            sys.stderr.flush()
            # 在 task_group 里跑主循环，reader 继续在后台跑
            await run_agent_loop(client)

        # 主循环退出后，取消 reader
        tg.cancel_scope.cancel()


if __name__ == "__main__":
    anyio.run(main)
