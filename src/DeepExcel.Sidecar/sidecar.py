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
import threading
import time
from pathlib import Path

from claude_agent_sdk import (
    ClaudeAgentOptions,
    ClaudeSDKClient,
    create_sdk_mcp_server,
)
from claude_agent_sdk.types import AssistantMessage, ResultMessage, StreamEvent, TextBlock, ToolUseBlock

from excel_tools import register_all_tools
from ipc import _message_buffer, read_message, route_message, write_message
from ipc import _init_buffer
from system_prompt import SYSTEM_PROMPT


def _parent_watchdog():
    """★ 父进程心跳检测：Excel 被强制 kill 时 OnDisconnection 不会触发，
    sidecar 会变成孤儿进程继续 spin。这个线程每 2 秒检查父进程（Excel）是否还在，
    不在则立即退出 sidecar，避免残留进程占用资源。
    """
    parent_pid = os.getppid()
    while True:
        time.sleep(2.0)
        try:
            new_ppid = os.getppid()
            # Windows 下父进程退出后 getppid() 可能返回相同值或 1，
            # 但 os.kill(pid, 0) 会抛 ProcessLookupError
            if new_ppid != parent_pid:
                sys.stderr.write(f"[sidecar] parent process changed ({parent_pid} -> {new_ppid}), exiting\n")
                sys.stderr.flush()
                os._exit(1)
            # 尝试检查父进程是否还活着（Windows 不支持 os.kill(pid, 0)，用 OpenProcess 替代）
            try:
                import ctypes
                kernel32 = ctypes.windll.kernel32
                PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
                handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, parent_pid)
                if handle == 0:
                    sys.stderr.write(f"[sidecar] parent process {parent_pid} no longer accessible, exiting\n")
                    sys.stderr.flush()
                    os._exit(1)
                kernel32.CloseHandle(handle)
            except Exception:
                pass
        except Exception:
            pass


# 可作为文本读取的附件扩展名
_TEXT_EXTENSIONS = {
    ".txt", ".md", ".markdown", ".csv", ".json", ".xml", ".html", ".htm",
    ".py", ".cs", ".js", ".ts", ".tsx", ".jsx", ".java", ".c", ".cpp", ".h",
    ".log", ".yml", ".yaml", ".ini", ".cfg", ".conf", ".sh", ".bat", ".ps1",
    ".sql", ".rb", ".go", ".rs", ".php", ".lua", ".r", ".m",
}

# ★ 图片附件扩展名 — 直接以 base64 image content block 发送给 Claude vision
_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}

# 图片 MIME 类型映射
_IMAGE_MIME = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".webp": "image/webp",
}

# ★ PDF 附件扩展名 — Claude 原生支持 PDF document block
_PDF_EXTENSIONS = {".pdf"}

# 单张图片最大 5MB（Claude API 限制 base64 不超过 5MB）
_MAX_IMAGE_SIZE = 5 * 1024 * 1024

# PDF 最大 32MB（Claude API 限制）
_MAX_PDF_SIZE = 32 * 1024 * 1024

_MAX_ATTACHMENT_TEXT_SIZE = 100 * 1024  # 单附件最多读 100KB 文本

# ★ 模型能力检测：判断当前模型是否支持 vision（image block）和 document（PDF block）
# DeepSeek 等兼容端点不支持 vision，需要降级处理（PDF 提取文本，图片提示不支持）
_VISION_UNSUPPORTED_KEYWORDS = ("deepseek", "qwen", "llama", "yi-", "glm")


def _supports_vision(base_url: str, model: str) -> bool:
    """检测当前模型/端点是否支持 image vision 和 PDF document block。
    通过 base_url 和模型名关键词判断。
    """
    combined = (base_url or "").lower() + " " + (model or "").lower()
    for kw in _VISION_UNSUPPORTED_KEYWORDS:
        if kw in combined:
            return False
    # 默认认为 Claude/OpenAI/GPT 系列支持 vision
    return True


def _extract_pdf_text(path: str, max_chars: int = 50000) -> str:
    """用 PyPDF2 提取 PDF 文本内容（降级方案，用于不支持 vision 的模型）。
    超过 max_chars 时截断。
    """
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(path)
        pages = []
        total = 0
        for page in reader.pages:
            text = page.extract_text() or ""
            pages.append(text)
            total += len(text)
            if total >= max_chars:
                break
        content = "\n\n".join(pages)
        if len(content) > max_chars:
            content = content[:max_chars] + f"\n... [PDF 过长，已截断，共 {len(reader.pages)} 页]"
        return content if content.strip() else "[PDF 文本提取为空，可能是扫描件/图片型 PDF]"
    except Exception as e:
        return f"[PDF 文本提取失败: {e}]"


def _build_attachment_context(attachments: list, supports_vision: bool = True) -> str:
    """从附件列表构建上下文文本。文本文件读内容，其他只列元信息。"""
    if not attachments:
        return ""

    lines = []
    lines.append("=== 用户上传的附件 ===")
    lines.append(f"共 {len(attachments)} 个附件：")

    for i, att in enumerate(attachments, 1):
        name = att.get("name", f"附件{i}")
        size = att.get("size", 0)
        path = att.get("path", "")
        ext = Path(name).suffix.lower()

        lines.append("")
        lines.append(f"[{i}] 文件名: {name}")
        lines.append(f"    大小: {size} 字节")

        # 文本文件尝试读取内容
        if ext in _TEXT_EXTENSIONS and path and os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    content = f.read(_MAX_ATTACHMENT_TEXT_SIZE)
                if len(content) >= _MAX_ATTACHMENT_TEXT_SIZE:
                    content += "\n... [文件过长，已截断]"
                lines.append(f"    内容如下：")
                lines.append(f"    ```{ext.lstrip('.') or 'text'}")
                for line in content.splitlines():
                    lines.append(f"    {line}")
                lines.append(f"    ```")
            except Exception as e:
                lines.append(f"    [读取失败: {e}]")
        elif ext in _IMAGE_EXTENSIONS:
            if supports_vision:
                # ★ 图片附件：直接以 base64 image content block 发送给 Claude vision
                lines.append(f"    [图片已直接发送给 AI，您可以直接看到图片内容]")
            else:
                # ★ 当前模型不支持 vision，图片无法识别
                lines.append(f"    [图片附件 - 当前模型不支持图片识别，请提示用户换用支持 vision 的模型（如 Claude）或手动输入图片中的数据]")
        elif ext in _PDF_EXTENSIONS:
            if supports_vision:
                # ★ PDF 附件：直接以 base64 document content block 发送给 Claude
                lines.append(f"    [PDF 已直接发送给 AI，您可以直接看到 PDF 内容]")
            else:
                # ★ 降级方案：用 PyPDF2 提取文本拼到 prompt
                if path and os.path.exists(path):
                    pdf_text = _extract_pdf_text(path)
                    lines.append(f"    PDF 文本内容如下：")
                    lines.append(f"    ```pdf")
                    for line in pdf_text.splitlines():
                        lines.append(f"    {line}")
                    lines.append(f"    ```")
                else:
                    lines.append(f"    [PDF 文件路径无效，无法提取文本]")
        else:
            # ★ 非文本非图片非 PDF 文件（xlsx/docx 等）：提示 AI 用 read_attachment 工具读取
            lines.append(f"    ★ 内容需用 read_attachment 工具读取（传入 file_name=\"{name}\"）")

    lines.append("")
    lines.append("=== 附件结束 ===")
    lines.append("")
    lines.append("提示：如需读取附件文件内容，请调用 read_attachment 工具，传入附件文件名。")
    lines.append("")
    return "\n".join(lines)


def _collect_image_blocks(attachments: list) -> list:
    """★ 从附件列表中提取图片文件，读取 base64 构造 Claude image content blocks。
    返回 image block 列表，供拼接到用户消息的 content 数组中。
    超过 _MAX_IMAGE_SIZE 的图片会被跳过并记录警告。
    """
    import base64

    blocks = []
    for att in attachments:
        name = att.get("name", "")
        path = att.get("path", "")
        size = att.get("size", 0)
        ext = Path(name).suffix.lower()

        if ext not in _IMAGE_EXTENSIONS:
            continue
        if not path or not os.path.exists(path):
            sys.stderr.write(f"[sidecar] image attachment not found: {name} -> {path}\n")
            sys.stderr.flush()
            continue
        if size > _MAX_IMAGE_SIZE:
            sys.stderr.write(f"[sidecar] image too large ({size} bytes > {_MAX_IMAGE_SIZE}): {name}\n")
            sys.stderr.flush()
            continue

        try:
            with open(path, "rb") as f:
                raw = f.read()
            b64 = base64.b64encode(raw).decode("ascii")
            mime = _IMAGE_MIME.get(ext, "image/png")
            blocks.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": mime,
                    "data": b64,
                },
            })
            sys.stderr.write(f"[sidecar] image block created: {name} ({size} bytes, {mime})\n")
            sys.stderr.flush()
        except Exception as e:
            sys.stderr.write(f"[sidecar] failed to read image {name}: {e}\n")
            sys.stderr.flush()

    return blocks


def _collect_pdf_blocks(attachments: list) -> list:
    """★ 从附件列表中提取 PDF 文件，读取 base64 构造 Claude document content blocks。
    Claude 原生支持 PDF，能直接理解 PDF 中的文字、表格、图表。
    单个 PDF 最大 32MB，最多 100 页（Claude API 限制）。
    """
    import base64

    blocks = []
    for att in attachments:
        name = att.get("name", "")
        path = att.get("path", "")
        size = att.get("size", 0)
        ext = Path(name).suffix.lower()

        if ext not in _PDF_EXTENSIONS:
            continue
        if not path or not os.path.exists(path):
            sys.stderr.write(f"[sidecar] pdf attachment not found: {name} -> {path}\n")
            sys.stderr.flush()
            continue
        if size > _MAX_PDF_SIZE:
            sys.stderr.write(f"[sidecar] pdf too large ({size} bytes > {_MAX_PDF_SIZE}): {name}\n")
            sys.stderr.flush()
            continue

        try:
            with open(path, "rb") as f:
                raw = f.read()
            b64 = base64.b64encode(raw).decode("ascii")
            blocks.append({
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": b64,
                },
            })
            sys.stderr.write(f"[sidecar] pdf block created: {name} ({size} bytes)\n")
            sys.stderr.flush()
        except Exception as e:
            sys.stderr.write(f"[sidecar] failed to read pdf {name}: {e}\n")
            sys.stderr.flush()

    return blocks


def _build_history_summary(history: list) -> str:
    """把历史对话列表构建为上下文摘要文本。
    只取 user/assistant 文本，跳过 tool/clarify。
    限制总长度避免 token 爆炸。
    """
    if not history:
        return ""

    lines = []
    lines.append("=== 之前的对话历史 ===")
    lines.append("（用户继续了之前的对话，以下是历史记录摘要，请基于此上下文回答）")
    lines.append("")

    MAX_CHARS = 8000  # 历史摘要最多 8000 字符，避免 token 爆炸
    total = 0
    for msg in history:
        role = msg.get("role", "")
        content = msg.get("content", "") or ""
        if not content or role not in ("user", "assistant"):
            continue
        # 跳过 clarify 类型的 assistant 消息
        if role == "assistant" and msg.get("type") == "clarify":
            continue

        label = "用户" if role == "user" else "助手"
        line = f"{label}: {content}"
        if total + len(line) > MAX_CHARS:
            remaining = MAX_CHARS - total
            if remaining > 50:
                line = line[:remaining] + "... [截断]"
                lines.append(line)
            break
        lines.append(line)
        total += len(line)

    lines.append("")
    lines.append("=== 历史结束 ===")
    lines.append("")
    return "\n".join(lines)


async def stdin_reader_loop():
    """唯一的 stdin 读取协程，把消息分发到 _message_buffer。整个生命周期内运行。"""
    while True:
        msg = await read_message()
        if msg is None:
            # stdin 关闭，往 user_message queue 发 None 让主循环退出
            await _message_buffer["user_message"].put(None)
            return
        route_message(msg)


# ★ 流式标志：本轮是否已通过 StreamEvent 发送过 token 级 delta。
# 每次新 user message 开始时重置为 False。
_had_partial_text = False


async def handle_sdk_message(response):
    """处理 ClaudeSDKClient.receive_response() 产生的流式消息"""
    global _had_partial_text
    if isinstance(response, StreamEvent):
        # ★ token 级流式：解析 Anthropic 原生 content_block_delta 事件
        # event 格式: {"type": "content_block_delta", "delta": {"type": "text_delta", "text": "..."}}
        evt = response.event if isinstance(response.event, dict) else {}
        if evt.get("type") == "content_block_delta":
            delta = evt.get("delta", {})
            if isinstance(delta, dict) and delta.get("type") == "text_delta":
                text = delta.get("text", "")
                if text:
                    _had_partial_text = True
                    await write_message({"type": "stream_delta", "text": text})
    elif isinstance(response, AssistantMessage):
        for block in response.content:
            if isinstance(block, TextBlock):
                # ★ 如果已通过 StreamEvent 发送过 token delta，这里不再发完整文本（避免重复）
                # 如果没有（某些 provider 不支持流式 / include_partial_messages 无效），发完整文本兜底
                if not _had_partial_text:
                    await write_message({"type": "stream_delta", "text": block.text})
            elif isinstance(block, ToolUseBlock):
                # SDK 内部已调度 @tool 函数，工具函数会通过 call_csharp() 发 tool_call 给 C#
                # 这里只发一个轻量通知给 C# 用于 UI 显示
                await write_message({
                    "type": "tool_use",
                    "tool": block.name,
                    "args": block.input if isinstance(block.input, dict) else {},
                })
        # ★ 一条 AssistantMessage 处理完后重置标志，为下一条消息（工具调用后的最终回复）做准备
        _had_partial_text = False
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


async def run_agent_loop(client, supports_vision: bool = True):
    """主循环：从 user_message queue 取消息，发给 SDK 处理"""
    while True:
        msg = await _message_buffer["user_message"].get()
        if msg is None:
            return

        # 重置 cancel 标志（每次新对话开始前）
        if _message_buffer["cancel"].is_set():
            _message_buffer["cancel"].clear()

        # ★ 重置流式标志：每轮新对话开始时，尚未收到任何 token delta
        global _had_partial_text
        _had_partial_text = False

        user_text = msg.get("text", "")
        context = msg.get("context") or {}
        attachments = context.get("attachments") if isinstance(context, dict) else None

        sys.stderr.write(f"[sidecar] run_agent_loop: processing user message (len={len(user_text)}), supports_vision={supports_vision}\n")
        sys.stderr.flush()

        # ★ 附件上下文注入：如果有附件，把附件信息拼到用户消息前面
        final_text = user_text
        if attachments and isinstance(attachments, list) and len(attachments) > 0:
            att_context = _build_attachment_context(attachments, supports_vision=supports_vision)
            if att_context:
                final_text = att_context + "\n\n用户问题：\n" + user_text
                sys.stderr.write(f"[sidecar] attachment context injected ({len(attachments)} attachments)\n")
                sys.stderr.flush()

        # ★ 历史上下文注入：如果 C# 发了 restore_history，把历史对话摘要拼到首条用户消息前。
        # 这样 AI 能"记得"之前聊过什么。只在首条消息注入一次，注入后清空。
        history = _message_buffer.get("restore_history")
        if history and isinstance(history, list) and len(history) > 0:
            summary = _build_history_summary(history)
            if summary:
                final_text = summary + "\n\n用户问题：\n" + final_text
                sys.stderr.write(f"[sidecar] history context injected ({len(history)} messages)\n")
                sys.stderr.flush()
            # 注入后清空，避免后续消息重复注入
            _message_buffer["restore_history"] = None

        try:
            # ★ 附件处理：仅当模型支持 vision 时才发送 image/document block
            direct_blocks = []
            if supports_vision and attachments and isinstance(attachments, list) and len(attachments) > 0:
                image_blocks = _collect_image_blocks(attachments)
                pdf_blocks = _collect_pdf_blocks(attachments)
                direct_blocks = image_blocks + pdf_blocks

            if direct_blocks:
                # 有图片/PDF：用 content blocks 数组（image/document blocks + text block）
                content_blocks = direct_blocks + [{"type": "text", "text": final_text}]

                async def _msg_gen():
                    yield {
                        "type": "user",
                        "message": {"role": "user", "content": content_blocks},
                        "parent_tool_use_id": None,
                        "session_id": "default",
                    }

                await client.query(_msg_gen())
                sys.stderr.write(f"[sidecar] client.query() done (with {len(image_blocks)} image(s) + {len(pdf_blocks)} pdf(s)), receiving response...\n")
            else:
                # 无图片/PDF 或不支持 vision：传纯文本（PDF 文本已在 _build_attachment_context 中提取并拼入）
                await client.query(final_text)
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
                "read_workbook", "read_selection", "read_range", "read_attachment",
                "write_formula", "write_value", "write_range", "fill_formula_down", "replace_formula",
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
            # ★ 启用 token 级流式：receive_response() 会产出 StreamEvent（含 content_block_delta），
            # 让 UI 逐 token 显示文本，而不是整段一次性出现
            include_partial_messages=True,
        )

        # ★ 检测模型是否支持 vision（image/document block）
        base_url_for_check = env_config.get("ANTHROPIC_BASE_URL", "")
        vision_ok = _supports_vision(base_url_for_check, model)
        sys.stderr.write(f"[sidecar] creating ClaudeSDKClient, model={model}, base_url={base_url_for_check}, supports_vision={vision_ok}\n")
        sys.stderr.flush()
        async with ClaudeSDKClient(options=options) as client:
            sys.stderr.write("[sidecar] ClaudeSDKClient connected, entering agent loop\n")
            sys.stderr.flush()
            # 在 task_group 里跑主循环，reader 继续在后台跑
            await run_agent_loop(client, supports_vision=vision_ok)

        # 主循环退出后，取消 reader
        tg.cancel_scope.cancel()


if __name__ == "__main__":
    # 启动父进程心跳检测线程（daemon，主进程退出时自动结束）
    t = threading.Thread(target=_parent_watchdog, daemon=True)
    t.start()
    anyio.run(main)
