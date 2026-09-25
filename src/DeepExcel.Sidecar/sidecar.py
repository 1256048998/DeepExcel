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
# Windows 下 Python 子进程的 stdin/stdout 默认是 cp1252，写入中文/emoji 会抛 UnicodeEncodeError
# C# 端 PythonSidecar.cs 已设置 StandardOutputEncoding=UTF8，Python 端必须匹配
# ★ 同时 reconfigure stdin：系统语言为英文时 ANSI 代码页 cp1252 不支持中文，
# C# → Python 方向的消息（如 tool_result）中的中文会在 stdin 读取时丢失
import io

# 每个流独立处理：任何一个流失败都不能连累另外两个。
# 早期版本把三个 reconfigure 放在同一个 try 里，stdin 失败（如 pytest 的
# DontReadFromInput 没有 reconfigure）会导致已经成功的 stdout/stderr 被重新包一层
# TextIOWrapper；新旧 wrapper 之一被 GC 时会关掉底层 buffer，
# 表现为 "I/O operation on closed file"。
_original_streams = []


def _force_utf8(name: str) -> None:
    stream = getattr(sys, name, None)
    if stream is None:
        return
    try:
        stream.reconfigure(encoding='utf-8', line_buffering=True)
        return
    except Exception:
        # Python < 3.7 没有 reconfigure；或该流是测试替身
        pass
    buffer = getattr(stream, 'buffer', None)
    if buffer is None:
        # 没有底层字节流可包（测试替身等），保持原样，绝不替换
        return
    try:
        wrapped = io.TextIOWrapper(buffer, encoding='utf-8', line_buffering=True)
    except Exception:
        return
    # 保留原 stream 引用：否则它被 GC 时会连带关闭 buffer
    _original_streams.append(stream)
    setattr(sys, name, wrapped)


for _stream_name in ('stdout', 'stderr', 'stdin'):
    _force_utf8(_stream_name)

import anyio
import base64
import json
import os
import re
import threading
import time
from pathlib import Path

from claude_agent_sdk import (
    ClaudeAgentOptions,
    ClaudeSDKClient,
    create_sdk_mcp_server,
)
from claude_agent_sdk.types import (
    AssistantMessage,
    HookMatcher,
    ResultMessage,
    StreamEvent,
    SystemMessage,
    TextBlock,
    ToolResultBlock,
    ToolUseBlock,
    UserMessage,
)

import explorer
import ui_events
from excel_tools import host_tool_note, register_all_tools
from model_windows import apply_context_window
from ipc import _message_buffer, read_message, route_message, write_message
from ipc import _init_buffer, request_permission, take_steer_messages
from system_prompt import SYSTEM_PROMPT


def _load_wps_local_config():
    """Load the current Excel configuration for the WPS host.

    WPS starts the same per-user sidecar without the C# bridge that normally
    sends a config message. API keys remain DPAPI-protected and are decrypted
    only in memory for the current Windows user.
    """
    try:
        appdata = os.environ.get("APPDATA", "")
        config_path = os.path.join(appdata, "DeepExcel", "config.json")
        with open(config_path, "r", encoding="utf-8") as stream:
            config = json.load(stream)
        provider = config.get("CurrentProvider") or config.get("currentProvider") or "anthropic"
        if not re.fullmatch(r"[A-Za-z0-9_-]+", provider):
            raise ValueError("invalid provider key")
        providers = config.get("Providers") or config.get("providers") or {}
        provider_config = providers.get(provider) or {}
        model = config.get("CurrentModel") or config.get("currentModel") or provider_config.get("DefaultModel")
        base_url = provider_config.get("BaseUrl") or provider_config.get("baseUrl") or "https://api.anthropic.com"

        credential_path = os.path.join(appdata, "DeepExcel", "credentials", f"key_{provider}.crypt")
        with open(credential_path, "r", encoding="utf-8") as stream:
            protected = base64.b64decode(stream.read().strip(), validate=True)
        import win32crypt
        api_key = win32crypt.CryptUnprotectData(protected, None, None, None, 0)[1].decode("utf-8")
        if not api_key:
            raise ValueError("empty API key")
        general = config.get("General") or config.get("general") or {}
        max_turns = general.get("MaxTurns") or general.get("maxTurns")
        return {"base_url": base_url, "model": model, "api_key": api_key, "max_turns": max_turns}
    except Exception as exc:
        sys.stderr.write(f"[sidecar] WPS local config unavailable: {type(exc).__name__}: {exc}\n")
        sys.stderr.flush()
        return None


# 需要经过确认流程的工具。
#
# 注意这里只是"要不要问 C#"，不是"一定会弹窗"。C# 侧会先算出具体变更集，
# 只有值得打扰用户时才真的弹；改 3 个单元格或根本没有变化的操作会被自动放行。
# 把这个判断放在 C# 而不是这里，是因为只有那边能看到工作簿的当前状态。
_HIGH_RISK_TOOLS = {
    # 任意代码：无法预演，只能快照 + 明确告知
    "execute_vba",
    "execute_jsa",
    "execute_python",
    "send_keys",
    # 破坏性结构操作
    "rollback",
    "delete_rows", "delete_columns", "delete_sheet", "delete_blank_rows",
    "clear_range",
    # 大范围写入：可能覆盖公式
    "write_range", "write_table", "replace_formula", "fill_formula_down",
    # 数据清洗：会批量改写单元格
    "clean_data", "remove_duplicates", "split_text_to_columns",
    "fill_blank_cells", "remove_special_chars", "clean_amount",
    "merge_columns", "rename_columns", "collapse_spaces",
    "text_to_number", "unify_date", "trim_spaces",
    # 合并单元格会丢数据
    "merge_cells", "unmerge_cells",
}

# "允许并记住"只适用于授予一种能力，不适用于逐次判断影响范围的操作。
#
# 批准过一次 delete_rows，不等于批准了之后任何一次删除——下一次可能删 200 行。
# 只有 VBA/Python 这类"我信任它能跑代码"的授权适合记住。
_REMEMBERABLE_TOOLS = {
    "execute_vba",
    "execute_jsa",
    "execute_python",
    "send_keys",
}

# ★ 会话级"允许并记住"：仅对 _REMEMBERABLE_TOOLS 生效
_allowed_tools_session = set()

# ★ 最近一条用户消息文本（用于 Computer Use 工具的触发门槛检查）
# run_agent_loop 每次收到用户消息时更新，PreToolUse hook 据此判断
# screenshot_excel/send_keys 是否由用户主动要求
_last_user_text = ""

# ★ Computer Use 触发关键词：用户消息含这些词才允许调用 screenshot_excel/send_keys
# 匹配到即放行，不匹配则 deny 并提示 AI"需要用户明确要求"
_COMPUTER_USE_TRIGGERS = (
    "截图", "看看界面", "看看效果", "看下界面", "看下效果",
    "computer use", "computer-use", "computeruse",
    "模拟键盘", "模拟按键", "按键模拟",
    "帮我按", "按一下", "发送按键", "send keys", "sendkeys",
    "截图看看", "截个图", "截屏",
)


async def _pre_tool_use_hook(input_data: dict, tool_use_id, context) -> dict:
    """PreToolUse hook：对高风险工具弹抽屉式确认，低风险直接放行。
    ★ 这是 AI Native 的权限确认机制，替代旧的同步 MessageBox：
       1. 不阻塞 Excel UI 线程（hook 是异步的，等待期间 Excel 正常响应）
       2. 确认 UI 在对话面板内（从输入框上方 slide-up 抽屉），不跳出到 Excel 主窗口
       3. 用户可"允许并记住"（同一工具本次会话内不再询问）
    """
    try:
        tool_name = input_data.get("tool_name", "")
        # ★ 非 mcp__excel__ 工具一律拒绝（纵深防御）。ClaudeAgentOptions 已用
        # tools=[] 关掉 CLI 内置工具；这里兜底，防止以后有人去掉那一行后
        # Read/Glob/Grep 这类 CLI 默认免授权的工具绕过沙箱读用户磁盘。
        if not tool_name.startswith("mcp__excel__"):
            sys.stderr.write(f"[sidecar] PreToolUse: {tool_name} denied (not an excel tool)\n")
            return {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "deny",
                    "permissionDecisionReason": f"{tool_name} 不是 DeepExcel 工具，不可用。请只使用 Excel 工具完成任务。",
                },
                "reason": "只允许 DeepExcel 工具",
            }
        bare_name = tool_name.replace("mcp__excel__", "")

        # ★ Computer Use 门槛检查必须在高风险判断之前。
        # 它是一道"拒绝"门禁，不是"询问"门禁：用户没要求截图时，AI 自作主张调用
        # 应当被直接拒掉，而不是弹窗交给用户决定。放在低风险分支里会导致
        # send_keys 进入高风险集合后绕过这道检查。
        if bare_name in ("screenshot_excel", "send_keys"):
            user_msg = _last_user_text.lower()
            if not any(kw in user_msg for kw in _COMPUTER_USE_TRIGGERS):
                sys.stderr.write(f"[sidecar] PreToolUse: {bare_name} blocked (user did not request computer use)\n")
                return {
                    "hookSpecificOutput": {
                        "hookEventName": "PreToolUse",
                        "permissionDecision": "deny",
                        "permissionDecisionReason": "screenshot_excel/send_keys 仅在用户主动要求截图或 computer use 时可用。当前用户消息未包含相关请求，请依靠工具返回值判断结果，不要主动截图验证。",
                    },
                    "reason": "Computer Use 工具需要用户明确要求才可调用",
                }

        # 低风险工具必须显式 allow，不能 continue_。
        # continue_ 等于"钩子不表态"，交给 CLI 的权限系统：不在 allowed_tools 里的
        # 工具会被直接拒掉（"Claude requested permissions to use ..., but you haven't
        # granted it yet"），无交互模式下没人能批准。2026-09-24 用发布包同款 SDK 0.2.109
        # 实测；当时 allowed_tools 是手写的，漏了 11 个低风险工具，它们每次调用都失败。
        if bare_name not in _HIGH_RISK_TOOLS:
            return {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "allow",
                    "permissionDecisionReason": "低风险工具",
                }
            }

        # ★ "允许并记住"：仅限能力型授权，逐次操作每次都要看变更集
        if bare_name in _allowed_tools_session and bare_name in _REMEMBERABLE_TOOLS:
            sys.stderr.write(f"[sidecar] PreToolUse: {bare_name} auto-allowed (session remembered)\n")
            return {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "allow",
                    "permissionDecisionReason": "用户已允许并记住",
                }
            }

        # 高风险工具：向 C#/前端请求权限确认
        sys.stderr.write(f"[sidecar] PreToolUse: {bare_name} requires permission, requesting...\n")
        tool_input = input_data.get("tool_input", {})
        await write_message(ui_events.envelope("status", text="等待你确认", tool=bare_name))
        decision = await request_permission(bare_name, tool_input)
        sys.stderr.write(f"[sidecar] PreToolUse: {bare_name} decision={decision}\n")

        if decision == "allow":
            # ★ 用户允许 → 记住，本次会话内不再询问此工具
            _allowed_tools_session.add(bare_name)
            return {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "allow",
                    "permissionDecisionReason": "用户已确认允许",
                }
            }
        else:
            return {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "deny",
                    "permissionDecisionReason": "用户拒绝执行此操作",
                },
                "reason": "用户拒绝了 " + bare_name + " 的执行",
            }
    except Exception as e:
        try:
            failed_tool = str(input_data.get("tool_name", "")).replace("mcp__excel__", "")
        except Exception:
            failed_tool = ""
        sys.stderr.write(f"[sidecar] PreToolUse hook error: {failed_tool}: {e}\n")
        sys.stderr.flush()
        # 高风险工具：确认流程坏了就不能当作用户已同意
        if failed_tool in _HIGH_RISK_TOOLS:
            return {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "deny",
                    "permissionDecisionReason": "确认流程出错，为安全起见未执行。请告诉用户重试。",
                },
                "reason": "权限确认出错",
            }
        return {"continue_": True}


def format_steer_context(messages: list) -> str:
    texts = [str(m.get("text") or "").strip() for m in messages]
    texts = [t for t in texts if t]
    if not texts:
        return ""
    body = "\n".join(f"- {t}" for t in texts)
    return (
        "<user-interjection>\n"
        "用户在你执行任务的过程中补充了下面的话。它比原来的指令更新：先据此调整接下来的步骤；"
        "如果它和你正在做的事冲突，停下来先回应用户。\n"
        f"{body}\n"
        "</user-interjection>"
    )


async def _post_tool_use_hook(input_data: dict, tool_use_id, context) -> dict:
    """任务进行中用户又发了消息：在这个工具结果之后交给模型（Claude Code 的排队消息）。

    以前任务进行中输入框是禁用的，用户发现方向不对只能按停止、再重新说一遍，
    已经做完的步骤和上下文都白费。"""
    try:
        pending = take_steer_messages()
        text = format_steer_context(pending)
        if not text:
            return {}
        await write_message(ui_events.envelope("steer_delivered", count=len(pending)))
        return {"hookSpecificOutput": {"hookEventName": "PostToolUse", "additionalContext": text}}
    except Exception as exc:
        sys.stderr.write(f"[sidecar] PostToolUse hook error: {exc}\n")
        return {}


# 以前这里有一个「纯文本回答」缓存（5 分钟、按用户原话做键）。工作簿一变，
# 「这张表有多少行」「A 列合计是多少」这类回答就是错的，而缓存照样原样返回，
# 连模型都没被问到。Claude Code 从不缓存回答，理由相同：答案取决于此刻的状态。
# 2026-09-25 删除。


DEFAULT_BASE_URL = "https://api.anthropic.com"
DEFAULT_MODEL = "claude-sonnet-4"


def build_env_config(cfg: dict, environ) -> tuple:
    """Resolve the SDK environment from the config the C# bridge sent.

    Two outbound modes, decided by the server and resolved on the C# side by
    SidecarRoutingResolver:

        byok   -- the user's own provider and API key (x-api-key header)
        hosted -- the DeepExcel proxy and a short-lived bearer token
                  (Authorization header)

    The rule that matters: exactly one credential is ever set. In hosted mode
    the user's provider key is NOT forwarded -- the proxy authenticates the
    request itself, so sending the key would expose it to one more party and
    bypass metering at the same time. In BYOK mode no proxy token exists, so
    there is nothing that could leak to a third-party provider.

    Returns (env_config, model).
    """
    if not cfg:
        # No config message: fall back to the ambient environment. This is the
        # development path and the 30-second-timeout path, never the normal one.
        return (
            {
                "ANTHROPIC_BASE_URL": environ.get("ANTHROPIC_BASE_URL", DEFAULT_BASE_URL),
                "ANTHROPIC_API_KEY": environ.get("ANTHROPIC_API_KEY", ""),
            },
            environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL),
        )

    env_config = {"ANTHROPIC_BASE_URL": cfg.get("base_url") or DEFAULT_BASE_URL}
    auth_token = cfg.get("auth_token") or ""
    if cfg.get("routing_mode") == "hosted" and auth_token:
        env_config["ANTHROPIC_AUTH_TOKEN"] = auth_token
    else:
        # Includes the case of hosted-without-a-token: falling back to the local
        # key would silently bill the user's own provider account. The C# side
        # already refuses to produce that combination; this is defence in depth.
        env_config["ANTHROPIC_API_KEY"] = cfg.get("api_key") or ""

    return env_config, (cfg.get("model") or DEFAULT_MODEL)


DEFAULT_MAX_TURNS = 20
MAX_TURNS_CEILING = 50


def resolve_max_turns(cfg) -> int:
    """设置里的 MaxTurns → SDK 的 max_turns。

    以前侧车写死 20，设置面板里的 MaxTurns 存下来却从没生效。
    上限与 C# 端 HandleSaveModelConfig 的 50 一致（防 AI 无限循环烧 API 费用）；
    缺省、非数字、越界都回落而不是报错——这个值错了不该让整个会话起不来。
    """
    raw = (cfg or {}).get("max_turns")
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return DEFAULT_MAX_TURNS
    if value < 1:
        return DEFAULT_MAX_TURNS
    return min(value, MAX_TURNS_CEILING)


def stale_env_keys(env_config: dict) -> list:
    """Environment variables that must be removed before creating the client.

    Model defaults inherited from ~/.claude/settings.json would override the
    model we were told to use. The credential we are NOT using must also go:
    when both are present the SDK's choice depends on implementation details,
    which could send the user's provider key to the DeepExcel proxy or the proxy
    token to a third-party provider. Both hand a credential to the wrong party.
    """
    stale = [
        "ANTHROPIC_MODEL",
        "ANTHROPIC_DEFAULT_HAIKU_MODEL",
        "ANTHROPIC_DEFAULT_SONNET_MODEL",
        "ANTHROPIC_DEFAULT_OPUS_MODEL",
    ]
    if "ANTHROPIC_AUTH_TOKEN" in env_config:
        stale.append("ANTHROPIC_API_KEY")
    else:
        stale.append("ANTHROPIC_AUTH_TOKEN")
    # 压缩窗口只能由我们按当前模型设置；从外部环境继承来的值属于别的模型
    for key in ("CLAUDE_CODE_AUTO_COMPACT_WINDOW", "CLAUDE_AUTOCOMPACT_PCT_OVERRIDE"):
        if key not in env_config:
            stale.append(key)
    return stale


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


def _build_excel_context_lite(context: dict) -> str:
    """从 context 中提取轻量级 Excel 上下文（地址、sheet 名、行列数），
    拼成简短文本注入用户消息。这样 AI 不需要额外调 read_workbook/read_selection。
    """
    if not isinstance(context, dict):
        return ""
    lines = []
    wb = context.get("workbook")
    if isinstance(wb, dict):
        sheets = wb.get("sheets")
        active = wb.get("activeSheet")
        if active:
            lines.append(f"[当前工作表] {active}")
        if sheets and isinstance(sheets, list) and len(sheets) > 1:
            lines.append(f"[所有工作表] {', '.join(str(s) for s in sheets[:10])}")
    sel = context.get("selection")
    if isinstance(sel, dict):
        addr = sel.get("address")
        sheet = sel.get("sheet")
        rows = sel.get("rowCount")
        cols = sel.get("columnCount")
        if addr:
            parts = [f"[选中区域] {addr}"]
            if rows and cols:
                parts.append(f"({rows}行×{cols}列)")
            lines.append(" ".join(parts))
    wb_name = context.get("workbookName")
    if wb_name:
        lines.append(f"[工作簿] {wb_name}")

    header = (" ".join(lines) + "\n") if lines else ""

    # 两轮之间用户手动改过的区域：之前读到的这些内容已经过期
    edits = context.get("userEdits")
    if isinstance(edits, list) and edits:
        shown = [str(e) for e in edits[:20] if e]
        if shown:
            header += f"[用户改动] 上一轮之后用户手动改了 {'、'.join(shown)}，要用这些内容先重新读取\n"

    notices = context.get("hostNotices")
    if isinstance(notices, list):
        for notice in notices[:5]:
            if isinstance(notice, str) and notice.strip():
                header += f"[提示] {notice.strip()}\n"

    # Workbook structure summary: column types, value ranges, blank positions,
    # named ranges and cross-sheet links. Roughly 300-600 tokens per sheet, and
    # it replaces the read_range round-trips the model would otherwise make just
    # to work out what the sheet contains -- which on a large sheet it cannot do
    # completely anyway.
    structure = context.get("structure")
    if isinstance(structure, str) and structure.strip():
        header += structure.rstrip() + "\n"

    return header


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

# ★ 本轮是否有工具调用，以及收集的完整文本（诊断日志用）
_tool_calls_in_turn = 0
_collected_text = ""

# ★ 本轮工具调用账本：tool_start / tool_end 按 tool_use_id 配对，终态行据此汇总
_run = ui_events.RunTracker()
# ★ 正在生成参数的工具调用（代码边写边显示）
_gen = ui_events.ToolGenTracker()

WATCHDOG_TICK_SECONDS = 5.0


async def status_watchdog(tick: float | None = None) -> None:
    """很久没动静时发状态行：在等模型、还是在等 Excel 执行（可能被弹窗挡住）。

    同一句话不重复发；恢复活动后发一个空状态让面板清掉。"""
    last_text = None
    while True:
        await anyio.sleep(WATCHDOG_TICK_SECONDS if tick is None else tick)
        if _message_buffer.get("awaiting_user"):
            # 在等用户确认或回答：不是卡住，也不该把「等待你确认」盖掉
            _run.touch()
            continue
        text = ui_events.watchdog_status(_run.idle_seconds(), _run.running_tool(), ui_events.PROGRESS.current())
        if text != last_text:
            await write_message(ui_events.envelope("status", text=text or ""))
            last_text = text

# ★ 压缩检测兜底：记录上一轮 context usage percentage。CLI 正常会发
# system/compact_boundary；没收到时 percentage 骤降也说明压缩发生了。
_prev_context_percentage = None


def _usage_tokens(usage) -> tuple:
    if isinstance(usage, dict):
        return usage.get("input_tokens", 0) or 0, usage.get("output_tokens", 0) or 0
    if usage:
        return getattr(usage, "input_tokens", 0) or 0, getattr(usage, "output_tokens", 0) or 0
    return 0, 0


async def _emit_tool_start(block) -> None:
    global _tool_calls_in_turn
    _tool_calls_in_turn += 1
    name = ui_events.bare_tool_name(block.name)
    args = block.input if isinstance(block.input, dict) else {}
    _run.start(block.id, name)
    # 旧消息：C# / WPS 用它记对话历史和任务轨迹
    await write_message({"type": "tool_use", "tool": block.name, "args": args})
    await write_message(ui_events.envelope(
        "tool_start", id=block.id, name=name, args=ui_events.summarize_args(args)))


async def _emit_tool_end(block) -> None:
    name, duration_ms = _run.end(block.tool_use_id)
    result = ui_events.parse_tool_result(block.content, block.is_error)
    if not result["ok"]:
        _run.failed_calls += 1
    await write_message(ui_events.envelope(
        "tool_end", id=block.tool_use_id, name=name, duration_ms=duration_ms, **result))


async def _close_unfinished_tools(code: str, message: str) -> None:
    """中断或出错时，给还没收到结果的调用补 tool_end，面板上的那一行才会停止转圈。"""
    for tool_use_id, name in _run.unfinished():
        _run.failed_calls += 1
        await write_message(ui_events.envelope(
            "tool_end", id=tool_use_id, name=name, ok=False,
            error={"code": code, "message": message}))


async def _emit_run_summary(outcome: str, in_tok: int = 0, out_tok: int = 0, num_turns=None) -> None:
    if _run.summarized:
        return
    _run.summarized = True
    await write_message(ui_events.envelope(
        "run_summary", outcome=outcome, tool_calls=_run.tool_calls,
        failed_calls=_run.failed_calls, duration_ms=_run.elapsed_ms(),
        num_turns=num_turns, input_tokens=in_tok, output_tokens=out_tok))


async def handle_sdk_message(response, client=None):
    """处理 ClaudeSDKClient.receive_response() 产生的流式消息"""
    global _had_partial_text, _collected_text, _prev_context_percentage
    _run.touch()
    if isinstance(response, StreamEvent):
        evt = response.event if isinstance(response.event, dict) else {}
        etype = evt.get("type")
        if etype == "content_block_start":
            block = evt.get("content_block") or {}
            if block.get("type") == "tool_use":
                _gen.start(evt.get("index", -1), block.get("id", ""), block.get("name", ""))
        elif etype == "content_block_stop":
            _gen.stop(evt.get("index", -1))
        elif etype == "content_block_delta":
            delta = evt.get("delta", {})
            if isinstance(delta, dict) and delta.get("type") == "text_delta":
                text = delta.get("text", "")
                if text:
                    _had_partial_text = True
                    _collected_text += text
                    await write_message({"type": "stream_delta", "text": text})
            elif isinstance(delta, dict) and delta.get("type") == "input_json_delta":
                gen_event = _gen.feed(evt.get("index", -1), delta.get("partial_json", ""))
                if gen_event:
                    await write_message(gen_event)
    elif isinstance(response, AssistantMessage):
        api_error = getattr(response, "error", None)
        if api_error:
            # CLI 把 API 错误包装成一条助手消息（文本是英文原始报错）。原样贴给用户
            # 看不懂也不知道下一步；换成分类后的错误卡片。
            raw = " ".join(b.text for b in response.content if isinstance(b, TextBlock))
            await write_message(ui_events.envelope(
                "error", **ui_events.classify_error(f"{api_error} {raw}")))
            _had_partial_text = False
            return
        for block in response.content:
            if isinstance(block, TextBlock):
                if not _had_partial_text:
                    _collected_text += block.text
                    await write_message({"type": "stream_delta", "text": block.text})
            elif isinstance(block, ToolUseBlock):
                await _emit_tool_start(block)
        _had_partial_text = False
    elif isinstance(response, UserMessage):
        content = response.content if isinstance(response.content, list) else []
        for block in content:
            if isinstance(block, ToolResultBlock):
                await _emit_tool_end(block)
    elif isinstance(response, SystemMessage):
        if response.subtype == "compact_boundary":
            meta = (response.data or {}).get("compact_metadata") or {}
            _run.compacted = True
            await write_message(ui_events.envelope(
                "compaction", trigger=meta.get("trigger") or "auto", pre_tokens=meta.get("pre_tokens")))
    elif isinstance(response, ResultMessage):
        in_tok, out_tok = _usage_tokens(getattr(response, "usage", None))
        await _close_unfinished_tools("no_result", "没有收到这个工具的结果")
        outcome = {
            "success": "success",
            "error_max_turns": "max_turns",
        }.get(getattr(response, "subtype", ""), "error" if getattr(response, "is_error", False) else "success")
        if _run.interrupted:
            outcome = "interrupted"
        await _emit_run_summary(outcome, in_tok, out_tok, getattr(response, "num_turns", None))
        await write_message({
            "type": "stream_end",
            "input_tokens": in_tok,
            "output_tokens": out_tok,
        })
        # ★ 压缩检测兜底：本轮没收到 compact_boundary，但 percentage 骤降（超过 40%）。
        if client is not None:
            try:
                context_usage = await client.get_context_usage()
                # ContextUsageResponse 可能是 dict 或对象，统一取 percentage
                if isinstance(context_usage, dict):
                    curr_pct = context_usage.get("percentage", 0)
                else:
                    curr_pct = getattr(context_usage, "percentage", 0)
                if (not _run.compacted and _prev_context_percentage is not None
                        and curr_pct < _prev_context_percentage * 0.6):
                    await write_message(ui_events.envelope(
                        "compaction", trigger="detected",
                        prev_pct=_prev_context_percentage, curr_pct=curr_pct))
                _prev_context_percentage = curr_pct
            except Exception as ctx_e:
                sys.stderr.write(f"[sidecar] get_context_usage failed: {type(ctx_e).__name__}: {ctx_e}\n")
                sys.stderr.flush()


# 按停止后等 CLI 自己收尾（发出 ResultMessage）的时长。超时就硬切，并在下一轮开始前
# 把残留的输出读干净——否则下一轮的 receive_response 会先读到上一轮剩下的消息，
# 在旧的 ResultMessage 处提前结束，用户看到的是上一个问题的回答。
INTERRUPT_GRACE_SECONDS = 15.0
DRAIN_TIMEOUT_SECONDS = 20.0

# 上一轮被硬切，CLI 的输出流里可能还有它的残留
_needs_drain = False


async def drain_stale_responses(client, timeout: float | None = None) -> bool:
    """读掉上一轮残留的消息，直到它的 ResultMessage。返回是否读到了结尾。"""
    with anyio.move_on_after(DRAIN_TIMEOUT_SECONDS if timeout is None else timeout):
        async for response in client.receive_response():
            if isinstance(response, ResultMessage):
                return True
    return False


async def interrupt_and_wait(client, cancel_scope, grace: float | None = None) -> None:
    """Claude Code 的 Esc：让 CLI 停下当前回合，而不只是不再读它的输出。

    interrupt 之后 CLI 会补发工具结果和 ResultMessage，主循环照常处理（面板上的步骤
    正常收尾）；等不到就硬切，并标记下一轮开始前先排空。"""
    global _needs_drain
    _run.interrupted = True
    await write_message(ui_events.envelope("status", text="正在停止…"))
    try:
        with anyio.fail_after(5):
            await client.interrupt()
    except Exception as exc:
        sys.stderr.write(f"[sidecar] interrupt failed: {type(exc).__name__}: {exc}\n")
        sys.stderr.flush()
    await anyio.sleep(INTERRUPT_GRACE_SECONDS if grace is None else grace)
    sys.stderr.write("[sidecar] interrupt grace expired, cutting the stream\n")
    sys.stderr.flush()
    _needs_drain = True
    cancel_scope.cancel()


async def defer_leftover_steer() -> None:
    """本轮结束时还没来得及注入的插话，作为下一条普通消息处理，不会丢。"""
    leftover = take_steer_messages()
    texts = [str(m.get("text") or "").strip() for m in leftover]
    texts = [t for t in texts if t]
    if not texts:
        return
    await write_message(ui_events.envelope("steer_deferred", count=len(texts)))
    _message_buffer["user_message"].put_nowait({
        "type": "user_message",
        "text": "\n".join(texts),
        "context": leftover[-1].get("context") or {},
    })


async def run_agent_loop(client, supports_vision: bool = True, model: str = "", base_url: str = ""):
    """主循环：从 user_message queue 取消息，发给 SDK 处理"""
    global _needs_drain
    while True:
        msg = await _message_buffer["user_message"].get()
        if msg is None:
            return

        # 重置 cancel 标志（每次新对话开始前）
        if _message_buffer["cancel"].is_set():
            _message_buffer["cancel"].clear()

        if _needs_drain:
            drained = await drain_stale_responses(client)
            sys.stderr.write(f"[sidecar] drained stale responses before new turn: reached_end={drained}\n")
            sys.stderr.flush()
            _needs_drain = False

        # ★ 重置流式标志 + 本轮工具账本
        global _had_partial_text, _tool_calls_in_turn, _collected_text, _run, _gen
        _had_partial_text = False
        _tool_calls_in_turn = 0
        _collected_text = ""
        _run = ui_events.RunTracker()
        _gen = ui_events.ToolGenTracker()

        user_text = msg.get("text", "")
        context = msg.get("context") or {}
        attachments = context.get("attachments") if isinstance(context, dict) else None

        # ★ 记录最近一条用户消息，供 PreToolUse hook 的 Computer Use 门槛检查使用
        global _last_user_text
        _last_user_text = user_text or ""

        sys.stderr.write(f"[sidecar] run_agent_loop: processing user message (len={len(user_text)}), supports_vision={supports_vision}\n")
        sys.stderr.flush()

        history = _message_buffer.get("restore_history")
        has_history = history and isinstance(history, list) and len(history) > 0
        has_attachments = attachments and isinstance(attachments, list) and len(attachments) > 0

        # ★ 附件上下文注入：如果有附件，把附件信息拼到用户消息前面
        final_text = user_text
        if has_attachments:
            att_context = _build_attachment_context(attachments, supports_vision=supports_vision)
            if att_context:
                final_text = att_context + "\n\n用户问题：\n" + user_text
                sys.stderr.write(f"[sidecar] attachment context injected ({len(attachments)} attachments)\n")
                sys.stderr.flush()

        # ★ 轻量级 Excel 上下文注入：把选中区域地址、sheet 名等信息拼到用户消息前
        # 这样 AI 不需要额外调 read_workbook/read_selection 就能知道基本信息
        excel_ctx = _build_excel_context_lite(context)
        if excel_ctx:
            final_text = excel_ctx + "\n" + final_text
            sys.stderr.write(f"[sidecar] excel context lite injected\n")
            sys.stderr.flush()

        # ★ 历史上下文注入：如果 C# 发了 restore_history，把历史对话摘要拼到首条用户消息前。
        if has_history:
            summary = _build_history_summary(history)
            if summary:
                final_text = summary + "\n\n用户问题：\n" + final_text
                sys.stderr.write(f"[sidecar] history context injected ({len(history)} messages)\n")
                sys.stderr.flush()
            _message_buffer["restore_history"] = None

        _message_buffer["turn_active"] = True
        try:
            direct_blocks = []
            if supports_vision and has_attachments:
                image_blocks = _collect_image_blocks(attachments)
                pdf_blocks = _collect_pdf_blocks(attachments)
                direct_blocks = image_blocks + pdf_blocks

            if direct_blocks:
                content_blocks = direct_blocks + [{"type": "text", "text": final_text}]

                async def _msg_gen():
                    yield {
                        "type": "user",
                        "message": {"role": "user", "content": content_blocks},
                        "parent_tool_use_id": None,
                        "session_id": "default",
                    }

                await client.query(_msg_gen())
                sys.stderr.write(f"[sidecar] client.query() done (with images + pdfs), receiving response...\n")
            else:
                await client.query(final_text)
                sys.stderr.write("[sidecar] client.query() done, receiving response...\n")
            sys.stderr.flush()

            async with anyio.create_task_group() as inner_tg:
                async def _cancel_watchdog():
                    while not _message_buffer["cancel"].is_set():
                        await anyio.sleep(0.1)
                    sys.stderr.write("[sidecar] cancel detected by watchdog, interrupting the CLI\n")
                    sys.stderr.flush()
                    await interrupt_and_wait(client, inner_tg.cancel_scope)

                inner_tg.start_soon(_cancel_watchdog)
                inner_tg.start_soon(status_watchdog)

                try:
                    async for response in client.receive_response():
                        await handle_sdk_message(response, client)
                except Exception as inner_e:
                    sys.stderr.write(f"[sidecar] receive_response exception: {type(inner_e).__name__}: {inner_e}\n")
                    sys.stderr.flush()
                    raise

                inner_tg.cancel_scope.cancel()

            sys.stderr.write(f"[sidecar] receive_response completed, tool_calls={_tool_calls_in_turn}, text_len={len(_collected_text)}\n")
            sys.stderr.flush()

        except Exception as e:
            sys.stderr.write(f"[sidecar] run_agent_loop exception: {type(e).__name__}: {e}\n")
            sys.stderr.flush()
            if isinstance(e, (KeyboardInterrupt, SystemExit)):
                raise
            await _close_unfinished_tools("aborted", "任务出错，这一步没有完成")
            await write_message(ui_events.envelope("error", **ui_events.classify_error(e)))
            await _emit_run_summary("error")
            await write_message({"type": "stream_end", "input_tokens": 0, "output_tokens": 0})

        _message_buffer["turn_active"] = False
        if _message_buffer["cancel"].is_set() or _run.interrupted:
            # CLI 正常收尾时 ResultMessage 已经发过终态行和 stream_end；只有被硬切时才补
            if not _run.summarized:
                sys.stderr.write("[sidecar] turn was cut, sending stream_end to unblock UI\n")
                sys.stderr.flush()
                await _close_unfinished_tools("interrupted", "已中断")
                await _emit_run_summary("interrupted")
                await write_message({"type": "stream_end", "input_tokens": 0, "output_tokens": 0})
            _message_buffer["cancel"].clear()
            # 用户按了停止：插话一起作废，不自动开始下一轮
            take_steer_messages()
        else:
            await defer_leftover_steer()


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

        # Excel sends config over stdin. WPS has no C# bridge, so it reuses the
        # same current-user config and DPAPI credential written by Excel.
        cfg = _load_wps_local_config() if os.environ.get("DEEPEXCEL_HOST") == "wps" else None
        if cfg is None and os.environ.get("DEEPEXCEL_HOST") != "wps":
            try:
                with anyio.fail_after(30.0):
                    cfg = await _message_buffer["config"].get()
                sys.stderr.write(
                    f"[sidecar] config received: base_url={cfg.get('base_url')}, "
                    f"model={cfg.get('model')}, mode={cfg.get('routing_mode', 'byok')}, "
                    f"hasKey={bool(cfg.get('api_key'))}, hasToken={bool(cfg.get('auth_token'))}\n")
                sys.stderr.flush()
            except TimeoutError:
                sys.stderr.write("[sidecar] WARNING: config timeout (30s), falling back to env vars\n")
                sys.stderr.flush()

        routing_mode = (cfg or {}).get("routing_mode", "byok")
        env_config, model = build_env_config(cfg, os.environ)
        # ★ 模型真实的上下文窗口：CLI 对不认识的模型一律按 200K、在 167K 时压缩（见 model_windows.py）
        model, window_env = apply_context_window(model, (cfg or {}).get("context_window"))
        env_config.update(window_env)
        sys.stderr.write(f"[sidecar] cli model={model}, window_env={window_env}\n")
        max_turns = resolve_max_turns(cfg)
        sys.stderr.write(f"[sidecar] max_turns={max_turns} (configured={(cfg or {}).get('max_turns')})\n")

        # ★ 关键修复：显式同步到 os.environ，覆盖 ~/.claude/settings.json 的 env 配置。
        # SDK 在 ClaudeSDKClient 创建时会读 settings.json 的 env 字段并 merge 进 process env，
        # 优先级高于 ClaudeAgentOptions.env。如果不显式 set os.environ，settings.json 中的
        # ANTHROPIC_BASE_URL（如 deepseek）会覆盖我们传的 stepfun base_url，导致请求发到错误端点。
        # 现象：sidecar 日志显示 base_url=stepfun，但实际请求打到 deepseek，model=step-3.7-flash
        # 被 deepseek 拒绝（400 "supported: deepseek-v4-pro/flash, but you passed step-3.7-flash"）。
        for k, v in env_config.items():
            os.environ[k] = v
        for k in stale_env_keys(env_config):
            os.environ.pop(k, None)

        # ★ 诊断日志：打印 os.environ 实际值（脱敏 api_key），验证 settings.json 是否被覆盖
        _diag_env_base = os.environ.get("ANTHROPIC_BASE_URL", "<unset>")
        _diag_env_model = os.environ.get("ANTHROPIC_MODEL", "<unset>")
        # ANTHROPIC_AUTH_TOKEN 只报告存在与否。以前它总是被 pop 掉，恒为 <unset>，
        # 打印全值没有后果；托管模式下它是一个活的 bearer 令牌，而这些日志会被
        # DeepExcel.Repair.exe 的诊断包收集并发给支持人员。
        _diag_has_token = "ANTHROPIC_AUTH_TOKEN" in os.environ
        _diag_env_sonnet = os.environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "<unset>")
        _diag_has_key = "ANTHROPIC_API_KEY" in os.environ
        sys.stderr.write(f"[sidecar][diag] os.environ after override: "
                         f"ANTHROPIC_BASE_URL={_diag_env_base}, "
                         f"ANTHROPIC_MODEL={_diag_env_model}, "
                         f"routing_mode={routing_mode}, "
                         f"ANTHROPIC_AUTH_TOKEN_present={_diag_has_token}, "
                         f"ANTHROPIC_DEFAULT_SONNET_MODEL={_diag_env_sonnet}, "
                         f"ANTHROPIC_API_KEY_present={_diag_has_key}\n")
        sys.stderr.flush()

        # ★ KV Cache 优化：保持原始模型名称（model 参数会被 SDK 直接发送给 API，
        # 修改它会导致 DeepSeek/Kimi 等提供商拒绝请求）。
        # 真正的优化在于扩展系统提示词长度，系统提示词占总 token 的比例越大，
        # KV Cache 命中率越高。当前 SYSTEM_PROMPT 约 1000 tokens，建议扩展到 5000+。

        # 注册工具到 MCP server。按宿主注册：WPS 专用的 execute_jsa 不出现在 Excel 会话里
        host = "wps" if os.environ.get("DEEPEXCEL_HOST") == "wps" else "excel"
        host_tools = register_all_tools(host)
        # 分头摸底的子会话用同一个模型、同一个出口
        explorer.configure(env_config, model, host)
        server = create_sdk_mcp_server(name="excel", tools=host_tools)
        system_prompt = SYSTEM_PROMPT + host_tool_note(host, [t.name for t in host_tools])

        options = ClaudeAgentOptions(
            model=model,
            # ★ 关掉 CLI 全部内置工具（→ --tools ""），只留下面的 MCP excel 工具。
            # 不设时模型能看到 Bash/Read/Write/Edit/Glob/Grep/WebFetch/Task 等 25 个
            # 内置工具（2026-09-23 实测 init 消息），完全绕过 CodeSandbox 与权限抽屉；
            # 其中 Task* 被调用后无人处理，面板会一直停在 "..."。
            # allowed_tools 只是"免确认"名单，不限制可见工具，所以必须用 tools。
            tools=[],
            mcp_servers={"excel": server},
            # 免确认名单从注册表生成，不再手写（手写版本曾漏掉 20 个工具）。
            # 真正的放行/确认由 PreToolUse 钩子决定，这里只是兜底。
            allowed_tools=[f"mcp__excel__{t.name}" for t in host_tools],
            system_prompt=system_prompt,
            max_turns=max_turns,  # ★ 来自设置 MaxTurns，见 resolve_max_turns
            env=env_config,  # ★ DeepSeek 配置必须在这里传
            # ★ 关键修复：setting_sources=[] 禁用 SDK 读取 ~/.claude/settings.json 等
            # 文件系统 settings。否则 SDK CLI 子进程会读 settings.json 的 env 字段
            # 并 merge 到自己的 process env（优先级高于父进程传来的 env），
            # 导致我们传的 stepfun base_url/key 被 settings.json 里的 deepseek 配置覆盖。
            # 现象：os.environ 显示 stepfun，但实际请求打到 deepseek（400 错误）。
            setting_sources=[],
            # ★ 禁用 thinking：DeepSeek anthropic 兼容端点不返回 thinking block 的 signature 字段，
            # SDK 在 message_parser.py:104 硬编码 block["signature"] 会抛 MessageParseError，
            # 导致 tool_use 后第二轮 API 响应解析崩溃，最终文本永不返回。
            # 参考：https://github.com/anthropics/claude-agent-sdk-python/issues/949
            thinking={"type": "disabled"},
            # ★ 启用 token 级流式：receive_response() 会产出 StreamEvent（含 content_block_delta），
            # 让 UI 逐 token 显示文本，而不是整段一次性出现
            include_partial_messages=True,
            # ★ PreToolUse hook：AI Native 权限确认机制
            # 对高风险工具（execute_vba/execute_python/rollback/clean_data/remove_duplicates）
            # 在面板内抽屉式确认，替代旧的同步 MessageBox（阻塞 UI 线程导致 Excel 崩溃）
            hooks={
                "PreToolUse": [HookMatcher(matcher=None, hooks=[_pre_tool_use_hook])],
                # ★ 插话注入：任务进行中用户发来的消息，在下一个工具结果后交给模型
                "PostToolUse": [HookMatcher(matcher=None, hooks=[_post_tool_use_hook])],
            },
        )

        # ★ 检测模型是否支持 vision（image/document block）
        base_url_for_check = env_config.get("ANTHROPIC_BASE_URL", "")
        vision_ok = _supports_vision(base_url_for_check, model)
        sys.stderr.write(f"[sidecar] creating ClaudeSDKClient, model={model}, base_url={base_url_for_check}, supports_vision={vision_ok}\n")
        sys.stderr.flush()
        async with ClaudeSDKClient(options=options) as client:
            # ★ 诊断日志：SDK 创建 client 后再次打印 os.environ，确认 SDK 是否把
            # settings.json 的 env 重新 merge 回 os.environ（如果是，run_agent_loop 期间
            # 这些值会被 SDK 实际使用，覆盖我们的 env_config）
            _post_base = os.environ.get("ANTHROPIC_BASE_URL", "<unset>")
            _post_model = os.environ.get("ANTHROPIC_MODEL", "<unset>")
            # Presence only. This is a live bearer token in hosted mode, and
            # these logs are collected into the support bundle.
            _post_has_token = "ANTHROPIC_AUTH_TOKEN" in os.environ
            _post_sonnet = os.environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "<unset>")
            sys.stderr.write(f"[sidecar][diag] os.environ AFTER client created: "
                             f"ANTHROPIC_BASE_URL={_post_base}, "
                             f"ANTHROPIC_MODEL={_post_model}, "
                             f"ANTHROPIC_AUTH_TOKEN_present={_post_has_token}, "
                             f"ANTHROPIC_DEFAULT_SONNET_MODEL={_post_sonnet}\n")
            sys.stderr.flush()
            sys.stderr.write("[sidecar] ClaudeSDKClient connected, entering agent loop\n")
            sys.stderr.flush()
            # 在 task_group 里跑主循环，reader 继续在后台跑
            await run_agent_loop(client, supports_vision=vision_ok, model=model, base_url=base_url_for_check)

        # 主循环退出后，取消 reader
        tg.cancel_scope.cancel()


if __name__ == "__main__":
    # 启动父进程心跳检测线程（daemon，主进程退出时自动结束）
    t = threading.Thread(target=_parent_watchdog, daemon=True)
    t.start()
    anyio.run(main)
