# src/DeepExcel.Sidecar/ui_events.py
"""面板事件信封 ui_event（协议见 docs/ui-event-protocol.md）。

以前侧车只发 stream_delta / tool_use / stream_end 三种消息：面板只知道「调用了
哪个工具」，不知道它什么时候结束、成没成功、错在哪，于是只能显示一串工具名。
Claude Code 的做法是每个工具调用都有开始和结束两行（⏺ 读取 A1:D20 → ⎿ 20 行），
失败时直接写出原因和下一步。这里给出同样的信息：

    {"type": "ui_event", "event": {"v": 1, "kind": ..., ...}}

kind:
    tool_start   {id, name, args}                  模型发出一次工具调用
    tool_end     {id, name, ok, duration_ms, summary?, error?, check?, checkpoint_id?, changes?}
    status       {text}                            当前在做什么（思考中、等待确认…）
    compaction   {trigger, pre_tokens?, prev_pct?, curr_pct?}
    error        {code, message, hint, retryable}  整轮失败
    run_summary  {outcome, tool_calls, failed_calls, duration_ms, num_turns,
                  input_tokens, output_tokens}
    steer_delivered {count}                        任务中插话已交给模型
    steer_deferred  {count}                        本轮没来得及注入，作为下一条消息处理

宿主（C# / WPS）原样转发，不解释内容；只有面板渲染它。
"""

from __future__ import annotations

import json
import time
from typing import Any

PROTOCOL_VERSION = 1

# 参数里的长字符串（VBA 代码、整段 JSON）只截取开头给面板显示；完整参数已经
# 在工具调用本身里，面板不需要第二份。
_MAX_ARG_STRING = 4000
_MAX_ARG_ITEMS = 20


def envelope(kind: str, **fields: Any) -> dict:
    event = {"v": PROTOCOL_VERSION, "kind": kind, "ts": int(time.time() * 1000)}
    event.update({k: v for k, v in fields.items() if v is not None})
    return {"type": "ui_event", "event": event}


def bare_tool_name(name: str) -> str:
    return (name or "").replace("mcp__excel__", "")


def summarize_args(value: Any, depth: int = 0) -> Any:
    """面板显示用的参数副本：长字符串截断，二维数组只留形状和前几行。"""
    if isinstance(value, str):
        if len(value) > _MAX_ARG_STRING:
            return value[:_MAX_ARG_STRING] + f"…（共 {len(value)} 字符）"
        return value
    if isinstance(value, dict):
        if depth > 3:
            return "{…}"
        return {str(k): summarize_args(v, depth + 1) for k, v in value.items()}
    if isinstance(value, list):
        if depth > 3:
            return "[…]"
        if value and all(isinstance(row, list) for row in value):
            cols = max((len(row) for row in value), default=0)
            return {
                "__shape": [len(value), cols],
                "head": [summarize_args(row, depth + 1) for row in value[:3]],
            }
        items = [summarize_args(v, depth + 1) for v in value[:_MAX_ARG_ITEMS]]
        if len(value) > _MAX_ARG_ITEMS:
            items.append(f"…（共 {len(value)} 项）")
        return items
    return value


def _result_text(content: Any) -> str:
    if content is None:
        return ""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, dict) and block.get("type") == "text":
                parts.append(str(block.get("text", "")))
        return "\n".join(parts)
    return str(content)


def parse_tool_result(content: Any, is_error: bool | None) -> dict:
    """把 ToolResultBlock 的内容变成 {ok, summary?, error?}。

    我们自己的工具返回 C# ToolResult 的 JSON（success/data/error/suggestion）；
    被 PreToolUse 拒绝或 CLI 自己报错时是一段纯文本。两种都要能读。
    """
    text = _result_text(content)
    payload = None
    try:
        payload = json.loads(text) if text else None
    except (ValueError, TypeError):
        payload = None

    if isinstance(payload, dict) and ("success" in payload or "error" in payload):
        ok = payload.get("success") is not False and not is_error
        out: dict = {"ok": ok}
        if ok:
            summary = summarize_result_data(payload.get("data"))
            if summary:
                out["summary"] = summary
            check = summarize_verification(payload.get("verification"))
            if check:
                out["check"] = check
            checkpoint = payload.get("checkpoint_id")
            if isinstance(checkpoint, str) and checkpoint:
                out["checkpoint_id"] = checkpoint
            changes = summarize_changes(payload.get("changes"))
            if changes:
                out["changes"] = changes
        else:
            out["error"] = {
                "code": str(payload.get("error_code") or "tool_failed"),
                "message": str(payload.get("error") or "工具执行失败"),
                "hint": str(payload.get("suggestion") or "") or None,
            }
        return out

    if is_error:
        message = text.strip() or "工具执行失败"
        lowered = message.lower()
        # 按停止后 CLI 给正在执行的调用回的是英文原文
        if "doesn't want to proceed" in lowered or "interrupted" in lowered or "用户已中断" in message:
            return {"ok": False, "error": {"code": "interrupted", "message": "已中断", "hint": None}}
        code = "denied" if ("拒绝" in message or "denied" in lowered) else "tool_failed"
        return {"ok": False, "error": {"code": code, "message": message[:500], "hint": None}}
    return {"ok": True}


def summarize_verification(verification: Any) -> dict | None:
    """C# 写后体检的结论 → {ok, summary}。面板只在没通过时显示；通过的一句话留给模型。"""
    if not isinstance(verification, dict):
        return None
    summary = verification.get("summary")
    if not isinstance(summary, str) or not summary.strip():
        return None
    return {"ok": verification.get("ok") is not False, "summary": summary.strip()[:300]}


def summarize_changes(changes: Any) -> dict | None:
    """C# 内联 diff → {changed, sheet?, samples: [{address, before, after}]}；没改动就不给。"""
    if not isinstance(changes, dict):
        return None
    changed = changes.get("changed")
    if not isinstance(changed, int) or changed <= 0:
        return None
    samples = []
    for item in changes.get("samples") or []:
        if isinstance(item, dict) and item.get("address"):
            samples.append({
                "address": str(item["address"]),
                "before": str(item.get("before") or ""),
                "after": str(item.get("after") or ""),
            })
    out = {"changed": changed, "samples": samples[:20]}
    if isinstance(changes.get("sheet"), str) and changes["sheet"]:
        out["sheet"] = changes["sheet"]
    return out


def summarize_result_data(data: Any) -> str | None:
    """结果的一句话概括（「20 行 × 4 列」「已写入 12 个单元格」），拿不准就不说。"""
    if isinstance(data, dict):
        for key in ("message", "summary"):
            if isinstance(data.get(key), str) and data[key].strip():
                return data[key].strip()[:200]
        values = data.get("values")
        if isinstance(values, list) and values and isinstance(values[0], list):
            return f"{len(values)} 行 × {max(len(r) for r in values)} 列"
        for key, unit in (("cells_written", "个单元格"), ("rows_affected", "行"),
                          ("affected", "处"), ("count", "项")):
            if isinstance(data.get(key), int):
                return f"{data[key]} {unit}"
    if isinstance(data, list):
        return f"{len(data)} 项"
    return None


# 整轮失败的分类。面板按 code 给出下一步，而不是把英文异常原样贴给用户。
_ERROR_RULES = [
    # (code, 匹配片段（小写）, 中文说明, 下一步, 可重试)
    ("auth", ("401", "authentication", "invalid x-api-key", "invalid api key", "unauthorized"),
     "模型服务拒绝了凭据。", "到「模型设置」检查 API Key 是否正确、是否过期。", False),
    ("quota", ("402", "quota_exhausted", "insufficient", "余额", "billing"),
     "额度已用完或账户欠费。", "充值或升级套餐后再试。", False),
    ("task_limit", ("task_call_limit",),
     "这个任务的模型调用次数达到上限。", "把任务拆小一点，或在新对话里继续。", False),
    ("rate_limit", ("429", "rate limit", "rate_limit", "overloaded", "529"),
     "模型服务繁忙，请求被限流。", "稍等几十秒再发送一次。", True),
    ("context_too_long", ("prompt is too long", "context length", "maximum context", "too many tokens"),
     "对话太长，超出了模型的上下文。", "新建对话，或让我先总结再继续。", False),
    ("model_not_found", ("model_not_found", "does not exist", "unknown model", "not supported model", "supported:"),
     "当前模型名不被服务商接受。", "到「模型设置」换一个模型。", False),
    ("timeout", ("timeout", "timed out"),
     "等待模型响应超时。", "检查网络后重试；复杂任务可以拆成几步。", True),
    ("network", ("connection", "connecterror", "network", "getaddrinfo", "ssl", "proxy"),
     "连不上模型服务。", "检查网络或代理设置后重试。", True),
    ("cli_missing", ("cli not found", "claude code not found", "cannot find claude"),
     "AI 引擎文件缺失。", "运行 DeepExcel 修复工具，或重新安装。", False),
]


def classify_error(exc_or_text: Any) -> dict:
    text = exc_or_text if isinstance(exc_or_text, str) else f"{type(exc_or_text).__name__}: {exc_or_text}"
    lowered = text.lower()
    for code, needles, message, hint, retryable in _ERROR_RULES:
        if any(n in lowered for n in needles):
            return {"code": code, "message": message, "hint": hint, "retryable": retryable,
                    "detail": text[:500]}
    return {"code": "unknown", "message": "处理时出错了。", "hint": "可以重试一次；反复出现请导出诊断包反馈。",
            "retryable": True, "detail": text[:500]}


class RunTracker:
    """一轮对话（一条用户消息到 stream_end）里的工具调用账本。"""

    def __init__(self) -> None:
        self.started = time.monotonic()
        self._pending: dict[str, tuple[str, float]] = {}
        self.tool_calls = 0
        self.failed_calls = 0
        self.compacted = False
        self.summarized = False
        self.interrupted = False
        # 最近一次从 CLI 收到任何东西的时刻；看门狗据此判断「很久没动静」
        self.last_activity = self.started

    def touch(self) -> None:
        self.last_activity = time.monotonic()

    def idle_seconds(self) -> float:
        return time.monotonic() - self.last_activity

    def running_tool(self) -> tuple[str, float] | None:
        """正在等宿主返回结果的调用里最早的那个：(工具名, 已等待秒数)。"""
        if not self._pending:
            return None
        name, t0 = min(self._pending.values(), key=lambda item: item[1])
        return name, time.monotonic() - t0

    def start(self, tool_use_id: str, name: str) -> None:
        self.tool_calls += 1
        self._pending[tool_use_id] = (name, time.monotonic())

    def end(self, tool_use_id: str) -> tuple[str, int]:
        name, t0 = self._pending.pop(tool_use_id, ("", time.monotonic()))
        return name, int((time.monotonic() - t0) * 1000)

    def unfinished(self) -> list[tuple[str, str]]:
        """中断或出错时还没收到结果的调用，要逐个补一个 tool_end，面板才不会一直转圈。"""
        items = [(tid, name) for tid, (name, _t0) in self._pending.items()]
        self._pending.clear()
        return items

    def elapsed_ms(self) -> int:
        return int((time.monotonic() - self.started) * 1000)


# ---------------------------------------------------------------------------
# 看门狗：很久没动静时告诉用户在等什么
# ---------------------------------------------------------------------------

# (空闲秒数阈值, 文案)。模型在长思考、写大段代码时可能几十秒不出字，
# 用户看到的只有三个跳动的点，分不清「在想」还是「卡死了」。
WATCHDOG_MODEL_WAIT = 20
WATCHDOG_TOOL_WAIT = 15
WATCHDOG_STUCK = 120


def watchdog_status(idle_s: float, running_tool: tuple[str, float] | None) -> str | None:
    """根据空闲时长和正在执行的工具，给出状态行；不需要提示时返回 None。"""
    if running_tool is not None:
        _name, waited = running_tool
        if waited >= WATCHDOG_STUCK:
            return f"这一步已执行 {int(waited)} 秒——Excel 可能弹出了对话框，请切到 Excel 看一下；也可以按停止"
        if waited >= WATCHDOG_TOOL_WAIT:
            return f"这一步执行中（{int(waited)} 秒）…"
        return None
    if idle_s >= WATCHDOG_STUCK:
        return f"模型已 {int(idle_s)} 秒没有响应，可能是网络或服务商繁忙；可以按停止后重试"
    if idle_s >= WATCHDOG_MODEL_WAIT:
        return f"仍在等待模型响应（{int(idle_s)} 秒）…"
    return None


# ---------------------------------------------------------------------------
# 代码边写边显示（tool_gen）
# ---------------------------------------------------------------------------

# 模型写一段 60 行的 VBA 或一张 500 行的表，参数要一个 token 一个 token 生成，
# 可能要几十秒。以前这段时间面板什么都没有；现在从流式事件里取出正在生成的
# 参数，边写边显示（Claude Code 写文件时也是这样）。
CODE_FIELDS = {"execute_vba": "code", "execute_jsa": "code", "execute_python": "code"}
TOOL_GEN_INTERVAL = 0.25
_CODE_PREVIEW_CHARS = 4000


def partial_json_string(buffer: str, field: str) -> str | None:
    """从不完整的 JSON 里取出字符串字段目前已生成的部分（处理转义，末尾可以是半个转义）。"""
    key = f'"{field}"'
    start = buffer.find(key)
    if start < 0:
        return None
    i = start + len(key)
    while i < len(buffer) and buffer[i] in " \t\r\n:":
        i += 1
    if i >= len(buffer) or buffer[i] != '"':
        return None
    i += 1
    out = []
    escapes = {"n": "\n", "t": "\t", "r": "\r", '"': '"', "\\": "\\", "/": "/", "b": "\b", "f": "\f"}
    while i < len(buffer):
        ch = buffer[i]
        if ch == '"':
            break
        if ch == "\\":
            if i + 1 >= len(buffer):
                break
            nxt = buffer[i + 1]
            if nxt == "u":
                digits = buffer[i + 2:i + 6]
                if len(digits) < 4:
                    break
                try:
                    out.append(chr(int(digits, 16)))
                except ValueError:
                    pass
                i += 6
                continue
            out.append(escapes.get(nxt, nxt))
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


class ToolGenTracker:
    """按 content block index 累积 tool_use 的 input_json_delta，节流后产出 tool_gen 事件。"""

    def __init__(self) -> None:
        self._blocks: dict[int, dict] = {}

    def start(self, index: int, tool_use_id: str, name: str) -> None:
        self._blocks[index] = {"id": tool_use_id, "name": bare_tool_name(name), "buf": "", "last": 0.0}

    def feed(self, index: int, partial: str, now: float | None = None) -> dict | None:
        block = self._blocks.get(index)
        if block is None or not partial:
            return None
        block["buf"] += partial
        now = time.monotonic() if now is None else now
        if now - block["last"] < TOOL_GEN_INTERVAL:
            return None
        block["last"] = now
        return self._event(block)

    def stop(self, index: int) -> None:
        self._blocks.pop(index, None)

    @staticmethod
    def _event(block: dict) -> dict:
        fields = {"id": block["id"], "name": block["name"], "chars": len(block["buf"])}
        field = CODE_FIELDS.get(block["name"])
        if field:
            code = partial_json_string(block["buf"], field)
            if code:
                fields["lines"] = code.count("\n") + 1
                fields["preview"] = code[-_CODE_PREVIEW_CHARS:]
        return envelope("tool_gen", **fields)
