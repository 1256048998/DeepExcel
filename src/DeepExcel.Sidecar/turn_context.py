"""Assemble the user message the model sees each turn.

Layout (one message per turn):

    <context>
    ...history summary / workbook memory / current sheet and selection /
    user edits / attachments / host / workbook structure...
    </context>

    <permission-mode>...</permission-mode>      (only when not the default mode)
    <request>
    what the user typed
    </request>
    （请用中文思考。）                              (when thinking is on)

The request goes after the context, not before it: Anthropic's long-context
guidance is to put long material first and the query last, which on a large
structure summary is exactly the case here. Decided 2026-11-25: follow
Claude's guidance and keep the request last. REQUEST_LAST flips it back.

What is sent once instead of every turn:
  * the workbook structure summary -- re-sent only when it changed, or after the
    conversation was compacted; otherwise one line says the earlier one still holds;
  * the host line (Excel / WPS version, dynamic arrays) -- first turn and after
    compaction only.
The current sheet, selection, user edits and host notices are short and change
between turns, so they are always sent.
"""
from __future__ import annotations

import hashlib

REQUEST_LAST = True

STRUCTURE_UNCHANGED = "[工作簿结构] 和前面发过的一样，沿用那一份"

_state = {"structure": None, "host": None}


def reset() -> None:
    """New conversation or compaction: everything is sent again on the next turn."""
    _state["structure"] = None
    _state["host"] = None


def _digest(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def host_line(host) -> str:
    """{name, version, build, dynamicArrays} → 一行；什么都不知道时返回空串。"""
    if not isinstance(host, dict):
        return ""
    name = str(host.get("name") or "").strip()
    version = str(host.get("version") or "").strip()
    build = str(host.get("build") or "").strip()
    if not name and not version:
        return ""
    label = " ".join(p for p in (name, version) if p)
    if build:
        label += f"（build {build}）"
    dynamic = host.get("dynamicArrays")
    if dynamic is True:
        arrays = "支持动态数组（FILTER / UNIQUE / SORT / SEQUENCE 会溢出）"
    elif dynamic is False:
        arrays = "不支持动态数组：不要写 FILTER / UNIQUE / SORT / SEQUENCE / XLOOKUP 这类新函数，用 INDEX / MATCH / SUMIFS 等老写法"
    else:
        arrays = "动态数组支持情况未知"
    return f"[宿主] {label}；{arrays}"


class Turn:
    """One turn's pieces. pending_* are what this turn will send; commit() records them once the query went out."""

    def __init__(self, structure: str | None, host) -> None:
        self.structure_text = ""
        self.host_text = ""
        self._pending_structure = None
        self._pending_host = None

        structure = structure.strip() if isinstance(structure, str) else ""
        if structure:
            digest = _digest(structure)
            if digest != _state["structure"]:
                self.structure_text = structure
                self._pending_structure = digest
            else:
                self.structure_text = STRUCTURE_UNCHANGED

        line = host_line(host)
        if line and line != _state["host"]:
            self.host_text = line
            self._pending_host = line

    @property
    def structure_sent(self) -> bool:
        """这一轮带了完整的结构摘要（第一次、变了、或压缩过后）"""
        return self._pending_structure is not None

    def commit(self) -> None:
        if self._pending_structure is not None:
            _state["structure"] = self._pending_structure
        if self._pending_host is not None:
            _state["host"] = self._pending_host


def compose(request: str, context_parts: list, *, instruction: str = "", trailer: str = "") -> str:
    """把上下文块、模式说明和用户的请求拼成一条消息。空块跳过；没有上下文时只发请求。"""
    parts = [p.strip() for p in context_parts if isinstance(p, str) and p.strip()]
    instruction = (instruction or "").strip()
    if not parts and not instruction:
        return (request or "") + (trailer or "")
    request_block = f"<request>\n{(request or '').strip()}\n</request>"
    head = []
    if parts:
        head.append("<context>\n" + "\n\n".join(parts) + "\n</context>")
    if REQUEST_LAST:
        body = head + ([instruction] if instruction else []) + [request_block]
    else:
        body = ([instruction] if instruction else []) + [request_block] + head
    return "\n\n".join(body) + (trailer or "")
