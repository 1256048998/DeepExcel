# src/DeepExcel.Sidecar/workbook_memory.py
"""工作簿记忆：Claude Code 的 CLAUDE.md 的工作簿版。

每个工作簿一个目录 %LOCALAPPDATA%\\DeepExcel\\workbooks\\<sha256(workbookKey) 前 16 字节 hex>\\：
- NOTES.md       agent 维护（update_workbook_notes）、用户可在面板里查看和修改：
                 结构怪癖、用户偏好、做过的改动、禁区
- history.jsonl  程序追加：每次写入 / 执行的工具、目标、成败和错误码（不存单元格内容）
- meta.json      工作簿名和 key，面板列表显示用

目录名与对话历史（ConversationHistory）同一口径：同一个文件在 Excel 和 WPS 里共用一份记忆。
只存本地，不上传。没保存过的新工作簿（key 不是路径）没有记忆：关掉就没了，记了也找不回来。

禁区（NOTES.md 的「## 禁区」一节，每行一个表名或 表名!区域）由 PreToolUse 在代码层拒绝写入，
见 protected_zones / check_write。模型不能自己删掉禁区条目：只有用户能在面板里解除。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

MAX_NOTES_CHARS = 6000
MAX_HISTORY_LINES = 500
KEEP_HISTORY_LINES = 300
SUMMARY_HISTORY = 8
SUMMARY_NOTES_CHARS = 3000

TEMPLATE = """# 工作簿记忆

## 结构怪癖

## 用户偏好

## 做过的改动

## 禁区
"""

# 只记会改工作簿、或执行代码的工具；读操作不进历史
READ_ONLY_TOOLS = frozenset({
    "read_workbook", "read_selection", "read_range", "find", "list", "inspect_sheet",
    "explore_workbook", "read_attachment", "clarify_intent", "todo_write", "update_workbook_notes",
    "load_skill", "present_plan", "screenshot_excel", "create_snapshot", "export_chart",
})

# 各工具里表示写入目标的参数（按优先级）
TARGET_ARGS = ("address", "range_address", "dest_address", "from_address", "source_range",
               "data_range", "sheet_name", "sheet", "name", "old_name")


def root_dir() -> Path:
    override = os.environ.get("DEEPEXCEL_MEMORY_DIR")
    if override:
        return Path(override)
    return Path(os.environ.get("LOCALAPPDATA", "")) / "DeepExcel" / "workbooks"


def key_hash(workbook_key: str) -> str:
    """与 C# ConversationHistory.GetFileName / WPS conversation-store 一致：SHA256 前 16 字节小写 hex"""
    return hashlib.sha256(str(workbook_key).encode("utf-8")).digest()[:16].hex()


def is_persistent_key(workbook_key: Any) -> bool:
    return isinstance(workbook_key, str) and ("\\" in workbook_key or "/" in workbook_key)


# ---------------------------------------------------------------------------
# 禁区
# ---------------------------------------------------------------------------

MAX_ROW = 1048576
MAX_COL = 16384

# 区域里有英文冒号（A1:D20），所以只在全角标点、括号、空白处截断；「- 汇总: 说明」末尾的冒号另外去掉
_ZONE_LINE = re.compile(r"^\s*[-*]\s*(?P<ref>'(?:[^']|'')+'(?:![^\s（(：，,；;]+)?|[^\s（(：，,；;]+)")
_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COL = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW = re.compile(r"^\$?(\d+)$")


def _col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_range(text: str) -> tuple[int, int, int, int] | None:
    """A1 / A1:D20 / A:C（整列）/ 3:5（整行）→ (行1, 列1, 行2, 列2)，1 起"""
    parts = (text or "").strip().split(":")
    if len(parts) > 2 or not parts[0]:
        return None
    a, b = parts[0], parts[-1]
    ca, cb = _CELL.match(a), _CELL.match(b)
    if ca and cb:
        r1, c1, r2, c2 = int(ca.group(2)), _col_number(ca.group(1)), int(cb.group(2)), _col_number(cb.group(1))
    elif _COL.match(a) and _COL.match(b):
        r1, r2 = 1, MAX_ROW
        c1, c2 = _col_number(_COL.match(a).group(1)), _col_number(_COL.match(b).group(1))
    elif _ROW.match(a) and _ROW.match(b):
        c1, c2 = 1, MAX_COL
        r1, r2 = int(_ROW.match(a).group(1)), int(_ROW.match(b).group(1))
    else:
        return None
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def split_address(address: str) -> tuple[str | None, str]:
    """'Sheet 1'!A1:B2 → ('Sheet 1', 'A1:B2')；没有表名时表名为 None"""
    text = (address or "").strip()
    if "!" not in text:
        return None, text
    sheet, _, rng = text.rpartition("!")
    if sheet.startswith("'") and sheet.endswith("'") and len(sheet) >= 2:
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, rng


@dataclass(frozen=True)
class Zone:
    sheet: str
    rect: tuple[int, int, int, int] | None  # None = 整张表
    text: str  # 原始条目（给用户看）

    def describe(self) -> str:
        return self.text


def _section(notes: str, keyword: str) -> list[str]:
    lines, inside = [], False
    for line in (notes or "").splitlines():
        if line.lstrip().startswith("#"):
            inside = keyword in line
            continue
        if inside:
            lines.append(line)
    return lines


def protected_zones(notes: str) -> list[Zone]:
    zones = []
    for line in _section(notes, "禁区"):
        m = _ZONE_LINE.match(line)
        if not m:
            continue
        ref = m.group("ref").rstrip(":")
        sheet, rng = split_address(ref) if "!" in ref else (ref.strip("'"), "")
        if not sheet:
            continue
        rect = parse_range(rng) if rng else None
        if rng and rect is None:
            continue  # 看不懂的区域不当禁区（宁可不拦，也不能把整张表误锁）
        zones.append(Zone(sheet=sheet, rect=rect, text=line.strip().lstrip("-* ").strip()))
    return zones


def _intersects(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def zone_hits(zones: list[Zone], sheet: str | None, rect: tuple[int, int, int, int] | None) -> list[Zone]:
    """写入目标（表名 + 矩形；矩形为 None 表示整张表或未知区域）碰到了哪些禁区"""
    if not sheet:
        return []
    hits = []
    for zone in zones:
        if zone.sheet.casefold() != sheet.casefold():
            continue
        if zone.rect is None or rect is None or _intersects(zone.rect, rect):
            hits.append(zone)
    return hits


# 写入目标：这些参数是被写的区域（source_address / data_range / source_range 是读的，不算）
_WRITE_ADDRESS_ARGS = ("address", "range_address", "dest_address", "from_address")
# 整张表级别的操作 → 参数名
_SHEET_ARGS = {
    "delete_sheet": "name", "rename_sheet": "old_name", "create_pivot_table": "destination_sheet",
    "refresh_pivot": "sheet_name", "group_pivot_date": "sheet_name", "set_pivot_value_display": "sheet_name",
    "set_pivot_totals": "sheet_name", "add_pivot_slicer": "sheet_name",
}
# 在活动表上加对象
_ACTIVE_SHEET_TOOLS = frozenset({"create_chart", "create_combo_chart"})
# 只改视图、不改内容
_VIEW_ONLY_TOOLS = frozenset({"freeze_panes"})
_CODE_TOOLS = frozenset({"execute_vba", "execute_jsa"})


def _write_targets(tool: str, args: dict, active_sheet: str | None) -> list[tuple[str | None, tuple | None, str]]:
    """(表名, 矩形或 None=整表/不明, 原始地址)。表名为 None 表示地址没带表名、也不知道活动表。"""
    targets = []
    if tool in _SHEET_ARGS:
        name = args.get(_SHEET_ARGS[tool])
        targets.append((name if isinstance(name, str) and name else active_sheet, None, str(name or "")))
        return targets
    if tool in _ACTIVE_SHEET_TOOLS:
        return [(active_sheet, None, "")]
    if tool in ("insert_rows", "delete_rows", "insert_columns", "delete_columns"):
        # 插入 / 删除会让后面的行（列）整体移动：从这里到表尾都算被改动
        try:
            if tool.endswith("rows"):
                start = int(args.get("row"))
                rect = (start, 1, MAX_ROW, MAX_COL)
            else:
                column = args.get("column")
                start = _col_number(column) if isinstance(column, str) and column.isalpha() else int(column)
                rect = (1, start, MAX_ROW, MAX_COL)
        except (TypeError, ValueError):
            rect = None
        return [(active_sheet, rect, "")]
    for key in _WRITE_ADDRESS_ARGS:
        value = args.get(key)
        if not isinstance(value, str) or not value.strip():
            continue
        sheet, rng = split_address(value)
        rect = parse_range(rng)
        if rect is None:
            continue  # 名称之类解析不了的地址：没法判断，不拦
        if tool == "fill_formula_down" and key == "from_address":
            try:
                rect = (rect[0], rect[1], rect[2] + max(0, int(args.get("row_count") or 0)), rect[3])
            except (TypeError, ValueError):
                pass
        targets.append((sheet or active_sheet, rect, value.strip()))
    return targets


def _code_mentions(code: str, sheet: str) -> bool:
    """代码里按名字引用了这张表（"汇总" / '汇总' / [汇总]）"""
    lowered = code.casefold()
    name = sheet.casefold()
    return any(token in lowered for token in (f'"{name}"', f"'{name}'", f"[{name}]", f"'{name}'!"))


def check_write(tool: str, args: dict, notes: str, active_sheet: str | None) -> str | None:
    """PreToolUse 用：这次调用要写到禁区就返回拒绝理由，否则 None"""
    zones = protected_zones(notes)
    if not zones or not tool or tool in READ_ONLY_TOOLS or tool in _VIEW_ONLY_TOOLS:
        return None
    args = args if isinstance(args, dict) else {}
    if tool in _CODE_TOOLS:
        code = str(args.get("code") or "")
        hits = [z for z in zones if _code_mentions(code, z.sheet)]
        if not hits and active_sheet and "activesheet" in code.casefold():
            hits = zone_hits(zones, active_sheet, None)
        if hits:
            return _refusal(hits)
        return None
    for sheet, rect, address in _write_targets(tool, args, active_sheet):
        if sheet is None:
            if zones:
                return ("工作簿里有禁区，但这个地址没有带表名、也不知道当前活动表，没法确认会不会写到禁区，本次未执行。"
                        "请在地址里写上表名（如 明细!A1:B2）后重试")
            continue
        hits = zone_hits(zones, sheet, rect)
        if hits:
            return _refusal(hits)
    return None


def _refusal(hits: list[Zone]) -> str:
    names = "；".join(z.text for z in hits)
    return (f"「{names}」是用户在工作簿记忆里标记的禁区，系统拒绝写入，本次未执行。"
            "不要换工具或用代码绕过：告诉用户这里是禁区，需要改的话请用户先在「工作簿记忆」里解除")


# ---------------------------------------------------------------------------
# 存储
# ---------------------------------------------------------------------------

class WorkbookMemory:
    def __init__(self, workbook_key: str, workbook_name: str | None = None, root: Path | None = None):
        self.key = workbook_key
        self.name = workbook_name or Path(workbook_key.replace("\\", "/")).name
        self.dir = (root or root_dir()) / key_hash(workbook_key)
        self.notes_path = self.dir / "NOTES.md"
        self.history_path = self.dir / "history.jsonl"
        self.meta_path = self.dir / "meta.json"

    # -- NOTES.md ----------------------------------------------------------
    def notes(self) -> str:
        try:
            return self.notes_path.read_text(encoding="utf-8")
        except OSError:
            return ""

    def notes_mtime(self) -> float:
        try:
            return self.notes_path.stat().st_mtime
        except OSError:
            return 0.0

    def save_notes(self, text: str, by_agent: bool = True) -> str | None:
        """写入 NOTES.md；不允许时返回原因（模型能读懂的一句话）"""
        text = (text or "").replace("\r\n", "\n").strip() + "\n"
        if len(text) > MAX_NOTES_CHARS:
            return f"记忆太长（{len(text)} 字），上限 {MAX_NOTES_CHARS} 字：只留以后还用得上的，做过的改动只记结论"
        if by_agent:
            kept = {z.text for z in protected_zones(text)}
            dropped = [z.text for z in protected_zones(self.notes()) if z.text not in kept]
            if dropped:
                return ("禁区只能由用户在「工作簿记忆」里解除，不能由你删除。请保留这些条目后重新提交："
                        + "；".join(dropped))
        self.dir.mkdir(parents=True, exist_ok=True)
        tmp = self.notes_path.with_suffix(".tmp")
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, self.notes_path)
        self._touch_meta()
        return None

    def clear(self) -> None:
        for path in (self.notes_path, self.history_path, self.meta_path):
            try:
                path.unlink()
            except OSError:
                pass
        try:
            self.dir.rmdir()
        except OSError:
            pass

    # -- history.jsonl -----------------------------------------------------
    def append_history(self, entry: dict) -> None:
        try:
            self.dir.mkdir(parents=True, exist_ok=True)
            with self.history_path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
            self._touch_meta()
            self._rotate()
        except OSError:
            pass

    def history(self, limit: int | None = None) -> list[dict]:
        try:
            lines = self.history_path.read_text(encoding="utf-8").splitlines()
        except OSError:
            return []
        entries = []
        for line in lines[-limit:] if limit else lines:
            try:
                entry = json.loads(line)
            except ValueError:
                continue
            if isinstance(entry, dict):
                entries.append(entry)
        return entries

    def _rotate(self) -> None:
        try:
            lines = self.history_path.read_text(encoding="utf-8").splitlines()
        except OSError:
            return
        if len(lines) <= MAX_HISTORY_LINES:
            return
        tmp = self.history_path.with_suffix(".tmp")
        tmp.write_text("\n".join(lines[-KEEP_HISTORY_LINES:]) + "\n", encoding="utf-8")
        os.replace(tmp, self.history_path)

    def _touch_meta(self) -> None:
        try:
            self.meta_path.write_text(json.dumps(
                {"workbook_key": self.key, "workbook_name": self.name, "updated_at": int(time.time())},
                ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass

    # -- 注入 ---------------------------------------------------------------
    def summary(self) -> str | None:
        """会话开始注入给模型的记忆摘要；什么都没记过返回 None"""
        notes = self.notes().strip()
        recent = self.history(SUMMARY_HISTORY)
        if not notes and not recent:
            return None
        parts = ["<workbook-memory>",
                 f"这是你以前在「{self.name}」上留下的记忆（只存在用户电脑上）。先读它，别重复踩坑；"
                 "发现新的结构怪癖、用户偏好，或做完一件值得记住的改动后，用 update_workbook_notes 更新。"]
        if notes:
            if len(notes) > SUMMARY_NOTES_CHARS:
                notes = notes[:SUMMARY_NOTES_CHARS] + "\n…（后面还有，需要时整份重写前先想好保留什么）"
            parts.append(notes)
        zones = protected_zones(self.notes())
        if zones:
            parts.append("禁区（写入会被系统直接拒绝，不要尝试绕过）：" + "；".join(z.text for z in zones))
        if recent:
            parts.append("最近的操作记录：")
            for entry in recent:
                parts.append("- " + describe_history(entry))
        parts.append("</workbook-memory>")
        return "\n".join(parts)


def describe_history(entry: dict) -> str:
    when = time.strftime("%m-%d %H:%M", time.localtime(entry.get("ts", 0)))
    target = f" {entry['target']}" if entry.get("target") else ""
    outcome = "成功" if entry.get("ok") else f"失败（{entry.get('error_code') or 'tool_failed'}）"
    return f"{when} {entry.get('tool')}{target} {outcome}"


def history_entry(tool: str, args: dict, result: dict) -> dict | None:
    """一次工具调用 → history.jsonl 的一行。读操作返回 None。只记目标地址，不记写入的内容。"""
    if not tool or tool in READ_ONLY_TOOLS:
        return None
    target = None
    for key in TARGET_ARGS:
        value = (args or {}).get(key)
        if isinstance(value, str) and value.strip():
            target = value.strip()[:80]
            break
    entry = {"ts": int(time.time()), "tool": tool, "ok": bool(result.get("ok"))}
    if target:
        entry["target"] = target
    if not result.get("ok"):
        error = result.get("error") or {}
        entry["error_code"] = str(error.get("code") or "tool_failed")[:40]
    return entry


# ---------------------------------------------------------------------------
# 当前会话的记忆
# ---------------------------------------------------------------------------

_current: WorkbookMemory | None = None
_injected: dict[str, float] = {}  # key → 注入时 NOTES.md 的 mtime


_active_sheet: str | None = None


def active_sheet() -> str | None:
    """用户发消息时的活动表（地址不带表名时，写入落在这里）"""
    return _active_sheet


def use_context(context: Any) -> WorkbookMemory | None:
    """每条用户消息调用：按上下文里的 workbookKey 切换当前记忆，记下活动表"""
    global _current, _active_sheet
    _active_sheet = None
    if isinstance(context, dict):
        workbook = context.get("workbook")
        # Excel：workbook.activeSheet；WPS：activeSheet
        sheet = workbook.get("activeSheet") if isinstance(workbook, dict) else context.get("activeSheet")
        _active_sheet = sheet if isinstance(sheet, str) and sheet else None
    key = context.get("workbookKey") if isinstance(context, dict) else None
    if not is_persistent_key(key):
        _current = None
        return None
    if _current is None or _current.key != key:
        name = context.get("workbookName") if isinstance(context.get("workbookName"), str) else None
        _current = WorkbookMemory(key, name)
    return _current


def current() -> WorkbookMemory | None:
    return _current


def injection_for_turn(memory: WorkbookMemory | None, force: bool = False) -> str | None:
    """本轮要不要把记忆拼进用户消息：本会话第一次、用户在面板里改过、或上下文被压缩过（force）"""
    if memory is None:
        return None
    mtime = memory.notes_mtime()
    if not force and memory.key in _injected and _injected[memory.key] == mtime:
        return None
    summary = memory.summary()
    _injected[memory.key] = mtime
    return summary


def mark_seen(memory: WorkbookMemory) -> None:
    """模型自己刚写的记忆不用在下一轮再注入一遍"""
    _injected[memory.key] = memory.notes_mtime()


def reset_injection() -> None:
    _injected.clear()
