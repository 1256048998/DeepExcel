"""写公式之后的模式检查：并入 C# 的「写后自动体检」（verification）。

只看被写入的列所在的区块。候选落在这次写入的区域里，说明写入本身有问题（漏填一行、
某行公式跟上下不同）；落在同列但在写入区域外，多半是表里原有的问题，只作附带信息。
"""

from __future__ import annotations

import re

from .anomalies import find_anomalies
from .blocks import detect_blocks
from .formulas import scan_formulas
from .grid import Rect, Snapshot
from .headers import analyze_header

_A1 = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')+'|[^!]+)!)?"
    r"\$?(?P<c1>[A-Za-z]{1,3})\$?(?P<r1>\d+)"
    r"(?::\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>\d+))?$"
)
_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
MAX_REPORTED = 5


def col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_a1(address: str) -> tuple[str | None, int, int, int, int] | None:
    """'Sheet1'!B2:D9 → (表名, 行1, 列1, 行2, 列2)，工作表坐标（1 起）；整列 / 整行等不支持的返回 None"""
    m = _A1.match((address or "").strip())
    if not m:
        return None
    sheet = m.group("sheet")
    if sheet and sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    r1, c1 = int(m.group("r1")), col_number(m.group("c1"))
    r2 = int(m.group("r2") or r1)
    c2 = col_number(m.group("c2") or m.group("c1"))
    return sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def _cell_of(address: str) -> tuple[int, int] | None:
    m = _CELL.match(address or "")
    return (int(m.group(2)), col_number(m.group(1))) if m else None


def pattern_check(snapshot: dict, written: tuple[int, int, int, int]) -> dict | None:
    """written 是被写入的矩形（工作表坐标 行1, 列1, 行2, 列2）。没有发现就返回 None。"""
    snap = Snapshot(snapshot)
    r1, c1, r2, c2 = written
    area = Rect(r1 - snap.row0, c1 - snap.col0, r2 - snap.row0, c2 - snap.col0)

    inside, nearby, inconsistent = [], [], []
    for block in detect_blocks(snap).blocks:
        rect = block.rect
        if block.kind != "table" or rect.c2 < area.c1 or rect.c1 > area.c2 or rect.r2 < area.r1 or rect.r1 > area.r2:
            continue
        header = analyze_header(snap, block)
        scan = scan_formulas(snap, block, header)
        for candidate in find_anomalies(snap, block, header, scan)["candidates"]:
            cell = _cell_of(candidate["address"])
            if cell is None or not (c1 <= cell[1] <= c2):
                continue
            (inside if r1 <= cell[0] <= r2 else nearby).append(candidate)
        for inc in scan.inconsistent:
            if area.c1 <= inc.col <= area.c2:
                label = header.label_of(inc.col)
                inconsistent.append(f"{snap.column_letter(inc.col)} 列" + (f"（{label}）" if label else ""))

    if not (inside or nearby or inconsistent):
        return None
    out: dict = {"complete": not (snap.truncated or snap.formulas_truncated)}
    if inside:
        out["in_written_range"] = inside[:MAX_REPORTED]
    if nearby:
        out["elsewhere_in_column"] = nearby[:MAX_REPORTED]
    if inconsistent:
        out["inconsistent_columns"] = inconsistent
    return out


def summary_line(check: dict) -> str | None:
    """给 verification.summary 追加的一句；写入区域里没有问题就不追加"""
    inside = check.get("in_written_range") or []
    if not inside:
        return None
    cells = "、".join(c["address"] for c in inside[:3])
    first = inside[0]["detail"]
    more = f" 等 {len(inside)} 处" if len(inside) > 3 else ""
    return f"公式模式检查：{cells}{more} 可疑（{first}）"
