"""多级表头：规则优先，拿不准就说拿不准。

逐行看填充率、文本 / 数字 / 日期占比、横向合并与纵向合并，先认标题行（整行只有一段文字），
再认最多 3 行表头；每列的标签自上而下拼成路径（「扣款/养老/公司」）。判断不了时
uncertain=True 并写明原因，交给 agent 用 read_range 实读，而不是硬给一个结论。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .blocks import Block
from .grid import DATE, EMPTY, ERROR, NUMBER, TEXT, Rect, Snapshot

MAX_TITLE_ROWS = 2
MAX_HEADER_ROWS = 3
HEADER_TEXT_RATIO = 0.7
HEADER_FILL = 0.5
SUB_HEADER_FILL = 0.3
TYPE_DOMINANCE = 0.8

TOTAL_WORDS = re.compile(r"^\s*(合\s*计|总\s*计|小\s*计|共\s*计|累\s*计|总额|汇总|grand\s+total|sub\s*total|total|sum)\b",
                         re.IGNORECASE)


@dataclass
class RowProfile:
    filled: int  # 非空格（合并覆盖的格也算）
    values: int  # 有值的格（合并区域只算左上角）
    text: int
    number: int
    date: int
    formulas: int
    hmerge: bool  # 该行起始的横向合并（跨 ≥2 列）
    vmerge: bool  # 该行起始、向下延伸的纵向合并

    def ratio(self, part: int) -> float:
        return part / self.values if self.values else 0.0


@dataclass
class ColumnInfo:
    col: int  # 窗口内列号
    letter: str
    label: str | None  # 表头路径，如「扣款/养老/公司」；没有表头时为 None
    type: str  # number / date / text / mixed / empty
    formula: bool  # 数据行基本都是公式
    filled: int  # 数据行中的非空格数


@dataclass
class HeaderInfo:
    title_rows: list[int] = field(default_factory=list)  # 窗口内行号
    header_rows: list[int] = field(default_factory=list)
    data_start: int = 0
    data_end: int = 0
    total_rows: list[int] = field(default_factory=list)  # 合计 / 小计行
    columns: list[ColumnInfo] = field(default_factory=list)
    uncertain: bool = False
    reasons: list[str] = field(default_factory=list)
    title: str | None = None

    def label_of(self, col: int) -> str | None:
        for info in self.columns:
            if info.col == col:
                return info.label
        return None


def _profile(snap: Snapshot, rect: Rect, r: int) -> RowProfile:
    filled = values = text = number = date = formulas = 0
    hmerge = vmerge = False
    for c in range(rect.c1, rect.c2 + 1):
        merge = snap.merge_at(r, c)
        kind = snap.kind(r, c)
        if kind != EMPTY or merge is not None:
            filled += 1
        if merge is not None and (merge.r1 != r or merge.c1 != c):
            continue  # 合并区域只看左上角
        if merge is not None:
            hmerge = hmerge or merge.cols >= 2
            vmerge = vmerge or merge.rows >= 2
        if kind == EMPTY:
            continue
        values += 1
        if kind == TEXT:
            text += 1
        elif kind == NUMBER:
            number += 1
        elif kind == DATE:
            date += 1
        if (r, c) in snap.formulas:
            formulas += 1
    return RowProfile(filled, values, text, number, date, formulas, hmerge, vmerge)


def _is_title_row(rect: Rect, profile: RowProfile) -> bool:
    return rect.cols >= 2 and profile.values == 1 and profile.text == 1


def _looks_like_header(profile: RowProfile, width: int, min_fill: float) -> bool:
    if profile.values == 0 or profile.formulas:
        return False
    return profile.ratio(profile.text) >= HEADER_TEXT_RATIO and profile.filled / width >= min_fill


def _is_data_like(profile: RowProfile) -> bool:
    return profile.values > 0 and (profile.ratio(profile.number + profile.date) >= 0.3 or profile.formulas > 0)


def _fills_gaps(snap: Snapshot, rect: Rect, upper: int, lower: int) -> bool:
    """下面这行在上面那行的空列上有值（「扣款」下面的「医疗」「失业」）"""
    def empty(r, c):
        return snap.kind(r, c) == EMPTY and snap.merge_at(r, c) is None
    return any(empty(upper, c) and not empty(lower, c) for c in range(rect.c1, rect.c2 + 1))


def _merge_reaches(snap: Snapshot, rect: Rect, header_rows: list[int], r: int) -> bool:
    """表头里有纵向合并从上面延伸到第 r 行"""
    return any(m.r1 in header_rows and m.r1 < r <= m.r2 and rect.c1 <= m.c1 <= rect.c2 for m in snap.merges)


def _label(snap: Snapshot, r: int, c: int, header_rows: list[int]) -> str:
    merge = snap.merge_at(r, c)
    if merge is not None and merge.r1 in header_rows:
        return snap.text(merge.r1, merge.c1)
    return snap.text(r, c)


def _column_paths(snap: Snapshot, rect: Rect, header_rows: list[int]) -> dict[int, str | None]:
    grid: dict[tuple[int, int], str] = {}
    for r in header_rows:
        for c in range(rect.c1, rect.c2 + 1):
            grid[(r, c)] = _label(snap, r, c, header_rows)

    # 「跨列居中」而不是合并：上层标签只写在第一列，右边留空，下层每列都有标签。
    # 这种空位继承左边最近的同行标签（直到遇到下一个上层标签）。
    for i, r in enumerate(header_rows[:-1]):
        below = header_rows[i + 1]
        current = ""
        for c in range(rect.c1, rect.c2 + 1):
            if grid[(r, c)]:
                current = grid[(r, c)]
            elif current and grid[(below, c)] and snap.merge_at(r, c) is None:
                grid[(r, c)] = current
            elif not grid[(below, c)]:
                current = ""

    paths: dict[int, str | None] = {}
    for c in range(rect.c1, rect.c2 + 1):
        parts: list[str] = []
        for r in header_rows:
            label = grid[(r, c)]
            if label and (not parts or parts[-1] != label):
                parts.append(label)
        paths[c] = "/".join(parts) if parts else None
    return paths


def _column_type(snap: Snapshot, c: int, rows: list[int]) -> tuple[str, bool, int]:
    counts = {TEXT: 0, NUMBER: 0, DATE: 0, "other": 0}
    formulas = 0
    filled = 0
    for r in rows:
        kind = snap.kind(r, c)
        if kind == EMPTY:
            if (r, c) in snap.formulas:
                formulas += 1
                filled += 1
            continue
        filled += 1
        if (r, c) in snap.formulas:
            formulas += 1
        if kind in counts:
            counts[kind] += 1
        elif kind != ERROR:
            counts["other"] += 1
    if filled == 0:
        return "empty", False, 0
    kind, top = max(counts.items(), key=lambda kv: kv[1])
    type_ = kind if top / filled >= TYPE_DOMINANCE and kind != "other" else "mixed"
    return type_, formulas / filled >= TYPE_DOMINANCE, filled


def is_total_row(snap: Snapshot, rect: Rect, r: int) -> bool:
    """前三格里有「合计 / 小计 / Total」之类的文字"""
    for c in range(rect.c1, min(rect.c2, rect.c1 + 2) + 1):
        if snap.kind(r, c) == TEXT and TOTAL_WORDS.match(snap.text(r, c)):
            return True
    return False


def analyze_header(snap: Snapshot, block: Block) -> HeaderInfo:
    rect = block.rect
    info = HeaderInfo(data_end=rect.r2)
    width = rect.cols
    r = rect.r1

    # 标题行：整行只有一段文字（合并或不合并都算）
    while r <= rect.r2 and len(info.title_rows) < MAX_TITLE_ROWS and rect.rows >= 3:
        profile = _profile(snap, rect, r)
        if not _is_title_row(rect, profile):
            break
        info.title_rows.append(r)
        r += 1
    if info.title_rows:
        info.title = " ".join(snap.text(t, c) for t in info.title_rows
                              for c in range(rect.c1, rect.c2 + 1) if snap.text(t, c))

    # 表头行
    first = _profile(snap, rect, r) if r <= rect.r2 else None
    if first is not None and _looks_like_header(first, width, HEADER_FILL):
        info.header_rows.append(r)
        previous = first
        r += 1
        while r <= rect.r2 and len(info.header_rows) < MAX_HEADER_ROWS:
            profile = _profile(snap, rect, r)
            if not _looks_like_header(profile, width, SUB_HEADER_FILL):
                break
            nxt = _profile(snap, rect, r + 1) if r + 1 <= rect.r2 else None
            # 下一层表头的证据：上一层有横向合并、纵向合并延伸下来（强证据）；
            # 或上一层有空位而这一层补齐（弱证据：数据首行也常比表头满，所以还要求紧跟着的是数据行）
            strong = previous.hmerge or previous.vmerge or _merge_reaches(snap, rect, info.header_rows, r)
            weak = _fills_gaps(snap, rect, r - 1, r) and nxt is not None and _is_data_like(nxt)
            if not (strong or weak):
                break
            if (nxt is not None and len(info.header_rows) + 1 == MAX_HEADER_ROWS
                    and _looks_like_header(nxt, width, SUB_HEADER_FILL) and (profile.hmerge or profile.vmerge)):
                info.uncertain = True
                info.reasons.append("表头可能超过 3 行")
            info.header_rows.append(r)
            previous = profile
            r += 1
        if r <= rect.r2 and len(info.header_rows) < MAX_HEADER_ROWS:
            nxt = _profile(snap, rect, r)
            if nxt.hmerge and _looks_like_header(nxt, width, SUB_HEADER_FILL):
                info.uncertain = True
                info.reasons.append(f"第 {snap.sheet_row(r)} 行有横向合并，可能也是表头")
    elif first is not None:
        info.uncertain = True
        info.reasons.append("没有识别到表头行" if not _is_data_like(first) else "首行像数据，可能没有表头")

    info.data_start = r
    data_rows = list(range(info.data_start, rect.r2 + 1))
    info.total_rows = [row for row in data_rows if is_total_row(snap, rect, row)]
    body = [row for row in data_rows if row not in info.total_rows]

    paths = _column_paths(snap, rect, info.header_rows) if info.header_rows else {}
    for c in range(rect.c1, rect.c2 + 1):
        type_, formula, filled = _column_type(snap, c, body)
        info.columns.append(ColumnInfo(
            col=c, letter=snap.column_letter(c), label=paths.get(c),
            type=type_, formula=formula, filled=filled,
        ))
    if info.header_rows:
        unnamed = [col.letter for col in info.columns if col.label is None and col.filled]
        if unnamed:
            info.uncertain = True
            info.reasons.append("这些列有数据但没有表头：" + "、".join(unnamed[:8]))
    return info
