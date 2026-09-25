"""R1C1 公式模式：一列 500 个 =D2*E2、=D3*E3 … 在 R1C1 里是同一个 =RC[-2]*RC[-1]。

按列求相同 R1C1 的连续游程，再把相邻列上行范围相同的游程合并成矩形；模式 id 取 R1C1
文本的哈希；渲染时把同行引用换成表头名（[数量]*[单价]），其余引用换回 A1。
同一列的数据行出现多种模式时标「公式不一致」。快照被截断时只说「统计不完整」，不给数字。
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from .blocks import Block
from .grid import Rect, Snapshot, col_letters
from .headers import HeaderInfo

MAX_PATTERNS = 30

# R1C1 单元格引用：R、C 各自可带 [相对偏移] 或绝对行列号。前面不能紧挨标识符字符（排除函数名 / 名称），
# 后面不能是字母数字或左括号。
_REF = re.compile(
    r"(?<![A-Za-z0-9_.\]])"
    r"R(?:\[(?P<rrel>-?\d+)\]|(?P<rabs>\d+))?"
    r"C(?:\[(?P<crel>-?\d+)\]|(?P<cabs>\d+))?"
    r"(?![A-Za-z0-9_(!'])"
)


@dataclass(frozen=True)
class Run:
    col: int
    r1: int
    r2: int
    formula: str

    @property
    def rows(self) -> int:
        return self.r2 - self.r1 + 1


@dataclass
class Pattern:
    id: str
    formula: str
    rendered: str
    rects: list[Rect] = field(default_factory=list)

    @property
    def cells(self) -> int:
        return sum(r.area for r in self.rects)


@dataclass
class Inconsistency:
    col: int
    patterns: list[tuple[str, int]]  # (模式 id, 行数)


@dataclass
class FormulaScan:
    patterns: list[Pattern]
    runs_by_col: dict[int, list[Run]]
    inconsistent: list[Inconsistency]
    total_formulas: int
    complete: bool
    patterns_omitted: int = 0


def pattern_id(formula: str) -> str:
    return hashlib.sha1(formula.encode("utf-8")).hexdigest()[:8]


def _split_strings(formula: str) -> list[tuple[bool, str]]:
    """按双引号切段：(是否字符串字面量, 文本)。字符串里的 "RC" 不是引用。"""
    parts, buf, in_str, i = [], "", False, 0
    while i < len(formula):
        ch = formula[i]
        if ch == '"':
            if in_str and i + 1 < len(formula) and formula[i + 1] == '"':
                buf += '""'
                i += 2
                continue
            if in_str:
                parts.append((True, buf + '"'))
                buf = ""
            else:
                if buf:
                    parts.append((False, buf))
                buf = '"'
            in_str = not in_str
            i += 1
            continue
        buf += ch
        i += 1
    if buf:
        parts.append((in_str, buf))
    return parts


def resolve_ref(match: re.Match, anchor_row: int, anchor_col: int, suffix: str = "") -> tuple[int, bool, int, bool]:
    """(工作表行号, 行是否绝对, 工作表列号, 列是否绝对)。suffix 用来读区域引用里第二个端点的组"""
    g = lambda name: match.group(name + suffix)  # noqa: E731
    if g("rabs"):
        row, row_abs = int(g("rabs")), True
    else:
        row, row_abs = anchor_row + int(g("rrel") or 0), False
    if g("cabs"):
        col, col_abs = int(g("cabs")), True
    else:
        col, col_abs = anchor_col + int(g("crel") or 0), False
    return row, row_abs, col, col_abs


def to_a1(formula: str, anchor_row: int, anchor_col: int, label_for=None) -> str:
    """把 R1C1 公式换成锚点格上的 A1 形式。label_for(工作表列号) 返回表头名时，
    同一张表上的同行相对引用渲染成 [表头名]。"""
    out = []
    for is_str, text in _split_strings(formula):
        if is_str:
            out.append(text)
            continue

        def replace(m: re.Match) -> str:
            row, row_abs, col, col_abs = resolve_ref(m, anchor_row, anchor_col)
            cross_sheet = m.start() > 0 and text[m.start() - 1] == "!"
            same_row = not row_abs and int(m.group("rrel") or 0) == 0
            if label_for is not None and same_row and not col_abs and not cross_sheet:
                label = label_for(col)
                if label:
                    return f"[{label}]"
            if row < 1 or col < 1:
                return "#REF!"
            return f"{'$' if col_abs else ''}{col_letters(col)}{'$' if row_abs else ''}{row}"

        out.append(_REF.sub(replace, text))
    return "".join(out)


def column_runs(snap: Snapshot, rect: Rect) -> dict[int, list[Run]]:
    runs: dict[int, list[Run]] = {}
    cells = sorted((c, r, f) for (r, c), f in snap.formulas.items() if rect.covers(r, c))
    for c, r, f in cells:
        col_runs = runs.setdefault(c, [])
        last = col_runs[-1] if col_runs else None
        if last is not None and last.formula == f and last.r2 + 1 == r:
            col_runs[-1] = Run(c, last.r1, r, f)
        else:
            col_runs.append(Run(c, r, r, f))
    return runs


def _rectangles(runs_by_col: dict[int, list[Run]]) -> dict[str, list[Rect]]:
    """行范围相同、公式相同、列相邻的游程合并成矩形"""
    groups: dict[tuple[int, int, str], list[int]] = {}
    for c, runs in runs_by_col.items():
        for run in runs:
            groups.setdefault((run.r1, run.r2, run.formula), []).append(c)
    by_formula: dict[str, list[Rect]] = {}
    for (r1, r2, f), cols in groups.items():
        cols.sort()
        start = prev = cols[0]
        for c in cols[1:] + [None]:
            if c is not None and c == prev + 1:
                prev = c
                continue
            by_formula.setdefault(f, []).append(Rect(r1, start, r2, prev))
            if c is not None:
                start = prev = c
    return by_formula


def label_lookup(snap: Snapshot, header: HeaderInfo | None, rect: Rect):
    """工作表列号 → 表头路径的最后一段（重名时用完整路径）"""
    if header is None or not header.columns:
        return None
    by_sheet_col: dict[int, str] = {}
    leaves: dict[str, int] = {}
    for info in header.columns:
        if info.label:
            leaf = info.label.split("/")[-1]
            leaves[leaf] = leaves.get(leaf, 0) + 1
    for info in header.columns:
        if info.label:
            leaf = info.label.split("/")[-1]
            by_sheet_col[snap.col0 + info.col] = leaf if leaves[leaf] == 1 else info.label
    return lambda col: by_sheet_col.get(col) if rect.c1 <= col - snap.col0 <= rect.c2 else None


def scan_formulas(snap: Snapshot, block: Block, header: HeaderInfo | None) -> FormulaScan:
    rect = block.rect
    runs_by_col = column_runs(snap, rect)
    label_for = label_lookup(snap, header, rect)

    patterns = []
    for f, rects in _rectangles(runs_by_col).items():
        rects.sort(key=lambda r: (r.r1, r.c1))
        anchor = rects[0]
        rendered = to_a1(f, snap.sheet_row(anchor.r1), snap.col0 + anchor.c1, label_for)
        patterns.append(Pattern(pattern_id(f), f, rendered, rects))
    patterns.sort(key=lambda p: -p.cells)

    skip_rows = set(header.header_rows + header.total_rows + header.title_rows) if header else set()
    inconsistent = []
    for c, runs in sorted(runs_by_col.items()):
        counts: dict[str, int] = {}
        for run in runs:
            rows = sum(1 for r in range(run.r1, run.r2 + 1) if r not in skip_rows)
            if rows:
                pid = pattern_id(run.formula)
                counts[pid] = counts.get(pid, 0) + rows
        if len(counts) > 1:
            inconsistent.append(Inconsistency(c, sorted(counts.items(), key=lambda kv: -kv[1])))

    total = sum(run.rows for runs in runs_by_col.values() for run in runs)
    return FormulaScan(
        patterns=patterns[:MAX_PATTERNS],
        runs_by_col=runs_by_col,
        inconsistent=inconsistent,
        total_formulas=total,
        complete=not (snap.truncated or snap.formulas_truncated),
        patterns_omitted=max(0, len(patterns) - MAX_PATTERNS),
    )
