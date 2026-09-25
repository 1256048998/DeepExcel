"""公式异常候选：只报候选、不自动修复。

- isolated     孤立偏离：一列公式里夹着一两格模式不同的（上下是同一个模式）
- hardcoded    疑似被改成死值：一列公式中间有一格是常量数字，上下公式是同一个模式
- total_gap    合计范围漏行：合计行的 SUM 没覆盖到它上面那段数据的首行或末行
- broken_ref   断链引用：公式里有 #REF!，或引用外部工作簿且结果是错误值

排除表头、合计行（前两类）、合并格、动态数组溢出区域。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .blocks import Block
from .formulas import _REF, FormulaScan, resolve_ref, to_a1
from .grid import ERROR, NUMBER, Snapshot
from .headers import HeaderInfo

MAX_CANDIDATES = 30
MAX_PER_COLUMN = 5
# 孤立偏离：中间那段最多几行，上下同模式加起来至少几行
ISOLATED_MAX_ROWS = 2
ISOLATED_MIN_CONTEXT = 3

_RANGE = re.compile(_REF.pattern + r"\s*:\s*" + _REF.pattern.replace("?P<rrel>", "?P<rrel2>")
                    .replace("?P<rabs>", "?P<rabs2>").replace("?P<crel>", "?P<crel2>").replace("?P<cabs>", "?P<cabs2>"))


@dataclass
class Candidate:
    type: str
    address: str
    detail: str
    formula: str | None = None

    def to_dict(self) -> dict:
        out = {"type": self.type, "address": self.address, "detail": self.detail}
        if self.formula:
            out["formula"] = self.formula
        return out


class _Collector:
    def __init__(self):
        self.items: list[Candidate] = []
        self.counts: dict[str, int] = {}
        self.per_column: dict[tuple[str, int], int] = {}

    def add(self, col: int, candidate: Candidate):
        self.counts[candidate.type] = self.counts.get(candidate.type, 0) + 1
        key = (candidate.type, col)
        self.per_column[key] = self.per_column.get(key, 0) + 1
        if self.per_column[key] <= MAX_PER_COLUMN and len(self.items) < MAX_CANDIDATES:
            self.items.append(candidate)

    @property
    def total(self) -> int:
        return sum(self.counts.values())


def _excluded(snap: Snapshot, r: int, c: int) -> bool:
    return snap.merge_at(r, c) is not None or snap.in_spill(r, c)


def _column_name(header: HeaderInfo, snap: Snapshot, c: int) -> str:
    label = header.label_of(c) if header else None
    return f"{snap.column_letter(c)} 列（{label}）" if label else f"{snap.column_letter(c)} 列"


def _a1(snap: Snapshot, formula: str, r: int, c: int) -> str:
    return to_a1(formula, snap.sheet_row(r), snap.col0 + c)


def _isolated(snap, header, runs_by_col, skip_rows, out: _Collector):
    for c, runs in runs_by_col.items():
        for i in range(1, len(runs) - 1):
            before, run, after = runs[i - 1], runs[i], runs[i + 1]
            if before.formula != after.formula or run.formula == before.formula:
                continue
            if before.r2 + 1 != run.r1 or run.r2 + 1 != after.r1:
                continue
            if run.rows > ISOLATED_MAX_ROWS or before.rows + after.rows < ISOLATED_MIN_CONTEXT:
                continue
            rows = [r for r in range(run.r1, run.r2 + 1) if r not in skip_rows and not _excluded(snap, r, c)]
            for r in rows:
                out.add(c, Candidate(
                    "isolated", snap.address(r, c),
                    f"{_column_name(header, snap, c)}上下都是 {_a1(snap, before.formula, r, c)}，这一格不同",
                    snap.formulas.get((r, c)) and _a1(snap, snap.formulas[(r, c)], r, c),
                ))


def _hardcoded(snap, header, runs_by_col, skip_rows, out: _Collector):
    for c, runs in runs_by_col.items():
        formula_rows = {r: run for run in runs for r in range(run.r1, run.r2 + 1)}
        if len(formula_rows) < 3:
            continue
        first, last = min(formula_rows), max(formula_rows)
        for r in range(first + 1, last):
            if r in formula_rows or r in skip_rows or _excluded(snap, r, c):
                continue
            if snap.kind(r, c) != NUMBER:
                continue
            above = next((formula_rows[x] for x in range(r - 1, first - 1, -1) if x in formula_rows), None)
            below = next((formula_rows[x] for x in range(r + 1, last + 1) if x in formula_rows), None)
            if above is None or below is None or above.formula != below.formula:
                continue
            out.add(c, Candidate(
                "hardcoded", snap.address(r, c),
                f"{_column_name(header, snap, c)}上下都是公式 {_a1(snap, above.formula, r, c)}，这一格是数值 {snap.text(r, c)}",
            ))


def _segment_rows(header: HeaderInfo, total_row: int) -> tuple[int, int]:
    """合计行上面那一段数据：从上一个合计行（或数据起始行）到合计行前一行"""
    previous = [t for t in header.total_rows if t < total_row]
    start = (max(previous) + 1) if previous else header.data_start
    return start, total_row - 1


def _total_gap(snap, header, out: _Collector, block: Block):
    for t in header.total_rows:
        start, end = _segment_rows(header, t)
        if end < start:
            continue
        for c in range(block.rect.c1, block.rect.c2 + 1):
            formula = snap.formulas.get((t, c))
            if not formula:
                continue
            filled = [r for r in range(start, end + 1)
                      if snap.kind(r, c) != "empty" or (r, c) in snap.formulas]
            if not filled:
                continue
            for m in _RANGE.finditer(formula):
                row_a, _, col_a, _ = resolve_ref(m, snap.sheet_row(t), snap.col0 + c)
                row_b, _, col_b, _ = resolve_ref(m, snap.sheet_row(t), snap.col0 + c, suffix="2")
                if col_a != snap.col0 + c or col_b != snap.col0 + c:
                    continue  # 只看对本列的纵向求和
                lo, hi = sorted((row_a - snap.row0, row_b - snap.row0))
                if hi < start or lo > end:
                    continue  # 不是在汇总这一段
                missing = [r for r in filled if r < lo or r > hi]
                if missing:
                    rows = "、".join(str(snap.sheet_row(r)) for r in missing[:5])
                    more = f" 等 {len(missing)} 行" if len(missing) > 5 else ""
                    out.add(c, Candidate(
                        "total_gap", snap.address(t, c),
                        f"合计只汇总了第 {snap.sheet_row(lo)}–{snap.sheet_row(hi)} 行，漏了第 {rows} 行{more}",
                        _a1(snap, formula, t, c),
                    ))
                break


def _broken(snap, header, block: Block, out: _Collector):
    for (r, c), formula in snap.formulas.items():
        if not block.rect.covers(r, c):
            continue
        if "#REF!" in formula:
            out.add(c, Candidate("broken_ref", snap.address(r, c), "公式里有 #REF!（引用的单元格或工作表已被删除）",
                                 _a1(snap, formula, r, c)))
        elif "[" in formula.replace("R[", "").replace("C[", "") and snap.kind(r, c) == ERROR:
            out.add(c, Candidate("broken_ref", snap.address(r, c),
                                 f"引用了外部工作簿，结果是 {snap.text(r, c)}（外部文件可能已移动或删除）",
                                 _a1(snap, formula, r, c)))


def find_anomalies(snap: Snapshot, block: Block, header: HeaderInfo, scan: FormulaScan) -> dict:
    out = _Collector()
    skip_rows = set(header.header_rows + header.title_rows + header.total_rows)
    _broken(snap, header, block, out)
    _total_gap(snap, header, out, block)
    _isolated(snap, header, scan.runs_by_col, skip_rows, out)
    _hardcoded(snap, header, scan.runs_by_col, skip_rows, out)
    return {
        "total": out.total,
        "by_type": out.counts,
        "candidates": [c.to_dict() for c in out.items],
        "truncated": out.total > len(out.items),
    }

