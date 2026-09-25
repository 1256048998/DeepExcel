"""区块检测：一张表上往往不止一个表格（标题、主表、旁边的小计表、底部备注）。

做法：
1. 非空掩码（有值 / 有公式 / 落在合并区域里）上做 8 连通区域；
2. 包围盒被另一块完全包含的孤岛并入外层块（表格中间的一小段空白不会把它切碎）；
3. 只隔 1 个空行、列跨度高度重合的上下两块合并（表格里常见的一行空隔断）；
4. 块太多时只详细保留主要块，其余只列范围。
"""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field

from .grid import Rect, Snapshot

MAX_BLOCKS = 12
MAX_LISTED_OTHERS = 40
# 上下两块只隔一行空行时，列跨度重合比例达到这个值才合并
GAP_MERGE_OVERLAP = 0.8


@dataclass
class Block:
    rect: Rect
    cells: int  # 非空格数（合并区域的每一格都计入）
    kind: str = "table"  # table | note
    extra: dict = field(default_factory=dict)

    @property
    def density(self) -> float:
        return self.cells / self.rect.area if self.rect.area else 0.0


@dataclass
class BlockScan:
    blocks: list[Block]
    others: list[Rect]  # 块太多时，没有详细分析的次要块
    others_total: int
    complete: bool  # False：快照只覆盖了已用区域的一部分，块可能被截断


def _components(snap: Snapshot) -> list[tuple[Rect, int]]:
    """8 连通区域。visited 保证每格只访问、只计数一次。"""
    visited = [[False] * snap.n_cols for _ in range(snap.n_rows)]
    found = []
    for r0 in range(snap.n_rows):
        for c0 in range(snap.n_cols):
            if visited[r0][c0] or not snap.filled(r0, c0):
                continue
            visited[r0][c0] = True
            queue = deque([(r0, c0)])
            r1, c1, r2, c2 = r0, c0, r0, c0
            count = 0
            while queue:
                r, c = queue.popleft()
                count += 1
                r1, c1, r2, c2 = min(r1, r), min(c1, c), max(r2, r), max(c2, c)
                for dr in (-1, 0, 1):
                    for dc in (-1, 0, 1):
                        nr, nc = r + dr, c + dc
                        if (dr or dc) and 0 <= nr < snap.n_rows and 0 <= nc < snap.n_cols \
                                and not visited[nr][nc] and snap.filled(nr, nc):
                            visited[nr][nc] = True
                            queue.append((nr, nc))
            found.append((Rect(r1, c1, r2, c2), count))
    return found


def _absorb_contained(parts: list[list]) -> list[list]:
    """包围盒被别的块包含的孤岛并入外层。parts 元素是 [rect, cells]"""
    parts = sorted(parts, key=lambda p: -p[0].area)
    kept: list[list] = []
    for rect, cells in parts:
        host = next((k for k in kept if k[0].contains(rect)), None)
        if host is not None:
            host[1] += cells
        else:
            kept.append([rect, cells])
    return kept


def _overlap_ratio(a: Rect, b: Rect) -> float:
    overlap = min(a.c2, b.c2) - max(a.c1, b.c1) + 1
    if overlap <= 0:
        return 0.0
    return overlap / min(a.cols, b.cols)


def _merge_one_row_gaps(parts: list[list]) -> list[list]:
    changed = True
    while changed:
        changed = False
        parts.sort(key=lambda p: (p[0].r1, p[0].c1))
        for i, upper in enumerate(parts):
            for j, lower in enumerate(parts):
                if i == j:
                    continue
                if lower[0].r1 - upper[0].r2 == 2 and _overlap_ratio(upper[0], lower[0]) >= GAP_MERGE_OVERLAP \
                        and min(upper[0].cols, lower[0].cols) >= 2:
                    upper[0] = upper[0].union(lower[0])
                    upper[1] += lower[1]
                    parts.pop(j)
                    parts[:] = _absorb_contained(parts)
                    changed = True
                    break
            if changed:
                break
    return parts


def _classify(snap: Snapshot, rect: Rect, cells: int) -> str:
    # 单独一格或一行寥寥几格的文字：标题、说明、单位注释
    if rect.rows == 1 and cells <= 2:
        return "note"
    if rect.area == 1:
        return "note"
    return "table"


def detect_blocks(snap: Snapshot) -> BlockScan:
    parts = [[rect, cells] for rect, cells in _components(snap)]
    parts = _absorb_contained(parts)
    parts = _merge_one_row_gaps(parts)

    ranked = sorted(parts, key=lambda p: -p[1])
    main = ranked[:MAX_BLOCKS]
    rest = ranked[MAX_BLOCKS:]
    blocks = [Block(rect, cells, _classify(snap, rect, cells)) for rect, cells in main]
    blocks.sort(key=lambda b: (b.rect.r1, b.rect.c1))
    others = sorted((p[0] for p in rest), key=lambda r: (r.r1, r.c1))
    return BlockScan(
        blocks=blocks,
        others=others[:MAX_LISTED_OTHERS],
        others_total=len(others),
        complete=not snap.truncated,
    )
