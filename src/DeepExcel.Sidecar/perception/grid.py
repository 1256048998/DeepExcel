"""宿主快照的解析层：把 sheet_snapshot 返回的 JSON 变成可按 (行, 列) 查询的网格。

快照格式（C# ExcelActionsImpl.SheetSnapshot 与 WPS wps-actions.sheetSnapshot 一致）::

    {
      "sheet": "Data",
      "used": "A1:H500",          # 已用区域
      "origin": [1, 1],           # 快照窗口左上角在工作表上的行、列（1 起）
      "total_rows": 500, "total_columns": 8,   # 已用区域尺寸
      "truncated": false,         # 窗口比已用区域小（只读了前面一部分行 / 列）
      "cells": [[...], ...],      # 每格：null / 数字 / 布尔 / "文本" / {"d": "2024-01-31"} / {"e": "#N/A"}
      "formulas": [[r, c, "=RC[-2]*RC[-1]"], ...],   # 窗口内 0 起行列 + R1C1 公式，稀疏
      "formulas_truncated": false,
      "merges": [[r1, c1, r2, c2], ...],             # 窗口内 0 起，含两端
      "merges_truncated": false,
      "spills": [[r1, c1, r2, c2], ...],             # 动态数组溢出区域（宿主读不到时为空）
      "objects": {"tables": [...], "charts": [...], "pivots": [...]}
    }

行列在本包内一律是窗口内 0 起坐标；只有输出给模型的地址才换成工作表上的 A1 地址。
"""

from __future__ import annotations

from dataclasses import dataclass

EMPTY = "empty"
TEXT = "text"
NUMBER = "number"
DATE = "date"
BOOL = "bool"
ERROR = "error"


@dataclass(frozen=True)
class Rect:
    """窗口内的矩形（0 起，含两端）。"""

    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def rows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def cols(self) -> int:
        return self.c2 - self.c1 + 1

    @property
    def area(self) -> int:
        return self.rows * self.cols

    def contains(self, other: "Rect") -> bool:
        return self.r1 <= other.r1 and self.c1 <= other.c1 and self.r2 >= other.r2 and self.c2 >= other.c2

    def covers(self, r: int, c: int) -> bool:
        return self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2

    def union(self, other: "Rect") -> "Rect":
        return Rect(min(self.r1, other.r1), min(self.c1, other.c1), max(self.r2, other.r2), max(self.c2, other.c2))


def col_letters(n: int) -> str:
    """1 → A，27 → AA"""
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_kind(value) -> str:
    if value is None:
        return EMPTY
    if isinstance(value, bool):
        return BOOL
    if isinstance(value, (int, float)):
        return NUMBER
    if isinstance(value, str):
        return TEXT if value.strip() else EMPTY
    if isinstance(value, dict):
        if "d" in value:
            return DATE
        if "e" in value:
            return ERROR
    return TEXT


class Snapshot:
    def __init__(self, data: dict):
        data = data or {}
        self.sheet: str = str(data.get("sheet") or "")
        origin = data.get("origin") or [1, 1]
        self.row0 = int(origin[0])
        self.col0 = int(origin[1])
        self.cells: list = [row if isinstance(row, list) else [] for row in (data.get("cells") or [])]
        self.n_rows = len(self.cells)
        self.n_cols = max((len(row) for row in self.cells), default=0)
        self.truncated = bool(data.get("truncated"))
        self.total_rows = int(data.get("total_rows") or self.n_rows)
        self.total_columns = int(data.get("total_columns") or self.n_cols)
        self.used = data.get("used") or ""

        self.formulas: dict[tuple[int, int], str] = {}
        for entry in data.get("formulas") or []:
            try:
                r, c, f = int(entry[0]), int(entry[1]), str(entry[2])
            except (TypeError, ValueError, IndexError):
                continue
            self.formulas[(r, c)] = f
        self.formulas_truncated = bool(data.get("formulas_truncated"))

        self.merges: list[Rect] = []
        self._merge_at: dict[tuple[int, int], Rect] = {}
        for entry in data.get("merges") or []:
            try:
                rect = Rect(*(int(v) for v in entry[:4]))
            except (TypeError, ValueError):
                continue
            if rect.r2 < rect.r1 or rect.c2 < rect.c1:
                continue
            self.merges.append(rect)
            for r in range(rect.r1, min(rect.r2, self.n_rows - 1) + 1):
                for c in range(rect.c1, min(rect.c2, self.n_cols - 1) + 1):
                    self._merge_at[(r, c)] = rect
        self.merges_truncated = bool(data.get("merges_truncated"))

        self.spills: list[Rect] = []
        for entry in data.get("spills") or []:
            try:
                self.spills.append(Rect(*(int(v) for v in entry[:4])))
            except (TypeError, ValueError):
                continue

        self.objects: dict = data.get("objects") or {}

    # ---- 单格 ----

    def value(self, r: int, c: int):
        if 0 <= r < self.n_rows:
            row = self.cells[r]
            if 0 <= c < len(row):
                return row[c]
        return None

    def kind(self, r: int, c: int) -> str:
        return cell_kind(self.value(r, c))

    def text(self, r: int, c: int) -> str:
        value = self.value(r, c)
        if value is None:
            return ""
        if isinstance(value, dict):
            return str(value.get("d") or value.get("e") or "")
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def formula(self, r: int, c: int) -> str | None:
        return self.formulas.get((r, c))

    def merge_at(self, r: int, c: int) -> Rect | None:
        return self._merge_at.get((r, c))

    def in_spill(self, r: int, c: int) -> bool:
        return any(s.covers(r, c) for s in self.spills)

    def filled(self, r: int, c: int) -> bool:
        """区块检测用的非空掩码：有值、有公式，或落在合并区域里（合并格视为连通）"""
        return self.kind(r, c) != EMPTY or (r, c) in self.formulas or (r, c) in self._merge_at

    # ---- 地址 ----

    def address(self, r: int, c: int) -> str:
        return f"{col_letters(self.col0 + c)}{self.row0 + r}"

    def rect_address(self, rect: Rect) -> str:
        first = self.address(rect.r1, rect.c1)
        if rect.r1 == rect.r2 and rect.c1 == rect.c2:
            return first
        return f"{first}:{self.address(rect.r2, rect.c2)}"

    def column_letter(self, c: int) -> str:
        return col_letters(self.col0 + c)

    def sheet_row(self, r: int) -> int:
        return self.row0 + r
