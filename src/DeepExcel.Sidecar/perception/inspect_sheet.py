"""分层 inspect_sheet：blocks / formulas / objects / dependencies 按需返回。

每层独立标注 status：
- complete  快照覆盖了整张表的已用区域，这一层的结论是全表结论
- partial   快照只读了一部分（大表），结论只针对已读部分；计数类数字不给
- error     这一层算的时候出错了，其他层不受影响
"""

from __future__ import annotations

import re

from .anomalies import find_anomalies
from .blocks import detect_blocks
from .formulas import scan_formulas
from .grid import Snapshot
from .headers import analyze_header

LAYERS = ("blocks", "formulas", "objects", "dependencies")
DEFAULT_LAYERS = ("blocks", "formulas")
MAX_COLUMNS_LISTED = 40

_SHEET_REF = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.一-鿿]+))!")
# 外部引用整段：[book.xlsx]Sheet1! 或 '[book.xlsx]My Sheet'!，先整段剥掉再找本簿的跨表引用
_EXTERNAL = re.compile(r"'?\[([^\]]+\.(?:xlsx|xlsm|xlsb|xls|csv|et))\][^!]*!", re.IGNORECASE)


def normalize_layers(layers) -> list[str]:
    if not layers:
        return list(DEFAULT_LAYERS)
    if isinstance(layers, str):
        layers = [p.strip() for p in re.split(r"[,，\s]+", layers) if p.strip()]
    wanted = [str(x).strip().lower() for x in layers]
    if "all" in wanted:
        return list(LAYERS)
    picked = [x for x in LAYERS if x in wanted]
    return picked or list(DEFAULT_LAYERS)


def coverage(snap: Snapshot) -> dict:
    info = {"complete": not snap.truncated, "used_range": snap.used}
    if snap.n_rows and snap.n_cols:
        info["read_range"] = f"{snap.address(0, 0)}:{snap.address(snap.n_rows - 1, snap.n_cols - 1)}"
    if snap.truncated:
        info["note"] = (f"大表只读了前 {snap.n_rows} 行 × {snap.n_cols} 列（已用区域 {snap.total_rows} 行 × "
                        f"{snap.total_columns} 列），以下结论只针对已读部分；需要后面的内容请用 find 或 read_range(offset)")
    return info


def _status(complete: bool) -> str:
    return "complete" if complete else "partial"


def _analyze_blocks(snap: Snapshot):
    scan = detect_blocks(snap)
    analyzed = []
    for block in scan.blocks:
        header = analyze_header(snap, block) if block.kind == "table" else None
        analyzed.append((block, header))
    return scan, analyzed


def _block_dict(snap: Snapshot, block, header) -> dict:
    out = {
        "range": snap.rect_address(block.rect),
        "kind": block.kind,
        "rows": block.rect.rows,
        "columns": block.rect.cols,
        "density": round(block.density, 2),
    }
    if block.kind == "note":
        out["text"] = " ".join(
            snap.text(r, c) for r in range(block.rect.r1, block.rect.r2 + 1)
            for c in range(block.rect.c1, block.rect.c2 + 1) if snap.text(r, c))[:120]
        return out
    if header is None:
        return out
    if header.title:
        out["title"] = header.title[:120]
    out["header_rows"] = [snap.sheet_row(r) for r in header.header_rows]
    if header.data_start <= header.data_end:
        out["data_rows"] = f"{snap.sheet_row(header.data_start)}-{snap.sheet_row(header.data_end)}"
    if header.total_rows:
        out["total_rows"] = [snap.sheet_row(r) for r in header.total_rows]
    if header.uncertain:
        out["header_uncertain"] = True
        out["header_note"] = "；".join(header.reasons) + "。请用 read_range 读前几行确认表头"
    columns = []
    for col in header.columns[:MAX_COLUMNS_LISTED]:
        entry = {"column": col.letter, "label": col.label, "type": col.type}
        if col.formula:
            entry["formula"] = True
        columns.append(entry)
    out["columns_detail"] = columns
    if len(header.columns) > MAX_COLUMNS_LISTED:
        out["columns_omitted"] = len(header.columns) - MAX_COLUMNS_LISTED
    return out


def _blocks_layer(snap: Snapshot, scan, analyzed) -> dict:
    layer = {
        "status": _status(scan.complete),
        "items": [_block_dict(snap, b, h) for b, h in analyzed],
    }
    if scan.others:
        layer["other_blocks"] = [snap.rect_address(r) for r in scan.others]
        layer["other_blocks_total"] = scan.others_total
    if not scan.complete:
        layer["note"] = "只分析了已读部分，下方可能还有别的区块"
    elif not analyzed:
        layer["note"] = "这张表是空的"
    return layer


def _formulas_layer(snap: Snapshot, analyzed) -> dict:
    complete = not (snap.truncated or snap.formulas_truncated)
    patterns, inconsistent, anomalies = [], [], []
    anomaly_total, anomaly_by_type, anomalies_truncated = 0, {}, False
    total = 0
    for block, header in analyzed:
        if block.kind != "table":
            continue
        scan = scan_formulas(snap, block, header)
        total += scan.total_formulas
        for p in scan.patterns:
            entry = {
                "id": p.id,
                "formula": p.rendered,
                "ranges": [snap.rect_address(r) for r in p.rects[:6]],
            }
            if len(p.rects) > 6:
                entry["ranges_omitted"] = len(p.rects) - 6
            if complete:
                entry["cells"] = p.cells
            patterns.append(entry)
        for inc in scan.inconsistent:
            label = header.label_of(inc.col) if header else None
            item = {"column": snap.column_letter(inc.col), "patterns": [pid for pid, _ in inc.patterns]}
            if label:
                item["label"] = label
            if complete:
                item["rows_per_pattern"] = [rows for _, rows in inc.patterns]
            inconsistent.append(item)
        if header is not None:
            found = find_anomalies(snap, block, header, scan)
            anomaly_total += found["total"]
            for k, v in found["by_type"].items():
                anomaly_by_type[k] = anomaly_by_type.get(k, 0) + v
            anomalies.extend(found["candidates"])
            anomalies_truncated = anomalies_truncated or found["truncated"]

    layer = {"status": _status(complete), "patterns": patterns[:30]}
    if complete:
        layer["formula_cells"] = total
    else:
        layer["note"] = "统计不完整：快照没有覆盖整张表（或公式太多只取了一部分），不给计数"
    if inconsistent:
        layer["inconsistent_columns"] = inconsistent
    layer["anomalies"] = {
        "candidates": anomalies[:30],
        "note": "只是候选，请先向用户确认再修改",
    }
    if complete:
        layer["anomalies"]["total"] = anomaly_total
        layer["anomalies"]["by_type"] = anomaly_by_type
    if anomalies_truncated or len(anomalies) > 30:
        layer["anomalies"]["truncated"] = True
    return layer


def _objects_layer(snap: Snapshot) -> dict:
    objects = snap.objects or {}
    return {
        "status": "complete" if not objects.get("error") else "error",
        "tables": objects.get("tables") or [],
        "charts": objects.get("charts") or [],
        "pivots": objects.get("pivots") or [],
        **({"error": objects["error"]} if objects.get("error") else {}),
    }


def _dependencies_layer(snap: Snapshot) -> dict:
    sheets: dict[str, int] = {}
    external: dict[str, int] = {}
    for formula in snap.formulas.values():
        for m in _EXTERNAL.finditer(formula):
            external[m.group(1)] = external.get(m.group(1), 0) + 1
        stripped = _EXTERNAL.sub("", formula)
        for m in _SHEET_REF.finditer(stripped):
            name = (m.group(1) or m.group(2) or "").replace("''", "'")
            if name and name != snap.sheet:
                sheets[name] = sheets.get(name, 0) + 1
    complete = not (snap.truncated or snap.formulas_truncated)
    layer = {
        "status": _status(complete),
        "sheets_referenced": sorted(sheets, key=lambda k: -sheets[k]),
        "external_workbooks": sorted(external, key=lambda k: -external[k]),
    }
    if complete:
        layer["reference_counts"] = sheets
    return layer


def inspect(snapshot: dict, layers=None) -> dict:
    snap = Snapshot(snapshot)
    wanted = normalize_layers(layers)
    result = {"sheet": snap.sheet, "coverage": coverage(snap), "layers": {}}

    analyzed = None
    if "blocks" in wanted or "formulas" in wanted:
        try:
            scan, analyzed = _analyze_blocks(snap)
            if "blocks" in wanted:
                result["layers"]["blocks"] = _blocks_layer(snap, scan, analyzed)
        except Exception as exc:  # noqa: BLE001 — 每层独立，一层坏了不拖垮其他层
            result["layers"]["blocks"] = {"status": "error", "error": f"{type(exc).__name__}: {exc}"}
    if "formulas" in wanted:
        try:
            if analyzed is None:
                raise RuntimeError("区块检测失败，无法按区块归纳公式")
            result["layers"]["formulas"] = _formulas_layer(snap, analyzed)
        except Exception as exc:  # noqa: BLE001
            result["layers"]["formulas"] = {"status": "error", "error": f"{type(exc).__name__}: {exc}"}
    if "objects" in wanted:
        result["layers"]["objects"] = _objects_layer(snap)
    if "dependencies" in wanted:
        try:
            result["layers"]["dependencies"] = _dependencies_layer(snap)
        except Exception as exc:  # noqa: BLE001
            result["layers"]["dependencies"] = {"status": "error", "error": f"{type(exc).__name__}: {exc}"}
    return result
