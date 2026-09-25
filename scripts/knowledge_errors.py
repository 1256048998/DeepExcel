#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rewrite each knowledge skill's "errors users actually hit" section from telemetry.

The hand-written 常见报错 tables in the skills are guesses. This replaces a
marked block in each SKILL.md with what the field reports: tool failures
grouped by (tool, error code), exported by the admin API.

    # 1. export (admin token; the export holds only tool names and fixed codes)
    curl -H "Authorization: Bearer $ADMIN" https://api.example.com/admin/api/tool-errors?days=30 > tool-errors.json
    # 2. rewrite the blocks, review the diff
    python scripts/knowledge_errors.py --errors tool-errors.json
    git diff src/DeepExcel.Sidecar/knowledge
    # 3. ship: python scripts/knowledge_pack.py build ...

A skill opts in by listing the tools it covers in its frontmatter
(``tools: write_formula, fill_formula_down``). A changed block bumps the
skill's version, so clients that already have the pack pick up the new one.
Codes without an entry in CODE_HINTS are still listed, and reported on stdout
so someone writes the advice.
"""

import argparse
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_SOURCE = os.path.join(ROOT, "src", "DeepExcel.Sidecar", "knowledge")

BEGIN = "<!-- telemetry-errors:begin -->"
END = "<!-- telemetry-errors:end -->"

# error_code -> (类别, 怎么办). Codes come from MessageBridge.ClassifyError and
# the sidecar's tool_end error codes (docs/ui-event-protocol.md).
CODE_HINTS = {
    "formula_name": ("公式 #NAME?", "查全角标点、函数名拼写、新函数在旧版本里不存在、文本没加引号"),
    "formula_value": ("公式 #VALUE!", "文本参与了运算：先查文本型数字、带单位的金额"),
    "formula_ref": ("公式 #REF!", "引用的行列或工作表被删了：读出错格的公式，改引用"),
    "formula_div0": ("公式 #DIV/0!", "分母为 0 或空：用 IF 先判断"),
    "formula_na": ("公式 #N/A", "查找值不存在：先查两边类型、首尾空格、精确匹配参数"),
    "formula_spill": ("公式 #SPILL!", "动态数组的溢出区域被占用：换位置或清空溢出区"),
    "unread_target": ("目标区域没读过", "先 read_range 看一眼目标区域再写"),
    "protected_zone": ("写到了禁区", "告诉用户这里被锁定，不要换工具绕过"),
    "bad_arguments": ("参数不对", "对照工具说明补齐必填参数，可选参数不要传占位值"),
    "type_mismatch": ("类型不匹配", "值的类型和预期不符：日期、数字、文本先统一再处理"),
    "vba_compile": ("VBA 编译错误", "按报错行修；注意中文字符串、续行和 End Sub"),
    "vba_error": ("VBA 运行错误", "执行后读回该改的区域，确认改到了哪一步"),
    "range_out_of_bounds": ("区域越界", "先 list / read_range 确认已用区域的边界"),
    "not_found": ("找不到对象", "表名、名称、图表名先 list 确认，注意全角半角和空格"),
    "timeout": ("超时", "缩小处理范围，或分批做"),
    "com_error": ("Excel 拒绝了操作", "可能有对话框、单元格在编辑或工作表受保护：请用户看一眼 Excel"),
    "permission_denied": ("没有权限", "工作表 / 工作簿受保护，或 VBA 工程访问未开启：告诉用户怎么开"),
    "denied": ("用户拒绝了这一步", "换一种更保守的做法，或先问清楚用户要什么"),
    "no_result": ("没有返回结果", "Excel 可能被对话框挡住：请用户看一眼"),
    "aborted": ("中途出错", "从最后一个成功的步骤重新检查"),
}

_FRONTMATTER = re.compile(r"^---\s*\n(?P<meta>.*?)\n---\s*\n", re.S)
_VERSION_LINE = re.compile(r"^version:\s*(\d+)\s*$", re.M)


def skill_tools(text):
    match = _FRONTMATTER.match(text.replace("\r\n", "\n"))
    if not match:
        return []
    for line in match.group("meta").splitlines():
        key, sep, value = line.partition(":")
        if sep and key.strip() == "tools":
            return [name.strip() for name in value.split(",") if name.strip()]
    return []


def render_block(errors, tools, days, min_count=3, top=8):
    rows = [e for e in errors if e.get("tool_name") in tools and int(e.get("count") or 0) >= min_count]
    rows.sort(key=lambda e: (-int(e["count"]), e["tool_name"], e.get("error_code") or ""))
    rows = rows[:top]
    if not rows:
        return None, []
    lines = [
        BEGIN,
        "## 用户实际遇到的报错（近 %d 天）" % days,
        "",
        "这一节由遥测统计生成（工具 × 报错类别 × 次数），排在前面的最常见，动手时优先避开。",
        "",
        "| 工具 | 报错类别 | 次数 | 怎么办 |",
        "|---|---|---|---|",
    ]
    unknown = []
    for row in rows:
        code = row.get("error_code") or "unknown"
        label, hint = CODE_HINTS.get(code, (code, "看工具返回的 error 和 suggestion"))
        if code not in CODE_HINTS:
            unknown.append(code)
        lines.append("| %s | %s | %d | %s |" % (row["tool_name"], label, int(row["count"]), hint))
    lines.append(END)
    return "\n".join(lines), unknown


def apply_block(text, block):
    """Returns (new_text, changed). block=None removes an existing block."""
    normalized = text.replace("\r\n", "\n")
    start, end = normalized.find(BEGIN), normalized.find(END)
    if start >= 0 and end > start:
        current = normalized[start:end + len(END)]
        if block == current:
            return normalized, False
        before, after = normalized[:start].rstrip("\n"), normalized[end + len(END):].lstrip("\n")
        if block is None:
            updated = before + "\n" + (("\n" + after) if after else "")
        else:
            updated = before + "\n\n" + block + "\n" + (("\n" + after) if after else "")
    else:
        if block is None:
            return normalized, False
        updated = normalized.rstrip("\n") + "\n\n" + block + "\n"
    return bump_version(updated), True


def bump_version(text):
    match = _VERSION_LINE.search(text)
    if not match:
        return text
    return text[:match.start()] + "version: %d" % (int(match.group(1)) + 1) + text[match.end():]


def run(errors_path, source=DEFAULT_SOURCE, min_count=3, dry_run=False):
    with open(errors_path, "r", encoding="utf-8") as stream:
        export = json.load(stream)
    errors, days = export.get("errors") or [], int(export.get("days") or 30)
    changed, unknown = [], set()
    for name in sorted(os.listdir(source)):
        path = os.path.join(source, name, "SKILL.md")
        if not os.path.isfile(path):
            continue
        with open(path, "r", encoding="utf-8") as stream:
            text = stream.read()
        tools = skill_tools(text)
        if not tools:
            continue
        block, missing = render_block(errors, tools, days, min_count)
        unknown.update(missing)
        updated, did_change = apply_block(text, block)
        if did_change:
            changed.append(name)
            if not dry_run:
                with open(path, "w", encoding="utf-8", newline="\n") as stream:
                    stream.write(updated)
    return changed, sorted(unknown)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--errors", required=True, help="output of GET /admin/api/tool-errors")
    parser.add_argument("--source", default=DEFAULT_SOURCE)
    parser.add_argument("--min-count", type=int, default=3)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    changed, unknown = run(args.errors, args.source, args.min_count, args.dry_run)
    print("[knowledge-errors] %s: %s" % ("would update" if args.dry_run else "updated",
                                         ", ".join(changed) if changed else "nothing"))
    if unknown:
        print("[knowledge-errors] no advice yet for: %s -- add them to CODE_HINTS" % ", ".join(unknown))
    return 0


if __name__ == "__main__":
    sys.exit(main())
