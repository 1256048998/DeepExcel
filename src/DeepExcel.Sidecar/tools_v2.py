"""Converged tool set (roadmap reserve item 1, second half) -- behind a switch.

With DEEPEXCEL_TOOLS_V2=1 the model sees four tools in place of 30:

    write_cells     values / formulas / text / number formats / styles / merges /
                    column widths / a new sheet, several regions in one call;
                    if a step fails, the steps before it are rolled back
    edit_structure  add / delete / rename sheets, insert / delete rows and
                    columns, freeze panes, clear a range
    chart           create, combo, data labels, title, colours, export
    pivot           create, refresh, group dates, value display, totals, slicer

Off by default: nothing changes for users until it has been tried on a real
model. Everything below reuses the existing tools -- each step runs the same
handler, and the PreToolUse hook checks each step exactly as if the model had
called the old tool (protected zones, plan mode, the confirmation drawer with
its preview). No second copy of the safety rules.

write_cells is one host call (ToolDispatcher.ExecuteBatch): every step's
read-before-write check first, then one checkpoint, then the steps through the
usual write path; a failing later step restores that checkpoint. The hook still
checks every step (protected zones, plan mode) but asks once, and the host
shows one combined preview (PreviewBuilder.BuildWriteCells).

Known gap: data validation is not offered -- the host has no handler for it yet.
"""
from __future__ import annotations

import json
import os

ENV_SWITCH = "DEEPEXCEL_TOOLS_V2"


def enabled() -> bool:
    return str(os.environ.get(ENV_SWITCH, "")).strip().lower() in ("1", "on", "true", "yes")


# action -> the existing tool that does it
CHART_ACTIONS = {
    "create": "create_chart",
    "combo": "create_combo_chart",
    "data_labels": "add_data_labels",
    "title": "set_chart_title",
    "colors": "set_chart_colors",
    "export": "export_chart",
}
PIVOT_ACTIONS = {
    "create": "create_pivot_table",
    "refresh": "refresh_pivot",
    "group_date": "group_pivot_date",
    "value_display": "set_pivot_value_display",
    "totals": "set_pivot_totals",
    "slicer": "add_pivot_slicer",
}
STRUCTURE_ACTIONS = {
    "add_sheet": "add_sheet",
    "delete_sheet": "delete_sheet",
    "rename_sheet": "rename_sheet",
    "insert_rows": "insert_rows",
    "delete_rows": "delete_rows",
    "insert_columns": "insert_columns",
    "delete_columns": "delete_columns",
    "freeze_panes": "freeze_panes",
    "clear": "clear_range",
}
WRITE_OPS = {
    "values": "write_range",
    "formula": "write_formula",
    "text": "write_value",
    "number_format": "set_number_format",
    "style": "set_cell_style",
    "merge": "merge_cells",
    "unmerge": "unmerge_cells",
    "column_width": "set_column_width",
    "add_sheet": "add_sheet",
}

COMPOSITES = {
    "chart": CHART_ACTIONS,
    "pivot": PIVOT_ACTIONS,
    "edit_structure": STRUCTURE_ACTIONS,
}

REPLACED = frozenset(set(CHART_ACTIONS.values()) | set(PIVOT_ACTIONS.values())
                     | set(STRUCTURE_ACTIONS.values()) | set(WRITE_OPS.values()))

# Old name -> what to use instead; the system prompt prints this so rules that
# still name an old tool stay meaningful.
RENAMES = {}
for _action, _old in WRITE_OPS.items():
    RENAMES[_old] = f"write_cells（type={_action}）"
for _name, _actions in COMPOSITES.items():
    for _action, _old in _actions.items():
        RENAMES.setdefault(_old, f"{_name}（action={_action}）")

MAX_WRITE_OPS = 20
SAMPLE_ROWS = 5
# One host call runs up to MAX_WRITE_OPS writes plus a checkpoint of a possibly large
# workbook. The default 60 s would report a timeout while the host is still writing,
# and the model's retry would then write everything twice.
BATCH_TIMEOUT_SECONDS = 180.0


def is_composite(name: str) -> bool:
    return name == "write_cells" or name in COMPOSITES


def resolve(name: str, args) -> list:
    """composite call -> [(existing tool, its args)]. Raises ValueError with a message for the model."""
    args = args if isinstance(args, dict) else {}
    if name in COMPOSITES:
        table = COMPOSITES[name]
        action = str(args.get("action") or "").strip()
        if action not in table:
            raise ValueError(f"{name} 的 action 必须是：{' / '.join(table)}")
        rest = {k: v for k, v in args.items() if k != "action"}
        return [(table[action], rest)]
    if name == "write_cells":
        ops = args.get("ops")
        if not isinstance(ops, list) or not ops:
            raise ValueError("ops 至少要有一步，例如 [{type: values, address: A1, values: [[...]]}]")
        if len(ops) > MAX_WRITE_OPS:
            raise ValueError(f"一次最多 {MAX_WRITE_OPS} 步，分成几次调用")
        out = []
        for i, op in enumerate(ops, 1):
            if not isinstance(op, dict):
                raise ValueError(f"第 {i} 步不是对象")
            kind = str(op.get("type") or "").strip()
            if kind not in WRITE_OPS:
                raise ValueError(f"第 {i} 步的 type 必须是：{' / '.join(WRITE_OPS)}")
            out.append((WRITE_OPS[kind], {k: v for k, v in op.items() if k != "type"}))
        return out
    raise ValueError(f"{name} 不是合并工具")


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

def _tool(name: str):
    import excel_tools
    for t in excel_tools.base_tools():
        if t.name == name:
            return t
    raise KeyError(name)


def _payload(wrapped) -> dict:
    try:
        return json.loads(wrapped["content"][0]["text"])
    except Exception:  # noqa: BLE001
        return {"success": False, "error": "工具没有返回可读的结果"}


async def _run_step(tool_name: str, tool_args: dict) -> dict:
    """Run one existing tool. A missing parameter raises inside the old handler
    (the schema check that would have caught it only covers the composite), so
    turn that into an ordinary failure instead of escaping past the rollback."""
    try:
        return _payload(await _tool(tool_name).handler(tool_args))
    except (KeyError, TypeError, ValueError) as exc:
        return {"success": False, "error": f"参数缺失或不对（{type(exc).__name__}: {exc}）",
                "suggestion": "照这一步 type 对应的参数写全后重试"}


def _column_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _sample_address(kind: str, tool_args: dict) -> str | None:
    """What to read back after a step: the written block (at most SAMPLE_ROWS rows)."""
    address = tool_args.get("address")
    if not isinstance(address, str) or not address.strip():
        return None
    values = tool_args.get("values")
    if kind != "values" or not isinstance(values, list) or not values:
        return address
    try:
        from perception.postwrite import parse_a1
        parsed = parse_a1(address)
    except Exception:  # noqa: BLE001
        parsed = None
    if parsed is None:
        return address
    sheet, r1, c1, _, _ = parsed
    width = max((len(row) for row in values if isinstance(row, list)), default=1)
    r2 = r1 + min(len(values), SAMPLE_ROWS) - 1
    c2 = c1 + max(1, width) - 1
    block = f"{_column_letters(c1)}{r1}:{_column_letters(c2)}{r2}"
    if sheet:
        quoted = sheet if sheet.replace("_", "").isalnum() else "'" + sheet.replace("'", "''") + "'"
        block = f"{quoted}!{block}"
    return block


async def run_composite(name: str, args) -> dict:
    import excel_tools
    try:
        [(tool_name, tool_args)] = resolve(name, args)
    except ValueError as exc:
        return excel_tools._wrap_result({"success": False, "error": str(exc)})
    return excel_tools._wrap_result(await _run_step(tool_name, tool_args))


async def run_write_cells(args) -> dict:
    """One host call (ToolDispatcher.ExecuteBatch): one checkpoint for the whole call,
    every step through the usual write path, the whole batch undone if a later step
    fails. The sidecar then adds what the host cannot: the formula pattern check
    and a readback of what was written."""
    import excel_tools
    from ipc import call_csharp
    try:
        ops = resolve("write_cells", args)
    except ValueError as exc:
        return excel_tools._wrap_result({"success": False, "error": str(exc)})

    result = await call_csharp("write_cells", {"ops": args["ops"]}, timeout=BATCH_TIMEOUT_SECONDS)
    if not isinstance(result, dict) or result.get("success") is not True:
        return excel_tools._wrap_result(result if isinstance(result, dict)
                                        else {"success": False, "error": "宿主没有返回结果"})

    kinds = [str(op.get("type")) for op in args["ops"]]
    for (tool_name, tool_args), kind in zip(ops, kinds):
        address = tool_args.get("address")
        if not isinstance(address, str):
            continue
        if kind == "formula":
            result = await excel_tools._pattern_check(result, address)
        elif kind == "values" and excel_tools._has_formula(tool_args.get("values")):
            values = tool_args["values"]
            width = max((len(row) for row in values if isinstance(row, list)), default=1)
            result = await excel_tools._pattern_check(result, address, rows=len(values), columns=width)

    samples = []
    for (tool_name, tool_args), kind in zip(ops, kinds):
        address = _sample_address(kind, tool_args) if kind in ("values", "formula", "text") else None
        if address:
            read = await _run_step("read_range", {"address": address, "limit": SAMPLE_ROWS})
            data = read.get("data") if isinstance(read.get("data"), dict) else {}
            if read.get("success") is not False and data.get("values") is not None:
                samples.append({"address": address, "values": data.get("values")})
    data = result.get("data") if isinstance(result.get("data"), dict) else {}
    result["data"] = {**data, "samples": samples}
    return excel_tools._wrap_result(result)


# ---------------------------------------------------------------------------
# Tool objects
# ---------------------------------------------------------------------------

_PY_TO_JSON = {str: "string", int: "integer", float: "number", bool: "boolean", list: "array", dict: "object"}


def _properties(tool_name: str) -> dict:
    schema = _tool(tool_name).input_schema or {}
    if isinstance(schema, dict) and schema.get("type") == "object":
        return dict(schema.get("properties") or {})
    return {key: {"type": _PY_TO_JSON.get(value, "string")} for key, value in schema.items()}


def _union(tool_names) -> dict:
    props = {}
    for name in tool_names:
        for key, spec in _properties(name).items():
            props.setdefault(key, spec)
    return props


def _action_doc(table: dict) -> str:
    parts = []
    for action, old in table.items():
        keys = "、".join(_properties(old)) or "无"
        parts.append(f"{action}（{keys}）")
    return "；".join(parts)


def build_tools() -> list:
    from claude_agent_sdk import tool
    from excel_tools import _CHECK_PLAIN, _CHECK_WRITE, _describe

    made = []
    for name, table, what, use in (
        ("chart", CHART_ACTIONS, "图表的创建和修改，action 选一种",
         "建图、组合图、加数据标签、改标题和配色、导出图片"),
        ("pivot", PIVOT_ACTIONS, "数据透视表的创建和调整，action 选一种",
         "按维度汇总、刷新、按年/季/月分组、占比/累计/排名、总计、切片器"),
        ("edit_structure", STRUCTURE_ACTIONS, "工作表和行列结构的修改，action 选一种",
         "增删改工作表、插入删除行列、冻结窗格、清空区域"),
    ):
        props = {"action": {"type": "string", "enum": list(table)}}
        props.update(_union(table.values()))
        description = _describe(
            what + "。各 action 的参数：" + _action_doc(table),
            use=use,
            check=_CHECK_PLAIN + "；删除类 action 会让引用它的公式变 #REF!，删前先确认",
        )

        def make(bound_name):
            async def handler(args):
                return await run_composite(bound_name, args)
            return handler

        made.append(tool(name, description, {"type": "object", "properties": props, "required": ["action"]})(make(name)))

    op_props = {"type": {"type": "string", "enum": list(WRITE_OPS)}}
    op_props.update(_union(WRITE_OPS.values()))
    write_cells_description = _describe(
        "一次调用写多步：ops 按顺序执行，每步 {type, ...参数}。type 的参数：" + _action_doc(WRITE_OPS)
        + "。values 的 values 是二维数组，以 = 开头的元素按公式写入；text 写纯文本，不解析为公式",
        use="一次把表头、数据、公式、格式写好（如新建汇总表：add_sheet → values → style → number_format → column_width）",
        avoid="整列套同一公式（写一格后用 fill_formula_down）",
        check=_CHECK_WRITE + "；data.samples 是写完后重新读到的前几行，核对值对不对；checkpoint_id 撤销的是整个这次调用",
        on_fail="error 写明第几步失败；前面的步骤会自动撤销，改正那一步后整批重试",
    )

    async def write_cells(args):
        return await run_write_cells(args)

    made.append(tool("write_cells", write_cells_description, {
        "type": "object",
        "properties": {"ops": {"type": "array", "items": {"type": "object", "properties": op_props, "required": ["type"]}}},
        "required": ["ops"],
    })(write_cells))
    return made


AVAILABLE_TOOLS_V2 = """<available-tools>
查找与浏览：find（按文本搜全工作簿的值或公式）/ list（列出工作表、名称、表格、透视表、图表）/ inspect_sheet（分析一张表的区块、多级表头、公式模式和异常候选）/ explore_workbook（表多或表大时，拆成子任务派只读子 agent 并行摸底，只交回结论）
数据读取：read_workbook / read_selection / read_range / read_attachment
写入：write_cells（值、公式、文本、数字格式、样式、合并、列宽、新建表，一次可写多步）/ fill_formula_down / replace_formula / copy_range
数据处理：clean_data（去重用它的 remove_duplicates 操作）/ sort_data / filter_data
数据清洗：delete_blank_rows / split_text_to_columns / fill_blank_cells / highlight_duplicates / remove_special_chars / clean_amount / merge_columns / rename_columns / collapse_spaces
格式化：apply_conditional_format / write_table
结构操作：edit_structure（增删改工作表、插入删除行列、冻结窗格、清空区域）
图表：chart（创建、组合图、数据标签、标题、配色、导出）
透视表：pivot（创建、刷新、日期分组、值显示方式、总计、切片器）
代码执行：execute_vba / execute_python（会弹安全确认窗；execute_python 只能做纯计算，碰不到工作簿）
快照：修改类工具执行前会自动备份（结果里的 backup_snapshot_id）；撤销用 rollback，create_snapshot 只在需要额外检查点时用
Computer Use：screenshot_excel / send_keys（截图 Excel 界面 + 模拟键盘，用于操作对话框/快捷键/弹窗）
其他：clarify_intent / todo_write（三步以上的任务先列计划，边做边更新状态）/ present_plan（提交变更方案等用户批准；「只出方案」模式下必须用它收尾）
记忆：update_workbook_notes（整份重写这个工作簿的记忆 NOTES.md，下次打开会话时自动带给你）
知识：load_skill（读 <knowledge-skills> 里列出的专业知识；任务涉及其中的主题时动手前先读）
</available-tools>"""


def rename_note() -> str:
    rows = "\n".join(f"- {old} → {new}" for old, new in sorted(RENAMES.items()))
    return ("<tool-renames>\n下面的规则和示例里还会出现旧的工具名，它们已经合并成新工具，按这张表换着用"
            "（参数名不变，放进 ops 的一步或加上 action）：\n" + rows + "\n</tool-renames>")


def system_prompt(base: str) -> str:
    """把 v1 的提示词改成 v2：换掉工具清单，并附上旧名对照表。"""
    head, rest = base.split("<available-tools>", 1)
    _, tail = rest.split("</available-tools>", 1)
    return head + AVAILABLE_TOOLS_V2 + "\n\n" + rename_note() + tail
