# src/DeepExcel.Sidecar/excel_tools.py
import json
from claude_agent_sdk import tool
from ipc import call_csharp, call_csharp_clarify


def _wrap_result(csharp_result: dict) -> dict:
    """把 C# 返回的 dict 包装成 MCP 工具返回格式。
    ★ 当 C# 返回 success=false 时设置 is_error=true，
    防止 LLM 把失败当成功继续推理，触发死循环重试。"""
    is_error = False
    if isinstance(csharp_result, dict) and csharp_result.get("success") is False:
        is_error = True
    return {
        "content": [
            {"type": "text", "text": json.dumps(csharp_result, ensure_ascii=False)}
        ],
        "is_error": is_error,
    }


# 写后模式检查只读这么多格：写入后每次都要多取一次快照，要快
POSTWRITE_SNAPSHOT_CELLS = 20000


async def _pattern_check(result: dict, address: str, rows: int = 1, columns: int = 1) -> dict:
    """写公式成功后，对被写入的列跑一遍公式模式检查，结论并入 verification。

    写入区域里出现孤立偏离 / 死值 / 合计漏行时，verification.ok 置为 false 并在 summary
    里点名，模型按规则 8 会先修再汇报。检查本身失败（快照读不到等）一律静默跳过，
    不能让体检把一次成功的写入变成失败。"""
    if not isinstance(result, dict) or result.get("success") is not True:
        return result
    try:
        from perception.postwrite import parse_a1, pattern_check, summary_line
        parsed = parse_a1(address)
        if parsed is None:
            return result
        sheet, r1, c1, r2, c2 = parsed
        r2 = max(r2, r1 + max(1, rows) - 1)
        c2 = max(c2, c1 + max(1, columns) - 1)
        payload = {"max_cells": POSTWRITE_SNAPSHOT_CELLS}
        if sheet:
            payload["sheet"] = sheet
        snapshot = await call_csharp("sheet_snapshot", payload)
        if not isinstance(snapshot, dict) or snapshot.get("success") is not True:
            return result
        check = pattern_check(snapshot.get("data") or {}, (r1, c1, r2, c2))
    except Exception:  # noqa: BLE001
        return result
    if not check:
        return result
    verification = result.get("verification")
    if not isinstance(verification, dict):
        verification = {"ok": True}
    verification["pattern_check"] = check
    line = summary_line(check)
    if line:
        verification["ok"] = False
        previous = verification.get("summary")
        verification["summary"] = f"{previous}；{line}" if previous else line
    result["verification"] = verification
    return result


def _has_formula(values) -> bool:
    return isinstance(values, list) and any(
        isinstance(v, str) and v.startswith("=")
        for row in values for v in (row if isinstance(row, list) else [row]))


def _optional_int(args: dict, key: str):
    value = args.get(key)
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


@tool(
    "read_range",
    "读取指定范围的单元格数据。整列/整行（A:A、1:1）会自动收到已用区域；一次最多 200 行（limit 可放大到 500），"
    "结果里 paging.next_offset 不为空就说明还有下一页，用同一个 address 加 offset 继续读。"
    "找东西在哪用 find，看有哪些表/名称/图表用 list，不要为了找数据把整张表读一遍。",
    {
        "type": "object",
        "properties": {
            "address": {"type": "string", "description": "区域地址，如 A1:D50、Sheet2!A:C、命名区域"},
            "offset": {"type": "integer", "description": "从区域第几行（0 起）开始读，翻页用"},
            "limit": {"type": "integer", "description": "本页最多读几行，默认 200，上限 500"},
        },
        "required": ["address"],
    },
)
async def read_range(args):
    payload = {"address": args["address"]}
    for key in ("offset", "limit"):
        value = _optional_int(args, key)
        if value is not None:
            payload[key] = value
    result = await call_csharp("read_range", payload)
    return _wrap_result(result)


@tool(
    "find",
    "在工作簿里搜索文本（相当于 Ctrl+F 的「查找全部」）：返回每个匹配的工作表、地址、值和公式。"
    "定位某个科目/客户/关键字、查哪些公式引用了某张表时先用它，比逐页 read_range 快得多。"
    "total 是真实命中数，truncated=true 表示只列出了前 max_results 个。",
    {
        "type": "object",
        "properties": {
            "query": {"type": "string", "description": "要找的文本（不区分大小写）"},
            "scope": {"type": "string", "enum": ["values", "formulas"], "description": "values 搜显示的值（默认）；formulas 搜公式文本，如搜 Sheet2! 找跨表引用"},
            "match": {"type": "string", "enum": ["contains", "exact"], "description": "contains 包含即算（默认）；exact 整个单元格等于 query"},
            "sheets": {"type": "array", "items": {"type": "string"}, "description": "只搜这些工作表；不传搜全部"},
            "max_results": {"type": "integer", "description": "最多列出几个匹配，默认 50，上限 200"},
        },
        "required": ["query"],
    },
)
async def find(args):
    payload = {"query": args.get("query", "")}
    for key in ("scope", "match"):
        if args.get(key):
            payload[key] = args[key]
    if args.get("sheets"):
        payload["sheets"] = args["sheets"]
    max_results = _optional_int(args, "max_results")
    if max_results is not None:
        payload["max_results"] = max_results
    result = await call_csharp("find", payload)
    return _wrap_result(result)


@tool(
    "list",
    "列出工作簿里的对象：sheets（工作表、可见性、已用区域）、names（定义的名称及引用）、tables（表格及列名）、"
    "pivots（数据透视表及数据源）、charts（图表及所在表）。弄清工作簿结构、找表名/名称时用它。",
    {
        "type": "object",
        "properties": {
            "kind": {"type": "string", "enum": ["sheets", "names", "tables", "pivots", "charts"], "description": "列哪一类，默认 sheets"},
        },
        "required": [],
    },
)
async def list_objects(args):
    result = await call_csharp("list", {"kind": args.get("kind") or "sheets"})
    return _wrap_result(result)


@tool(
    "inspect_sheet",
    "分析一张工作表的结构（接手陌生或复杂的表时先用它）：blocks 层给出表上有几块数据、每块的标题、"
    "多级表头（列名路径如「扣款/养老」）、数据行范围、合计行、每列类型；formulas 层按 R1C1 归纳公式模式"
    "（如 =[数量]*[单价] 覆盖 D2:D500）、标出同列公式不一致，并列出异常候选（合计漏行、被改成死值、"
    "孤立偏离、断链引用）；objects 层列表格/图表/透视表；dependencies 层列跨表和外部工作簿引用。"
    "每层带 status：complete 是全表结论，partial 表示大表只读了一部分；header_uncertain=true 时"
    "先用 read_range 读前几行确认表头。异常只是候选，改之前先问用户。",
    {
        "type": "object",
        "properties": {
            "sheet": {"type": "string", "description": "工作表名；不传用当前活动表"},
            "layers": {
                "type": "array",
                "items": {"type": "string", "enum": ["blocks", "formulas", "objects", "dependencies", "all"]},
                "description": "要哪些层，默认 blocks + formulas；all 表示全部",
            },
        },
        "required": [],
    },
)
async def inspect_sheet(args):
    import perception
    payload = {}
    if args.get("sheet"):
        payload["sheet"] = args["sheet"]
    snapshot = await call_csharp("sheet_snapshot", payload)
    if not isinstance(snapshot, dict) or snapshot.get("success") is False:
        return _wrap_result(snapshot)
    try:
        report = perception.inspect(snapshot.get("data") or {}, args.get("layers"))
    except Exception as exc:  # noqa: BLE001 — 分析失败不能拖垮会话，给出可继续的路
        return _wrap_result({
            "success": False,
            "error": f"结构分析失败：{type(exc).__name__}: {exc}",
            "suggestion": "改用 list(kind=sheets) 和 read_range 直接读取",
        })
    return _wrap_result({"success": True, "data": report})


@tool(
    "explore_workbook",
    "大工作簿分头摸底：把要查的问题拆成最多 4 个互不重叠的子任务（通常按表分组），每个子任务由一个"
    "独立的只读子 agent 并行去查（它们能用 list / find / inspect_sheet / read_range），只把结论交回来，"
    "你的上下文不会被几十次读取的原始数据塞满。适合表多（5 张以上）或表大、而问题涉及多张表的时候；"
    "一两张表、或只查一个具体位置时直接自己查，不要用它。结论未经你核实，写入前对关键位置再确认。",
    {
        "type": "object",
        "properties": {
            "tasks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "question": {"type": "string", "description": "这个子任务要查清什么（写清楚要交回哪些信息）"},
                        "sheets": {"type": "array", "items": {"type": "string"}, "description": "只查这些表；不传查整个工作簿"},
                    },
                    "required": ["question"],
                },
                "description": "1–4 个子任务",
            },
        },
        "required": ["tasks"],
    },
)
async def explore_workbook(args):
    import explorer
    import ui_events
    from ipc import _cancelled, write_message

    tasks = explorer.normalize_tasks(args.get("tasks"))
    if not tasks:
        return _wrap_result({"success": False, "error": "tasks 为空",
                             "suggestion": "传 1–4 个子任务：[{question: 要查什么, sheets: [表名]}]"})
    if not explorer.configured():
        return _wrap_result({"success": False, "error": "子 agent 还没配置好（会话尚未初始化）",
                             "suggestion": "直接用 list / find / inspect_sheet 自己查"})
    host = explorer._config.get("host", "excel")
    tools = [t for t in register_all_tools(host) if t.name in explorer.EXPLORER_TOOLS]

    async def publish(text: str):
        ui_events.PROGRESS.set(text)
        await write_message(ui_events.envelope("status", text=text, tool="explore_workbook"))

    try:
        outcomes = await explorer.explore(tasks, tools, publish, cancelled=_cancelled)
    finally:
        ui_events.PROGRESS.clear()
        await write_message(ui_events.envelope("status", text=""))
    finished = sum(1 for o in outcomes if o.status == "ok")
    return _wrap_result({
        "success": finished > 0,
        "data": {
            "tasks": [o.to_dict() for o in outcomes],
            "note": "以上是子 agent 的结论，未经核实；要写入或向用户下结论前，对关键位置用 read_range 确认",
        },
        **({} if finished else {"error": "所有子任务都没有完成",
                                "suggestion": "改用 list / find / inspect_sheet 自己查"}),
    })


@tool("write_formula", "向指定单元格写入 Excel 公式（以 = 开头）", {"address": str, "formula": str})
async def write_formula(args):
    result = await call_csharp("write_formula", {
        "address": args["address"],
        "formula": args["formula"],
    })
    return _wrap_result(await _pattern_check(result, args["address"]))


@tool("write_value", "向指定单元格写入纯文本/数字值（不解析为公式，写入张三会显示张三而不是=\"张三\"）", {"address": str, "value": str})
async def write_value(args):
    result = await call_csharp("write_value", {
        "address": args["address"],
        "value": args["value"],
    })
    return _wrap_result(result)


@tool("write_range", "批量写入二维数组到指定起始单元格（比逐个 write_value 快 100 倍，无弹窗）。address 是左上角单元格（如 A1 或 Sheet3!A1），values 是二维数组。", {"address": str, "values": list})
async def write_range(args):
    result = await call_csharp("write_range", {
        "address": args["address"],
        "values": args["values"],
    })
    values = args.get("values")
    if _has_formula(values):
        width = max((len(row) for row in values if isinstance(row, list)), default=1)
        result = await _pattern_check(result, args["address"], rows=len(values), columns=width)
    return _wrap_result(result)


MAX_CLARIFY_QUESTIONS = 4
MAX_CLARIFY_OPTIONS = 6


def normalize_questions(args: dict) -> list:
    """clarify_intent 的参数整理成提问卡：[{question, header, options: [{label, description}], multi_select}]。

    兼容只问一个问题的旧写法（question + options 字符串数组）；选项可以是字符串或 {label, description}。"""
    def text(value, limit):
        return str(value or "").strip()[:limit]

    def options_of(raw):
        out = []
        for opt in (raw or [])[:MAX_CLARIFY_OPTIONS]:
            if isinstance(opt, dict):
                label = text(opt.get("label"), 80)
                if label:
                    out.append({"label": label, "description": text(opt.get("description"), 160)})
            elif text(opt, 80):
                out.append({"label": text(opt, 80), "description": ""})
        return out

    raw_questions = args.get("questions")
    if not isinstance(raw_questions, list) or not raw_questions:
        raw_questions = [{"question": args.get("question"), "options": args.get("options"),
                          "multi_select": args.get("multi_select")}]
    questions = []
    for raw in raw_questions[:MAX_CLARIFY_QUESTIONS]:
        if not isinstance(raw, dict):
            continue
        question = text(raw.get("question"), 300)
        if question:
            questions.append({"question": question, "header": text(raw.get("header"), 12),
                              "options": options_of(raw.get("options")),
                              "multi_select": bool(raw.get("multi_select"))})
    return questions


_CLARIFY_OPTION = {
    "anyOf": [
        {"type": "string"},
        {"type": "object", "properties": {"label": {"type": "string"}, "description": {"type": "string"}},
         "required": ["label"]},
    ]
}


@tool(
    "clarify_intent",
    "向用户提问以澄清模糊指令。面板把问题渲染成提问卡，用户点选或自己填写后一次提交。"
    "只问一件事：question + options；有几处不确定时用 questions 一次问清（最多 4 题），不要一轮一轮地问。"
    "每题：question 问题；header 2~6 字的短标签（如「汇总口径」）；options 2~4 个选项，"
    "可写成 {label: 选项, description: 一句话说明后果}；multi_select 为 true 时可多选。"
    "用户总能选「其他」自己填，不必把「其他」列进选项",
    {
        "type": "object",
        "properties": {
            "question": {"type": "string"},
            "options": {"type": "array", "items": _CLARIFY_OPTION},
            "multi_select": {"type": "boolean"},
            "questions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "question": {"type": "string"},
                        "header": {"type": "string"},
                        "options": {"type": "array", "items": _CLARIFY_OPTION},
                        "multi_select": {"type": "boolean"},
                    },
                    "required": ["question"],
                },
            },
        },
    },
)
async def clarify_intent(args):
    questions = normalize_questions(args)
    if not questions:
        return _wrap_result({"success": False, "error": "没有给出要问的问题",
                             "suggestion": "传 question（只问一件事）或 questions（最多 4 题）"})
    user_answer = await call_csharp_clarify(questions)
    return _wrap_result({"success": True, "data": {"user_answer": user_answer}})


_TODO_STATUSES = ("pending", "in_progress", "completed")


def normalize_todos(raw) -> list:
    """校验并规整计划条目：[{content, status}]，最多 20 条，最多一条 in_progress。"""
    items = []
    for entry in raw if isinstance(raw, list) else []:
        if isinstance(entry, str):
            entry = {"content": entry}
        if not isinstance(entry, dict):
            continue
        content = str(entry.get("content") or "").strip()
        if not content:
            continue
        status = entry.get("status") if entry.get("status") in _TODO_STATUSES else "pending"
        items.append({"content": content[:120], "status": status})
    items = items[:20]
    seen_active = False
    for item in items:
        if item["status"] == "in_progress":
            if seen_active:
                item["status"] = "pending"
            seen_active = True
    return items


@tool(
    "todo_write",
    "维护本次任务的计划清单（面板顶部显示进度）。三步以上的任务开始前先列出计划；"
    "每开始一步把它标为 in_progress（同一时间只有一条），做完立刻标 completed；发现新步骤就加进去。"
    "todos 是完整清单（每次都传全部条目）：[{content: 步骤描述, status: pending|in_progress|completed}]",
    {"todos": list},
)
async def todo_write(args):
    import ui_events
    from ipc import write_message
    items = normalize_todos(args.get("todos"))
    await write_message(ui_events.envelope("plan", items=items))
    done = sum(1 for item in items if item["status"] == "completed")
    return _wrap_result({"success": True, "data": {"message": f"计划已更新（{done}/{len(items)} 完成）"}})


def normalize_plan(args: dict) -> dict:
    """present_plan 的参数整理成面板要显示的方案；长度都有上限"""
    def text(value, limit):
        return str(value or "").strip()[:limit]

    steps = []
    for raw in (args.get("steps") or [])[:20]:
        if isinstance(raw, dict):
            action = text(raw.get("action"), 200)
            if action:
                steps.append({"action": action, "target": text(raw.get("target"), 120),
                              "detail": text(raw.get("detail"), 400)})
        elif str(raw or "").strip():
            steps.append({"action": text(raw, 200), "target": "", "detail": ""})
    risks = [text(r, 200) for r in (args.get("risks") or [])[:10] if str(r or "").strip()]
    return {"summary": text(args.get("summary"), 400), "steps": steps, "risks": risks}


@tool(
    "present_plan",
    "提交一份变更方案给用户批准（「只出方案」模式下必须用它收尾；其他模式下改动很大时也可以先用它对齐）。"
    "summary：一句话说要做什么；steps：每一步 {action: 做什么, target: 表和区域, detail: 怎么做}；"
    "risks：风险点（会覆盖什么、有什么不确定）。提交后本轮结束，等用户批准，不要接着执行",
    {
        "type": "object",
        "properties": {
            "summary": {"type": "string"},
            "steps": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "action": {"type": "string"},
                        "target": {"type": "string"},
                        "detail": {"type": "string"},
                    },
                    "required": ["action"],
                },
            },
            "risks": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["summary", "steps"],
    },
)
async def present_plan(args):
    import ui_events
    from ipc import write_message
    plan = normalize_plan(args)
    if not plan["summary"] or not plan["steps"]:
        return _wrap_result({"success": False, "error": "方案缺少 summary 或 steps",
                             "suggestion": "summary 写一句话目标，steps 至少一步 {action, target, detail}"})
    await write_message(ui_events.envelope("plan_proposal", **plan))
    return _wrap_result({"success": True, "data": {
        "message": f"方案已提交给用户（{len(plan['steps'])} 步），等待批准。本轮到此结束："
                   "不要再调用工具，用一两句话说明方案要点，请用户批准。"}})


@tool(
    "update_workbook_notes",
    "整份重写这个工作簿的记忆（NOTES.md，只存在用户电脑上，下次会话自动带给你）。"
    "保留四个小节：## 结构怪癖 / ## 用户偏好 / ## 做过的改动 / ## 禁区。禁区每行一个「- 表名」或「- 表名!A1:D20」，"
    "只能加不能删。notes 是完整的新内容（Markdown），不是增量",
    {"notes": str},
)
async def update_workbook_notes(args):
    import workbook_memory
    memory = workbook_memory.current()
    if memory is None:
        return _wrap_result({"success": False,
                             "error": "这个工作簿还没保存过，没有地方存记忆",
                             "suggestion": "告诉用户先保存工作簿；这次要记的内容可以直接在回复里说明"})
    notes = args.get("notes")
    if not isinstance(notes, str) or not notes.strip():
        return _wrap_result({"success": False, "error": "notes 不能为空",
                             "suggestion": "传入完整的记忆内容（保留四个小节）"})
    refusal = memory.save_notes(notes, by_agent=True)
    if refusal:
        return _wrap_result({"success": False, "error": refusal})
    workbook_memory.mark_seen(memory)
    zones = workbook_memory.protected_zones(memory.notes())
    message = f"已更新「{memory.name}」的记忆（{len(memory.notes())} 字）"
    if zones:
        message += "；禁区：" + "；".join(z.text for z in zones)
    return _wrap_result({"success": True, "data": {"message": message}})


@tool(
    "load_skill",
    "读取一份专业知识（<knowledge-skills> 里列出的技能）。先只传 name 读正文；"
    "正文里提到补充文件（如 wps.md）时再传 file 读它",
    {
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "技能名，如 cn-data-cleaning"},
            "file": {"type": "string", "description": "可选：补充文件名，如 wps.md；读正文时不要传"},
        },
        "required": ["name"],
    },
)
async def load_skill(args):
    import knowledge_skills
    file = args.get("file")
    # 模型偶尔把「不传」写成字符串 null / None
    if not isinstance(file, str) or file.strip().lower() in ("", "null", "none", "skill.md"):
        file = None
    return _wrap_result(knowledge_skills.load(args.get("name") or "", file))


@tool("read_workbook", "读取当前工作簿的结构信息", {})
async def read_workbook(args):
    result = await call_csharp("read_workbook", {})
    return _wrap_result(result)


@tool("read_attachment", "读取用户上传的附件文件内容（xlsx/xls/csv/txt/json 等）。当用户消息提到附件或要求处理附件文件时使用此工具。", {"file_name": str})
async def read_attachment(args):
    result = await call_csharp("read_attachment", {"file_name": args["file_name"]})
    return _wrap_result(result)


@tool("read_selection", "读取当前选中的单元格信息", {})
async def read_selection(args):
    result = await call_csharp("read_selection", {})
    return _wrap_result(result)


@tool("fill_formula_down", "将公式向下填充到指定行数", {"from_address": str, "row_count": int})
async def fill_formula_down(args):
    result = await call_csharp("fill_formula_down", {
        "from_address": args["from_address"],
        "row_count": args["row_count"],
    })
    # 与 C# 的写入区域一致：起始行 + row_count 行
    rows = (_optional_int(args, "row_count") or 0) + 1
    return _wrap_result(await _pattern_check(result, args["from_address"], rows=rows))


@tool("replace_formula", "在指定范围内批量替换公式中的字符串", {"range_address": str, "find": str, "replace": str})
async def replace_formula(args):
    result = await call_csharp("replace_formula", {
        "range_address": args["range_address"],
        "find": args["find"],
        "replace": args["replace"],
    })
    return _wrap_result(await _pattern_check(result, args["range_address"]))


@tool("clean_data", "执行数据清洗操作（unify_date/remove_duplicates/highlight_missing/trim_spaces/text_to_number）", {"range_address": str, "operations": list})
async def clean_data(args):
    result = await call_csharp("clean_data", {
        "range_address": args["range_address"],
        "operations": args["operations"],
    })
    return _wrap_result(result)


@tool("create_chart", "基于数据范围创建图表", {"data_range": str, "chart_type": str, "title": str, "x_label": str, "y_label": str})
async def create_chart(args):
    result = await call_csharp("create_chart", {
        "data_range": args["data_range"],
        "chart_type": args.get("chart_type", "column"),
        "title": args.get("title", ""),
        "x_label": args.get("x_label", ""),
        "y_label": args.get("y_label", ""),
    })
    return _wrap_result(result)


@tool("create_pivot_table", "基于数据范围创建数据透视表", {"source_range": str, "destination_sheet": str, "pivot_table_name": str, "row_fields": list, "column_fields": list, "value_fields": list, "value_function": str})
async def create_pivot_table(args):
    result = await call_csharp("create_pivot_table", {
        "source_range": args["source_range"],
        "destination_sheet": args["destination_sheet"],
        "pivot_table_name": args.get("pivot_table_name", "PivotTable1"),
        "row_fields": args.get("row_fields", []),
        "column_fields": args.get("column_fields", []),
        "value_fields": args.get("value_fields", []),
        "value_function": args.get("value_function", "Sum"),
    })
    return _wrap_result(result)


@tool("execute_vba", "执行 VBA 代码（Sub DeepExcel_TempMacro）", {"code": str})
async def execute_vba(args):
    result = await call_csharp("execute_vba", {"code": args["code"]})
    return _wrap_result(result)


@tool("execute_jsa", "执行 WPS JSA 代码（ES6 语法，对象模型与 VBA 一致：Application/Workbook/Worksheet/Range）", {"code": str})
async def execute_jsa(args):
    """★ WPS 端专用：当宿主是 WPS 表格时，用此工具替代 execute_vba。
    JSA 语法为 ES6（let/const/箭头函数），对象模型与 VBA 完全一致。
    Excel 端不支持此工具，会返回错误。"""
    result = await call_csharp("execute_jsa", {"code": args["code"]})
    return _wrap_result(result)


@tool("execute_python", "执行纯计算的 Python 代码（字符串处理、正则、数学计算）。碰不到工作簿：读写 Excel 请用对应的工具或 execute_vba。脚本失败不会影响工作簿。", {"code": str})
async def execute_python(args):
    result = await call_csharp("execute_python", {"code": args["code"]})
    return _wrap_result(result)


@tool("create_snapshot", "创建当前工作簿快照（用于回滚）", {})
async def create_snapshot(args):
    result = await call_csharp("create_snapshot", {})
    return _wrap_result(result)


@tool("add_sheet", "添加新的空白工作表（sheet）到当前工作簿", {"name": str})
async def add_sheet(args):
    result = await call_csharp("add_sheet", {"name": args["name"]})
    return _wrap_result(result)


@tool("delete_sheet", "删除指定名称的工作表", {"name": str})
async def delete_sheet(args):
    result = await call_csharp("delete_sheet", {"name": args["name"]})
    return _wrap_result(result)


@tool("rename_sheet", "重命名工作表", {"old_name": str, "new_name": str})
async def rename_sheet(args):
    result = await call_csharp("rename_sheet", {
        "old_name": args["old_name"],
        "new_name": args["new_name"],
    })
    return _wrap_result(result)


@tool("set_number_format", "设置单元格区域的数字格式（如 #,##0.00 / 0% / yyyy-mm-dd）", {"address": str, "format": str})
async def set_number_format(args):
    result = await call_csharp("set_number_format", {
        "address": args["address"],
        "format": args["format"],
    })
    return _wrap_result(result)


@tool("set_column_width", "设置列宽（auto_fit=True 时自动适应宽度，忽略 width）", {"address": str, "width": float, "auto_fit": bool})
async def set_column_width(args):
    result = await call_csharp("set_column_width", {
        "address": args["address"],
        "width": args.get("width", 10.0),
        "auto_fit": args.get("auto_fit", False),
    })
    return _wrap_result(result)


@tool("sort_data", "对指定区域按行排序（不会交换列）。range_address: 数据区域（如 'A1:D100'，含表头则从表头行开始）；sort_column: 列字母如 'B' 或列序号如 '2'；descending: true降序/false升序；has_header: 【必传】第一行是否为列标题，是→true(表头不参与排序)，否→false。必须根据 read_range 结果判断后传入", {"range_address": str, "sort_column": str, "descending": bool, "has_header": bool})
async def sort_data(args):
    result = await call_csharp("sort_data", {
        "range_address": args["range_address"],
        "sort_column": args["sort_column"],
        "descending": args.get("descending", False),
        "has_header": args.get("has_header", False),
    })
    return _wrap_result(result)


@tool("filter_data", "对指定区域应用自动筛选（column_index 从 1 开始，criteria 如 '>100' 或 '北京'）", {"range_address": str, "column_index": int, "criteria": str})
async def filter_data(args):
    result = await call_csharp("filter_data", {
        "range_address": args["range_address"],
        "column_index": args["column_index"],
        "criteria": args["criteria"],
    })
    return _wrap_result(result)


@tool("merge_cells", "合并指定区域的单元格", {"address": str})
async def merge_cells(args):
    result = await call_csharp("merge_cells", {"address": args["address"]})
    return _wrap_result(result)


@tool("unmerge_cells", "拆分指定区域内所有合并单元格", {"address": str})
async def unmerge_cells(args):
    result = await call_csharp("unmerge_cells", {"address": args["address"]})
    return _wrap_result(result)


@tool("set_cell_style", "设置单元格样式（颜色支持 hex 如 #FF0000 或颜色名 red/blue/green）", {"address": str, "font_name": str, "font_size": float, "bold": bool, "italic": bool, "font_color": str, "bg_color": str, "h_align": str, "v_align": str, "wrap_text": bool})
async def set_cell_style(args):
    result = await call_csharp("set_cell_style", {
        "address": args["address"],
        "font_name": args.get("font_name", ""),
        "font_size": args.get("font_size"),
        "bold": args.get("bold"),
        "italic": args.get("italic"),
        "font_color": args.get("font_color", ""),
        "bg_color": args.get("bg_color", ""),
        "h_align": args.get("h_align", ""),
        "v_align": args.get("v_align", ""),
        "wrap_text": args.get("wrap_text"),
    })
    return _wrap_result(result)


@tool("copy_range", "复制源区域到目标位置（含格式和公式）", {"source_address": str, "dest_address": str})
async def copy_range(args):
    result = await call_csharp("copy_range", {
        "source_address": args["source_address"],
        "dest_address": args["dest_address"],
    })
    # 复制会带着公式过去：目标区域至少和源一样大（D2 复制到 D3:D7 时是整个 D3:D7）
    from perception.postwrite import parse_a1
    source = parse_a1(args.get("source_address") or "")
    if source is not None:
        _, r1, c1, r2, c2 = source
        result = await _pattern_check(result, args["dest_address"], rows=r2 - r1 + 1, columns=c2 - c1 + 1)
    return _wrap_result(result)


@tool("clear_range", "清空区域（clear_type: contents=只清内容, formats=只清格式, all=全部）", {"address": str, "clear_type": str})
async def clear_range(args):
    result = await call_csharp("clear_range", {
        "address": args["address"],
        "clear_type": args.get("clear_type", "all"),
    })
    return _wrap_result(result)


@tool("insert_rows", "在第 row 行前插入 count 行", {"row": int, "count": int})
async def insert_rows(args):
    result = await call_csharp("insert_rows", {
        "row": args["row"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("delete_rows", "从第 row 行开始删除 count 行", {"row": int, "count": int})
async def delete_rows(args):
    result = await call_csharp("delete_rows", {
        "row": args["row"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("insert_columns", "在第 column 列前插入 count 列（column 从 1 开始）", {"column": int, "count": int})
async def insert_columns(args):
    result = await call_csharp("insert_columns", {
        "column": args["column"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("delete_columns", "从第 column 列开始删除 count 列（column 从 1 开始）", {"column": int, "count": int})
async def delete_columns(args):
    result = await call_csharp("delete_columns", {
        "column": args["column"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("freeze_panes", "冻结窗格（在指定单元格左上角冻结，如 'B2' 冻结 A 列和 1 行）", {"address": str})
async def freeze_panes(args):
    result = await call_csharp("freeze_panes", {"address": args["address"]})
    return _wrap_result(result)


@tool("apply_conditional_format", "应用条件格式（rule_type: color_scale/data_bar/highlight_rules/cell_value）", {"address": str, "rule_type": str, "rule_args": dict})
async def apply_conditional_format(args):
    result = await call_csharp("apply_conditional_format", {
        "address": args["address"],
        "rule_type": args["rule_type"],
        "rule_args": args.get("rule_args", {}),
    })
    return _wrap_result(result)


@tool("write_table", "将指定区域转换为 Excel 表格（ListObject，自带筛选和样式）", {"address": str, "table_name": str})
async def write_table(args):
    result = await call_csharp("write_table", {
        "address": args["address"],
        "table_name": args.get("table_name", ""),
    })
    return _wrap_result(result)


@tool("rollback", "把工作簿恢复到指定快照。修改类工具的结果里带 backup_snapshot_id（本回合第一次修改前自动做的备份），传入它即可撤销本回合的修改；结果里的 checkpoint_id 是这一步执行前的检查点，传入它只撤销这一步及之后的修改；恢复前会自动另存当前状态，恢复本身也可撤销。", {"snapshot_id": str})
async def rollback(args):
    result = await call_csharp("rollback", {"snapshot_id": args["snapshot_id"]})
    return _wrap_result(result)


@tool("delete_blank_rows", "删除指定区域内的空行（整行都为空的行将被删除，下方数据自动上移）", {"range_address": str})
async def delete_blank_rows(args):
    result = await call_csharp("delete_blank_rows", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("split_text_to_columns", "按分隔符拆分文本到多列（从指定列开始向右扩展）", {"range_address": str, "delimiter": str})
async def split_text_to_columns(args):
    result = await call_csharp("split_text_to_columns", {
        "range_address": args["range_address"],
        "delimiter": args.get("delimiter", ","),
    })
    return _wrap_result(result)


@tool("fill_blank_cells", "向下填充空白单元格（用上方最近的非空值填充空白单元格）", {"range_address": str})
async def fill_blank_cells(args):
    result = await call_csharp("fill_blank_cells", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("highlight_duplicates", "高亮标记重复值（相同值的单元格标为浅红色背景）", {"range_address": str, "color": str})
async def highlight_duplicates(args):
    result = await call_csharp("highlight_duplicates", {
        "range_address": args["range_address"],
        "color": args.get("color", "#FFC7CE"),
    })
    return _wrap_result(result)


@tool("remove_special_chars", "去除文本中的特殊字符（非打印字符、不可见字符、不间断空格等）", {"range_address": str})
async def remove_special_chars(args):
    result = await call_csharp("remove_special_chars", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("clean_amount", "清洗金额数据：去除货币符号（¥/$/€）、千分位逗号，转为数字格式", {"range_address": str})
async def clean_amount(args):
    result = await call_csharp("clean_amount", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("merge_columns", "合并多列内容为一列（用指定分隔符连接，跳过空值）", {"range_address": str, "delimiter": str, "target_column": str})
async def merge_columns(args):
    result = await call_csharp("merge_columns", {
        "range_address": args["range_address"],
        "delimiter": args.get("delimiter", " "),
        "target_column": args.get("target_column", ""),
    })
    return _wrap_result(result)


@tool("rename_columns", "批量重命名列标题（修改第一行的列名）", {"range_address": str, "new_names": list})
async def rename_columns(args):
    result = await call_csharp("rename_columns", {
        "range_address": args["range_address"],
        "new_names": args["new_names"],
    })
    return _wrap_result(result)


@tool("collapse_spaces", "压缩内部多余空格（多个连续空格变为一个）", {"range_address": str})
async def collapse_spaces(args):
    result = await call_csharp("collapse_spaces", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("add_data_labels", "为图表添加数据标签（chart_name 为空则操作当前工作表第一个图表）", {"chart_name": str, "position": str, "show_value": bool, "show_category_name": bool, "show_percentage": bool})
async def add_data_labels(args):
    result = await call_csharp("add_data_labels", {
        "chart_name": args.get("chart_name", ""),
        "position": args.get("position", "outside_end"),
        "show_value": args.get("show_value", True),
        "show_category_name": args.get("show_category_name", False),
        "show_percentage": args.get("show_percentage", False),
    })
    return _wrap_result(result)


@tool("set_chart_title", "设置/修改图表标题（chart_name 为空则操作当前工作表第一个图表）", {"chart_name": str, "title": str})
async def set_chart_title(args):
    result = await call_csharp("set_chart_title", {
        "chart_name": args.get("chart_name", ""),
        "title": args.get("title", ""),
    })
    return _wrap_result(result)


@tool("set_chart_colors", "设置图表系列颜色（按顺序为每个系列设置颜色，传入 hex 颜色数组如 ['#FF0000','#00FF00']）", {"chart_name": str, "colors": list})
async def set_chart_colors(args):
    result = await call_csharp("set_chart_colors", {
        "chart_name": args.get("chart_name", ""),
        "colors": args.get("colors", []),
    })
    return _wrap_result(result)


@tool("create_combo_chart", "创建组合图（柱状+折线双轴图，line_series_index 指定哪个系列作为折线图）", {"data_range": str, "title": str, "x_label": str, "y_label": str, "secondary_y_label": str, "line_series_index": int})
async def create_combo_chart(args):
    result = await call_csharp("create_combo_chart", {
        "data_range": args["data_range"],
        "title": args.get("title", ""),
        "x_label": args.get("x_label", ""),
        "y_label": args.get("y_label", ""),
        "secondary_y_label": args.get("secondary_y_label", ""),
        "line_series_index": args.get("line_series_index", 2),
    })
    return _wrap_result(result)


@tool("export_chart", "导出图表为图片文件（PNG/JPG/GIF/BMP），返回输出路径", {"chart_name": str, "output_path": str, "format": str})
async def export_chart(args):
    result = await call_csharp("export_chart", {
        "chart_name": args.get("chart_name", ""),
        "output_path": args.get("output_path", ""),
        "format": args.get("format", "png"),
    })
    return _wrap_result(result)


@tool("refresh_pivot", "刷新数据透视表（数据变化后更新透视表结果）", {"pivot_table_name": str, "sheet_name": str})
async def refresh_pivot(args):
    result = await call_csharp("refresh_pivot", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("group_pivot_date", "透视表日期字段分组（按年/季度/月/日组合，group_by 如 'year,month'）", {"pivot_table_name": str, "field_name": str, "group_by": str, "sheet_name": str})
async def group_pivot_date(args):
    result = await call_csharp("group_pivot_date", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "field_name": args["field_name"],
        "group_by": args.get("group_by", "month"),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("set_pivot_value_display", "设置透视表值显示方式（normal/percent_of_column/percent_of_row/percent_of_total/running_total/rank）", {"pivot_table_name": str, "value_field": str, "display_type": str, "base_field": str, "sheet_name": str})
async def set_pivot_value_display(args):
    result = await call_csharp("set_pivot_value_display", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "value_field": args["value_field"],
        "display_type": args.get("display_type", "normal"),
        "base_field": args.get("base_field", ""),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("set_pivot_totals", "控制透视表总计显示（行总计/列总计）", {"pivot_table_name": str, "show_row_totals": bool, "show_column_totals": bool, "sheet_name": str})
async def set_pivot_totals(args):
    result = await call_csharp("set_pivot_totals", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "show_row_totals": args.get("show_row_totals", True),
        "show_column_totals": args.get("show_column_totals", True),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("add_pivot_slicer", "为透视表添加切片器（交互式筛选）", {"pivot_table_name": str, "field_name": str, "sheet_name": str})
async def add_pivot_slicer(args):
    result = await call_csharp("add_pivot_slicer", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "field_name": args["field_name"],
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


# ★ auto_analyze 已移除：宿主端没有对应实现（调用必然返回"未知工具"，白白浪费一轮）。
# tests/test_excel_tools.py 的守卫会拦住下一个只在这里注册、C# 侧没有实现的工具。


# ★ smart_chart 已移除：宿主端没有对应实现（调用必然失败），
# 且"自动选图表类型"本就该由 agent 结合数据自己判断，不需要宿主用规则替它决定。
# agent 用 create_chart 显式传 chart_type 即可。


# ============================ Computer Use 工具 ============================

@tool("screenshot_excel", "截图 Excel 主窗口（Computer Use）。仅在用户主动要求截图/computer use 时调用，禁止主动截图验证工具执行效果。返回 base64 JPEG（缩放到最大宽 1280px，<300KB）。无需传参。", {})
async def screenshot_excel(args):
    result = await call_csharp("screenshot_excel", {})
    return _wrap_result(result)


@tool("send_keys", "模拟键盘输入到 Excel 窗口（Computer Use）。仅用于操作 Excel 原生对话框/弹窗（如 {ESC} 关闭弹窗）。禁止用快捷键替代专门工具（保存/复制/撤销等用对应工具）。keys 语法：{ENTER}/{ESC}/{TAB}/{UP}/{DOWN} 等；+Shift ^Ctrl %Alt 前缀。", {"keys": str})
async def send_keys(args):
    result = await call_csharp("send_keys", {"keys": args["keys"]})
    return _wrap_result(result)


# 只在某个宿主实现的工具。另一个宿主注册它，模型调用就必然拿到"未知工具"。
_HOST_ONLY_TOOLS = {
    "execute_jsa": "wps",
}

# WPS 宿主（src/DeepExcel.Wps/tool-dispatcher.js 的 case 分支）实际实现的工具。
# WPS 的工具分发只实现了 Excel 的一小半：图表、透视、快照、VBA、清洗等在 WPS 里
# 调用永远拿到「WPS 端暂未实现」，白白浪费一轮还让用户以为功能坏了。所以 WPS 会话
# 只注册这里列出的工具。tests/test_host_tools.py 逐条比对 JS 源码，两边不一致就变红。
WPS_HOST_TOOLS = frozenset({
    "read_range", "find", "list", "write_formula", "write_value", "write_range",
    "read_workbook", "read_selection", "sort_data", "filter_data",
    "merge_cells", "unmerge_cells", "add_sheet", "delete_sheet", "rename_sheet",
    "set_number_format", "set_column_width", "freeze_panes", "fill_formula_down",
    "copy_range", "clear_range", "insert_rows", "delete_rows", "insert_columns",
    "delete_columns", "set_cell_style", "write_table", "execute_jsa",
})

# 不经过宿主工具分支的工具：走独立消息通道（clarify），两个宿主都能用
SIDECAR_CHANNEL_TOOLS = frozenset({"clarify_intent", "todo_write", "update_workbook_notes", "load_skill",
                                   "present_plan"})

# 宿主原语：宿主分发器里有这个分支，但不注册给模型，只由侧车工具调用
WPS_HOST_PRIMITIVES = frozenset({"sheet_snapshot"})

# 在侧车里计算、只依赖宿主原语的工具 → 它需要的原语
SIDECAR_COMPUTED_TOOLS = {"inspect_sheet": "sheet_snapshot"}

# 在侧车里起子会话、只用只读工具的工具：两个宿主都能用（子会话的工具再按宿主过滤）
SIDECAR_SESSION_TOOLS = frozenset({"explore_workbook"})


def host_supports_tool(host: str, name: str) -> bool:
    if _HOST_ONLY_TOOLS.get(name, host) != host:
        return False
    if name in SIDECAR_SESSION_TOOLS:
        return True
    if host == "wps":
        if name in SIDECAR_COMPUTED_TOOLS:
            return SIDECAR_COMPUTED_TOOLS[name] in WPS_HOST_PRIMITIVES
        return name in WPS_HOST_TOOLS or name in SIDECAR_CHANNEL_TOOLS
    return True


def register_all_tools(host: str = "excel") -> list:
    """返回当前宿主可用的 @tool 工具对象列表。

    这份列表同时决定 MCP 注册和 allowed_tools，是工具清单的唯一来源；
    system_prompt.py 的 <available-tools> 由 tests/test_excel_tools.py 与它对齐。
    """
    tools = [
        read_workbook, read_selection, read_range, find, list_objects, inspect_sheet, explore_workbook,
        read_attachment,
        write_formula, write_value, write_range, fill_formula_down, replace_formula,
        clean_data,
        delete_blank_rows, split_text_to_columns, fill_blank_cells,
        highlight_duplicates, remove_special_chars, clean_amount,
        merge_columns, rename_columns, collapse_spaces,
        create_chart, create_combo_chart,
        add_data_labels, set_chart_title, set_chart_colors, export_chart,
        create_pivot_table,
        refresh_pivot, group_pivot_date,
        set_pivot_value_display, set_pivot_totals, add_pivot_slicer,
        execute_vba, execute_jsa, execute_python,
        create_snapshot, rollback,
        add_sheet, delete_sheet, rename_sheet,
        set_number_format, set_column_width,
        sort_data, filter_data,
        merge_cells, unmerge_cells,
        set_cell_style, copy_range, clear_range,
        insert_rows, delete_rows, insert_columns, delete_columns,
        freeze_panes,
        apply_conditional_format, write_table,
        clarify_intent, todo_write, present_plan, update_workbook_notes, load_skill,
        # ★ Computer Use 工具
        screenshot_excel, send_keys,
    ]
    return [t for t in tools if host_supports_tool(host, t.name)]


def host_tool_note(host: str, registered: list) -> str:
    """系统提示词里的 <available-tools> 是按 Excel 写的全集。WPS 会话只注册了
    WPS 真正能执行的那部分，这里补一段说明，免得模型按提示词去找不存在的工具。"""
    if host != "wps":
        return ""
    return (
        "\n\n<host-tools>\n当前宿主是 WPS 表格。本会话只提供以下工具，"
        "提示词其他地方提到、但不在此列的工具在 WPS 下不可用，不要尝试调用；"
        "需要时改用 execute_jsa 或告诉用户该功能请在 Excel 中完成：\n"
        + ", ".join(sorted(registered))
        + "\n</host-tools>"
    )
