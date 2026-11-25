# src/DeepExcel.Sidecar/excel_tools.py
import json
from claude_agent_sdk import tool
from ipc import call_csharp, call_csharp_clarify


# ---------------------------------------------------------------------------
# Result envelope
#
# Every tool result the model reads starts with the same keys, in this order:
#   ok             bool   -- did the call do what was asked
#   status         str    -- done | partial | check_failed | failed
#   checkpoint_id  str    -- only for steps that changed the workbook: rollback
#                            to it undoes this step and everything after it
#   next           str    -- only when the model has something to do about the
#                            result: the next call to make or the thing to stop on
#   warnings       [str]  -- only when non-empty: things to mention to the user
# followed by the host's original fields (success / data / error / suggestion /
# verification / changes / backup_snapshot_id ...). Those stay because the
# panel parses the same text (ui_events.parse_tool_result reads success/error).
# ---------------------------------------------------------------------------

STATUS_DONE = "done"
STATUS_PARTIAL = "partial"
STATUS_CHECK_FAILED = "check_failed"
STATUS_FAILED = "failed"

ENVELOPE_KEYS = ("ok", "status", "checkpoint_id", "next", "warnings")

_NEXT_ON_FAILURE = "按 error 改正参数后可以再试一次；同样的错误不要原样重试，第二次还失败就停下来，向用户说明原因和可选方案"
_NEXT_ON_INTERRUPT = "用户已中断：停下来等用户的新指示，不要重试"
_NEXT_ON_DECLINE = "用户拒绝了这一步：不要重试，也不要换个工具或用代码绕过，问用户想怎么改"
_NEXT_ON_CHECK_FAILED = ("先按 verification.summary 修正这一步（改不好就用 rollback 传 checkpoint_id 回到这一步之前），"
                         "修好再向用户汇报")
_NEXT_ON_TRUNCATED = "只列出了一部分：缩小范围（sheets / 地址）或换更具体的条件再查"


def _is_decline(text: str) -> bool:
    # 只认用户自己点「否」的文字（SecurityGateway：「用户取消了高风险操作」）。
    # 「Excel 拒绝了这次操作」「已拒绝执行」是参数或代码的问题，要改了再试，不能当成用户拒绝。
    return any(marker in text for marker in ("用户取消", "用户拒绝"))


def _warnings_of(result: dict, data) -> list:
    warnings = []
    single = result.get("warning")
    if isinstance(single, str) and single.strip():
        warnings.append(single.strip())
    listed = data.get("warnings") if isinstance(data, dict) else None
    if isinstance(listed, list):
        warnings.extend(str(w).strip() for w in listed if str(w or "").strip())
    seen = set()
    return [w for w in warnings if not (w in seen or seen.add(w))]


def _partial_next(data) -> str | None:
    """读取类结果没读全时下一步怎么做；读全了返回 None。"""
    if not isinstance(data, dict):
        return None
    hint = data.get("hint") if isinstance(data.get("hint"), str) and data.get("hint").strip() else None
    paging = data.get("paging")
    if isinstance(paging, dict) and (paging.get("next_offset") is not None or paging.get("columns_truncated")):
        if hint:
            return hint.strip()
        offset = paging.get("next_offset")
        if offset is not None:
            return f"还有没读的行：需要时用同一个 address 加 offset={offset} 继续读；不需要就别读完整张表"
        return "只返回了前几列：其余列用更窄的地址读取"
    if data.get("truncated") is True:
        return hint.strip() if hint else _NEXT_ON_TRUNCATED
    if data.get("status") == "partial":
        return hint.strip() if hint else "结论只覆盖了一部分，下结论前对没覆盖到的部分再确认"
    return None


def build_envelope(result: dict) -> dict:
    """C# / 本地工具的原始结果 → 模型读到的统一信封（见上方说明）。不是 dict 的原样返回。"""
    if not isinstance(result, dict):
        return result
    data = result.get("data")
    ok = result.get("success") is not False
    warnings = _warnings_of(result, data)
    next_step = None

    if not ok:
        status = STATUS_FAILED
        error = str(result.get("error") or "")
        suggestion = result.get("suggestion")
        if result.get("error_code") == "interrupted":
            next_step = _NEXT_ON_INTERRUPT
        elif _is_decline(error):
            next_step = _NEXT_ON_DECLINE
        elif isinstance(suggestion, str) and suggestion.strip():
            next_step = suggestion.strip()
        else:
            next_step = _NEXT_ON_FAILURE
    else:
        verification = result.get("verification")
        if isinstance(verification, dict) and verification.get("ok") is False:
            status = STATUS_CHECK_FAILED
            next_step = _NEXT_ON_CHECK_FAILED
        else:
            next_step = _partial_next(data)
            status = STATUS_PARTIAL if next_step else STATUS_DONE

    envelope = {"ok": ok, "status": status}
    checkpoint = result.get("checkpoint_id")
    if isinstance(checkpoint, str) and checkpoint:
        envelope["checkpoint_id"] = checkpoint
    if next_step:
        envelope["next"] = next_step
    if warnings:
        envelope["warnings"] = warnings
    for key, value in result.items():
        # 单数的 warning 已并进 warnings，原样再带一份只会让模型读两遍
        if key not in ENVELOPE_KEYS and key != "warning":
            envelope[key] = value
    return envelope


# ---------------------------------------------------------------------------
# Tool descriptions
#
# Every description has the same shape so the model can find the same thing in
# the same place on all tools:
#   <what it does and its limits>。何时用：…。[不要用于：…。]看返回：…。[失败时：…。]
# 「失败时」 is only written when the tool has its own advice; the envelope's
# `next` already covers the generic case. tests/test_excel_tools.py checks the
# shape on every registered tool.
# ---------------------------------------------------------------------------

_CHECK_WRITE = "status=check_failed 时先按 verification.summary 修正再汇报；checkpoint_id 传给 rollback 只撤销这一步"
_CHECK_PLAIN = "ok=true 即完成"


def _describe(what: str, *, use: str, check: str, avoid: str | None = None, on_fail: str | None = None) -> str:
    def sentence(text: str) -> str:
        text = text.strip()
        return text if text.endswith("。") else text + "。"
    parts = [sentence(what), sentence("何时用：" + use)]
    if avoid:
        parts.append(sentence("不要用于：" + avoid))
    parts.append(sentence("看返回：" + check))
    if on_fail:
        parts.append(sentence("失败时：" + on_fail))
    return "".join(parts)


def _wrap_result(csharp_result: dict) -> dict:
    """把 C# 返回的 dict 包装成 MCP 工具返回格式（内容是统一信封，见 build_envelope）。
    ★ 当 C# 返回 success=false 时设置 is_error=true，
    防止 LLM 把失败当成功继续推理，触发死循环重试。"""
    is_error = False
    if isinstance(csharp_result, dict) and csharp_result.get("success") is False:
        is_error = True
    return {
        "content": [
            {"type": "text", "text": json.dumps(build_envelope(csharp_result), ensure_ascii=False)}
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
    _describe(
        '读一块区域的值和公式；整列/整行（A:A、1:1）自动收到已用区域，一次最多 200 行（limit 最大 500）',
        use='要具体单元格的值、核对某几行、写入前读一下要覆盖的区域',
        avoid='找东西在哪（用 find）、看有哪些表/名称/图表（用 list）、为了找数据把整张表逐页读完',
        check='status=partial 表示没读完，next 给出下一页怎么读（同一个 address 加 offset）；需要时再读',
    ),
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
    _describe(
        '在整个工作簿里搜文本（相当于 Ctrl+F 的「查找全部」），返回每处匹配的工作表、地址、值和公式',
        use='定位某个科目/客户/关键字，或查哪些公式引用了某张表（scope=formulas 搜 Sheet2!）',
        avoid='读取整块数据（用 read_range）',
        check='total 是真实命中数；status=partial 表示只列出了前 max_results 个，按 next 缩小范围',
    ),
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
    _describe(
        '列出工作簿里的对象：sheets（工作表、可见性、已用区域）、names（定义的名称及引用）、tables（表格及列名）、pivots（数据透视表及数据源）、charts（图表及所在表）',
        use='弄清工作簿结构、找表名/名称/图表名',
        check='data 里就是清单，ok=true 即可直接用',
    ),
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
    _describe(
        '分析一张工作表的结构：blocks 层给出有几块数据、每块的标题、多级表头（列名路径如「扣款/养老」）、数据行范围、合计行、每列类型；formulas 层按 R1C1 归纳公式模式（如 =[数量]*[单价] 覆盖 D2:D500）、标出同列公式不一致，并列出异常候选（合计漏行、被改成死值、孤立偏离、断链引用）；objects 层列表格/图表/透视表；dependencies 层列跨表和外部工作簿引用',
        use='接手陌生或复杂的表（多级表头、一张表上好几块、公式多）时先用它',
        avoid='只要几个具体单元格的值（用 read_range）',
        check='每层带 status：complete 是全表结论，partial 表示大表只读了一部分；header_uncertain=true 时先用 read_range 读前几行确认表头；异常只是候选，改之前先问用户',
        on_fail='按 next 改用 list(kind=sheets) 和 read_range 直接读',
    ),
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
    _describe(
        '大工作簿分头摸底：把问题拆成最多 4 个互不重叠的子任务（通常按表分组），每个由一个独立的只读子 agent 并行去查（它们能用 list / find / inspect_sheet / read_range），只把结论交回来，你的上下文不会被几十次读取的原始数据塞满',
        use='表多（5 张以上）或表大、而问题涉及多张表',
        avoid='一两张表、或只查一个具体位置（直接自己查）',
        check='结论未经你核实，写入或向用户下结论前对关键位置用 read_range 确认；有的子任务没完成时看 tasks 里各自的 status',
        on_fail='按 next 改用 list / find / inspect_sheet 自己查',
    ),
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
    tools = [t for t in register_all_tools() if t.name in explorer.EXPLORER_TOOLS]

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


@tool("write_formula", _describe(
        '向一个单元格写入公式（以 = 开头）；写后自动体检，并检查这一列公式模式是否和上下一致',
        use='用户要求写公式，或需要一个会随数据更新的计算结果',
        avoid='写纯文本或数字（用 write_value）、整列同一公式（写一格后用 fill_formula_down）',
        check=_CHECK_WRITE,
    ), {"address": str, "formula": str})
async def write_formula(args):
    result = await call_csharp("write_formula", {
        "address": args["address"],
        "formula": args["formula"],
    })
    return _wrap_result(await _pattern_check(result, args["address"]))


@tool("write_value", _describe(
        '向一个单元格写入纯文本或数字，不解析为公式（写「张三」显示张三而不是 ="张三"）',
        use='改一个单元格的值',
        avoid='写一片区域（用 write_range，快得多）、写公式（用 write_formula）',
        check=_CHECK_WRITE,
    ), {"address": str, "value": str})
async def write_value(args):
    result = await call_csharp("write_value", {
        "address": args["address"],
        "value": args["value"],
    })
    return _wrap_result(result)


@tool("write_range", _describe(
        '从左上角单元格起批量写入二维数组（比逐格 write_value 快 100 倍，无弹窗）；address 是左上角（如 A1 或 Sheet3!A1），values 是二维数组，以 = 开头的元素按公式写入',
        use='写表头和多行数据、一次填一整块',
        avoid='只改一个格（用 write_value）',
        check=_CHECK_WRITE,
    ), {"address": str, "values": list})
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
    _describe(
        '向用户提问以澄清模糊指令：面板把问题渲染成提问卡，用户点选或自己填写后一次提交。只问一件事用 question + options；几处不确定时用 questions 一次问清（最多 4 题）。每题：question 问题；header 2~6 字的短标签（如「汇总口径」）；options 2~4 个选项，可写成 {label: 选项, description: 一句话说明后果}；multi_select 为 true 时可多选；用户总能选「其他」自己填，不必把「其他」列进选项',
        use='指令有歧义、而猜错代价大（覆盖数据、口径不同结果不同）',
        avoid='一轮一轮地追问（一次问清）、能从表里看出答案的事',
        check='data.user_answer 是用户的回答，按它继续',
    ),
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
    _describe(
        '维护本次任务的计划清单（面板顶部显示进度）。todos 是完整清单（每次都传全部条目）：[{content: 步骤描述, status: pending|in_progress|completed}]',
        use='三步以上的任务开始前先列计划；每开始一步把它标为 in_progress（同一时间只有一条），做完立刻标 completed；发现新步骤就加进去',
        avoid='一两步就能做完的事',
        check=_CHECK_PLAIN,
    ),
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
    _describe(
        '提交一份变更方案给用户批准。summary 一句话说要做什么；steps 每一步 {action: 做什么, target: 表和区域, detail: 怎么做}；risks 风险点（会覆盖什么、有什么不确定）',
        use='「只出方案」模式下必须用它收尾；其他模式下改动很大时也可以先用它对齐',
        check='提交后本轮结束：不要再调用工具，用一两句话说明方案要点，请用户批准',
    ),
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
    _describe(
        '整份重写这个工作簿的记忆（NOTES.md，只存在用户电脑上，下次会话自动带给你）。保留四个小节：## 结构怪癖 / ## 用户偏好 / ## 做过的改动 / ## 禁区；禁区每行一个「- 表名」或「- 表名!A1:D20」，只能加不能删。notes 是完整的新内容（Markdown），不是增量',
        use='发现了值得下次记住的结构怪癖、用户偏好、重要改动或禁区',
        avoid='记具体数据、这次对话的过程',
        check=_CHECK_PLAIN,
        on_fail='工作簿没保存过时没有地方存：告诉用户先保存，这次要记的内容直接在回复里说明',
    ),
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
    _describe(
        '读取一份专业知识（<knowledge-skills> 里列出的技能）。先只传 name 读正文；正文里提到补充文件（如 wps.md）时再传 file 读它',
        use='任务涉及技能清单里的主题时，动手前先读',
        check='data.content 是正文',
        on_fail='名字写错时按 next 列出的可用名字再试',
    ),
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


@tool("read_workbook", _describe(
        '读取当前工作簿的结构信息（工作表、已用区域等）',
        use='消息里没有「工作簿结构」摘要、或摘要标了「索引不完整」时',
        avoid='摘要已经够用时（直接用摘要）',
        check=_CHECK_PLAIN,
    ), {})
async def read_workbook(args):
    result = await call_csharp("read_workbook", {})
    return _wrap_result(result)


@tool("read_attachment", _describe(
        '读取用户上传的附件内容（xlsx/xls/csv/txt/json 等）',
        use='用户消息提到附件或要求处理附件文件时',
        avoid='读工作簿本身的数据（用 read_range）',
        check=_CHECK_PLAIN,
        on_fail='file_name 用用户消息里给出的附件文件名，不要自己猜',
    ), {"file_name": str})
async def read_attachment(args):
    result = await call_csharp("read_attachment", {"file_name": args["file_name"]})
    return _wrap_result(result)


@tool("read_selection", _describe(
        '读取用户当前选中的单元格：地址、值和公式',
        use='用户说「这里」「选中的」「这几格」时',
        check=_CHECK_PLAIN,
    ), {})
async def read_selection(args):
    result = await call_csharp("read_selection", {})
    return _wrap_result(result)


@tool("fill_formula_down", _describe(
        '把 from_address 里的公式向下填充 row_count 行（相对引用跟着变）',
        use='整列套用同一公式：先 write_formula 写第一格，再用它填充',
        check=_CHECK_WRITE,
    ), {"from_address": str, "row_count": int})
async def fill_formula_down(args):
    result = await call_csharp("fill_formula_down", {
        "from_address": args["from_address"],
        "row_count": args["row_count"],
    })
    # 与 C# 的写入区域一致：起始行 + row_count 行
    rows = (_optional_int(args, "row_count") or 0) + 1
    return _wrap_result(await _pattern_check(result, args["from_address"], rows=rows))


@tool("replace_formula", _describe(
        '在指定区域内批量替换公式文本里的字符串（如把 Sheet1! 换成 Sheet2!）',
        use='批量改引用、改函数名',
        avoid='改单元格的值（用 write_value / write_range）',
        check=_CHECK_WRITE,
    ), {"range_address": str, "find": str, "replace": str})
async def replace_formula(args):
    result = await call_csharp("replace_formula", {
        "range_address": args["range_address"],
        "find": args["find"],
        "replace": args["replace"],
    })
    return _wrap_result(await _pattern_check(result, args["range_address"]))


@tool("clean_data", _describe(
        '对区域执行数据清洗操作：unify_date / remove_duplicates / highlight_missing / trim_spaces / text_to_number',
        use='去重（remove_duplicates）、统一日期、去首尾空格、文本型数字转数字、标出缺失值',
        check=_CHECK_PLAIN + '；remove_duplicates 会删行，先告诉用户删了多少',
    ), {"range_address": str, "operations": list})
async def clean_data(args):
    result = await call_csharp("clean_data", {
        "range_address": args["range_address"],
        "operations": args["operations"],
    })
    return _wrap_result(result)


@tool("create_chart", _describe(
        '基于数据区域创建图表（chart_type 如 column / line / pie / bar）',
        use='用户要图表；图表类型由你结合数据判断后显式传入',
        avoid='柱状+折线双轴（用 create_combo_chart）',
        check=_CHECK_PLAIN + '；之后改标题/颜色/标签要用图表名时，用 list(kind=charts) 查',
    ), {"data_range": str, "chart_type": str, "title": str, "x_label": str, "y_label": str})
async def create_chart(args):
    result = await call_csharp("create_chart", {
        "data_range": args["data_range"],
        "chart_type": args.get("chart_type", "column"),
        "title": args.get("title", ""),
        "x_label": args.get("x_label", ""),
        "y_label": args.get("y_label", ""),
    })
    return _wrap_result(result)


@tool("create_pivot_table", _describe(
        '基于数据区域在目标工作表创建数据透视表（行字段、列字段、值字段、汇总方式）',
        use='按维度汇总、交叉统计',
        avoid='只要一个简单合计（用 write_formula 写 SUM / SUMIF）',
        check=_CHECK_PLAIN + '；之后刷新或调整时传你起的 pivot_table_name',
    ), {"source_range": str, "destination_sheet": str, "pivot_table_name": str, "row_fields": list, "column_fields": list, "value_fields": list, "value_function": str})
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


@tool("execute_vba", _describe(
        '执行一段 VBA（入口过程按 Sub DeepExcel_TempMacro 运行，只写过程体也可以），会弹安全确认窗',
        use='现有工具做不到的操作',
        avoid='现有工具能完成的事（优先用专门工具，它们有备份和体检）',
        check='ok=true 只说明宏跑完了，不说明结果对：涉及数据时用 read_range 读回关键位置',
        on_fail='按 error 修代码后再试一次；用户取消了就不要重试',
    ), {"code": str})
async def execute_vba(args):
    result = await call_csharp("execute_vba", {"code": args["code"]})
    return _wrap_result(result)


@tool("execute_python", _describe(
        '执行纯计算的 Python 代码（字符串处理、正则、数学计算），碰不到工作簿，失败也不影响工作簿',
        use='复杂文本处理或计算，算完再用写入工具写回',
        avoid='读写工作簿（用对应工具或 execute_vba）、写文件到磁盘',
        check='ok=true 时 data 里是代码的输出',
    ), {"code": str})
async def execute_python(args):
    result = await call_csharp("execute_python", {"code": args["code"]})
    return _wrap_result(result)


@tool("create_snapshot", _describe(
        '手动给当前工作簿做一个快照',
        use='大改之前想多留一个回退点（修改类工具执行前已会自动备份，一般不需要）',
        check='返回的快照 ID 可传给 rollback',
    ), {})
async def create_snapshot(args):
    result = await call_csharp("create_snapshot", {})
    return _wrap_result(result)


@tool("add_sheet", _describe(
        '在当前工作簿添加一张新的空白工作表',
        use='要把结果放到新表、或用户要新建表',
        check=_CHECK_PLAIN,
        on_fail='重名时换个名字',
    ), {"name": str})
async def add_sheet(args):
    result = await call_csharp("add_sheet", {"name": args["name"]})
    return _wrap_result(result)


@tool("delete_sheet", _describe(
        '删除指定名称的工作表',
        use='用户明确要删一张表',
        check=_CHECK_PLAIN + '；其他表引用它的公式会变成 #REF!，删之前先向用户确认',
    ), {"name": str})
async def delete_sheet(args):
    result = await call_csharp("delete_sheet", {"name": args["name"]})
    return _wrap_result(result)


@tool("rename_sheet", _describe(
        '重命名工作表',
        use='改表名',
        check=_CHECK_PLAIN,
        on_fail='重名或含非法字符时换个名字',
    ), {"old_name": str, "new_name": str})
async def rename_sheet(args):
    result = await call_csharp("rename_sheet", {
        "old_name": args["old_name"],
        "new_name": args["new_name"],
    })
    return _wrap_result(result)


@tool("set_number_format", _describe(
        '设置区域的数字格式（如 #,##0.00 / 0% / yyyy-mm-dd）',
        use='改显示格式；值本身不变',
        avoid='把文本转成数字（用 clean_data 的 text_to_number 或 clean_amount）',
        check=_CHECK_PLAIN,
    ), {"address": str, "format": str})
async def set_number_format(args):
    result = await call_csharp("set_number_format", {
        "address": args["address"],
        "format": args["format"],
    })
    return _wrap_result(result)


@tool("set_column_width", _describe(
        '设置列宽；auto_fit=true 时自动适应内容宽度（忽略 width）',
        use='内容显示不全或排版需要',
        check=_CHECK_PLAIN,
    ), {"address": str, "width": float, "auto_fit": bool})
async def set_column_width(args):
    result = await call_csharp("set_column_width", {
        "address": args["address"],
        "width": args.get("width", 10.0),
        "auto_fit": args.get("auto_fit", False),
    })
    return _wrap_result(result)


@tool("sort_data", _describe(
        "对区域按某一列排序（整行一起移动，不会交换列）。range_address 含表头时从表头行开始；sort_column 写列字母如 'B' 或列序号如 '2'；descending 为 true 降序；has_header【必传】第一行是否表头，要根据读过的内容判断",
        use='按某列排序',
        avoid='只看前几名（读出来自己排）',
        check=_CHECK_PLAIN,
    ), {"range_address": str, "sort_column": str, "descending": bool, "has_header": bool})
async def sort_data(args):
    result = await call_csharp("sort_data", {
        "range_address": args["range_address"],
        "sort_column": args["sort_column"],
        "descending": args.get("descending", False),
        "has_header": args.get("has_header", False),
    })
    return _wrap_result(result)


@tool("filter_data", _describe(
        "对区域应用自动筛选（column_index 从 1 开始，criteria 如 '>100' 或 '北京'）",
        use='用户要在表上筛出符合条件的行',
        avoid='只是要统计符合条件的数量（用 COUNTIF 公式）',
        check=_CHECK_PLAIN,
    ), {"range_address": str, "column_index": int, "criteria": str})
async def filter_data(args):
    result = await call_csharp("filter_data", {
        "range_address": args["range_address"],
        "column_index": args["column_index"],
        "criteria": args["criteria"],
    })
    return _wrap_result(result)


@tool("merge_cells", _describe(
        '合并指定区域的单元格',
        use='做标题、排版',
        avoid='数据区域（合并会让排序、筛选、公式出问题，数据区先劝用户别合并）',
        check=_CHECK_PLAIN,
    ), {"address": str})
async def merge_cells(args):
    result = await call_csharp("merge_cells", {"address": args["address"]})
    return _wrap_result(result)


@tool("unmerge_cells", _describe(
        '拆分区域内所有合并单元格',
        use='要对数据排序、筛选、公式处理前',
        check=_CHECK_PLAIN + '；拆开后只有左上角有值，需要时接着用 fill_blank_cells 补齐',
    ), {"address": str})
async def unmerge_cells(args):
    result = await call_csharp("unmerge_cells", {"address": args["address"]})
    return _wrap_result(result)


@tool("set_cell_style", _describe(
        '设置单元格样式：字体、字号、粗体、斜体、字体颜色、背景色、对齐、自动换行（颜色可写 hex 如 #FF0000 或 red/blue/green）',
        use='表头加粗、标色、对齐',
        avoid='按条件着色（用 apply_conditional_format）',
        check=_CHECK_PLAIN,
    ), {"address": str, "font_name": str, "font_size": float, "bold": bool, "italic": bool, "font_color": str, "bg_color": str, "h_align": str, "v_align": str, "wrap_text": bool})
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


@tool("copy_range", _describe(
        '把源区域复制到目标位置（含格式和公式）；目标区域比源大时按源的形状重复填充',
        use='复制一块表、把公式复制到别处',
        check=_CHECK_WRITE,
    ), {"source_address": str, "dest_address": str})
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


@tool("clear_range", _describe(
        '清空区域：clear_type 为 contents 只清内容、formats 只清格式、all 全部清掉',
        use='用户要清掉一块内容或格式',
        avoid='删除整行整列（用 delete_rows / delete_columns）',
        check=_CHECK_PLAIN,
    ), {"address": str, "clear_type": str})
async def clear_range(args):
    result = await call_csharp("clear_range", {
        "address": args["address"],
        "clear_type": args.get("clear_type", "all"),
    })
    return _wrap_result(result)


@tool("insert_rows", _describe(
        '在第 row 行前插入 count 行',
        use='要在表中间加行',
        check=_CHECK_PLAIN,
    ), {"row": int, "count": int})
async def insert_rows(args):
    result = await call_csharp("insert_rows", {
        "row": args["row"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("delete_rows", _describe(
        '从第 row 行开始删除 count 行',
        use='用户明确要删某几行',
        avoid='删空行（用 delete_blank_rows）',
        check=_CHECK_PLAIN + '；删行会让下方数据上移、引用它的公式可能变 #REF!',
    ), {"row": int, "count": int})
async def delete_rows(args):
    result = await call_csharp("delete_rows", {
        "row": args["row"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("insert_columns", _describe(
        '在第 column 列前插入 count 列（column 从 1 开始）',
        use='要在表中间加列',
        check=_CHECK_PLAIN,
    ), {"column": int, "count": int})
async def insert_columns(args):
    result = await call_csharp("insert_columns", {
        "column": args["column"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("delete_columns", _describe(
        '从第 column 列开始删除 count 列（column 从 1 开始）',
        use='用户明确要删某几列',
        check=_CHECK_PLAIN + '；引用被删列的公式会变 #REF!',
    ), {"column": int, "count": int})
async def delete_columns(args):
    result = await call_csharp("delete_columns", {
        "column": args["column"],
        "count": args.get("count", 1),
    })
    return _wrap_result(result)


@tool("freeze_panes", _describe(
        "冻结窗格：在指定单元格的左上方冻结（'B2' 冻结 A 列和第 1 行）",
        use='表头或首列滚动时要保持可见',
        check=_CHECK_PLAIN,
    ), {"address": str})
async def freeze_panes(args):
    result = await call_csharp("freeze_panes", {"address": args["address"]})
    return _wrap_result(result)


@tool("apply_conditional_format", _describe(
        '给区域应用条件格式（rule_type: color_scale / data_bar / highlight_rules / cell_value）',
        use='按数值大小着色、数据条、标出超过阈值的格',
        avoid='固定颜色（用 set_cell_style）',
        check=_CHECK_PLAIN,
    ), {"address": str, "rule_type": str, "rule_args": dict})
async def apply_conditional_format(args):
    result = await call_csharp("apply_conditional_format", {
        "address": args["address"],
        "rule_type": args["rule_type"],
        "rule_args": args.get("rule_args", {}),
    })
    return _wrap_result(result)


@tool("write_table", _describe(
        '把区域转换为 Excel 表格（ListObject，自带筛选和样式）',
        use='用户要「套用表格格式」、要结构化引用',
        check=_CHECK_PLAIN,
        on_fail='区域和已有表格重叠时先用 list(kind=tables) 看已有表格',
    ), {"address": str, "table_name": str})
async def write_table(args):
    result = await call_csharp("write_table", {
        "address": args["address"],
        "table_name": args.get("table_name", ""),
    })
    return _wrap_result(result)


@tool("rollback", _describe(
        '把工作簿恢复到指定快照；恢复前会自动另存当前状态，恢复本身也可撤销',
        use='撤销修改：传 backup_snapshot_id 撤销本回合的全部修改；传某一步结果里的 checkpoint_id 只撤销这一步及之后',
        check='ok=true 后告诉用户恢复到了哪一步',
    ), {"snapshot_id": str})
async def rollback(args):
    result = await call_csharp("rollback", {"snapshot_id": args["snapshot_id"]})
    return _wrap_result(result)


@tool("delete_blank_rows", _describe(
        '删除区域内整行都为空的行，下方数据自动上移',
        use='清掉数据中间的空行',
        avoid='删除指定行（用 delete_rows）',
        check=_CHECK_PLAIN + '；会删行，汇报时说删了几行',
    ), {"range_address": str})
async def delete_blank_rows(args):
    result = await call_csharp("delete_blank_rows", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("split_text_to_columns", _describe(
        '按分隔符把一列文本拆到右侧多列',
        use='「张三,销售部」这类合在一格的信息要分开',
        check=_CHECK_PLAIN + '；拆出的列会覆盖右侧原有内容，右侧有数据时先确认',
    ), {"range_address": str, "delimiter": str})
async def split_text_to_columns(args):
    result = await call_csharp("split_text_to_columns", {
        "range_address": args["range_address"],
        "delimiter": args.get("delimiter", ","),
    })
    return _wrap_result(result)


@tool("fill_blank_cells", _describe(
        '向下填充空白单元格：每个空格填上方最近的非空值',
        use='合并单元格拆开后、或分组标签只写在第一行时补齐',
        check=_CHECK_PLAIN,
    ), {"range_address": str})
async def fill_blank_cells(args):
    result = await call_csharp("fill_blank_cells", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("highlight_duplicates", _describe(
        '把区域内重复的值标成浅红色背景（color 可改）',
        use='只想标出重复、不删除',
        avoid='直接删重复（用 clean_data 的 remove_duplicates）',
        check=_CHECK_PLAIN,
    ), {"range_address": str, "color": str})
async def highlight_duplicates(args):
    result = await call_csharp("highlight_duplicates", {
        "range_address": args["range_address"],
        "color": args.get("color", "#FFC7CE"),
    })
    return _wrap_result(result)


@tool("remove_special_chars", _describe(
        '去除文本里的特殊字符（非打印字符、不可见字符、不间断空格等）',
        use='从系统导出的数据匹配不上、看起来一样却不相等时',
        check=_CHECK_PLAIN,
    ), {"range_address": str})
async def remove_special_chars(args):
    result = await call_csharp("remove_special_chars", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("clean_amount", _describe(
        '清洗金额：去掉货币符号（¥/$/€）和千分位逗号，转成数字',
        use='金额列是文本、求和为 0 或报错时',
        check=_CHECK_PLAIN,
    ), {"range_address": str})
async def clean_amount(args):
    result = await call_csharp("clean_amount", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("merge_columns", _describe(
        '把多列内容用分隔符连成一列（跳过空值），结果写到 target_column',
        use='姓和名、省市区这类要拼成一列',
        avoid='合并单元格（用 merge_cells）',
        check=_CHECK_PLAIN,
    ), {"range_address": str, "delimiter": str, "target_column": str})
async def merge_columns(args):
    result = await call_csharp("merge_columns", {
        "range_address": args["range_address"],
        "delimiter": args.get("delimiter", " "),
        "target_column": args.get("target_column", ""),
    })
    return _wrap_result(result)


@tool("rename_columns", _describe(
        '批量重命名列标题（改区域第一行的列名）',
        use='统一或改写表头',
        check=_CHECK_PLAIN,
    ), {"range_address": str, "new_names": list})
async def rename_columns(args):
    result = await call_csharp("rename_columns", {
        "range_address": args["range_address"],
        "new_names": args["new_names"],
    })
    return _wrap_result(result)


@tool("collapse_spaces", _describe(
        '把文本里连续的多个空格压成一个',
        use='文本中间有多余空格导致匹配失败',
        check=_CHECK_PLAIN,
    ), {"range_address": str})
async def collapse_spaces(args):
    result = await call_csharp("collapse_spaces", {"range_address": args["range_address"]})
    return _wrap_result(result)


@tool("add_data_labels", _describe(
        '给图表加数据标签（chart_name 为空时操作当前表的第一个图表）',
        use='用户要在图上显示数值、类别或百分比',
        check=_CHECK_PLAIN,
        on_fail='找不到图表时先用 list(kind=charts) 查图表名',
    ), {"chart_name": str, "position": str, "show_value": bool, "show_category_name": bool, "show_percentage": bool})
async def add_data_labels(args):
    result = await call_csharp("add_data_labels", {
        "chart_name": args.get("chart_name", ""),
        "position": args.get("position", "outside_end"),
        "show_value": args.get("show_value", True),
        "show_category_name": args.get("show_category_name", False),
        "show_percentage": args.get("show_percentage", False),
    })
    return _wrap_result(result)


@tool("set_chart_title", _describe(
        '设置或修改图表标题（chart_name 为空时操作当前表的第一个图表）',
        use='改图表标题',
        check=_CHECK_PLAIN,
        on_fail='找不到图表时先用 list(kind=charts) 查图表名',
    ), {"chart_name": str, "title": str})
async def set_chart_title(args):
    result = await call_csharp("set_chart_title", {
        "chart_name": args.get("chart_name", ""),
        "title": args.get("title", ""),
    })
    return _wrap_result(result)


@tool("set_chart_colors", _describe(
        "按顺序设置图表各系列颜色（hex 数组，如 ['#FF0000','#00FF00']）",
        use='用户指定配色',
        check=_CHECK_PLAIN,
        on_fail='找不到图表时先用 list(kind=charts) 查图表名',
    ), {"chart_name": str, "colors": list})
async def set_chart_colors(args):
    result = await call_csharp("set_chart_colors", {
        "chart_name": args.get("chart_name", ""),
        "colors": args.get("colors", []),
    })
    return _wrap_result(result)


@tool("create_combo_chart", _describe(
        '创建组合图（柱状 + 折线双轴），line_series_index 指定哪个系列画成折线',
        use='两个量纲不同的指标要放在一张图（如销售额和增长率）',
        check=_CHECK_PLAIN + '；之后修改要用图表名时，用 list(kind=charts) 查',
    ), {"data_range": str, "title": str, "x_label": str, "y_label": str, "secondary_y_label": str, "line_series_index": int})
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


@tool("export_chart", _describe(
        '把图表导出为图片文件（PNG/JPG/GIF/BMP），返回输出路径',
        use='用户要图表图片',
        check='ok=true 时把返回的路径告诉用户',
    ), {"chart_name": str, "output_path": str, "format": str})
async def export_chart(args):
    result = await call_csharp("export_chart", {
        "chart_name": args.get("chart_name", ""),
        "output_path": args.get("output_path", ""),
        "format": args.get("format", "png"),
    })
    return _wrap_result(result)


@tool("refresh_pivot", _describe(
        '刷新数据透视表，让它反映源数据的最新变化',
        use='改过源数据之后',
        check=_CHECK_PLAIN,
        on_fail='找不到透视表时先用 list(kind=pivots) 查名称',
    ), {"pivot_table_name": str, "sheet_name": str})
async def refresh_pivot(args):
    result = await call_csharp("refresh_pivot", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("group_pivot_date", _describe(
        "把透视表的日期字段按年/季度/月/日分组（group_by 如 'year,month'）",
        use='按月、按季度汇总',
        check=_CHECK_PLAIN,
        on_fail='日期字段里有文本或空值时分组会失败，先清洗源数据再刷新',
    ), {"pivot_table_name": str, "field_name": str, "group_by": str, "sheet_name": str})
async def group_pivot_date(args):
    result = await call_csharp("group_pivot_date", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "field_name": args["field_name"],
        "group_by": args.get("group_by", "month"),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("set_pivot_value_display", _describe(
        '设置透视表值的显示方式：normal / percent_of_column / percent_of_row / percent_of_total / running_total / rank',
        use='要占比、累计或排名',
        check=_CHECK_PLAIN,
    ), {"pivot_table_name": str, "value_field": str, "display_type": str, "base_field": str, "sheet_name": str})
async def set_pivot_value_display(args):
    result = await call_csharp("set_pivot_value_display", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "value_field": args["value_field"],
        "display_type": args.get("display_type", "normal"),
        "base_field": args.get("base_field", ""),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("set_pivot_totals", _describe(
        '控制透视表的行总计和列总计是否显示',
        use='用户要隐藏或显示总计',
        check=_CHECK_PLAIN,
    ), {"pivot_table_name": str, "show_row_totals": bool, "show_column_totals": bool, "sheet_name": str})
async def set_pivot_totals(args):
    result = await call_csharp("set_pivot_totals", {
        "pivot_table_name": args.get("pivot_table_name", ""),
        "show_row_totals": args.get("show_row_totals", True),
        "show_column_totals": args.get("show_column_totals", True),
        "sheet_name": args.get("sheet_name", ""),
    })
    return _wrap_result(result)


@tool("add_pivot_slicer", _describe(
        '给透视表的某个字段加切片器（交互式筛选按钮）',
        use='用户要能点选筛选的报表',
        check=_CHECK_PLAIN,
    ), {"pivot_table_name": str, "field_name": str, "sheet_name": str})
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

@tool("screenshot_excel", _describe(
        '截图 Excel 主窗口（Computer Use），返回 base64 JPEG（最大宽 1280px，<300KB），无需传参',
        use='用户主动要求截图 / computer use 时',
        avoid='主动截图验证工具执行效果（用工具结果和 read_range）',
        check=_CHECK_PLAIN,
    ), {})
async def screenshot_excel(args):
    result = await call_csharp("screenshot_excel", {})
    return _wrap_result(result)


@tool("send_keys", _describe(
        '向 Excel 窗口模拟键盘输入（Computer Use）。keys 语法：{ENTER} {ESC} {TAB} {UP} {DOWN} 等，+Shift ^Ctrl %Alt 前缀',
        use='操作 Excel 原生对话框/弹窗（如 {ESC} 关掉弹窗）',
        avoid='用快捷键替代专门工具（保存、复制、撤销等都有对应工具）',
        check=_CHECK_PLAIN,
    ), {"keys": str})
async def send_keys(args):
    result = await call_csharp("send_keys", {"keys": args["keys"]})
    return _wrap_result(result)


def register_all_tools() -> list:
    """返回给模型的全部工具。

    这份列表同时决定 MCP 注册和 allowed_tools，是工具清单的唯一来源；
    system_prompt.py 的 <available-tools> 由 tests/test_excel_tools.py 与它对齐。
    打开 DEEPEXCEL_TOOLS_V2 时，30 个写入/结构/图表/透视工具换成 4 个合并工具（见 tools_v2）。
    """
    import tools_v2
    tools = base_tools()
    if not tools_v2.enabled():
        return tools
    return [t for t in tools if t.name not in tools_v2.REPLACED] + tools_v2.build_tools()


def base_tools() -> list:
    """一个个独立的工具（v1 工具集）；合并工具的每一步也是调用它们。"""
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
        execute_vba, execute_python,
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
    return tools
