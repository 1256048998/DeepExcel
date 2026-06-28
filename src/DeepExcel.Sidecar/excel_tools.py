# src/DeepExcel.Sidecar/excel_tools.py
import json
from claude_agent_sdk import tool
from ipc import call_csharp, call_csharp_clarify


def _wrap_result(csharp_result: dict) -> dict:
    """把 C# 返回的 dict 包装成 MCP 工具返回格式"""
    return {
        "content": [
            {"type": "text", "text": json.dumps(csharp_result, ensure_ascii=False)}
        ]
    }


@tool("echo", "回声测试工具，原样返回输入文本（用于验证 sidecar 管线）", {"text": str})
async def echo(args):
    return _wrap_result({"success": True, "data": {"echo": args["text"]}})


@tool("read_range", "读取指定范围的单元格数据", {"address": str})
async def read_range(args):
    result = await call_csharp("read_range", {"address": args["address"]})
    return _wrap_result(result)


@tool("write_formula", "向指定单元格写入 Excel 公式", {"address": str, "formula": str})
async def write_formula(args):
    result = await call_csharp("write_formula", {
        "address": args["address"],
        "formula": args["formula"],
    })
    return _wrap_result(result)


@tool("clarify_intent", "向用户提问以澄清模糊指令", {"question": str, "options": list})
async def clarify_intent(args):
    user_answer = await call_csharp_clarify(args["question"], args.get("options", []))
    return _wrap_result({"success": True, "data": {"user_answer": user_answer}})


@tool("read_workbook", "读取当前工作簿的结构信息", {})
async def read_workbook(args):
    result = await call_csharp("read_workbook", {})
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
    return _wrap_result(result)


@tool("replace_formula", "在指定范围内批量替换公式中的字符串", {"range_address": str, "find": str, "replace": str})
async def replace_formula(args):
    result = await call_csharp("replace_formula", {
        "range_address": args["range_address"],
        "find": args["find"],
        "replace": args["replace"],
    })
    return _wrap_result(result)


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


@tool("execute_python", "执行 Python 代码（操作 Excel）", {"code": str})
async def execute_python(args):
    result = await call_csharp("execute_python", {"code": args["code"]})
    return _wrap_result(result)


@tool("create_snapshot", "创建当前工作簿快照（用于回滚）", {})
async def create_snapshot(args):
    result = await call_csharp("create_snapshot", {})
    return _wrap_result(result)


@tool("rollback", "回滚到指定快照", {"snapshot_id": str})
async def rollback(args):
    result = await call_csharp("rollback", {"snapshot_id": args["snapshot_id"]})
    return _wrap_result(result)


def register_all_tools() -> list:
    """返回所有 @tool 装饰后的工具对象列表"""
    return [
        echo,
        read_workbook, read_selection, read_range,
        write_formula, fill_formula_down, replace_formula,
        clean_data, create_chart, create_pivot_table,
        execute_vba, execute_python,
        create_snapshot, rollback,
        clarify_intent,
    ]
