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


@tool("sort_data", "对指定区域排序（sort_column 可以是列字母如 'A' 或列序号如 '1'，descending=True 降序）", {"range_address": str, "sort_column": str, "descending": bool})
async def sort_data(args):
    result = await call_csharp("sort_data", {
        "range_address": args["range_address"],
        "sort_column": args["sort_column"],
        "descending": args.get("descending", False),
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
        add_sheet, delete_sheet, rename_sheet,
        set_number_format, set_column_width,
        sort_data, filter_data,
        merge_cells, unmerge_cells,
        set_cell_style, copy_range, clear_range,
        insert_rows, delete_rows, insert_columns, delete_columns,
        freeze_panes,
        apply_conditional_format, write_table,
        clarify_intent,
    ]
