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


@tool("echo", "回声测试工具，原样返回输入文本（用于验证 sidecar 管线）", {"text": str})
async def echo(args):
    return _wrap_result({"success": True, "data": {"echo": args["text"]}})


@tool("read_range", "读取指定范围的单元格数据", {"address": str})
async def read_range(args):
    result = await call_csharp("read_range", {"address": args["address"]})
    return _wrap_result(result)


@tool("write_formula", "向指定单元格写入 Excel 公式（以 = 开头）", {"address": str, "formula": str})
async def write_formula(args):
    result = await call_csharp("write_formula", {
        "address": args["address"],
        "formula": args["formula"],
    })
    return _wrap_result(result)


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
    return _wrap_result(result)


@tool("clarify_intent", "向用户提问以澄清模糊指令", {"question": str, "options": list})
async def clarify_intent(args):
    user_answer = await call_csharp_clarify(args["question"], args.get("options", []))
    return _wrap_result({"success": True, "data": {"user_answer": user_answer}})


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


@tool("auto_analyze", "自动分析数据范围并生成统计报告（包含基础统计、异常值检测、图表建议）。data_range 为数据区域地址（如 A1:C100）。", {"data_range": str})
async def auto_analyze(args):
    result = await call_csharp("auto_analyze", {"data_range": args["data_range"]})
    return _wrap_result(result)


@tool("quick_summary", "快速生成数据摘要：读取指定范围，计算基础统计（求和、平均、最大、最小、计数），并返回一句话摘要。address 为数据区域地址。", {"address": str})
async def quick_summary(args):
    result = await call_csharp("read_range", {"address": args["address"]})
    if not isinstance(result, dict) or result.get("success") is not True:
        return _wrap_result(result)
    data = result.get("data") or result.get("values") or []
    if not data or not isinstance(data, list) or len(data) == 0:
        return _wrap_result({"success": True, "data": {"summary": "数据为空", "stats": {}}})
    nums = []
    for row in data:
        if isinstance(row, list):
            for cell in row:
                try:
                    nums.append(float(cell))
                except (ValueError, TypeError):
                    pass
    if len(nums) == 0:
        return _wrap_result({"success": True, "data": {"summary": "无数字数据", "stats": {"count": 0}}})
    total = sum(nums)
    avg = total / len(nums)
    max_val = max(nums)
    min_val = min(nums)
    summary = f"共 {len(nums)} 个数字，总和 {total:.2f}，平均 {avg:.2f}，最大 {max_val:.2f}，最小 {min_val:.2f}"
    return _wrap_result({
        "success": True,
        "data": {
            "summary": summary,
            "stats": {
                "count": len(nums),
                "sum": total,
                "average": avg,
                "max": max_val,
                "min": min_val,
            }
        }
    })


@tool("smart_chart", "智能创建图表：自动分析数据类型并选择最合适的图表类型。data_range 为数据范围，title 为图表标题。", {"data_range": str, "title": str})
async def smart_chart(args):
    result = await call_csharp("smart_chart", {
        "data_range": args["data_range"],
        "title": args.get("title", ""),
    })
    return _wrap_result(result)


@tool("create_plan", "创建执行计划，用于复杂任务的分步执行。tasks 为任务列表（字符串数组），按顺序排列。description 为计划描述（可选）。", {"tasks": list, "description": str})
async def create_plan(args):
    tasks = args.get("tasks", [])
    description = args.get("description", "")
    plan_id = f"plan_{abs(hash(str(tasks) + description)) % 1000000}"
    return _wrap_result({
        "success": True,
        "data": {
            "plan_id": plan_id,
            "tasks": tasks,
            "description": description,
            "total": len(tasks),
            "completed": 0,
            "current": 0,
        }
    })


@tool("update_plan", "更新执行计划进度。plan_id 为计划 ID，task_index 为完成的任务索引（从 0 开始），status 为状态（completed/failed）。", {"plan_id": str, "task_index": int, "status": str})
async def update_plan(args):
    plan_id = args.get("plan_id", "")
    task_index = args.get("task_index", 0)
    status = args.get("status", "completed")
    return _wrap_result({
        "success": True,
        "data": {
            "plan_id": plan_id,
            "task_index": task_index,
            "status": status,
            "message": f"任务 {task_index + 1} 已{status}",
        }
    })


# ============================ Computer Use 工具 ============================

@tool("screenshot_excel", "截图 Excel 主窗口（Computer Use）。仅在用户主动要求截图/computer use 时调用，禁止主动截图验证工具执行效果。返回 base64 JPEG（缩放到最大宽 1280px，<300KB）。无需传参。", {})
async def screenshot_excel(args):
    result = await call_csharp("screenshot_excel", {})
    return _wrap_result(result)


@tool("send_keys", "模拟键盘输入到 Excel 窗口（Computer Use）。仅用于操作 Excel 原生对话框/弹窗（如 {ESC} 关闭弹窗）。禁止用快捷键替代专门工具（保存/复制/撤销等用对应工具）。keys 语法：{ENTER}/{ESC}/{TAB}/{UP}/{DOWN} 等；+Shift ^Ctrl %Alt 前缀。", {"keys": str})
async def send_keys(args):
    result = await call_csharp("send_keys", {"keys": args["keys"]})
    return _wrap_result(result)


def register_all_tools() -> list:
    """返回所有 @tool 装饰后的工具对象列表"""
    return [
        echo,
        read_workbook, read_selection, read_range, read_attachment,
        write_formula, write_value, write_range, fill_formula_down, replace_formula,
        clean_data,
        delete_blank_rows, split_text_to_columns, fill_blank_cells,
        highlight_duplicates, remove_special_chars, clean_amount,
        merge_columns, rename_columns, collapse_spaces,
        create_chart, create_combo_chart, smart_chart,
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
        clarify_intent,
        auto_analyze, quick_summary,
        create_plan, update_plan,
        # ★ Computer Use 工具
        screenshot_excel, send_keys,
    ]
