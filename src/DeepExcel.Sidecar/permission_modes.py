# src/DeepExcel.Sidecar/permission_modes.py
"""权限模式：Claude Code 的 default / acceptEdits / plan 的工作簿版。

- default（每步确认）：高风险工具逐个弹确认，看变更预览。默认。
- accept_writes（本次会话自动应用写入）：批量写入、清洗不再逐个确认；删除、清空、回滚、
  跑代码仍然单独确认。每一步照常有检查点可以回退。
- plan（只出方案）：只能读，写入类工具在钩子里直接拒绝；模型用 present_plan 提交方案，
  用户批准后按选定的模式执行。

模式只存在于这个进程的内存里，不写任何配置文件（「允许并记住」只适用于能力型授权，
「自动应用写入」不是一种能力）。每条用户消息都带着面板当前的模式：侧车重启回到默认时，
不会出现面板显示「只出方案」、侧车却在写的情况。
"""

from __future__ import annotations

DEFAULT = "default"
ACCEPT_WRITES = "accept_writes"
PLAN = "plan"
MODES = (DEFAULT, ACCEPT_WRITES, PLAN)

LABELS = {DEFAULT: "每步确认", ACCEPT_WRITES: "自动应用写入", PLAN: "只出方案"}

# 自动应用写入模式下仍然要逐次确认的：不可预演的代码、删除、清空、回滚
ALWAYS_CONFIRM = frozenset({
    "execute_vba", "execute_jsa", "execute_python", "send_keys",
    "delete_rows", "delete_columns", "delete_sheet", "delete_blank_rows",
    "clear_range", "rollback",
})

# 只出方案模式下能用的：读、问、列计划、查知识、记笔记（不碰工作簿）、提交方案
PLAN_MODE_TOOLS = frozenset({
    "read_workbook", "read_selection", "read_range", "find", "list", "inspect_sheet",
    "explore_workbook", "read_attachment", "screenshot_excel",
    "clarify_intent", "todo_write", "load_skill", "update_workbook_notes", "present_plan",
})

_current = DEFAULT


def current() -> str:
    return _current


def set_mode(mode) -> bool:
    """合法就切换并返回 True；不认识的值忽略（保持原模式）"""
    global _current
    if mode in MODES:
        _current = mode
        return True
    return False


def reset() -> None:
    global _current
    _current = DEFAULT


def plan_denial(tool: str) -> str | None:
    """只出方案模式下这个工具不能用时返回拒绝理由"""
    if _current != PLAN or tool in PLAN_MODE_TOOLS:
        return None
    return ("现在是「只出方案」模式，不能修改工作簿。不要换别的工具绕过：把这一步写进方案，"
            "读完需要的信息后用 present_plan 提交方案，等用户批准后再执行。")


def auto_applies(tool: str) -> bool:
    """自动应用写入模式下，这个高风险工具不用再弹确认"""
    return _current == ACCEPT_WRITES and tool not in ALWAYS_CONFIRM


def turn_note() -> str:
    """附在用户消息前的模式说明；默认模式不加"""
    if _current == PLAN:
        return (
            "<permission-mode>只出方案：这一轮只读不写。先读懂相关的表（需要时 inspect_sheet / read_range / "
            "load_skill），然后调用 present_plan 提交方案：涉及的表和区域、每一步做什么、风险点。"
            "提交后用一两句话说明方案要点、请用户批准，本轮到此结束，不要执行任何修改。</permission-mode>\n"
        )
    if _current == ACCEPT_WRITES:
        return (
            "<permission-mode>自动应用写入：写入不再逐个请用户确认，每一步都有检查点可以回退；"
            "删除、清空、回滚和执行代码仍然需要确认。动手前照样先读、写完照样回读核对。</permission-mode>\n"
        )
    return ""
