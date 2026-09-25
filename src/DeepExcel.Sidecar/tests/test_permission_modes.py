"""权限模式：每步确认 / 本次会话自动应用写入 / 只出方案"""

import json

import pytest

import ipc
import permission_modes
from excel_tools import normalize_plan, register_all_tools


def _decision(result):
    return result["hookSpecificOutput"]["permissionDecision"]


@pytest.fixture
def hook(monkeypatch):
    import sidecar
    asked, sent = [], []

    async def fake_request(name, tool_input):
        asked.append(name)
        return "allow"

    async def fake_write(msg):
        sent.append(msg)

    monkeypatch.setattr(sidecar, "request_permission", fake_request)
    monkeypatch.setattr(sidecar, "write_message", fake_write)
    monkeypatch.setattr(sidecar.workbook_memory, "current", lambda: None)
    sidecar._allowed_tools_session.clear()
    permission_modes.reset()

    async def call(tool, tool_input=None):
        return await sidecar._pre_tool_use_hook(
            {"tool_name": f"mcp__excel__{tool}", "tool_input": tool_input or {}}, "id", None)

    yield call, asked
    sidecar._allowed_tools_session.clear()
    permission_modes.reset()


@pytest.mark.asyncio
async def test_default_mode_asks_for_bulk_writes(hook):
    call, asked = hook
    assert _decision(await call("write_range", {"address": "A1"})) == "allow"
    assert asked == ["write_range"]


@pytest.mark.asyncio
async def test_accept_writes_skips_confirmation_for_writes_but_not_for_dangerous_tools(hook):
    call, asked = hook
    permission_modes.set_mode(permission_modes.ACCEPT_WRITES)
    for tool in ("write_range", "fill_formula_down", "clean_amount", "merge_cells", "write_table"):
        assert _decision(await call(tool, {"address": "A1"})) == "allow"
    assert asked == []
    for tool in ("delete_sheet", "delete_rows", "clear_range", "execute_vba", "rollback"):
        await call(tool)
    assert asked == ["delete_sheet", "delete_rows", "clear_range", "execute_vba", "rollback"]


@pytest.mark.asyncio
async def test_plan_mode_denies_every_write_including_low_risk_ones(hook):
    call, asked = hook
    permission_modes.set_mode(permission_modes.PLAN)
    for tool in ("write_formula", "write_value", "set_cell_style", "add_sheet", "write_range", "execute_vba",
                 "execute_python", "create_chart", "rollback", "create_snapshot"):
        result = await call(tool, {"address": "A1"})
        assert _decision(result) == "deny", tool
        assert "只出方案" in result["hookSpecificOutput"]["permissionDecisionReason"]
    assert asked == []


@pytest.mark.asyncio
async def test_plan_mode_still_reads_and_presents(hook):
    call, _ = hook
    permission_modes.set_mode(permission_modes.PLAN)
    for tool in ("read_range", "find", "list", "inspect_sheet", "load_skill", "todo_write", "present_plan",
                 "clarify_intent"):
        assert _decision(await call(tool, {"address": "A1"})) == "allow", tool


def test_plan_mode_tool_set_is_really_read_only():
    import workbook_memory
    extra = permission_modes.PLAN_MODE_TOOLS - workbook_memory.READ_ONLY_TOOLS
    assert extra == set(), extra


def test_mode_rides_on_every_user_message_and_can_switch_mid_turn():
    permission_modes.reset()
    ipc.route_message({"type": "user_message", "text": "x", "permission_mode": "plan"})
    assert permission_modes.current() == "plan"
    ipc.route_message({"type": "set_permission_mode", "mode": "accept_writes"})
    assert permission_modes.current() == "accept_writes"
    # 不认识的值不改变模式（尤其不能因为拼错就从「只出方案」掉回可写）
    ipc.route_message({"type": "set_permission_mode", "mode": "yolo"})
    assert permission_modes.current() == "accept_writes"
    # 旧宿主不带这个字段：保持
    ipc.route_message({"type": "user_message", "text": "y"})
    assert permission_modes.current() == "accept_writes"
    permission_modes.reset()
    while not ipc._message_buffer["user_message"].empty():
        ipc._message_buffer["user_message"].get_nowait()


def test_turn_note_only_outside_default():
    permission_modes.reset()
    assert permission_modes.turn_note() == ""
    permission_modes.set_mode("plan")
    assert "present_plan" in permission_modes.turn_note()
    permission_modes.set_mode("accept_writes")
    assert "检查点" in permission_modes.turn_note()
    permission_modes.reset()


def test_mode_is_never_written_to_disk():
    import inspect
    source = inspect.getsource(permission_modes)
    assert "open(" not in source and "write_text" not in source and "json" not in source


def test_normalize_plan_caps_and_cleans():
    plan = normalize_plan({
        "summary": "  把金额改成万元  ",
        "steps": [{"action": "在 E 列写公式", "target": "明细!E2:E500", "detail": "=ROUND(D2/10000,2)"},
                  "直接一句话的步骤", {"target": "没有 action 的步骤被丢掉"}] + [{"action": f"第{i}步"} for i in range(30)],
        "risks": ["会覆盖 E 列", "", None],
    })
    assert plan["summary"] == "把金额改成万元"
    assert plan["steps"][0]["target"] == "明细!E2:E500"
    assert plan["steps"][1] == {"action": "直接一句话的步骤", "target": "", "detail": ""}
    assert len(plan["steps"]) <= 20
    assert plan["risks"] == ["会覆盖 E 列"]


@pytest.mark.asyncio
async def test_present_plan_emits_a_plan_proposal_event(monkeypatch):
    import ipc as ipc_module
    sent = []

    async def fake_write(msg):
        sent.append(msg)

    monkeypatch.setattr(ipc_module, "write_message", fake_write)
    tool = next(t for t in register_all_tools("wps") if t.name == "present_plan")
    out = await tool.handler({"summary": "整理花名册", "steps": [{"action": "统一日期", "target": "花名册!D:D"}]})
    payload = json.loads(out["content"][0]["text"])
    assert payload["success"] and "等待批准" in payload["data"]["message"]
    event = sent[0]["event"]
    assert event["kind"] == "plan_proposal" and event["steps"][0]["target"] == "花名册!D:D"


@pytest.mark.asyncio
async def test_present_plan_refuses_an_empty_plan():
    tool = next(t for t in register_all_tools("excel") if t.name == "present_plan")
    payload = json.loads((await tool.handler({"summary": "", "steps": []}))["content"][0]["text"])
    assert payload["success"] is False
