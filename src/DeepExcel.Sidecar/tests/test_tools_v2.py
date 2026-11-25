"""Converged tool set behind DEEPEXCEL_TOOLS_V2: registration, prompt, per-step
gating through the existing PreToolUse checks, and write_cells' rollback."""
import json
import re

import pytest

import excel_tools
import permission_modes
import tools_v2


@pytest.fixture
def v2(monkeypatch):
    monkeypatch.setenv(tools_v2.ENV_SWITCH, "1")
    yield


def _names(tools):
    return {t.name for t in tools}


def test_off_by_default(monkeypatch):
    monkeypatch.delenv(tools_v2.ENV_SWITCH, raising=False)
    names = _names(excel_tools.register_all_tools())
    assert "write_cells" not in names and "write_range" in names


def test_on_replaces_thirty_tools_with_four(v2):
    names = _names(excel_tools.register_all_tools())
    assert {"write_cells", "edit_structure", "chart", "pivot"} <= names
    assert not (names & tools_v2.REPLACED)
    assert len(names) == len(excel_tools.base_tools()) - len(tools_v2.REPLACED) + 4


def test_every_replaced_tool_has_a_rename_entry():
    assert set(tools_v2.RENAMES) == set(tools_v2.REPLACED)


def test_the_v2_prompt_lists_exactly_the_v2_tools(v2):
    from system_prompt import SYSTEM_PROMPT
    prompt = tools_v2.system_prompt(SYSTEM_PROMPT)
    block = prompt.split("<available-tools>")[1].split("</available-tools>")[0]
    listed = set()
    for line in block.strip().splitlines():
        body = re.sub(r"（[^）]*）", "", line.split("：", 1)[-1])
        listed |= set(re.findall(r"\b[a-z]+(?:_[a-z]+)*\b", body))
    assert listed == _names(excel_tools.register_all_tools())
    assert "<tool-renames>" in prompt and "write_formula → write_cells（type=formula）" in prompt


def test_v2_descriptions_keep_the_shared_shape(v2):
    for t in tools_v2.build_tools():
        assert "何时用：" in t.description and "看返回：" in t.description, t.name


# ---- resolve -----------------------------------------------------------------

def test_resolve_maps_actions_and_steps():
    assert tools_v2.resolve("chart", {"action": "title", "chart_name": "c1", "title": "月度"}) == \
        [("set_chart_title", {"chart_name": "c1", "title": "月度"})]
    steps = tools_v2.resolve("write_cells", {"ops": [
        {"type": "add_sheet", "name": "汇总"},
        {"type": "values", "address": "汇总!A1", "values": [["月份", "金额"]]},
        {"type": "style", "address": "汇总!A1:B1", "bold": True},
    ]})
    assert [s[0] for s in steps] == ["add_sheet", "write_range", "set_cell_style"]
    assert steps[1][1] == {"address": "汇总!A1", "values": [["月份", "金额"]]}


@pytest.mark.parametrize("name,args", [
    ("chart", {"action": "explode"}),
    ("pivot", {}),
    ("write_cells", {"ops": []}),
    ("write_cells", {"ops": [{"type": "delete_everything"}]}),
    ("write_cells", {"ops": [{"type": "values"}] * (tools_v2.MAX_WRITE_OPS + 1)}),
])
def test_resolve_rejects_bad_calls(name, args):
    with pytest.raises(ValueError):
        tools_v2.resolve(name, args)


# ---- the hook checks each step like the old tool ------------------------------

@pytest.fixture
def hook(monkeypatch):
    import sidecar
    asked = []

    async def fake_request(name, tool_input):
        asked.append(name)
        return "allow"

    async def fake_write(msg):
        pass

    monkeypatch.setattr(sidecar, "request_permission", fake_request)
    monkeypatch.setattr(sidecar, "write_message", fake_write)
    monkeypatch.setattr(sidecar.workbook_memory, "current", lambda: None)
    sidecar._allowed_tools_session.clear()
    permission_modes.reset()

    async def call(tool, tool_input):
        out = await sidecar._pre_tool_use_hook({"tool_name": f"mcp__excel__{tool}", "tool_input": tool_input}, "id", None)
        return out["hookSpecificOutput"]["permissionDecision"], out
    yield call, asked
    sidecar._allowed_tools_session.clear()
    permission_modes.reset()


@pytest.mark.asyncio
async def test_risky_steps_get_their_own_confirmation(hook):
    call, asked = hook
    decision, _ = await call("write_cells", {"ops": [
        {"type": "style", "address": "A1", "bold": True},
        {"type": "values", "address": "A2", "values": [[1]]},
        {"type": "merge", "address": "A3:B3"},
    ]})
    assert decision == "allow"
    assert asked == ["write_range", "merge_cells"]


@pytest.mark.asyncio
async def test_plan_mode_refuses_a_composite_write(hook):
    call, asked = hook
    permission_modes.set_mode(permission_modes.PLAN)
    decision, out = await call("edit_structure", {"action": "delete_rows", "row": 3})
    assert decision == "deny" and "只出方案" in out["hookSpecificOutput"]["permissionDecisionReason"]


@pytest.mark.asyncio
async def test_a_protected_zone_refuses_the_whole_batch(hook, monkeypatch):
    import sidecar
    call, _ = hook

    class Memory:
        def notes(self):
            return "## 禁区\n- 汇总\n"

    monkeypatch.setattr(sidecar.workbook_memory, "current", lambda: Memory())
    monkeypatch.setattr(sidecar.workbook_memory, "active_sheet", lambda: "明细")
    decision, out = await call("write_cells", {"ops": [
        {"type": "values", "address": "明细!A1", "values": [[1]]},
        {"type": "values", "address": "汇总!A1", "values": [[2]]},
    ]})
    assert decision == "deny" and "禁区" in out["hookSpecificOutput"]["permissionDecisionReason"]


@pytest.mark.asyncio
async def test_a_bad_composite_call_is_refused_with_the_reason(hook):
    call, _ = hook
    decision, out = await call("chart", {"action": "explode"})
    assert decision == "deny" and "action" in out["hookSpecificOutput"]["permissionDecisionReason"]


# ---- write_cells execution -----------------------------------------------------

class _FakeTool:
    def __init__(self, name, results):
        self.name = name
        self._results = results
        self.calls = []

    async def handler(self, args):
        self.calls.append(args)
        payload = self._results.pop(0) if self._results else {"success": True}
        return excel_tools._wrap_result(payload)


@pytest.fixture
def fake_tools(monkeypatch):
    tools = {}
    rollbacks = []

    def get(name):
        return tools.setdefault(name, _FakeTool(name, []))

    async def fake_call_csharp(name, payload):
        assert name == "rollback"
        rollbacks.append(payload["snapshot_id"])
        return {"success": True}

    monkeypatch.setattr(tools_v2, "_tool", get)
    import ipc
    monkeypatch.setattr(ipc, "call_csharp", fake_call_csharp)
    return tools, get, rollbacks


@pytest.mark.asyncio
async def test_a_failed_step_rolls_back_the_earlier_ones(fake_tools):
    tools, get, rollbacks = fake_tools
    get("write_range")._results = [{"success": True, "checkpoint_id": "cp-1", "backup_snapshot_id": "bk"}]
    get("set_cell_style")._results = [{"success": False, "error": "颜色写错了", "suggestion": "用 #RRGGBB"}]
    out = json.loads((await tools_v2.run_write_cells({"ops": [
        {"type": "values", "address": "A1", "values": [[1]]},
        {"type": "style", "address": "A1", "bg_color": "紫红"},
    ]}))["content"][0]["text"])
    assert rollbacks == ["cp-1"]
    assert out["ok"] is False and "第 2 步（style）失败，前 1 步已撤销" in out["error"]
    assert out["next"] == "用 #RRGGBB" and out["backup_snapshot_id"] == "bk"


@pytest.mark.asyncio
async def test_a_failing_first_step_has_nothing_to_undo(fake_tools):
    tools, get, rollbacks = fake_tools
    get("write_range")._results = [{"success": False, "error": "地址不对"}]
    out = json.loads((await tools_v2.run_write_cells({"ops": [
        {"type": "values", "address": "??", "values": [[1]]}]}))["content"][0]["text"])
    assert rollbacks == [] and out["status"] == "failed" and "第 1 步" in out["error"]


@pytest.mark.asyncio
async def test_success_returns_samples_one_checkpoint_and_merged_checks(fake_tools):
    tools, get, rollbacks = fake_tools
    get("write_range")._results = [{"success": True, "checkpoint_id": "cp-1", "verification": {"ok": True}}]
    get("write_formula")._results = [{"success": True, "checkpoint_id": "cp-2",
                                      "verification": {"ok": False, "summary": "新增 1 个 #DIV/0!"}}]
    get("read_range")._results = [{"success": True, "data": {"values": [["月份", "金额"]]}},
                                  {"success": True, "data": {"values": [["#DIV/0!"]]}}]
    out = json.loads((await tools_v2.run_write_cells({"ops": [
        {"type": "values", "address": "A1:B1", "values": [["月份", "金额"]]},
        {"type": "formula", "address": "C2", "formula": "=B2/0"},
        {"type": "number_format", "address": "B:B", "format": "#,##0"},
    ]}))["content"][0]["text"])
    assert rollbacks == []
    assert out["checkpoint_id"] == "cp-1" and out["status"] == "check_failed"
    assert out["verification"]["summary"] == "新增 1 个 #DIV/0!"
    assert [s["address"] for s in out["data"]["samples"]] == ["A1:B1", "C2"]
    assert get("read_range").calls[0] == {"address": "A1:B1", "limit": tools_v2.SAMPLE_ROWS}


@pytest.mark.parametrize("address,values,expected", [
    ("A1", [["a", "b"], ["c", "d"]], "A1:B2"),
    ("汇总!C3", [[1, 2, 3]] * 9, "汇总!C3:E7"),
    ("'Q1 销售'!Z10", [[1, 2]], "'Q1 销售'!Z10:AA10"),
])
def test_values_steps_read_back_the_written_block(address, values, expected):
    assert tools_v2._sample_address("values", {"address": address, "values": values}) == expected


@pytest.mark.asyncio
async def test_a_step_missing_a_parameter_fails_cleanly_and_rolls_back(fake_tools):
    tools, get, rollbacks = fake_tools
    get("write_range")._results = [{"success": True, "checkpoint_id": "cp-1"}]

    class Broken:
        name = "write_formula"

        async def handler(self, args):
            return args["formula"]  # KeyError, like the real handler without it

    tools["write_formula"] = Broken()
    out = json.loads((await tools_v2.run_write_cells({"ops": [
        {"type": "values", "address": "A1", "values": [[1]]},
        {"type": "formula", "address": "B1"},
    ]}))["content"][0]["text"])
    assert rollbacks == ["cp-1"]
    assert out["status"] == "failed" and "参数缺失" in out["error"] and "重新读" in out["error"]
