"""explore_workbook：只读子 agent 并行摸底。用假的子会话（run_query）测编排逻辑。"""

import json
import time
from unittest.mock import AsyncMock, patch

import anyio
import pytest
from claude_agent_sdk import AssistantMessage, ResultMessage, TextBlock, ToolUseBlock

import explorer
import ui_events


def _assistant(*blocks):
    return AssistantMessage(content=list(blocks), model="fake")


def _result(text, is_error=False):
    return ResultMessage(subtype="success", duration_ms=1, duration_api_ms=1, is_error=is_error,
                         num_turns=2, session_id="s", result=text)


def _fake_query(script):
    """script(prompt) → [(delay_seconds, message), ...]"""
    async def run(prompt, tools):
        for delay, message in script(prompt):
            if delay:
                await anyio.sleep(delay)
            yield message
    return run


@pytest.fixture(autouse=True)
def _configured(monkeypatch):
    explorer.configure({"ANTHROPIC_BASE_URL": "https://example.invalid"}, "fake-model", "excel")
    yield
    monkeypatch.setattr(explorer, "run_query", explorer._default_query)


def test_tasks_are_normalized_and_capped():
    tasks = explorer.normalize_tasks([
        "应收在哪", {"question": " 看工资表 ", "sheets": "工资, 社保"}, {"question": ""}, 42,
        {"question": "a"}, {"question": "b"}, {"question": "c"},
    ])
    assert len(tasks) == explorer.MAX_TASKS
    assert tasks[0] == {"question": "应收在哪", "sheets": []}
    assert tasks[1] == {"question": "看工资表", "sheets": ["工资", "社保"]}
    assert "只看这些表：工资、社保" in explorer.task_prompt(tasks[1])
    assert "整个工作簿" in explorer.task_prompt(tasks[0])


@pytest.mark.anyio
async def test_tasks_run_in_parallel_and_return_only_conclusions(monkeypatch):
    def script(prompt):
        return [
            (0.2, _assistant(TextBlock("先看看"), ToolUseBlock(id="t1", name="mcp__excel__inspect_sheet",
                                                              input={"sheet": "工资"}))),
            (0.2, _assistant(TextBlock(f"结论：{prompt.splitlines()[0]} 在 工资!A1:F10"))),
            (0, _result(None)),
        ]
    monkeypatch.setattr(explorer, "run_query", _fake_query(script))
    published = []

    async def publish(text):
        published.append(text)

    started = time.monotonic()
    outcomes = await explorer.explore(explorer.normalize_tasks(["甲", "乙", "丙"]), [], publish)
    elapsed = time.monotonic() - started

    assert elapsed < 1.0, "三个子任务应当并行（串行要 1.2 秒）"
    assert [o.status for o in outcomes] == ["ok", "ok", "ok"]
    assert outcomes[1].answer == "结论：乙 在 工资!A1:F10"  # 只留最后一轮说的话
    assert outcomes[0].tool_calls == 1
    assert published[0] == "分头摸底：0/3 个子任务完成"
    assert any("#1 inspect_sheet 工资" in p for p in published)
    assert published[-1].startswith("分头摸底：3/3")


@pytest.mark.anyio
async def test_result_text_wins_and_long_answers_are_clipped(monkeypatch):
    monkeypatch.setattr(explorer, "run_query", _fake_query(
        lambda p: [(0, _assistant(TextBlock("中间过程"))), (0, _result("最终" + "字" * 3000))]))
    outcomes = await explorer.explore(explorer.normalize_tasks(["x"]), [], AsyncMock())
    assert outcomes[0].answer.startswith("最终")
    assert outcomes[0].answer.endswith("…（已截断）")
    assert len(outcomes[0].answer) <= explorer.ANSWER_LIMIT + 10


@pytest.mark.anyio
async def test_a_failing_or_slow_task_does_not_sink_the_others(monkeypatch):
    monkeypatch.setattr(explorer, "TASK_TIMEOUT_SECONDS", 0.3)

    def script(prompt):
        if prompt.startswith("慢"):
            return [(5, _result("来不及"))]
        if prompt.startswith("坏"):
            raise RuntimeError("子会话起不来")
        return [(0, _assistant(TextBlock("好的结论"))), (0, _result(None))]

    async def run(prompt, tools):
        for delay, message in script(prompt):
            await anyio.sleep(delay)
            yield message

    monkeypatch.setattr(explorer, "run_query", run)
    outcomes = await explorer.explore(explorer.normalize_tasks(["好", "慢", "坏"]), [], AsyncMock())
    assert [o.status for o in outcomes] == ["ok", "timeout", "error"]
    assert "子会话起不来" in outcomes[2].error
    assert outcomes[0].answer == "好的结论"


@pytest.mark.anyio
async def test_stop_cancels_the_remaining_tasks(monkeypatch):
    monkeypatch.setattr(explorer, "run_query", _fake_query(lambda p: [(10, _result("不该出现"))]))
    flag = {"stop": False}

    async def stop_soon():
        await anyio.sleep(0.2)
        flag["stop"] = True

    started = time.monotonic()
    async with anyio.create_task_group() as tg:
        tg.start_soon(stop_soon)
        outcomes = await explorer.explore(explorer.normalize_tasks(["a", "b"]), [], AsyncMock(),
                                          cancelled=lambda: flag["stop"])
    assert time.monotonic() - started < 2
    assert [o.status for o in outcomes] == ["cancelled", "cancelled"]


@pytest.mark.anyio
async def test_the_tool_returns_conclusions_and_clears_the_status(monkeypatch):
    from excel_tools import explore_workbook
    monkeypatch.setattr(explorer, "run_query", _fake_query(
        lambda p: [(0, _assistant(TextBlock("应收在 Data!B3"))), (0, _result(None))]))
    sent = []

    async def capture(msg):
        sent.append(msg)

    fn = getattr(explore_workbook, "handler", explore_workbook)
    with patch("ipc.write_message", new=capture):
        out = await fn({"tasks": [{"question": "应收账款在哪", "sheets": ["Data"]}]})
    parsed = json.loads(out["content"][0]["text"])
    assert parsed["success"] is True
    assert parsed["data"]["tasks"][0]["answer"] == "应收在 Data!B3"
    assert parsed["data"]["tasks"][0]["sheets"] == ["Data"]
    statuses = [m["event"]["text"] for m in sent if m.get("event", {}).get("kind") == "status"]
    assert statuses[0].startswith("分头摸底：0/1") and statuses[-1] == ""
    assert ui_events.PROGRESS.current() is None


@pytest.mark.anyio
async def test_the_tool_refuses_empty_tasks():
    from excel_tools import explore_workbook
    fn = getattr(explore_workbook, "handler", explore_workbook)
    out = await fn({"tasks": []})
    assert out["is_error"] is True


def test_fresh_progress_replaces_the_stuck_warning():
    stuck = ui_events.watchdog_status(0, ("explore_workbook", 130.0))
    assert "对话框" in stuck
    assert ui_events.watchdog_status(0, ("explore_workbook", 130.0), "分头摸底：1/3") == "分头摸底：1/3"
    board = ui_events.ProgressBoard()
    board.set("x")
    assert board.current() == "x"
    board.updated -= board.FRESH_SECONDS + 1
    assert board.current() is None  # 90 秒没进度：回到看门狗自己的判断


def test_subagents_get_only_read_only_tools_and_the_same_endpoint():
    from excel_tools import register_all_tools
    tools = [t for t in register_all_tools("excel") if t.name in explorer.EXPLORER_TOOLS]
    assert sorted(t.name for t in tools) == sorted(explorer.EXPLORER_TOOLS)
    options = explorer._options(tools)
    assert options.tools == []
    assert options.env == {"ANTHROPIC_BASE_URL": "https://example.invalid"}
    assert options.model == "fake-model"
    assert all(name.startswith("mcp__excel__") for name in options.allowed_tools)
    assert not any("write" in name or "execute" in name for name in options.allowed_tools)
