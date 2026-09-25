"""explore_workbook：大工作簿按子任务派出只读子 agent 并行摸底，只把结论交回主 agent。

这是 Claude Code 里 Task / Explore 子 agent 的工作簿版：每个子任务是一个独立的会话，
有自己的上下文，只能用 list / find / inspect_sheet / read_range / read_workbook 读；
几十次读取的原始输出留在子会话里，主会话只收到几百字的结论，上下文保持干净。

没有用 CLI 自带的子 agent 工具：它的每一步都会混进主会话的事件流和权限钩子，面板
无从区分；这里由侧车自己起子会话，超时、并发、进度显示都在掌控之内。

子会话复用主会话的 env（同一个出口、同一套凭据与请求头），不自己判断路由。
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field

import anyio

MAX_TASKS = 4
SUB_MAX_TURNS = 12
TASK_TIMEOUT_SECONDS = 150.0
ANSWER_LIMIT = 1500
EXPLORER_TOOLS = ("list", "find", "inspect_sheet", "read_range", "read_workbook")

EXPLORER_PROMPT = """你是 DeepExcel 的摸底子 agent，替主 agent 查清工作簿里的一部分内容。

<rules>
- 只能读，不能写。可用工具：list / find / inspect_sheet / read_range / read_workbook
- 先用 list 看有哪些表，用 inspect_sheet 看结构；找东西用 find；只有必须看具体值时才 read_range，而且只读需要的那几行
- 不要寒暄，不要复述工具的原始输出，不要给修改建议之外的长篇解释
</rules>

<answer>
查完用不超过 600 字交回结论，按这个顺序：
1. 直接回答问题
2. 关键位置：表名!区域、列名（多级表头写成路径）、公式模式
3. 疑点或拿不准的地方；inspect_sheet 标了 partial、表头 uncertain、或只看了部分行的，要写明是样本结论或推测
</answer>
"""

# 由 sidecar 在创建主会话时写入：子会话要用同一个模型、同一个出口
_config: dict = {}


def configure(env: dict, model: str | None, host: str) -> None:
    _config.clear()
    _config.update({"env": dict(env or {}), "model": model, "host": host})


def configured() -> bool:
    return bool(_config)


@dataclass
class TaskOutcome:
    index: int
    question: str
    sheets: list[str]
    status: str = "ok"  # ok / timeout / error / cancelled
    answer: str = ""
    tool_calls: int = 0
    error: str | None = None
    seconds: float = 0.0
    tools_used: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        out = {
            "task": self.index + 1,
            "question": self.question,
            "status": self.status,
            "answer": self.answer,
            "tool_calls": self.tool_calls,
            "seconds": round(self.seconds, 1),
        }
        if self.sheets:
            out["sheets"] = self.sheets
        if self.error:
            out["error"] = self.error
        return out


def normalize_tasks(raw) -> list[dict]:
    """[{question, sheets?}]，最多 MAX_TASKS 个；也接受字符串列表"""
    tasks = []
    for entry in raw if isinstance(raw, list) else []:
        if isinstance(entry, str):
            entry = {"question": entry}
        if not isinstance(entry, dict):
            continue
        question = str(entry.get("question") or "").strip()
        if not question:
            continue
        sheets = entry.get("sheets") or []
        if isinstance(sheets, str):
            sheets = [s.strip() for s in sheets.replace("，", ",").split(",") if s.strip()]
        tasks.append({"question": question[:500], "sheets": [str(s) for s in sheets][:20]})
    return tasks[:MAX_TASKS]


def task_prompt(task: dict) -> str:
    scope = ("只看这些表：" + "、".join(task["sheets"])) if task["sheets"] else "范围：整个工作簿"
    return f"{task['question']}\n\n{scope}"


def clip_answer(text: str) -> str:
    text = (text or "").strip()
    return text if len(text) <= ANSWER_LIMIT else text[:ANSWER_LIMIT] + "…（已截断）"


def _options(tools):
    from claude_agent_sdk import ClaudeAgentOptions, create_sdk_mcp_server

    server = create_sdk_mcp_server(name="excel", tools=list(tools))
    return ClaudeAgentOptions(
        model=_config.get("model"),
        tools=[],  # 与主会话相同：关掉 CLI 内置工具
        mcp_servers={"excel": server},
        allowed_tools=[f"mcp__excel__{t.name}" for t in tools],
        system_prompt=EXPLORER_PROMPT,
        max_turns=SUB_MAX_TURNS,
        env=dict(_config.get("env") or {}),
        setting_sources=[],
        thinking={"type": "disabled"},
    )


async def _default_query(prompt: str, tools):
    """一个子会话：产出 SDK 消息。测试里替换成假的。"""
    from claude_agent_sdk import query

    async for message in query(prompt=prompt, options=_options(tools)):
        yield message


# 测试替换点
run_query = _default_query


class Progress:
    """汇报给面板状态行：已完成几个、各自在做什么"""

    def __init__(self, total: int, publish):
        self.total = total
        self.done = 0
        self.current: dict[int, str] = {}
        self._publish = publish

    def text(self) -> str:
        head = f"分头摸底：{self.done}/{self.total} 个子任务完成"
        doing = [f"#{i + 1} {what}" for i, what in sorted(self.current.items())]
        return head + ("（" + "；".join(doing[:3]) + "）" if doing else "")

    async def update(self, index: int, doing: str | None):
        if doing is None:
            self.current.pop(index, None)
        else:
            self.current[index] = doing
        await self._publish(self.text())

    async def finish(self, index: int):
        self.done += 1
        self.current.pop(index, None)
        await self._publish(self.text())


def _describe_tool(name: str, args: dict) -> str:
    short = name.replace("mcp__excel__", "")
    target = args.get("sheet") or args.get("address") or args.get("query") or args.get("kind") or ""
    return f"{short} {target}".strip()[:40]


async def run_task(index: int, task: dict, tools, progress: Progress) -> TaskOutcome:
    from claude_agent_sdk import AssistantMessage, ResultMessage, TextBlock, ToolUseBlock

    outcome = TaskOutcome(index=index, question=task["question"], sheets=task["sheets"])
    started = time.monotonic()
    texts: list[str] = []
    result_text = None
    try:
        with anyio.fail_after(TASK_TIMEOUT_SECONDS):
            async for message in run_query(task_prompt(task), tools):
                if isinstance(message, AssistantMessage):
                    turn_text = []
                    for block in message.content:
                        if isinstance(block, ToolUseBlock):
                            outcome.tool_calls += 1
                            outcome.tools_used.append(block.name.replace("mcp__excel__", ""))
                            await progress.update(index, _describe_tool(block.name, block.input or {}))
                        elif isinstance(block, TextBlock) and block.text.strip():
                            turn_text.append(block.text)
                    if turn_text:
                        texts = turn_text  # 只留最后一轮说的话：那是结论
                elif isinstance(message, ResultMessage):
                    if getattr(message, "is_error", False):
                        outcome.status = "error"
                        outcome.error = str(getattr(message, "result", "") or "子会话出错")[:300]
                    result_text = getattr(message, "result", None)
    except TimeoutError:
        outcome.status = "timeout"
        outcome.error = f"{int(TASK_TIMEOUT_SECONDS)} 秒内没有查完"
    except anyio.get_cancelled_exc_class():
        outcome.status = "cancelled"
        raise
    except Exception as exc:  # noqa: BLE001 — 一个子任务坏了不拖垮其他子任务
        outcome.status = "error"
        outcome.error = f"{type(exc).__name__}: {exc}"[:300]
    finally:
        outcome.seconds = time.monotonic() - started
    answer = result_text if isinstance(result_text, str) and result_text.strip() else "\n".join(texts)
    outcome.answer = clip_answer(answer)
    if outcome.status == "ok" and not outcome.answer:
        outcome.status = "error"
        outcome.error = "子任务没有给出结论"
    await progress.finish(index)
    return outcome


async def explore(tasks: list[dict], tools, publish, cancelled=lambda: False) -> list[TaskOutcome]:
    """并行跑所有子任务。cancelled() 为真（用户按了停止）时取消剩下的。"""
    progress = Progress(len(tasks), publish)
    results: list[TaskOutcome | None] = [None] * len(tasks)
    await publish(progress.text())

    async def one(i: int, task: dict):
        results[i] = await run_task(i, task, tools, progress)

    async with anyio.create_task_group() as tg:
        async def watch_cancel():
            while True:
                await anyio.sleep(0.5)
                if cancelled():
                    tg.cancel_scope.cancel()
                    return
                if all(r is not None for r in results):
                    return

        tg.start_soon(watch_cancel)
        for i, task in enumerate(tasks):
            tg.start_soon(one, i, task)

    return [r if r is not None else TaskOutcome(index=i, question=tasks[i]["question"], sheets=tasks[i]["sheets"],
                                                status="cancelled", error="用户已中断")
            for i, r in enumerate(results)]
