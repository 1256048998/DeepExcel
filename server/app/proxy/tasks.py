"""What counts as one task.

The quota is sold in tasks, so the proxy has to answer "is this call the start of
a new task or the continuation of one". A single user request drives an agent
loop of a dozen or more model calls -- every tool result goes back to the model
as another request -- and the first version of the proxy counted each of those
as a task. A free tier of 50 "tasks" ran out after three to five real ones.

Two signals, in order of preference:

1. ``x-trace-id``: the client generates one random id per user task and sends it
   on every call of that task. A task is counted on the first successful call
   carrying an id not seen before (within ``TASK_WINDOW``).
2. Without it (older clients), the shape of the request: a call whose last
   message is a user turn with no ``tool_result`` blocks starts a new turn; a
   call that carries tool results back is a continuation.

Neither signal is trusted blindly. A client that reuses one trace id forever
would turn everything into a single task, so an id expires after
``TASK_WINDOW`` and carries at most ``max_calls_per_task`` calls.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import Any

# One task should finish well inside this. Past it, the same id starts a new
# task rather than extending an old one indefinitely.
TASK_WINDOW = dt.timedelta(hours=3)

# Random ids only: letters, digits, dash, underscore. Anything else is dropped
# rather than stored, so the column can never become a channel for user content.
_TRACE_ID = re.compile(r"^[A-Za-z0-9_-]{8,64}$")

# Context-window suffixes such as "deepseek-v4-pro[1m]". The CLI reads them to
# size its auto-compaction and strips them before calling the API; stripping
# again here means an upstream never sees one even if a client forgets.
_WINDOW_SUFFIX = re.compile(r"\[\d+[mk]\]$", re.IGNORECASE)


def clean_trace_id(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    value = value.strip()
    return value if _TRACE_ID.match(value) else None


def strip_window_suffix(model: str) -> str:
    return _WINDOW_SUFFIX.sub("", model or "")


# Auto-compaction. Observed on the wire from Claude Code 2.1.190 (the CLI bundled
# with claude-agent-sdk 0.2.109) against a fake upstream -- see
# scripts/probe_compaction_shapes.py:
#
#   1. The summary request is the conversation so far with this instruction
#      appended as an extra text block to the last user message. Same system
#      prompt as any other call, no tool_result -- so by shape alone it looked
#      like a new user turn and every compaction cost the user one task.
#   2. The call after it is one user message: reminder, the summary (starting
#      with the sentence below), then whatever the user just typed, if anything.
#      With new user text it is a new turn; without (compaction in the middle of
#      a task) it continues the task that was running.
_COMPACT_INSTRUCTION = "CRITICAL: Respond with TEXT ONLY. Do NOT call any tools."
_CONTINUED_SUMMARY = "This session is being continued from a previous conversation"
_REMINDER = "<system-reminder>"

NEW_TURN = "new"
CONTINUATION = "continuation"
COMPACTION = "compaction"


def _text_blocks(content) -> list[str]:
    if isinstance(content, str):
        return [content]
    if not isinstance(content, list):
        return []
    return [
        str(block.get("text") or "") for block in content
        if isinstance(block, dict) and block.get("type") == "text"
    ]


def classify_turn(payload: dict) -> str:
    """NEW_TURN, CONTINUATION (tool results going back), or COMPACTION.

    COMPACTION is a claim the client makes by the shape of its request, so the
    proxy only honours it for a user who already has a task running; see
    router._record_usage. Anything unrecognised is a new turn.
    """
    messages = payload.get("messages") if isinstance(payload, dict) else None
    if not isinstance(messages, list) or not messages:
        return NEW_TURN
    last = messages[-1]
    if not isinstance(last, dict) or last.get("role") != "user":
        # An assistant-prefill or malformed tail: not a continuation of tool use.
        return NEW_TURN
    content = last.get("content")
    if isinstance(content, list):
        for block in content:
            if isinstance(block, dict) and block.get("type") == "tool_result":
                return CONTINUATION
    texts = [t.strip() for t in _text_blocks(content)]
    if any(t.startswith(_COMPACT_INSTRUCTION) for t in texts):
        return COMPACTION
    if any(t.startswith(_CONTINUED_SUMMARY) for t in texts):
        typed = [t for t in texts if t and not t.startswith((_CONTINUED_SUMMARY, _REMINDER))]
        return NEW_TURN if typed else COMPACTION
    return NEW_TURN


def starts_new_user_turn(payload: dict) -> bool:
    """True unless the request carries tool results back or is compaction overhead."""
    return classify_turn(payload) == NEW_TURN
