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


def starts_new_user_turn(payload: dict) -> bool:
    """True unless the request is carrying tool results back to the model."""
    messages = payload.get("messages") if isinstance(payload, dict) else None
    if not isinstance(messages, list) or not messages:
        return True
    last = messages[-1]
    if not isinstance(last, dict) or last.get("role") != "user":
        # An assistant-prefill or malformed tail: not a continuation of tool use.
        return True
    content = last.get("content")
    if isinstance(content, list):
        for block in content:
            if isinstance(block, dict) and block.get("type") == "tool_result":
                return False
    return True
