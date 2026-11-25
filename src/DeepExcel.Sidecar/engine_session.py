"""Continue a past conversation with the engine's own session instead of a summary.

Before: continuing an old conversation restarted the sidecar and pasted up to
8000 characters of user/assistant text in front of the next message. Tool
calls, what was read and written, and anything past 8000 characters were gone.

Now: the sidecar reports the Claude Code CLI session id after every turn
(engine_session). The host stores it with the conversation and hands it back in
restore_history. If it can be resumed, the sidecar reconnects with
`resume=<id>` and the model sees the whole earlier conversation; the summary is
not sent.

Resume only when it is safe -- otherwise fall back to the summary, as before:
  * the fingerprint (model + endpoint) matches: thinking blocks from another
    model or provider can be rejected outright when replayed;
  * the CLI still has the transcript on disk;
  * the reconnect itself succeeds.
"""
from __future__ import annotations

import hashlib
import os
from pathlib import Path

_reported = {"session_id": None}


def fingerprint(model: str, base_url: str) -> str:
    raw = f"{(model or '').strip()}|{(base_url or '').strip().rstrip('/').lower()}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def should_report(session_id) -> bool:
    """每个新的会话 id 只报一次"""
    if not isinstance(session_id, str) or not session_id or session_id == _reported["session_id"]:
        return False
    _reported["session_id"] = session_id
    return True


def _config_dir(environ) -> Path:
    configured = environ.get("CLAUDE_CONFIG_DIR")
    if configured:
        return Path(configured)
    return Path(environ.get("USERPROFILE") or environ.get("HOME") or Path.home()) / ".claude"


def transcript_exists(session_id: str, environ=None) -> bool:
    """CLI 把会话记录存在 <config>/projects/<按目录编码>/<id>.jsonl；找得到才 resume。"""
    environ = os.environ if environ is None else environ
    if not session_id or any(ch in session_id for ch in "/\\:*?\"<>|"):
        return False
    projects = _config_dir(environ) / "projects"
    try:
        return any(projects.glob(f"*/{session_id}.jsonl"))
    except OSError:
        return False


def resume_target(request, current_fingerprint: str, environ=None):
    """restore_history 带来的 {session_id, fingerprint} → 可以 resume 的会话 id；不行返回 (None, 原因)。"""
    if not isinstance(request, dict):
        return None, "没有会话 id"
    session_id = request.get("session_id")
    if not isinstance(session_id, str) or not session_id:
        return None, "没有会话 id"
    if request.get("fingerprint") != current_fingerprint:
        return None, "模型或出口变了"
    if not transcript_exists(session_id, environ):
        return None, "会话记录不在了"
    return session_id, ""
