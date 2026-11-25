"""Probe: does `resume=<session id>` really put the earlier conversation back in
front of the model, with the options the sidecar uses?

A local fake upstream answers every request and records it. Session 1 says
something memorable; the client disconnects; session 2 connects with
`resume=<session 1's id>` and asks again. The second session's request must
contain session 1's user message. The sidecar's own transcript check
(engine_session.transcript_exists) must find the file in between.

No real model, no API key: CLAUDE_CONFIG_DIR points at a temporary folder and
setting_sources=[] keeps ~/.claude/settings.json out (it would otherwise
redirect the CLI to a real endpoint). Run after upgrading claude-agent-sdk.

Usage: python scripts/probe_resume.py
"""
import asyncio
import dataclasses
import json
import os
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src" / "DeepExcel.Sidecar"))
import engine_session  # noqa: E402

REQUESTS = []
MARKER = "口令是蓝莓七号"


def sse(event, data):
    return f"event: {event}\ndata: {json.dumps(data)}\n\n".encode()


class FakeUpstream(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_GET(self):
        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        body = self.rfile.read(int(self.headers.get("content-length") or 0))
        if not self.path.startswith("/v1/messages") or "count_tokens" in self.path:
            self.send_response(404)
            self.end_headers()
            return
        REQUESTS.append(json.loads(body or b"{}"))
        self.send_response(200)
        self.send_header("content-type", "text/event-stream")
        self.end_headers()
        w = self.wfile.write
        w(sse("message_start", {"type": "message_start", "message": {
            "id": f"msg_{len(REQUESTS)}", "type": "message", "role": "assistant", "model": "claude-sonnet-4-5",
            "content": [], "stop_reason": None, "stop_sequence": None,
            "usage": {"input_tokens": 100, "output_tokens": 1}}}))
        w(sse("content_block_start", {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}}))
        w(sse("content_block_delta", {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": "记住了。"}}))
        w(sse("content_block_stop", {"type": "content_block_stop", "index": 0}))
        w(sse("message_delta", {"type": "message_delta", "delta": {"stop_reason": "end_turn", "stop_sequence": None},
                                "usage": {"output_tokens": 3}}))
        w(sse("message_stop", {"type": "message_stop"}))
        self.wfile.flush()


def _texts(payload):
    out = []
    for msg in payload.get("messages") or []:
        content = msg.get("content")
        if isinstance(content, str):
            out.append(content)
        elif isinstance(content, list):
            out.extend(str(b.get("text") or "") for b in content if isinstance(b, dict))
    return out


async def run(config_dir: str):
    server = ThreadingHTTPServer(("127.0.0.1", 0), FakeUpstream)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    from claude_agent_sdk import ClaudeAgentOptions, ClaudeSDKClient
    from claude_agent_sdk.types import ResultMessage

    env = {
        "ANTHROPIC_BASE_URL": f"http://127.0.0.1:{server.server_address[1]}",
        "ANTHROPIC_API_KEY": "probe-not-a-real-key",
        "CLAUDE_CONFIG_DIR": config_dir,
        "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC": "1",
        "DISABLE_TELEMETRY": "1",
    }
    for key in ("ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_MODEL"):
        os.environ.pop(key, None)
    os.environ["CLAUDE_CONFIG_DIR"] = config_dir
    options = ClaudeAgentOptions(model="claude-sonnet-4-5", tools=[], max_turns=2, env=env,
                                 system_prompt="你是测试助手。", setting_sources=[])
    session_id = None
    try:
        async with ClaudeSDKClient(options=options) as client:
            await client.query(MARKER + "，先记住。")
            async for message in client.receive_response():
                if isinstance(message, ResultMessage):
                    session_id = message.session_id
        found = engine_session.transcript_exists(session_id, os.environ) if session_id else False
        first_count = len(REQUESTS)
        async with ClaudeSDKClient(options=dataclasses.replace(options, resume=session_id)) as client:
            await client.query("刚才的口令是什么？")
            async for _ in client.receive_response():
                pass
    finally:
        server.shutdown()
    return session_id, found, REQUESTS[first_count:]


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="deepexcel-resume-probe-") as config_dir:
        session_id, found, second = asyncio.run(run(config_dir))
    carried = any(MARKER in text for payload in second for text in _texts(payload))
    print(f"session_id={session_id} transcript_found={found} second_session_requests={len(second)} "
          f"earlier_message_carried={carried}")
    ok = bool(session_id) and found and second and carried
    print("OK" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
