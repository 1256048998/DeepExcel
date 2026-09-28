"""探针：CLI 自动压缩时发给模型服务的请求长什么样，服务端计费规则能不能认出来。

托管代理按「用户任务」扣额度（server/app/proxy/tasks.py）。CLI 的自动压缩会额外发请求，
如果被认成新任务，每压缩一次就多扣用户一次。这里起一个本机假上游（固定回复，
从第二轮起上报接近满的 input_tokens 让 CLI 以为上下文快满了），让真实的 CLI 连过去跑三句话，
第三句前会触发压缩；记录每个请求的结构并用服务端的 classify_turn 判一遍。

不接真模型、不需要 API Key；setting_sources=[] 和侧车一样不读 ~/.claude/settings.json。
升级 claude-agent-sdk 后跑一次：CLI 改了压缩请求的格式，这里会失败，计费规则要跟着改。

用法：python scripts/probe_compaction_shapes.py
"""
import asyncio
import json
import os
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "server"))
from app.proxy.tasks import classify_turn  # noqa: E402

REQUESTS = []


def shape(payload):
    msgs = payload.get("messages") or []
    last = msgs[-1] if msgs else {}
    content = last.get("content")
    blocks = content if isinstance(content, list) else [{"type": "text", "text": str(content or "")}]
    return {
        "n_messages": len(msgs),
        "last_blocks": [(b.get("type"), (b.get("text") or "")[:70].replace("\n", " "))
                        for b in blocks if isinstance(b, dict)],
        "turn": classify_turn(payload),
    }


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
        payload = json.loads(body or b"{}")
        REQUESTS.append(shape(payload))
        self.send_response(200)
        self.send_header("content-type", "text/event-stream")
        self.end_headers()
        w = self.wfile.write
        w(sse("message_start", {"type": "message_start", "message": {
            "id": f"msg_{len(REQUESTS)}", "type": "message", "role": "assistant", "model": payload.get("model"),
            "content": [], "stop_reason": None, "stop_sequence": None,
            # 第一轮报小值（只有一轮对话时 CLI 压不了：too_few_groups），之后报接近满
            "usage": {"input_tokens": 1_200 if len(REQUESTS) == 1 else 190_000, "output_tokens": 1}}}))
        w(sse("content_block_start", {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}}))
        w(sse("content_block_delta", {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": "好的。"}}))
        w(sse("content_block_stop", {"type": "content_block_stop", "index": 0}))
        w(sse("message_delta", {"type": "message_delta", "delta": {"stop_reason": "end_turn", "stop_sequence": None},
                                "usage": {"output_tokens": 3}}))
        w(sse("message_stop", {"type": "message_stop"}))
        self.wfile.flush()


async def run() -> list:
    server = ThreadingHTTPServer(("127.0.0.1", 0), FakeUpstream)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    from claude_agent_sdk import ClaudeAgentOptions, ClaudeSDKClient

    env = {
        "ANTHROPIC_BASE_URL": f"http://127.0.0.1:{server.server_address[1]}",
        "ANTHROPIC_API_KEY": "probe-not-a-real-key",
        "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC": "1",
        "DISABLE_TELEMETRY": "1",
        # 窗口调小，配合假上游上报的 190K 输入，保证下一句前一定压缩
        "CLAUDE_CODE_AUTO_COMPACT_WINDOW": "60000",
    }
    for key in ("ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_MODEL"):
        os.environ.pop(key, None)
    options = ClaudeAgentOptions(model="claude-sonnet-4-5", tools=[], max_turns=2, env=env,
                                 system_prompt="你是测试助手。", setting_sources=[])
    try:
        async with ClaudeSDKClient(options=options) as client:
            for prompt in ["第一句：你好", "第二句：继续", "第三句：再来"]:
                await client.query(prompt)
                async for _ in client.receive_response():
                    pass
    finally:
        server.shutdown()
    return REQUESTS


def main() -> int:
    requests = asyncio.run(run())
    for row in requests:
        print(json.dumps(row, ensure_ascii=False))
    turns = [r["turn"] for r in requests]
    # 三句话 = 三个任务；压缩请求必须被认成 compaction，而不是第四个任务
    ok = turns.count("new") == 3 and "compaction" in turns
    print(f"\nturns={turns}\n{'OK' if ok else 'FAIL'}：{turns.count('new')} 个新任务，"
          f"{turns.count('compaction')} 个压缩请求")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
