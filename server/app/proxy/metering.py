"""Token accounting.

The proxy is the only place a client cannot bypass, which is what makes it the
authoritative counter. The client-reported counter that exists today is fine
while nothing costs money; once it does, this is what the bill is built from.
"""

from __future__ import annotations

import json
from dataclasses import dataclass


@dataclass
class Usage:
    input_tokens: int = 0
    output_tokens: int = 0
    cache_read_tokens: int = 0
    cache_write_tokens: int = 0

    @property
    def total(self) -> int:
        return self.input_tokens + self.output_tokens

    def merge(self, other: "Usage") -> None:
        # Anthropic reports input usage on message_start and output usage on
        # message_delta. Taking the max rather than summing avoids double
        # counting when a field appears in both.
        self.input_tokens = max(self.input_tokens, other.input_tokens)
        self.cache_read_tokens = max(self.cache_read_tokens, other.cache_read_tokens)
        self.cache_write_tokens = max(self.cache_write_tokens, other.cache_write_tokens)
        self.output_tokens = max(self.output_tokens, other.output_tokens)


def usage_from_payload(payload: dict) -> Usage:
    raw = payload.get("usage") or {}
    if not isinstance(raw, dict):
        return Usage()
    return Usage(
        input_tokens=int(raw.get("input_tokens") or 0),
        output_tokens=int(raw.get("output_tokens") or 0),
        cache_read_tokens=int(raw.get("cache_read_input_tokens") or 0),
        cache_write_tokens=int(raw.get("cache_creation_input_tokens") or 0),
    )


class StreamUsageCollector:
    """Extracts usage from an SSE stream without buffering it.

    The stream is forwarded to the client byte for byte as it arrives; this only
    inspects it in passing. Buffering to read the totals at the end would turn a
    streaming response into a blocking one and destroy the reason streaming
    exists.
    """

    def __init__(self) -> None:
        self.usage = Usage()
        self.stop_reason: str | None = None
        self._buffer = b""

    def feed(self, chunk: bytes) -> None:
        self._buffer += chunk
        # SSE events are separated by a blank line; process whole events only.
        while b"\n\n" in self._buffer:
            event, self._buffer = self._buffer.split(b"\n\n", 1)
            self._consume_event(event)
        # A malformed or hostile upstream must not grow this without bound.
        if len(self._buffer) > 1_000_000:
            self._buffer = self._buffer[-100_000:]

    def _consume_event(self, event: bytes) -> None:
        for line in event.split(b"\n"):
            if not line.startswith(b"data:"):
                continue
            raw = line[5:].strip()
            if not raw or raw == b"[DONE]":
                continue
            try:
                payload = json.loads(raw.decode("utf-8"))
            except (ValueError, UnicodeDecodeError):
                continue
            if not isinstance(payload, dict):
                continue

            # message_start carries the request's input usage.
            message = payload.get("message")
            if isinstance(message, dict):
                self.usage.merge(usage_from_payload(message))
            # message_delta carries the running output usage.
            self.usage.merge(usage_from_payload(payload))

            delta = payload.get("delta")
            if isinstance(delta, dict) and delta.get("stop_reason"):
                self.stop_reason = delta["stop_reason"]


# Per-million-token prices, used to record cost at the time of use.
#
# Recorded per request rather than computed later from a current price list:
# prices change, and a bill must reflect what the call cost when it was made.
DEFAULT_PRICES: dict[str, tuple[float, float]] = {
    # model prefix -> (input per 1M, output per 1M) in USD
    "claude-opus": (15.0, 75.0),
    "claude-sonnet": (3.0, 15.0),
    "claude-haiku": (0.80, 4.0),
    "deepseek": (0.27, 1.10),
    "kimi": (0.60, 2.50),
    "qwen": (0.40, 1.20),
    "glm": (0.50, 1.50),
}


def estimate_cost_usd(model: str, usage: Usage) -> float:
    model = (model or "").lower()
    for prefix, (input_price, output_price) in DEFAULT_PRICES.items():
        if model.startswith(prefix):
            return round(
                usage.input_tokens / 1_000_000 * input_price
                + usage.output_tokens / 1_000_000 * output_price,
                6,
            )
    # An unpriced model records zero rather than guessing. A wrong number on an
    # invoice is worse than a missing one, and the gap is visible in the admin
    # dashboard.
    return 0.0
