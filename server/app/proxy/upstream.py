"""Upstream provider selection and failover.

The proxy holds the real provider keys; clients never see them. That is the
whole reason hosted routing exists, and it is what makes metering, quotas and
revocation possible at all.
"""

from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass(frozen=True)
class Upstream:
    name: str
    base_url: str
    api_key: str
    # Models this upstream accepts. Empty means "anything", used for
    # Anthropic-compatible gateways that proxy several vendors.
    models: tuple[str, ...] = ()
    priority: int = 100

    def accepts(self, model: str) -> bool:
        if not self.models:
            return True
        return any(model.startswith(prefix) for prefix in self.models)


def _parse_models(raw: str | None) -> tuple[str, ...]:
    if not raw:
        return ()
    return tuple(part.strip() for part in raw.split(",") if part.strip())


def load_upstreams() -> list[Upstream]:
    """Reads upstream definitions from the environment.

    Configured as numbered slots so a provider can be added or reordered
    without a deploy:

        UPSTREAM_1_NAME=anthropic
        UPSTREAM_1_BASE_URL=https://api.anthropic.com
        UPSTREAM_1_API_KEY=sk-ant-...
        UPSTREAM_1_MODELS=claude-
        UPSTREAM_1_PRIORITY=10
    """
    upstreams: list[Upstream] = []
    for index in range(1, 11):
        prefix = f"UPSTREAM_{index}_"
        base_url = os.environ.get(prefix + "BASE_URL")
        api_key = os.environ.get(prefix + "API_KEY")
        if not base_url or not api_key:
            continue
        upstreams.append(
            Upstream(
                name=os.environ.get(prefix + "NAME", f"upstream{index}"),
                base_url=base_url.rstrip("/"),
                api_key=api_key,
                models=_parse_models(os.environ.get(prefix + "MODELS")),
                priority=int(os.environ.get(prefix + "PRIORITY", "100")),
            )
        )
    upstreams.sort(key=lambda u: u.priority)
    return upstreams


def select(upstreams: list[Upstream], model: str) -> list[Upstream]:
    """Candidates for a model, best first.

    Returns a list rather than one choice so the caller can fail over: an
    upstream returning 5xx must not become the user's problem when another one
    can serve the same model.
    """
    return [upstream for upstream in upstreams if upstream.accepts(model)]
