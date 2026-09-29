"""The hosted model catalog: which models a hosted user may pick.

Hosted users have no provider configured locally, so without a catalog the
panel had nothing to offer and every hosted user ran whatever model name sat
in their local config. The server now says which models exist, what each costs
in points, and which one is the default; the client shows exactly that list
and never decides on its own (see EndpointConfig).

The proxy enforces the same list. A catalog the proxy did not enforce would be
a suggestion: any model an upstream happened to accept would still go through,
at whatever it costs.

Configured with HOSTED_MODELS (config.py). Not configured means no catalog: the
client falls back to its local model name and the proxy accepts any model an
upstream serves, which is how it worked before the catalog existed.

Every request the CLI makes, compaction included, uses the main model (checked
by scripts/probe_compaction_shapes.py), so enforcing the list does not break
background calls.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..config import get_settings
from .metering import model_weight
from .tasks import strip_window_suffix
from .upstream import load_upstreams, select


@dataclass(frozen=True)
class CatalogEntry:
    model: str
    label: str
    points: int
    default: bool


def is_configured() -> bool:
    return bool(get_settings().hosted_models)


def catalog_for(plan: str) -> list[CatalogEntry]:
    """The models this plan may use, in configured order.

    An entry no upstream serves is left out: offering it would only produce a
    502 when the user picked it.
    """
    settings = get_settings()
    upstreams = load_upstreams()
    entries: list[CatalogEntry] = []
    for item in settings.hosted_models:
        if item["plans"] is not None and plan not in item["plans"]:
            continue
        if not select(upstreams, item["model"]):
            continue
        entries.append(CatalogEntry(
            model=item["model"],
            label=item["label"],
            points=model_weight(item["model"], settings.model_weights),
            default=item["default"],
        ))
    if entries and not any(e.default for e in entries):
        # No default for this plan (or the default was filtered out): the
        # cheapest model is the one nobody is surprised by.
        cheapest = min(entries, key=lambda e: e.points)
        entries = [CatalogEntry(e.model, e.label, e.points, e is cheapest) for e in entries]
    return entries


def is_offered(model: str, plan: str) -> bool:
    if not is_configured():
        return True
    name = strip_window_suffix(model).lower()
    return any(e.model.lower() == name for e in catalog_for(plan))
