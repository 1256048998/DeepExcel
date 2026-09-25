# src/DeepExcel.Sidecar/model_windows.py
"""告诉 CLI 模型真实的上下文窗口，免得它过早或过晚压缩对话。

CLI 对不认识的模型一律按 200K 算，在 167K 时自动压缩（2026-09-25 用 get_context_usage
实测）。DeepSeek v4 的窗口是 1,048,576（同日用超长请求的报错实测，deepseek-flash /
deepseek-v4-flash / deepseek-v4-pro 相同），于是长对话在不到六分之一处就被压缩，丢掉
前面读过的表结构和用户交代过的要求。反过来，窗口小于 200K 的模型会在 CLI 压缩之前就
被服务商以「prompt is too long」拒绝。

CLI 的两个开关（同日实测）：
- 模型名加 `[1m]` 后缀：按 1M 窗口算（阈值 967K）；CLI 发请求前会去掉后缀，DeepSeek
  照常接受。托管模式下代理也会剥掉它（server/app/proxy/tasks.py）。
- 环境变量 CLAUDE_CODE_AUTO_COMPACT_WINDOW：只能调小，不能超过模型的上限
  （200K，或加了 `[1m]` 之后的 1M）。

所以：窗口 > 200K → 加 `[1m]`，窗口 < 1M 时再用环境变量调到真实大小；
窗口 < 200K → 只设环境变量；不知道 → 不动。
"""

from __future__ import annotations

import re

CLI_DEFAULT_WINDOW = 200_000
CLI_1M_WINDOW = 1_000_000
AUTO_COMPACT_ENV = "CLAUDE_CODE_AUTO_COMPACT_WINDOW"

# 只收录实测过的。猜错比不填更糟：填大了会在服务商那里报「prompt is too long」，
# 填小了又回到过早压缩。
KNOWN_WINDOWS = [
    (re.compile(r"^deepseek-(v4-)?(flash|pro)\b", re.IGNORECASE), 1_048_576),
]

_SUFFIX = re.compile(r"\[\d+[mk]\]$", re.IGNORECASE)


def known_window(model: str) -> int | None:
    base = _SUFFIX.sub("", model or "")
    # Claude 模型 CLI 自己认识（含 1M 版本），不要插手
    if base.lower().startswith("claude"):
        return None
    for pattern, window in KNOWN_WINDOWS:
        if pattern.search(base):
            return window
    return None


def apply_context_window(model: str, configured_window=None) -> tuple[str, dict]:
    """返回 (交给 CLI 的模型名, 额外环境变量)。

    configured_window 来自配置（托管模式下由服务端 EndpointConfig 给出），优先于内置表。
    """
    if not model or _SUFFIX.search(model):
        # 用户或服务端已经自己加了后缀：尊重它
        return model, {}
    try:
        window = int(configured_window) if configured_window else None
    except (TypeError, ValueError):
        window = None
    if not window or window <= 0:
        window = known_window(model)
    if not window:
        return model, {}
    if window > CLI_DEFAULT_WINDOW:
        extra = {AUTO_COMPACT_ENV: str(window)} if window < CLI_1M_WINDOW else {}
        return f"{model}[1m]", extra
    if window < CLI_DEFAULT_WINDOW:
        return model, {AUTO_COMPACT_ENV: str(window)}
    return model, {}
