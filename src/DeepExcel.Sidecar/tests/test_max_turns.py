# src/DeepExcel.Sidecar/tests/test_max_turns.py
# 设置面板的 MaxTurns 以前存下来却从没下发，侧车写死 20。
# resolve_max_turns 是它现在唯一的入口：错值回落，不能让会话起不来。
import pytest

import sidecar as sidecar_module
from sidecar import resolve_max_turns


@pytest.mark.parametrize("cfg, expected", [
    ({"max_turns": 35}, 35),
    ({"max_turns": "12"}, 12),       # WPS 路径直接读 config.json，类型不保证
    ({"max_turns": 1}, 1),
    ({"max_turns": 50}, 50),
    ({"max_turns": 200}, 50),        # 旧前端允许填到 200；上限与 C# 端一致
    ({"max_turns": 0}, 20),
    ({"max_turns": -3}, 20),
    ({"max_turns": None}, 20),
    ({"max_turns": "abc"}, 20),
    ({}, 20),                        # 旧版 C# 不发这个字段
    (None, 20),                      # 30 秒没收到 config，走环境变量兜底
])
def test_resolve_max_turns(cfg, expected):
    assert resolve_max_turns(cfg) == expected


def test_ceiling_matches_csharp_limit():
    # C# HandleSaveModelConfig 只接受 maxTurns <= 50；两边不一致时
    # 用户填的值会在其中一边被静默改掉
    assert sidecar_module.MAX_TURNS_CEILING == 50
