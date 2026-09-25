# src/DeepExcel.Sidecar/tests/test_model_windows.py
# 告诉 CLI 模型真实的上下文窗口（数值与 CLI 行为均为 2026-09-25 实测）。
import pytest

import sidecar
from model_windows import AUTO_COMPACT_ENV, apply_context_window, known_window


@pytest.mark.parametrize("model", ["deepseek-flash", "deepseek-v4-flash", "deepseek-v4-pro", "DeepSeek-V4-Pro"])
def test_deepseek_gets_the_1m_window(model):
    # 实测窗口 1,048,576；CLI 加 [1m] 后按 1M 算，阈值 967K。窗口不小于 1M，不需要再调小
    assert apply_context_window(model) == (f"{model}[1m]", {})


def test_windows_between_200k_and_1m_are_capped_by_env():
    assert apply_context_window("some-model", 400_000) == ("some-model[1m]", {AUTO_COMPACT_ENV: "400000"})


def test_small_windows_only_lower_the_threshold():
    # 128K 的模型：CLI 按 200K 算会先被服务商以 prompt is too long 拒绝
    assert apply_context_window("kimi-k2.6", 128_000) == ("kimi-k2.6", {AUTO_COMPACT_ENV: "128000"})


def test_unknown_models_are_left_alone():
    assert apply_context_window("qwen3.7-max") == ("qwen3.7-max", {})
    assert apply_context_window("qwen3.7-max", "not a number") == ("qwen3.7-max", {})


def test_claude_models_are_left_to_the_cli():
    assert known_window("claude-sonnet-5") is None
    assert apply_context_window("claude-sonnet-5") == ("claude-sonnet-5", {})


def test_an_existing_suffix_is_respected():
    assert apply_context_window("deepseek-v4-pro[1M]") == ("deepseek-v4-pro[1M]", {})


def test_configured_window_overrides_the_table():
    # 托管模式下由服务端给出窗口
    assert apply_context_window("deepseek-v4-flash", 128_000) == ("deepseek-v4-flash", {AUTO_COMPACT_ENV: "128000"})


def test_inherited_compaction_env_is_removed_unless_we_set_it():
    assert "CLAUDE_CODE_AUTO_COMPACT_WINDOW" in sidecar.stale_env_keys({"ANTHROPIC_API_KEY": "k"})
    assert "CLAUDE_CODE_AUTO_COMPACT_WINDOW" not in sidecar.stale_env_keys(
        {"ANTHROPIC_API_KEY": "k", "CLAUDE_CODE_AUTO_COMPACT_WINDOW": "128000"})
