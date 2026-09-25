"""scripts/knowledge_errors.py：用遥测聚合改写技能里的「用户实际遇到的报错」"""

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "..", "scripts"))
import knowledge_errors  # noqa: E402
import knowledge_skills  # noqa: E402

SKILL = """---
name: demo
title: 演示
description: 测试用
version: 2
tools: write_formula, fill_formula_down
---

# 演示

正文。
"""

EXPORT = {"days": 30, "errors": [
    {"tool_name": "write_formula", "error_code": "formula_name", "count": 40},
    {"tool_name": "fill_formula_down", "error_code": "formula_ref", "count": 5},
    {"tool_name": "write_formula", "error_code": "brand_new_code", "count": 4},
    {"tool_name": "write_formula", "error_code": "formula_na", "count": 2},
    {"tool_name": "clean_amount", "error_code": "type_mismatch", "count": 99},
]}


@pytest.fixture
def source(tmp_path):
    root = tmp_path / "knowledge"
    (root / "demo").mkdir(parents=True)
    (root / "demo" / "SKILL.md").write_text(SKILL, encoding="utf-8")
    (root / "untagged").mkdir()
    (root / "untagged" / "SKILL.md").write_text(SKILL.replace("tools: write_formula, fill_formula_down\n", "")
                                                .replace("name: demo", "name: untagged"), encoding="utf-8")
    errors = tmp_path / "errors.json"
    errors.write_text(json.dumps(EXPORT), encoding="utf-8")
    return root, errors


def _read(root, name="demo"):
    return (root / name / "SKILL.md").read_text(encoding="utf-8")


def test_block_lists_only_this_skills_tools_above_the_threshold(source):
    root, errors = source
    changed, unknown = knowledge_errors.run(str(errors), str(root))
    text = _read(root)
    assert changed == ["demo"]
    assert "| write_formula | 公式 #NAME? | 40 |" in text
    assert "| fill_formula_down | 公式 #REF! | 5 |" in text
    assert "clean_amount" not in text          # 别的技能的工具
    assert "#N/A" not in text                  # 低于阈值
    assert unknown == ["brand_new_code"]       # 没有建议的类别要报出来
    assert text.index("40 |") < text.index("5 |")


def test_changed_block_bumps_version_and_rerun_is_stable(source):
    root, errors = source
    knowledge_errors.run(str(errors), str(root))
    once = _read(root)
    assert "version: 3" in once
    changed, _ = knowledge_errors.run(str(errors), str(root))
    assert changed == []
    assert _read(root) == once


def test_block_is_replaced_not_appended(source):
    root, errors = source
    knowledge_errors.run(str(errors), str(root))
    errors.write_text(json.dumps({"days": 7, "errors": [
        {"tool_name": "write_formula", "error_code": "formula_spill", "count": 9}]}), encoding="utf-8")
    knowledge_errors.run(str(errors), str(root))
    text = _read(root)
    assert text.count(knowledge_errors.BEGIN) == 1
    assert "#SPILL!" in text and "#NAME?" not in text and "近 7 天" in text
    assert "version: 4" in text


def test_block_disappears_when_errors_drop_below_threshold(source):
    root, errors = source
    knowledge_errors.run(str(errors), str(root))
    errors.write_text(json.dumps({"days": 30, "errors": []}), encoding="utf-8")
    knowledge_errors.run(str(errors), str(root))
    text = _read(root)
    assert knowledge_errors.BEGIN not in text and text.rstrip().endswith("正文。")


def test_untagged_skills_are_left_alone(source):
    root, errors = source
    before = _read(root, "untagged")
    knowledge_errors.run(str(errors), str(root))
    assert _read(root, "untagged") == before


def test_dry_run_writes_nothing(source):
    root, errors = source
    changed, _ = knowledge_errors.run(str(errors), str(root), dry_run=True)
    assert changed == ["demo"] and _read(root) == SKILL


def test_generated_skill_still_parses_for_the_sidecar(source, monkeypatch):
    root, errors = source
    knowledge_errors.run(str(errors), str(root))
    monkeypatch.setenv("DEEPEXCEL_KNOWLEDGE_DIR", str(root))
    skill = knowledge_skills.catalog()["demo"]
    assert skill.version == 3 and "公式 #NAME?" in skill.body


def test_bundled_skills_declare_their_tools():
    from excel_tools import register_all_tools
    registered = {t.name for host in ("excel", "wps") for t in register_all_tools(host)}
    for skill in knowledge_skills._scan(knowledge_skills.BUNDLED_DIR).values():
        tools = knowledge_errors.skill_tools((skill.directory / "SKILL.md").read_text(encoding="utf-8"))
        assert tools, skill.name
        assert set(tools) <= registered, (skill.name, set(tools) - registered)
