"""知识技能：内置目录、缓存覆盖、load_skill 的读取边界、索引注入"""

import json

import pytest

import knowledge_skills
from excel_tools import register_all_tools

BUNDLED = ("delivery-quality", "cn-data-cleaning", "cn-formula-writing", "cn-financial-reconciliation",
           "cn-payroll-tax", "cn-attendance-roster", "sales-reporting", "ar-aging-reconciliation",
           "vba-writing-debugging")


@pytest.fixture(autouse=True)
def empty_cache(tmp_path, monkeypatch):
    cache = tmp_path / "knowledge-cache"
    monkeypatch.setenv("DEEPEXCEL_KNOWLEDGE_DIR", str(cache))
    return cache


def _write_skill(root, name, version=1, description="测试用", body="正文", extra=None):
    directory = root / name
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "SKILL.md").write_text(
        f"---\nname: {name}\ntitle: 标题{name}\ndescription: {description}\nversion: {version}\n---\n\n{body}\n",
        encoding="utf-8")
    for filename, text in (extra or {}).items():
        (directory / filename).write_text(text, encoding="utf-8")
    return directory


def test_bundled_skills_all_parse():
    skills = knowledge_skills.catalog()
    for name in BUNDLED:
        assert name in skills, name
        skill = skills[name]
        assert skill.description and skill.title != name
        assert len(skill.body) > 500
    assert skills["cn-formula-writing"].files == ["wps.md"]


def test_index_lists_every_skill_with_description():
    index = knowledge_skills.index_prompt()
    assert index.startswith("\n<knowledge-skills>") and index.endswith("</knowledge-skills>")
    for name in BUNDLED:
        assert f"- {name}（" in index


def test_index_skips_skills_cached_for_a_host_that_no_longer_exists(empty_cache):
    """旧知识包里只给 WPS 宿主模式用的技能（hosts: wps）还可能在缓存里，不能列给模型"""
    directory = _write_skill(empty_cache, "old-wps-only")
    text = (directory / "SKILL.md").read_text(encoding="utf-8")
    (directory / "SKILL.md").write_text(text.replace("version: 1\n", "version: 1\nhosts: wps\n"), encoding="utf-8")
    assert knowledge_skills.catalog()["old-wps-only"].hosts == ["wps"]
    index = knowledge_skills.index_prompt()
    assert "- old-wps-only（" not in index
    assert "- vba-writing-debugging（" in index and "- cn-data-cleaning（" in index


def test_index_stays_small():
    # 索引常驻 system prompt，正文按需读；技能多了以后索引本身也不能膨胀
    assert len(knowledge_skills.index_prompt()) < 2500


def test_index_is_empty_without_skills():
    assert knowledge_skills.index_prompt({}) == ""


def test_load_body_mentions_extra_files():
    result = knowledge_skills.load("cn-formula-writing")
    assert result["success"]
    assert "wps.md" in result["data"]["content"]
    assert result["data"]["title"] == "中文用户的公式写入"


def test_load_extra_file():
    result = knowledge_skills.load("cn-formula-writing", "wps.md")
    assert result["success"]
    assert "WPS" in result["data"]["content"]


def test_unknown_skill_lists_what_exists():
    result = knowledge_skills.load("no-such-skill")
    assert not result["success"]
    assert "delivery-quality" in result["suggestion"]


@pytest.mark.parametrize("file", ["../delivery-quality/SKILL.md", "..\\x.md", "SKILL.md", "/etc/passwd",
                                  "C:\\Windows\\win.ini", "wps.txt", "missing.md"])
def test_file_argument_cannot_escape_the_skill(file):
    result = knowledge_skills.load("cn-formula-writing", file)
    assert not result["success"]
    assert "wps.md" in result["suggestion"]


def test_cache_overrides_bundled_when_version_not_lower(empty_cache):
    _write_skill(empty_cache, "delivery-quality", version=2, body="服务端下发的新版正文")
    skills = knowledge_skills.catalog()
    assert skills["delivery-quality"].version == 2
    assert knowledge_skills.load("delivery-quality")["data"]["content"] == "服务端下发的新版正文"


def test_stale_cache_does_not_override_bundled(empty_cache):
    _write_skill(empty_cache, "delivery-quality", version=0, body="旧缓存")
    assert knowledge_skills.catalog()["delivery-quality"].body != "旧缓存"


def test_cache_can_add_new_skills(empty_cache):
    _write_skill(empty_cache, "cn-tax", description="个税计算", extra={"table.md": "税率表"})
    assert "- cn-tax（" in knowledge_skills.index_prompt()
    assert knowledge_skills.load("cn-tax", "table.md")["data"]["content"] == "税率表"


def test_malformed_skills_are_skipped(empty_cache):
    bad = empty_cache / "broken"
    bad.mkdir(parents=True)
    (bad / "SKILL.md").write_text("没有元数据", encoding="utf-8")
    _write_skill(empty_cache, "renamed").joinpath("SKILL.md").write_text(
        "---\nname: other\ndescription: x\n---\n正文", encoding="utf-8")
    skills = knowledge_skills.catalog()
    assert "broken" not in skills and "renamed" not in skills and "other" not in skills


def test_long_extra_file_is_truncated(empty_cache):
    _write_skill(empty_cache, "big", extra={"long.md": "x" * (knowledge_skills.MAX_FILE_CHARS + 50)})
    content = knowledge_skills.load("big", "long.md")["data"]["content"]
    assert content.endswith("已截断）") and len(content) < knowledge_skills.MAX_FILE_CHARS + 50


def test_load_skill_is_registered():
    assert "load_skill" in {t.name for t in register_all_tools()}


@pytest.mark.asyncio
async def test_load_skill_tool_returns_content():
    tool = next(t for t in register_all_tools() if t.name == "load_skill")
    out = await tool.handler({"name": "cn-data-cleaning"})
    payload = json.loads(out["content"][0]["text"])
    assert payload["success"] and "科学计数法" in payload["data"]["content"]


@pytest.mark.asyncio
@pytest.mark.parametrize("file", [None, "", "null", "None", "SKILL.md"])
async def test_load_skill_treats_placeholder_file_as_body(file):
    tool = next(t for t in register_all_tools() if t.name == "load_skill")
    args = {"name": "cn-formula-writing"} if file is None else {"name": "cn-formula-writing", "file": file}
    payload = json.loads((await tool.handler(args))["content"][0]["text"])
    assert payload["success"] and "版本兼容" in payload["data"]["content"]


def test_load_skill_file_is_optional_in_schema():
    tool = next(t for t in register_all_tools() if t.name == "load_skill")
    assert tool.input_schema["required"] == ["name"]


def test_load_skill_never_hits_protected_zone_checks():
    import workbook_memory
    assert workbook_memory.check_write("load_skill", {"name": "x"}, "## 禁区\n- 汇总\n", "汇总") is None
