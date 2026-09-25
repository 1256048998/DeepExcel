# src/DeepExcel.Sidecar/knowledge_skills.py
"""知识技能：Claude Code Skills 的工作簿版。

专业知识（中文数据清洗、财务勾稽、公式写法……）不常驻上下文：system prompt 末尾只放一份索引
（名字 + 一句话说明什么时候该读），模型需要时调用 load_skill(name) 读正文，正文里提到的
补充文件再用 load_skill(name, file) 读。保持 tools=[]：不放开 CLI 内置的 Skill / Read 工具。

来源两处，后者覆盖前者（版本号不低于内置时）：
- 内置：侧车目录下 knowledge/<name>/SKILL.md（随安装包发布）
- 缓存：%LOCALAPPDATA%\\DeepExcel\\knowledge\\<name>\\（服务端下发的新版本落在这里）

SKILL.md 开头是一段 --- 包起来的元数据（name / title / description / version），
其余是正文。技能目录里只认 .md 文件，file 参数不能带路径。
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

BUNDLED_DIR = Path(__file__).resolve().parent / "knowledge"
MAX_FILE_CHARS = 20000
_FRONTMATTER = re.compile(r"^---\s*\n(?P<meta>.*?)\n---\s*\n(?P<body>.*)$", re.S)
_SAFE_FILE = re.compile(r"^[A-Za-z0-9_\-]+\.md$")


def cache_dir() -> Path:
    override = os.environ.get("DEEPEXCEL_KNOWLEDGE_DIR")
    if override:
        return Path(override)
    return Path(os.environ.get("LOCALAPPDATA", "")) / "DeepExcel" / "knowledge"


@dataclass
class Skill:
    name: str
    title: str
    description: str
    version: int
    directory: Path
    body: str
    files: list[str] = field(default_factory=list)
    # 这份知识讲的是哪些工具（frontmatter 的 tools:）；也用于遥测改写「用户实际遇到的报错」
    tools: list[str] = field(default_factory=list)
    # 只适用于哪些宿主（frontmatter 的 hosts:），空表示都适用
    hosts: list[str] = field(default_factory=list)


def _parse(directory: Path) -> Skill | None:
    path = directory / "SKILL.md"
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return None
    m = _FRONTMATTER.match(text.replace("\r\n", "\n"))
    if not m:
        return None
    meta = {}
    for line in m.group("meta").splitlines():
        key, sep, value = line.partition(":")
        if sep:
            meta[key.strip()] = value.strip()
    name = meta.get("name") or directory.name
    if name != directory.name or not meta.get("description"):
        return None
    try:
        version = int(meta.get("version") or 1)
    except ValueError:
        version = 1
    files = sorted(p.name for p in directory.glob("*.md") if p.name != "SKILL.md" and _SAFE_FILE.match(p.name))
    def listed(key):
        return [t.strip() for t in (meta.get(key) or "").split(",") if t.strip()]

    return Skill(name=name, title=meta.get("title") or name, description=meta["description"],
                 version=version, directory=directory, body=m.group("body").strip(), files=files,
                 tools=listed("tools"), hosts=listed("hosts"))


def _scan(root: Path) -> dict[str, Skill]:
    found = {}
    try:
        entries = sorted(p for p in root.iterdir() if p.is_dir())
    except OSError:
        return found
    for directory in entries:
        skill = _parse(directory)
        if skill:
            found[skill.name] = skill
    return found


def catalog() -> dict[str, Skill]:
    """内置 + 缓存；缓存里的同名技能版本不低于内置时用缓存（服务端下发的更新）"""
    skills = _scan(BUNDLED_DIR)
    for name, cached in _scan(cache_dir()).items():
        if name not in skills or cached.version >= skills[name].version:
            skills[name] = cached
    return skills


def index_prompt(skills: dict[str, Skill] | None = None, host: str | None = None) -> str:
    """host：只列适用于这个宿主的技能（frontmatter hosts: excel / wps；不写就是两边都适用）。
    WPS 会话里不列 VBA 技能、Excel 会话里不列 JSA 技能，省得模型去读用不上的知识。
    不按 tools: 过滤——WPS 没注册清洗工具，但中文数据清洗的知识照样用得上。"""
    skills = catalog() if skills is None else skills
    if host is not None:
        skills = {name: s for name, s in skills.items() if not s.hosts or host in s.hosts}
    if not skills:
        return ""
    lines = [
        "",
        "<knowledge-skills>",
        "下面是按需查阅的专业知识，里面是这类任务常见的坑和检查清单，凭常识做很容易踩坑。"
        "用户的请求涉及其中某个主题时，第一个动作就是 load_skill(name) 读它，读完再读表、再动手；"
        "不相关的不要读。",
    ]
    for skill in skills.values():
        lines.append(f"- {skill.name}（{skill.title}）：{skill.description}")
    lines.append("</knowledge-skills>")
    return "\n".join(lines)


def load(name: str, file: str | None = None) -> dict:
    """load_skill 工具的结果（C# ToolResult 的形状）"""
    skills = catalog()
    skill = skills.get((name or "").strip())
    if skill is None:
        return {"success": False, "error": f"没有叫「{name}」的知识技能",
                "suggestion": ("可用的有：" + "、".join(skills)) if skills else "当前没有可用的知识技能"}
    if not file:
        content = skill.body
        if skill.files:
            content += "\n\n（补充文件：" + "、".join(skill.files) + "，需要时用 load_skill(name, file) 读取）"
        return {"success": True, "data": {"name": skill.name, "title": skill.title,
                                          "version": skill.version, "content": content}}
    if not _SAFE_FILE.match(file) or file not in skill.files:
        return {"success": False, "error": f"技能「{skill.name}」里没有文件「{file}」",
                "suggestion": ("可读的补充文件：" + "、".join(skill.files)) if skill.files else "这个技能没有补充文件"}
    try:
        text = (skill.directory / file).read_text(encoding="utf-8")
    except OSError as exc:
        return {"success": False, "error": f"读取失败：{exc}"}
    if len(text) > MAX_FILE_CHARS:
        text = text[:MAX_FILE_CHARS] + "\n…（文件过长，已截断）"
    return {"success": True, "data": {"name": skill.name, "file": file, "content": text}}
