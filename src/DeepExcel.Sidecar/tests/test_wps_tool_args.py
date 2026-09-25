"""WPS 端读的参数名必须和侧车转给宿主的一致。

侧车用 call_csharp(name, {...}) 把参数转给宿主；Excel 端（C#）和 WPS 端（tool-dispatcher.js）
各自按名字取值。WPS 端曾有 11 个工具取的名字侧车从来不发（fill_formula_down 读 from/to、
insert_rows 读 address……），取到空值，工具在 WPS 上每次都失败，而任何测试都看不出来。
"""
import ast
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TOOLS = ROOT / "DeepExcel.Sidecar" / "excel_tools.py"
DISPATCHER = ROOT / "DeepExcel.Wps" / "tool-dispatcher.js"

# 侧车原样转发参数（不是字面量字典）的工具：没法静态比对
DYNAMIC = "<dynamic>"


def forwarded_keys():
    tree = ast.parse(TOOLS.read_text(encoding="utf-8"))
    out = {}
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Call) and getattr(node.func, "id", None) == "call_csharp" and node.args):
            continue
        name = node.args[0]
        if not (isinstance(name, ast.Constant) and isinstance(name.value, str)):
            continue
        keys = set()
        if len(node.args) > 1:
            payload = node.args[1]
            if isinstance(payload, ast.Dict):
                keys = {k.value for k in payload.keys if isinstance(k, ast.Constant)}
            else:
                keys = {DYNAMIC}
        out.setdefault(name.value, set()).update(keys)
    return out


def wps_reads():
    source = DISPATCHER.read_text(encoding="utf-8")
    out = {}
    for chunk in re.split(r"\n\s*case '", source)[1:]:
        name = chunk.split("'", 1)[0]
        body = re.split(r"\n\s*case '|\n\s*default:", chunk, maxsplit=1)[0]
        out[name] = set(re.findall(r"_get(?:Arg|Int|StringList)\(args,\s*'(\w+)'", body))
    return out


def test_parsers_see_the_tools():
    forwarded, reads = forwarded_keys(), wps_reads()
    assert "fill_formula_down" in forwarded and "fill_formula_down" in reads
    assert len(reads) >= 25


def test_wps_only_reads_arguments_the_sidecar_sends():
    forwarded = forwarded_keys()
    problems = []
    for tool, used in wps_reads().items():
        sent = forwarded.get(tool)
        if sent is None or DYNAMIC in sent:
            continue
        missing = used - sent
        if missing:
            problems.append(f"{tool}: WPS 读 {sorted(missing)}，侧车只发 {sorted(sent)}")
    assert not problems, "\n".join(problems)


def test_wps_reads_every_argument_the_sidecar_sends():
    forwarded = forwarded_keys()
    problems = []
    for tool, used in wps_reads().items():
        sent = forwarded.get(tool)
        if sent is None or DYNAMIC in sent:
            continue
        ignored = sent - used
        if ignored:
            problems.append(f"{tool}: 侧车发了 {sorted(ignored)}，WPS 没用")
    assert not problems, "\n".join(problems)
