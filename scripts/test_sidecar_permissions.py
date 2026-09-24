#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for the sidecar's permission gate.

Run:  python scripts/test_sidecar_permissions.py

Two properties matter here, and both are the kind that fail silently:

  * "Allow and remember" must not apply to operations whose blast radius varies
    per call. Approving one delete_rows must not silently approve deleting 200
    rows later.
  * The Computer Use gate is a *deny*, not an *ask*. Moving a tool into the
    high-risk set must not turn a refusal into a prompt.
"""

import ast
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SIDECAR = os.path.join(ROOT, "src", "DeepExcel.Sidecar", "sidecar.py")
FAILURES = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" -- {detail}" if detail else ""))
    if not condition:
        FAILURES.append(name)


def load_sets():
    """Extract the two tool sets without importing the SDK."""
    with open(SIDECAR, "r", encoding="utf-8") as stream:
        tree = ast.parse(stream.read(), filename=SIDECAR)

    wanted = {"_HIGH_RISK_TOOLS", "_REMEMBERABLE_TOOLS", "_COMPUTER_USE_TRIGGERS"}
    namespace = {}
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id in wanted:
                namespace[target.id] = ast.literal_eval(node.value)
    missing = wanted - set(namespace)
    if missing:
        raise RuntimeError(f"sidecar.py no longer defines: {sorted(missing)}")
    return namespace


def hook_source():
    with open(SIDECAR, "r", encoding="utf-8") as stream:
        tree = ast.parse(stream.read(), filename=SIDECAR)
    for node in tree.body:
        if isinstance(node, ast.AsyncFunctionDef) and node.name == "_pre_tool_use_hook":
            return node
    raise RuntimeError("_pre_tool_use_hook not found")


def agent_options_call():
    with open(SIDECAR, "r", encoding="utf-8") as stream:
        tree = ast.parse(stream.read(), filename=SIDECAR)
    calls = [node for node in ast.walk(tree)
             if isinstance(node, ast.Call)
             and isinstance(node.func, ast.Name) and node.func.id == "ClaudeAgentOptions"]
    if len(calls) != 1:
        raise RuntimeError(f"expected exactly one ClaudeAgentOptions(...) call, found {len(calls)}")
    return calls[0]


def main():
    sets = load_sets()
    high_risk = sets["_HIGH_RISK_TOOLS"]
    rememberable = sets["_REMEMBERABLE_TOOLS"]

    print("=== sidecar permission gate ===")

    # ---- coverage --------------------------------------------------------
    for tool in ["execute_vba", "execute_python", "delete_rows", "delete_columns",
                 "clear_range", "write_range", "remove_duplicates", "clean_data",
                 "replace_formula", "merge_cells"]:
        check(f"{tool} requires confirmation", tool in high_risk)

    for tool in ["read_range", "read_workbook", "read_selection", "create_chart",
                 "set_cell_style", "freeze_panes", "list_snapshots"]:
        check(f"{tool} does not prompt", tool not in high_risk,
              "prompting for reads is the noise that makes real prompts ignored")

    # ---- remember scope --------------------------------------------------
    check("rememberable tools are a subset of high risk",
          rememberable <= high_risk)

    # The core rule: capability grants can be remembered, per-operation
    # approvals cannot, because each call has a different blast radius.
    for tool in ["execute_vba", "execute_python"]:
        check(f"{tool} can be remembered", tool in rememberable)

    for tool in ["delete_rows", "delete_columns", "clear_range", "write_range",
                 "remove_duplicates", "clean_data", "rollback"]:
        check(f"{tool} is re-confirmed every time", tool not in rememberable,
              "approving one call must not approve the next, larger one")

    # ---- gate ordering ---------------------------------------------------
    flat = ast.unparse(hook_source())

    computer_use_at = flat.find("_COMPUTER_USE_TRIGGERS")
    high_risk_at = flat.find("_HIGH_RISK_TOOLS")
    check("computer use gate runs before the high-risk check",
          0 <= computer_use_at < high_risk_at,
          "otherwise adding send_keys to the high-risk set turns a refusal into a prompt")

    # ---- send_keys specifically -----------------------------------------
    # send_keys is both a Computer Use tool (deny unless asked) and destructive
    # (confirm when asked). Both must apply, in that order.
    deny_branch = flat[computer_use_at:high_risk_at] if 0 <= computer_use_at < high_risk_at else ""
    check("send_keys is covered by the computer use gate",
          "send_keys" in deny_branch)
    check("send_keys is also high risk", "send_keys" in high_risk)
    check("send_keys grant is rememberable", "send_keys" in rememberable)

    # ---- CLI built-in tools ---------------------------------------------
    # Without tools=[] the model sees 25 Claude Code built-ins (Bash, Read,
    # Write, Glob, Grep, WebFetch, Task...). Read/Glob/Grep need no approval in
    # the CLI, so they bypass CodeSandbox and this hook entirely. allowed_tools
    # does not restrict visibility; only `tools` does.
    options_call = agent_options_call()
    tools_kw = next((kw for kw in options_call.keywords if kw.arg == "tools"), None)
    check("ClaudeAgentOptions disables CLI built-in tools (tools=[])",
          tools_kw is not None
          and isinstance(tools_kw.value, ast.List) and not tools_kw.value.elts,
          "otherwise Read/Glob/Grep can read the user's disk without a prompt")

    # Defence in depth: a non-excel tool reaching the hook must be denied,
    # never passed through.
    non_excel_branch = flat[:flat.find("bare_name")]
    check("hook denies tools outside mcp__excel__",
          "'deny'" in non_excel_branch and "continue_" not in non_excel_branch)

    print()
    if FAILURES:
        print(f"FAILED ({len(FAILURES)}): {', '.join(FAILURES)}")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
