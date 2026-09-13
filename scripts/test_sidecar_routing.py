#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for the sidecar's outbound routing resolution.

Run:  python scripts/test_sidecar_routing.py

The property under test is credential isolation: exactly one credential reaches
the SDK, and it is the right one for the destination. Getting this wrong sends
the user's provider key to the DeepExcel proxy, or the proxy token to a
third-party provider -- each hands a credential to a party that should not have
it, and neither would be visible in normal use.
"""

import importlib.util
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FAILURES = []


def load_sidecar():
    """Import sidecar.py without its heavy SDK dependencies.

    The module imports claude_agent_sdk at the top, which is not installed in a
    plain checkout. Only the two pure functions are needed, so the source is
    parsed and those definitions executed in isolation.
    """
    import ast

    path = os.path.join(ROOT, "src", "DeepExcel.Sidecar", "sidecar.py")
    with open(path, "r", encoding="utf-8") as stream:
        tree = ast.parse(stream.read(), filename=path)

    wanted = {"build_env_config", "stale_env_keys", "DEFAULT_BASE_URL", "DEFAULT_MODEL"}
    kept = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in wanted:
            kept.append(node)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id in wanted:
                    kept.append(node)
                    break

    module = ast.Module(body=kept, type_ignores=[])
    namespace = {}
    exec(compile(module, path, "exec"), namespace)  # noqa: S102 - test harness
    missing = wanted - set(namespace)
    if missing:
        raise RuntimeError(f"sidecar.py no longer defines: {sorted(missing)}")
    return namespace


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" -- {detail}" if detail else ""))
    if not condition:
        FAILURES.append(name)


def main():
    sidecar = load_sidecar()
    build_env_config = sidecar["build_env_config"]
    stale_env_keys = sidecar["stale_env_keys"]

    print("=== sidecar outbound routing ===")

    # ---- BYOK ----------------------------------------------------------
    env, model = build_env_config(
        {
            "routing_mode": "byok",
            "base_url": "https://api.deepseek.com/anthropic",
            "model": "deepseek-v4",
            "api_key": "sk-user-key",
            "auth_token": "",
        },
        {},
    )
    check("byok uses the local provider", env["ANTHROPIC_BASE_URL"] == "https://api.deepseek.com/anthropic")
    check("byok sends the user's api key", env.get("ANTHROPIC_API_KEY") == "sk-user-key")
    check("byok sends no auth token", "ANTHROPIC_AUTH_TOKEN" not in env)
    check("byok honours the chosen model", model == "deepseek-v4")

    # ---- hosted --------------------------------------------------------
    env, model = build_env_config(
        {
            "routing_mode": "hosted",
            "base_url": "https://api.deepexcel.com/v1",
            "model": "claude-opus-5",
            "api_key": "sk-user-secret",
            "auth_token": "proxy-token-abc",
        },
        {},
    )
    check("hosted uses the proxy base url", env["ANTHROPIC_BASE_URL"] == "https://api.deepexcel.com/v1")
    check("hosted sends the bearer token", env.get("ANTHROPIC_AUTH_TOKEN") == "proxy-token-abc")
    # The important one.
    check("hosted NEVER forwards the user's provider key", "ANTHROPIC_API_KEY" not in env,
          "found: " + repr(env.get("ANTHROPIC_API_KEY")))
    check("hosted honours the chosen model", model == "claude-opus-5")

    # ---- hosted without a token: defence in depth ----------------------
    env, _ = build_env_config(
        {
            "routing_mode": "hosted",
            "base_url": "https://api.deepexcel.com/v1",
            "model": "m",
            "api_key": "sk-user-key",
            "auth_token": "",
        },
        {},
    )
    # The C# resolver refuses to produce this; if it ever arrives, treating it
    # as BYOK is the safe reading -- it is at least a credential the user owns.
    check("hosted without a token falls back to the api key rather than sending nothing",
          env.get("ANTHROPIC_API_KEY") == "sk-user-key" and "ANTHROPIC_AUTH_TOKEN" not in env)

    # ---- no config -----------------------------------------------------
    env, model = build_env_config(
        None, {"ANTHROPIC_BASE_URL": "https://env.example", "ANTHROPIC_MODEL": "env-model"}
    )
    check("no config falls back to the ambient environment",
          env["ANTHROPIC_BASE_URL"] == "https://env.example" and model == "env-model")
    check("no-config path never invents an auth token", "ANTHROPIC_AUTH_TOKEN" not in env)

    # ---- defaults ------------------------------------------------------
    env, model = build_env_config({"routing_mode": "byok", "base_url": "", "model": "", "api_key": ""}, {})
    check("empty base url falls back to anthropic", env["ANTHROPIC_BASE_URL"] == sidecar["DEFAULT_BASE_URL"])
    check("empty model falls back to the default", model == sidecar["DEFAULT_MODEL"])

    # ---- stale key removal ---------------------------------------------
    stale = stale_env_keys({"ANTHROPIC_AUTH_TOKEN": "t", "ANTHROPIC_BASE_URL": "u"})
    check("hosted clears any inherited api key", "ANTHROPIC_API_KEY" in stale)
    check("hosted keeps its own auth token", "ANTHROPIC_AUTH_TOKEN" not in stale)

    stale = stale_env_keys({"ANTHROPIC_API_KEY": "k", "ANTHROPIC_BASE_URL": "u"})
    check("byok clears any inherited auth token", "ANTHROPIC_AUTH_TOKEN" in stale)
    check("byok keeps its own api key", "ANTHROPIC_API_KEY" not in stale)

    # settings.json model defaults would otherwise override the chosen model.
    for key in ("ANTHROPIC_MODEL", "ANTHROPIC_DEFAULT_SONNET_MODEL",
                "ANTHROPIC_DEFAULT_OPUS_MODEL", "ANTHROPIC_DEFAULT_HAIKU_MODEL"):
        check(f"inherited {key} is cleared", key in stale)

    # ---- the log must not contain a live token -------------------------
    source_path = os.path.join(ROOT, "src", "DeepExcel.Sidecar", "sidecar.py")
    with open(source_path, "r", encoding="utf-8") as stream:
        source = stream.read()
    # Diagnostic logs are collected into the support bundle, so a printed token
    # would travel with it.
    check("diagnostics never print the auth token value",
          "ANTHROPIC_AUTH_TOKEN={" not in source and "_diag_env_token" not in source)

    print()
    if FAILURES:
        print(f"FAILED ({len(FAILURES)}): {', '.join(FAILURES)}")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
