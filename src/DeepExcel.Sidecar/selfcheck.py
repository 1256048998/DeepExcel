# src/DeepExcel.Sidecar/selfcheck.py
"""claude.exe 启动自检：起不来时给出明确的诊断码，而不是面板上「没有回复」。

实测结论（CLI 2.1.190）：
- 找不到 Git Bash 时 CLI 退回 PowerShell，不会失败；我们传 tools=[]，根本不起 shell。
  唯一致命的情况是环境变量 CLAUDE_CODE_GIT_BASH_PATH 指向一个不存在的文件（卸载了 Git
  但变量还在）：CLI 直接 exit(1)。这里把这种失效的变量去掉。
- claude.exe 是 Bun 打包的可执行文件，Bun 要求 Windows 10 1809（build 17763）或更新；
  更老的 Win10 上它根本起不来，与 ConPTY 无关（SDK 走管道，不开伪终端）。

两段：
- preflight：不起进程，每次启动都做（零成本）——系统版本、引擎文件在不在、失效的 Git Bash 变量
- diagnose：只在连接失败后做——起一次 `claude.exe --version`，按错误码区分缺失 / 被拦截 / 不兼容 / 崩溃
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Mapping

MIN_WINDOWS_BUILD = 17763  # Windows 10 1809
GIT_BASH_ENV = "CLAUDE_CODE_GIT_BASH_PATH"
PROBE_TIMEOUT_S = 30


@dataclass
class Finding:
    code: str
    message: str
    hint: str
    fatal: bool = True
    detail: str | None = None

    def event_fields(self) -> dict:
        """ui_events.envelope("error", **fields) 用的字段"""
        fields = {"code": self.code, "message": self.message, "hint": self.hint, "retryable": False}
        if self.detail:
            fields["detail"] = self.detail[:500]
        return fields


def _finding(code: str, detail: str | None = None) -> Finding:
    message, hint, fatal = _CATALOG[code]
    return Finding(code=code, message=message, hint=hint, fatal=fatal, detail=detail)


_CATALOG = {
    # code: (中文说明, 下一步, 是否致命)
    "os_too_old": ("这台电脑的 Windows 版本太旧，AI 引擎无法运行。",
                   "AI 引擎需要 Windows 10 1809（版本号 17763）或更新的系统。请先升级 Windows。", True),
    "cli_missing": ("AI 引擎文件缺失。", "运行 DeepExcel 修复工具，或重新安装。", True),
    "cli_blocked": ("AI 引擎被系统拦截，无法启动。",
                    "杀毒软件或公司的应用管控拦截了 claude.exe。请把 DeepExcel 安装目录加入白名单，或联系 IT 放行。", True),
    "cli_incompatible": ("AI 引擎与这台电脑不兼容。",
                         "需要 64 位 Windows 10 1809 或更新的系统。文件也可能已损坏：运行修复工具重新安装。", True),
    "cli_timeout": ("AI 引擎启动超时。",
                    "电脑可能过于繁忙，或杀毒软件在逐个扫描。稍后重启 Excel 再试；反复出现请导出诊断包反馈。", True),
    "cli_crashed": ("AI 引擎一启动就退出了。", "运行修复工具；反复出现请导出诊断包反馈。", True),
    "connect_failed": ("AI 引擎启动了，但连接没有建立起来。", "重启 Excel 再试；反复出现请导出诊断包反馈。", True),
    "git_bash_path_stale": ("环境变量 CLAUDE_CODE_GIT_BASH_PATH 指向的文件不存在，已忽略它。",
                            "不影响使用。可以在系统环境变量里删掉这一项。", False),
}

# Windows 错误码 → 诊断码
_WINERROR = {
    2: "cli_missing", 3: "cli_missing",
    5: "cli_blocked", 225: "cli_blocked", 1260: "cli_blocked", 4551: "cli_blocked",
    193: "cli_incompatible", 216: "cli_incompatible",
}


def bundled_cli() -> str | None:
    """SDK 用的同一个 claude.exe：优先包里自带的，其次 PATH 上的"""
    try:
        import claude_agent_sdk
        name = "claude.exe" if sys.platform == "win32" else "claude"
        path = Path(claude_agent_sdk.__file__).parent / "_bundled" / name
        if path.is_file():
            return str(path)
    except Exception:
        pass
    return shutil.which("claude")


def windows_build() -> int | None:
    getter = getattr(sys, "getwindowsversion", None)
    return getter().build if getter else None


def preflight(environ: Mapping[str, str], cli_path: str | None, build: int | None) -> list[Finding]:
    """不起进程的检查。返回的 git_bash_path_stale 需要调用方从环境里去掉该变量。"""
    findings: list[Finding] = []
    if build is not None and build < MIN_WINDOWS_BUILD:
        findings.append(_finding("os_too_old", f"Windows build {build} < {MIN_WINDOWS_BUILD}"))
    if not cli_path or not os.path.isfile(cli_path):
        findings.append(_finding("cli_missing", f"cli_path={cli_path}"))
    git_bash = environ.get(GIT_BASH_ENV)
    if git_bash and not os.path.isfile(git_bash):
        findings.append(_finding("git_bash_path_stale", f"{GIT_BASH_ENV}={git_bash}"))
    return findings


def probe_cli(cli_path: str | None, env: Mapping[str, str] | None = None,
              runner: Callable = subprocess.run, timeout: float = PROBE_TIMEOUT_S) -> Finding | None:
    """起一次 `claude.exe --version`。能正常打印版本返回 None。"""
    if not cli_path:
        return _finding("cli_missing", "cli_path=None")
    try:
        proc = runner([cli_path, "--version"], env=dict(env) if env is not None else None,
                      capture_output=True, text=True, encoding="utf-8", errors="replace",
                      timeout=timeout, stdin=subprocess.DEVNULL,
                      creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        return _finding("cli_timeout", f"--version did not finish in {timeout}s")
    except FileNotFoundError as exc:
        return _finding("cli_missing", str(exc))
    except OSError as exc:
        code = _WINERROR.get(getattr(exc, "winerror", None) or 0)
        if code is None and isinstance(exc, PermissionError):
            code = "cli_blocked"
        return _finding(code or "cli_crashed", f"{type(exc).__name__}: {exc}")
    if proc.returncode != 0:
        output = ((proc.stderr or "") + (proc.stdout or "")).strip()
        return _finding("cli_crashed", f"exit={proc.returncode}: {output[:400]}")
    return None


def diagnose(exc: BaseException, cli_path: str | None, env: Mapping[str, str] | None = None,
             runner: Callable = subprocess.run) -> Finding:
    """连接失败后找原因：先看异常本身，再实际起一次引擎。"""
    name = type(exc).__name__
    text = f"{name}: {exc}"
    if name == "CLINotFoundError":
        return _finding("cli_missing", text)
    probed = probe_cli(cli_path, env, runner)
    if probed is not None:
        probed.detail = f"{text} / {probed.detail}"
        return probed
    exit_code = getattr(exc, "exit_code", None)
    stderr = getattr(exc, "stderr", None)
    if exit_code is not None:
        return _finding("cli_crashed", f"{text} (exit={exit_code}) {stderr or ''}".strip())
    return _finding("connect_failed", text)
