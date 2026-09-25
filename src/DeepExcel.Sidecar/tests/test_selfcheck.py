import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import selfcheck  # noqa: E402
import ui_events  # noqa: E402


def _codes(findings):
    return [f.code for f in findings]


def test_preflight_passes_on_a_healthy_machine(tmp_path):
    cli = tmp_path / "claude.exe"
    cli.write_bytes(b"")
    assert selfcheck.preflight({}, str(cli), 22631) == []


def test_old_windows_and_missing_engine_are_fatal(tmp_path):
    findings = selfcheck.preflight({}, str(tmp_path / "gone.exe"), 17134)  # Win10 1803
    assert _codes(findings) == ["os_too_old", "cli_missing"]
    assert all(f.fatal for f in findings)
    assert "17763" in findings[0].hint


def test_a_stale_git_bash_variable_is_flagged_but_not_fatal(tmp_path):
    cli = tmp_path / "claude.exe"
    cli.write_bytes(b"")
    findings = selfcheck.preflight({selfcheck.GIT_BASH_ENV: r"C:\nonexistent\bash.exe"}, str(cli), None)
    assert _codes(findings) == ["git_bash_path_stale"]
    assert not findings[0].fatal
    # 指向真实文件的变量不动
    assert selfcheck.preflight({selfcheck.GIT_BASH_ENV: str(cli)}, str(cli), None) == []


def _runner(result=None, raises=None):
    def run(args, **kwargs):
        assert args[1] == "--version"
        if raises is not None:
            raise raises
        return result
    return run


def test_probe_maps_windows_errors_to_codes():
    def oserror(winerror):
        exc = OSError("blocked")
        exc.winerror = winerror
        return exc

    assert selfcheck.probe_cli("c.exe", runner=_runner(raises=oserror(225))).code == "cli_blocked"  # 杀毒
    assert selfcheck.probe_cli("c.exe", runner=_runner(raises=oserror(1260))).code == "cli_blocked"  # 组策略
    assert selfcheck.probe_cli("c.exe", runner=_runner(raises=oserror(193))).code == "cli_incompatible"
    assert selfcheck.probe_cli("c.exe", runner=_runner(raises=PermissionError("denied"))).code == "cli_blocked"
    assert selfcheck.probe_cli("c.exe", runner=_runner(raises=FileNotFoundError("x"))).code == "cli_missing"
    timeout = subprocess.TimeoutExpired(["c.exe"], 30)
    assert selfcheck.probe_cli("c.exe", runner=_runner(raises=timeout)).code == "cli_timeout"


def test_probe_reports_a_crash_with_its_output():
    crashed = subprocess.CompletedProcess(["c.exe"], 3, stdout="", stderr="panic: boom")
    finding = selfcheck.probe_cli("c.exe", runner=_runner(crashed))
    assert finding.code == "cli_crashed"
    assert "panic: boom" in finding.detail
    ok = subprocess.CompletedProcess(["c.exe"], 0, stdout="2.1.190 (Claude Code)", stderr="")
    assert selfcheck.probe_cli("c.exe", runner=_runner(ok)) is None


class CLINotFoundError(Exception):
    pass


class ProcessError(Exception):
    def __init__(self, message, exit_code=None, stderr=None):
        super().__init__(message)
        self.exit_code = exit_code
        self.stderr = stderr


def test_diagnose_prefers_what_the_engine_itself_says():
    ok = _runner(subprocess.CompletedProcess(["c.exe"], 0, stdout="2.1.190", stderr=""))
    assert selfcheck.diagnose(CLINotFoundError("not found"), "c.exe", runner=ok).code == "cli_missing"
    # 引擎能打印版本，但会话进程退出了
    crashed = selfcheck.diagnose(ProcessError("Command failed", exit_code=1, stderr="x"), "c.exe", runner=ok)
    assert crashed.code == "cli_crashed"
    assert "exit=1" in crashed.detail
    # 引擎本身被拦截：以探测结果为准，原始异常留在 detail 里
    blocked = selfcheck.diagnose(RuntimeError("init timeout"), "c.exe", runner=_runner(raises=PermissionError("no")))
    assert blocked.code == "cli_blocked"
    assert "init timeout" in blocked.detail
    assert selfcheck.diagnose(RuntimeError("init timeout"), "c.exe", runner=ok).code == "connect_failed"


def test_findings_become_panel_error_events():
    finding = selfcheck.probe_cli("c.exe", runner=_runner(raises=PermissionError("no")))
    event = ui_events.envelope("error", **finding.event_fields())["event"]
    assert event["kind"] == "error"
    assert event["code"] == "cli_blocked"
    assert event["retryable"] is False
    assert event["hint"]


def test_the_bundled_engine_is_found_and_starts():
    cli = selfcheck.bundled_cli()
    if cli is None:
        return  # 这台机器没装 SDK 自带的引擎
    assert selfcheck.preflight({}, cli, selfcheck.windows_build()) == []
    assert selfcheck.probe_cli(cli) is None
