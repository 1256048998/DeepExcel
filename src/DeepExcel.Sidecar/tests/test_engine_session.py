"""Continuing a past conversation: resume the engine session when it is safe,
fall back to the history summary otherwise."""
import pytest

import engine_session
import ipc


@pytest.fixture(autouse=True)
def _clean():
    engine_session._reported["session_id"] = None
    ipc._init_buffer()
    ipc._message_buffer["restore_resume"] = None
    yield
    engine_session._reported["session_id"] = None
    ipc._message_buffer["restore_resume"] = None


@pytest.fixture
def config_dir(tmp_path):
    (tmp_path / "projects" / "C--Users-a-DeepExcel").mkdir(parents=True)
    (tmp_path / "projects" / "C--Users-a-DeepExcel" / "sess-1.jsonl").write_text("{}\n", encoding="utf-8")
    return {"CLAUDE_CONFIG_DIR": str(tmp_path)}


FP = engine_session.fingerprint("deepseek-v4", "https://api.deepseek.com/anthropic/")


def test_fingerprint_ignores_trailing_slash_and_case():
    assert FP == engine_session.fingerprint("deepseek-v4", "https://API.deepseek.com/anthropic")
    assert FP != engine_session.fingerprint("deepseek-v4-flash", "https://api.deepseek.com/anthropic")


def test_each_session_id_is_reported_once():
    assert engine_session.should_report("s1") and not engine_session.should_report("s1")
    assert engine_session.should_report("s2")
    assert not engine_session.should_report(None) and not engine_session.should_report("")


def test_resume_when_fingerprint_matches_and_transcript_exists(config_dir):
    assert engine_session.resume_target({"session_id": "sess-1", "fingerprint": FP}, FP, config_dir) == ("sess-1", "")


@pytest.mark.parametrize("request_,reason", [
    (None, "没有会话 id"),
    ({"fingerprint": FP}, "没有会话 id"),
    ({"session_id": "sess-1", "fingerprint": "other"}, "模型或出口变了"),
    ({"session_id": "gone", "fingerprint": FP}, "会话记录不在了"),
    ({"session_id": "../../etc", "fingerprint": FP}, "会话记录不在了"),
])
def test_no_resume_falls_back_with_a_reason(config_dir, request_, reason):
    assert engine_session.resume_target(request_, FP, config_dir) == (None, reason)


def test_restore_history_keeps_the_resume_request():
    ipc.route_message({"type": "restore_history", "messages": [{"role": "user", "content": "hi"}],
                       "resume_session_id": "sess-1", "resume_fingerprint": FP})
    assert ipc._message_buffer["restore_resume"] == {"session_id": "sess-1", "fingerprint": FP}


def test_old_hosts_without_a_session_id_change_nothing():
    ipc.route_message({"type": "restore_history", "messages": [{"role": "user", "content": "hi"}]})
    assert ipc._message_buffer["restore_resume"] is None


@pytest.mark.asyncio
async def test_maybe_resume_swaps_the_client(monkeypatch, config_dir):
    import sidecar
    monkeypatch.setattr(sidecar, "_engine_fingerprint", FP)
    monkeypatch.setattr(engine_session.os, "environ", config_dir)
    ipc._message_buffer["restore_resume"] = {"session_id": "sess-1", "fingerprint": FP}
    calls = []

    async def reconnect(session_id):
        calls.append(session_id)
        return "resumed-client"

    client, resumed = await sidecar._maybe_resume("fresh-client", reconnect)
    assert (client, resumed, calls) == ("resumed-client", True, ["sess-1"])
    assert ipc._message_buffer["restore_resume"] is None


@pytest.mark.asyncio
async def test_maybe_resume_keeps_the_fresh_client_when_reconnect_fails(monkeypatch, config_dir):
    import sidecar
    monkeypatch.setattr(sidecar, "_engine_fingerprint", FP)
    monkeypatch.setattr(engine_session.os, "environ", config_dir)
    ipc._message_buffer["restore_resume"] = {"session_id": "sess-1", "fingerprint": FP}

    async def reconnect(session_id):
        raise RuntimeError("No conversation found")

    assert await sidecar._maybe_resume("fresh-client", reconnect) == ("fresh-client", False)


@pytest.mark.asyncio
async def test_maybe_resume_without_a_request_does_nothing():
    import sidecar

    async def reconnect(session_id):
        raise AssertionError("must not reconnect")

    assert await sidecar._maybe_resume("fresh-client", reconnect) == ("fresh-client", False)
