# src/DeepExcel.Sidecar/tests/test_stream_encoding.py
# sidecar.py 在 import 时把 stdout/stderr/stdin 强制成 UTF-8。
# 早期版本三个流共用一个 try：stdin 失败会让已成功的 stdout 被重复包装，
# 旧 wrapper 被 GC 时关掉底层 buffer → "I/O operation on closed file"。
# 这里逐流验证 _force_utf8，不依赖 pytest 自身的捕获机制碰巧触发。
import gc
import io
import sys

import sidecar as sidecar_module


class _ReconfigurableStream:
    """reconfigure 可用的流（真实 Python 3.7+ 的 TextIOWrapper 行为）"""

    def __init__(self):
        self.reconfigured_with = None

    def reconfigure(self, **kwargs):
        self.reconfigured_with = kwargs


class _NoReconfigureNoBuffer:
    """既不能 reconfigure 也没有 buffer（如 pytest 的 DontReadFromInput）"""


def test_stream_without_reconfigure_or_buffer_is_left_alone(monkeypatch):
    stub = _NoReconfigureNoBuffer()
    monkeypatch.setattr(sys, "stdin", stub)

    sidecar_module._force_utf8("stdin")

    assert sys.stdin is stub


def test_reconfigurable_stream_is_reconfigured_in_place(monkeypatch):
    stream = _ReconfigurableStream()
    monkeypatch.setattr(sys, "stdout", stream)

    sidecar_module._force_utf8("stdout")

    assert sys.stdout is stream
    assert stream.reconfigured_with == {"encoding": "utf-8", "line_buffering": True}


def test_failure_on_one_stream_does_not_rewrap_the_others(monkeypatch):
    """回归：stdin 失败不能让已经 reconfigure 成功的 stdout/stderr 被替换"""
    out, err = _ReconfigurableStream(), _ReconfigurableStream()
    monkeypatch.setattr(sys, "stdout", out)
    monkeypatch.setattr(sys, "stderr", err)
    monkeypatch.setattr(sys, "stdin", _NoReconfigureNoBuffer())

    for name in ("stdout", "stderr", "stdin"):
        sidecar_module._force_utf8(name)

    assert sys.stdout is out
    assert sys.stderr is err
    assert out.reconfigured_with is not None
    assert err.reconfigured_with is not None


def test_fallback_wraps_buffer_and_survives_gc_of_original(monkeypatch):
    """无 reconfigure 但有 buffer：包成 UTF-8，且原 stream 被 GC 不会关掉 buffer"""
    buffer = io.BytesIO()

    class _LegacyStream:
        def __init__(self, buf):
            self.buffer = buf

        def __del__(self):
            # 模拟旧 TextIOWrapper 被回收时关闭底层 buffer
            self.buffer.close()

    monkeypatch.setattr(sys, "stdout", _LegacyStream(buffer))
    try:
        sidecar_module._force_utf8("stdout")
        wrapped = sys.stdout
        assert isinstance(wrapped, io.TextIOWrapper)
        assert wrapped.encoding.lower() == "utf-8"

        gc.collect()
        # 不带换行：Windows 下 TextIOWrapper 会把 \n 翻译成 \r\n
        wrapped.write("中文✓")
        wrapped.flush()
        assert buffer.getvalue() == "中文✓".encode("utf-8")
    finally:
        # 避免 wrapper 在 monkeypatch 还原后被回收时关闭 buffer 产生噪音
        if isinstance(sys.stdout, io.TextIOWrapper):
            sys.stdout.detach()
