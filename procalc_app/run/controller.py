"""Threaded engine runner with a debounce (autocalc) and stdout capture.

The engine is synchronous and prints progress; we run it on a QThread, capture
stdout into a Qt signal, and coalesce rapid grid edits with a debounce timer.
Cancellation is cooperative-at-boundary: a stale result is discarded and the
latest queued snapshot runs next.
"""
from __future__ import annotations

import io
import sys
import threading

from PySide6.QtCore import QObject, QThread, QTimer, Signal

import engine_api as api


class _StdoutTee(io.TextIOBase):
    def __init__(self, emit):
        self._emit = emit
        self._buf = ""

    def write(self, s):
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self._emit(line)
        return len(s)

    def flush(self):
        if self._buf:
            self._emit(self._buf)
            self._buf = ""


class RunWorker(QThread):
    finished_ok = Signal(list)         # list[str] output paths
    failed = Signal(str)
    log = Signal(str)

    def __init__(self, rows, hmb_path, case, flash_mode, meta, unit_system,
                 unit_overrides=None, stream_snapshot=None):
        super().__init__()
        self._args = (rows, hmb_path, case, flash_mode, meta, unit_system,
                      unit_overrides, stream_snapshot)

    def run(self):
        rows, hmb, case, mode, meta, units, overrides, snapshot = self._args
        tee = _StdoutTee(self.log.emit)
        old = sys.stdout
        sys.stdout = tee
        try:
            paths = api.run(rows, hmb, case=case, flash_mode=mode,
                            meta=meta, unit_system=units, unit_overrides=overrides,
                            stream_snapshot=snapshot)
            self.finished_ok.emit(paths)
        except SystemExit as e:
            self.failed.emit(f"engine stopped: {e}")
        except Exception as e:  # noqa: BLE001
            import traceback
            self.failed.emit(f"{type(e).__name__}: {e}")
            self.log.emit(traceback.format_exc())
        finally:
            tee.flush()
            sys.stdout = old


class RunController(QObject):
    started = Signal()
    finished = Signal(list)
    error = Signal(str)
    log = Signal(str)
    state = Signal(str)                 # "idle" | "running" | "queued"

    def __init__(self, get_rows, get_context, parent=None):
        super().__init__(parent)
        self._get_rows = get_rows          # () -> list[dict]
        self._get_context = get_context    # () -> dict(hmb_path, case, flash_mode, meta, units, auto)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(600)
        self._timer.timeout.connect(self._launch)
        self._worker: RunWorker | None = None
        self._queued = False
        self.auto = False

    # grid edited -> maybe schedule
    def schedule(self):
        if not self.auto:
            return
        self._timer.start()

    def run_now(self):
        self._timer.stop()
        self._launch()

    def _launch(self):
        if self._worker is not None and self._worker.isRunning():
            self._queued = True
            self.state.emit("queued")
            return
        ctx = self._get_context()
        rows = self._get_rows()
        if not rows or not ctx.get("hmb_path"):
            self.error.emit("Load an HMB file and add at least one row before running.")
            return
        self._worker = RunWorker(rows, ctx["hmb_path"], ctx.get("case", "Case 1"),
                                 ctx.get("flash_mode", "isothermal"),
                                 ctx.get("meta"), ctx.get("units", "FPS"),
                                 ctx.get("unit_overrides"),
                                 ctx.get("stream_snapshot"))
        self._worker.log.connect(self.log)
        self._worker.finished_ok.connect(self._done)
        self._worker.failed.connect(self._fail)
        self.state.emit("running")
        self.started.emit()
        self._worker.start()

    def _done(self, paths):
        self.finished.emit(paths)
        self.state.emit("idle")
        self._maybe_requeue()

    def _fail(self, msg):
        self.error.emit(msg)
        self.state.emit("idle")
        self._maybe_requeue()

    def _maybe_requeue(self):
        if self._queued:
            self._queued = False
            self._launch()
