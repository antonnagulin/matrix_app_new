"""Потокобезопасный мост диалогов tkinter для фонового run_check."""

from __future__ import annotations

import threading
from typing import Any


class TkDialogBridge:
    """Вызовы ask_* / choose_* из worker-потока через root.after."""

    DEFAULT_TIMEOUT_SEC = 600.0

    def __init__(self, root, *, timeout_sec: float = DEFAULT_TIMEOUT_SEC) -> None:
        self.root = root
        self.timeout_sec = timeout_sec
        self._closed = False

    def close(self) -> None:
        """Сигнал: UI закрывается — не ждать диалоги бесконечно."""
        self._closed = True

    def ask_include_missing(self, fio: str) -> str:
        from src.logic import ask_include_missing

        return self._call_on_ui(lambda: ask_include_missing(self.root, fio), default="no")

    def choose_actual_record(self, fio, records):
        from src.logic import choose_actual_record

        return self._call_on_ui(
            lambda: choose_actual_record(self.root, fio, records),
            default=None,
        )

    def _call_on_ui(self, fn, default: Any = None) -> Any:
        if self._closed:
            return default
        if threading.current_thread() is threading.main_thread():
            return fn()

        box: dict[str, Any] = {"value": default}
        done = threading.Event()

        def runner() -> None:
            try:
                if self._closed:
                    box["value"] = default
                else:
                    box["value"] = fn()
            except Exception:
                box["value"] = default
            finally:
                done.set()

        try:
            self.root.after(0, runner)
        except Exception:
            return default

        remaining = float(self.timeout_sec)
        slice_sec = 0.25
        while remaining > 0:
            if self._closed:
                return default
            if done.wait(timeout=min(slice_sec, remaining)):
                return box["value"]
            remaining -= slice_sec
        return default
