"""Short beeps so the operator can keep their eyes on the stack: ok() = slip entered, attention() = needs you.

Windows only (winsound.MessageBeep returns at once and plays the user's own system sounds); a silent no-op
elsewhere. Never raises.
"""
from __future__ import annotations

import sys


def _backend():
    if sys.platform != "win32":
        return None
    import winsound
    return winsound


def _beep(kind: str) -> None:
    try:
        ws = _backend()
        if ws is not None:
            ws.MessageBeep(getattr(ws, kind))
    except Exception:
        pass


def ok() -> None:
    _beep("MB_ICONASTERISK")


def attention() -> None:
    _beep("MB_ICONEXCLAMATION")
