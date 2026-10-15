"""Decides *when* to read a slip: wait until it sits still and in focus, read it once,
then wait for it to be taken away before arming again."""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from . import slip_vision as sv

IDLE = "idle"                # nothing in view
TRACKING = "tracking"        # slip seen, waiting for it to settle
WAIT_REMOVE = "wait_remove"  # captured; waiting until the slip is taken away
PAUSED = "paused"            # a scan is waiting for the operator


@dataclass
class Update:
    state: str
    message: str
    quad: np.ndarray | None = None
    progress: float = 0.0                     # 0..1 settling progress
    reading: sv.SlipReading | None = None     # set once, on the frame a scan completes
    error: str | None = None


class CaptureEngine:
    def __init__(self, n_rows: int = 9, stable_frames: int = 7, settings: sv.Settings | None = None):
        self.n_rows = n_rows
        self.stable_frames = max(3, stable_frames)
        self.settings = settings or sv.Settings()
        self.paused = False
        self.reset()

    def reset(self):
        self.state = IDLE
        self._prev = None
        self._stable = 0
        self._best_frame = None
        self._best_sharp = 0.0
        self._max_sharp = 0.0
        self._absent = 0
        self._captured_quad = None

    # ------------------------------------------------------------------
    @staticmethod
    def _drift(a: np.ndarray, b: np.ndarray, frame_shape) -> float:
        diag = float(np.hypot(frame_shape[0], frame_shape[1]))
        return float(np.mean(np.linalg.norm(a - b, axis=1)) / diag)

    def _read(self, frame: np.ndarray, quad) -> Update:
        try:
            reading = sv.read_slip(frame, self.n_rows, self.settings)
        except sv.SlipNotFound as e:
            self.state = WAIT_REMOVE
            self._captured_quad = quad
            return Update(WAIT_REMOVE, f"Slip seen but not recognised ({e}). Re-place it flat and fully in view.",
                          quad, 1.0, error=str(e))
        self.state = WAIT_REMOVE
        self._captured_quad = quad
        return Update(WAIT_REMOVE, "Captured. Remove the slip.", quad, 1.0, reading=reading)

    def capture_now(self, frame: np.ndarray) -> Update:
        """Manual capture button: skip the settling wait."""
        return self._read(frame, sv.find_slip_quad(frame))

    def feed(self, frame: np.ndarray) -> Update:
        quad = sv.find_slip_quad(frame)

        if self.paused:
            return Update(PAUSED, "Finish the current slip first (Enter = write, Esc = discard).", quad)

        if self.state == WAIT_REMOVE:
            gone = quad is None
            moved = (quad is not None and self._captured_quad is not None and
                     self._drift(quad, self._captured_quad, frame.shape) > 0.04)
            self._absent = self._absent + 1 if (gone or moved) else 0
            if self._absent >= 6:
                self.reset()
                return Update(IDLE, "Ready - place the next slip.", None)
            return Update(WAIT_REMOVE, "Captured. Remove the slip.", quad, 1.0)

        if quad is None:
            self.reset()
            return Update(IDLE, "Place a slip flat under the camera.", None)

        # --- tracking: is it still, and is the focus converged? ---
        if self._prev is not None and self._drift(quad, self._prev, frame.shape) < 0.006:
            self._stable += 1
        else:
            self._stable = 0
            self._best_frame, self._best_sharp, self._max_sharp = None, 0.0, 0.0
        self._prev = quad
        sharp = sv.sharpness(frame, quad)
        self._max_sharp = max(self._max_sharp, sharp)
        if sharp >= self._best_sharp:
            self._best_frame, self._best_sharp = frame, sharp

        progress = min(1.0, self._stable / self.stable_frames)
        if self._stable >= self.stable_frames:
            focused = self._best_sharp >= 0.85 * self._max_sharp
            if focused or self._stable >= 3 * self.stable_frames:
                best = self._best_frame if self._best_frame is not None else frame
                return self._read(best, quad)
            return Update(TRACKING, "Focusing...", quad, progress)
        return Update(TRACKING, "Hold still...", quad, progress)
