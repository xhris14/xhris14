"""Frame sources: the webcam (Logitech C920s) and file-based stand-ins for testing."""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

import cv2
import numpy as np


class CameraError(Exception):
    pass


def _api(backend: str) -> int:
    if sys.platform == "win32":
        return {"dshow": cv2.CAP_DSHOW, "msmf": cv2.CAP_MSMF}.get(backend, cv2.CAP_ANY)
    return cv2.CAP_ANY


class WebcamSource:
    """OpenCV webcam capture tuned for the C920s: 1080p MJPG, small buffer, focus control."""

    def __init__(self, index: int = 0, width: int = 1920, height: int = 1080,
                 backend: str = "dshow", autofocus: bool = True, focus: int = 30):
        self._lock = threading.Lock()            # cv2.VideoCapture is not thread-safe: read() vs the focus slider
        self.cap = cv2.VideoCapture(index, _api(backend))
        if not self.cap.isOpened():
            raise CameraError(f"cannot open camera {index} - is another app (Zoom, Teams, OBS) using it, "
                              "or is the C920s privacy shutter closed?")
        # Order matters on DirectShow: pixel format first, then size, then frame rate.
        self.cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*"MJPG"))
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
        self.cap.set(cv2.CAP_PROP_FPS, 30)
        self.cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        self.set_autofocus(autofocus, focus)
        for _ in range(5):          # let exposure settle
            self.cap.read()
        # what the driver says it will deliver (0 = unknown). In dim light a C920s drops to 10-15 fps on its own
        # (Low Light Compensation / RightLight); FrameGrabber.fps() measures what really arrives.
        try:
            self.fps_reported = float(self.cap.get(cv2.CAP_PROP_FPS) or 0.0)
        except (cv2.error, TypeError, ValueError):
            self.fps_reported = 0.0
        self.label = f"camera {index} ({int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))}x" \
                     f"{int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))})"

    def set_autofocus(self, on: bool, focus: int = 30):
        """Autofocus on, or off with a fixed focus value (0-255; the C920s moves in steps of 5)."""
        with self._lock:
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 1 if on else 0)
            if not on:
                self.cap.set(cv2.CAP_PROP_FOCUS, int(focus))

    def open_driver_settings(self):
        """DirectShow driver dialog (exposure, focus, brightness...). Windows only."""
        with self._lock:
            self.cap.set(cv2.CAP_PROP_SETTINGS, 1)

    def read(self):
        with self._lock:
            ok, frame = self.cap.read()
        return frame if ok else None

    def close(self):
        with self._lock:
            self.cap.release()


class ImageSource:
    """Serve a still image as if it were a camera. ``demo=True`` alternates empty table and
    slip so the whole place-slip / remove-slip cycle can be exercised without hardware."""

    def __init__(self, path: str, demo: bool = False, hold: int = 60, gap: int = 40):
        img = cv2.imread(str(path))
        if img is None:
            raise CameraError(f"cannot read image {path}")
        self.frame = img
        self.demo = demo
        self.hold, self.gap = hold, gap
        self.blank = np.full_like(img, 236)
        self.n = 0
        self.label = f"image {Path(path).name}" + (" (demo loop)" if demo else "")

    def set_autofocus(self, *a, **k):
        pass

    def open_driver_settings(self):
        pass

    def read(self):
        time.sleep(1 / 30)
        self.n += 1
        if self.demo and (self.n % (self.hold + self.gap)) >= self.hold:
            return self.blank.copy()
        return self.frame.copy()

    def close(self):
        pass


class FrameGrabber:
    """Reads ``source`` in its own thread and keeps only the newest frame.

    A consumer that is busy for a moment (reading a slip takes ~200 ms) then never works through a
    backlog of stale frames, and the driver queue cannot back up. ``get`` hands out each frame once."""

    FPS_ALPHA = 0.1             # weight of the newest gap in the moving average
    FPS_MIN_GAPS = 20
    MAX_GAP_S = 0.5

    def __init__(self, source):
        self.source = source
        self._cv = threading.Condition()
        self._frame = None
        self._id = 0
        self._t = time.monotonic()
        self._gap = 0.0                 # moving average of the seconds between two frames
        self._gaps = 0                  # ... made from this many
        self._stopped = False
        self._thread = threading.Thread(target=self._run, name="grabber", daemon=True)
        self._thread.start()

    def _run(self):
        while not self._stopped:
            try:
                frame = self.source.read()
            except Exception:
                frame = None
            if frame is None:
                time.sleep(0.05)
                continue
            now = time.monotonic()
            with self._cv:
                gap = now - self._t
                if self._id and 0 < gap < self.MAX_GAP_S:        # (a longer gap is a stall or an unplugged camera, not a rate)
                    self._gap = gap if not self._gaps else self._gap + self.FPS_ALPHA * (gap - self._gap)
                    self._gaps += 1
                self._frame, self._id, self._t = frame, self._id + 1, now
                self._cv.notify_all()

    def get(self, after: int = 0, timeout: float = 0.5):
        """(id, frame) of the newest frame if it is newer than ``after``; waits up to ``timeout`` s for one,
        else (after, None)."""
        end = time.monotonic() + timeout
        with self._cv:
            while self._id <= after and not self._stopped:
                left = end - time.monotonic()
                if left <= 0:
                    break
                self._cv.wait(left)
            if self._id > after and self._frame is not None:
                return self._id, self._frame
        return after, None

    def age(self) -> float:
        """Seconds since the last frame arrived."""
        return time.monotonic() - self._t

    def fps(self) -> float:
        """Frames per second that really arrive (moving average of the gaps between frames); 0.0 until about
        FPS_MIN_GAPS gaps have been seen. A slow camera (dim light, Low Light Compensation) is what makes the
        lift of a slip fall between two pictures."""
        with self._cv:
            return 1.0 / self._gap if self._gaps >= self.FPS_MIN_GAPS and self._gap > 0 else 0.0

    def stop(self):
        self._stopped = True
        with self._cv:
            self._cv.notify_all()
        if self._thread is not threading.current_thread():
            self._thread.join(1.0)


def list_cameras(max_index: int = 5, backend: str = "dshow") -> list:
    found = []
    for i in range(max_index + 1):
        cap = cv2.VideoCapture(i, _api(backend))
        if cap.isOpened():
            found.append(i)
        cap.release()
    return found
