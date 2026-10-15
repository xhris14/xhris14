"""Frame sources: the webcam (Logitech C920s) and file-based stand-ins for testing."""
from __future__ import annotations

import sys
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
        self.label = f"camera {index} ({int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))}x" \
                     f"{int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))})"

    def set_autofocus(self, on: bool, focus: int = 30):
        self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 1 if on else 0)
        if not on:
            self.cap.set(cv2.CAP_PROP_FOCUS, focus)

    def open_driver_settings(self):
        """DirectShow driver dialog (exposure, focus, brightness...). Windows only."""
        self.cap.set(cv2.CAP_PROP_SETTINGS, 1)

    def read(self):
        ok, frame = self.cap.read()
        return frame if ok else None

    def close(self):
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


def list_cameras(max_index: int = 5, backend: str = "dshow") -> list:
    found = []
    for i in range(max_index + 1):
        cap = cv2.VideoCapture(i, _api(backend))
        if cap.isOpened():
            found.append(i)
        cap.release()
    return found
