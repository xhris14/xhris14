"""Decides *when* to read a slip, for a stack of slips under a fixed camera.

Read a slip once it sits still (and in focus), enter it, then wait for the operator to lift it
off.  The next slip of the stack lands in the same place, so "removed" cannot be told from the
quad alone: the engine watches for a *disturbance* (a hand, the slip moving, the quad vanishing
or jumping, a frame-difference spike over the slip) and only then arms again.  A slip that
settles is fingerprinted from its pen ink (slip_vision.fingerprint) and compared with the last
one entered, so a slip that was lifted and put back - or never lifted at all - is never entered
twice, while a different slip with identical picks always is.
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field

import cv2
import numpy as np

from . import slip_vision as sv

IDLE = "idle"                # nothing in view
TRACKING = "tracking"        # slip seen, waiting for it to settle
WAIT_REMOVE = "wait_remove"  # captured (or seen again); waiting for the slip to be lifted or changed
PAUSED = "paused"            # a scan is waiting for the operator

_THUMB = 80                  # side of the grey thumbnail used for change detection


@dataclass
class Update:
    state: str
    message: str
    quad: np.ndarray | None = None
    progress: float = 0.0                     # 0..1 settling progress
    reading: sv.SlipReading | None = None     # set once, on the frame a NEW slip's scan completes
    error: str | None = None
    duplicate: bool = False                   # the slip that settled is the one already entered
    timings: dict = field(default_factory=dict)   # ms: 'read_ms' (read_slip), 'fingerprint_ms', 'feed_ms'


class CaptureEngine:
    # Change detection works on gain-normalised 80x80 thumbnails of the slip's surroundings.
    # A pixel counts as changed when it differs by more than max(PIX_MIN, PIX_K * noise) grey
    # levels beyond what edge wobble (focus breathing, 1 px of misregistration) explains, noise
    # being the measured sensor noise of the thumbnails of a slip lying still.
    PIX_MIN = 14.0
    PIX_K = 10.0
    EDGE_SLACK = 1.0         # px of misregistration tolerated along edges
    HI = 0.05                # changed fraction that is a disturbance on its own (hand, moving slip)
    STRONG = 0.10            # ... and one frame this different (a hand over the stack) re-arms on its own
    LO = 0.01                # changed fraction that counts as "the picture is different" ...
    LO_WEIGHT = 0.25         # ... worth this much of a disturbed frame (so ~8 such frames re-arm)
    CALM_DRIFT = 0.006       # quad movement (fraction of the frame diagonal) between still frames ...
    SLIDE_DRIFT = 0.010      # ... and since the slip last started to hold still (a slow slide is not stillness)
    JUMP_DRIFT = 0.04        # quad movement versus the captured one that means the slip was moved
    RETRY_EVERY = 30         # frames between re-reads of an unreadable slip that is just lying there

    def __init__(self, n_rows: int = 9, stable_frames: int = 4, settings: sv.Settings | None = None, *,
                 wait_focus: bool = True, focus_cap_frames: int = 12, focus_cap_s: float = 0.4,
                 rearm_frames: int = 2, dedupe: bool = True, retry_unreadable: int = 2, clock=time.monotonic):
        self.n_rows = n_rows
        self.stable_frames = max(3, stable_frames)
        self.settings = settings or sv.Settings()
        self.wait_focus = wait_focus
        self.focus_cap_frames = max(1, focus_cap_frames)
        self.focus_cap_s = focus_cap_s
        self.rearm_frames = max(1.0, float(rearm_frames))
        self.dedupe = dedupe
        self.retry_unreadable = retry_unreadable
        self.clock = clock
        self.paused = False
        self._lock = threading.RLock()
        self._last_fp: sv.Fingerprint | None = None
        self._shape = None
        self._noise = 1.0                        # grey levels: sensor noise of the thumbnails (see _diff)
        self.reset()

    def reset(self, forget: bool = False):
        """Back to IDLE (arms immediately). The last-entered fingerprint is kept unless ``forget``."""
        with self._lock:
            self.state = IDLE
            self._prev = None
            self._stable = 0
            self._best_frame = None
            self._best_sharp = 0.0
            self._max_sharp = 0.0
            self._sharps: list = []
            self._anchor = None
            self._wait_msg = ""
            self._roi = None
            self._thumbs: list = []
            self._focus_t0 = None
            self._focus_n = 0
            self._ref = None                     # post-capture reference: (roi, thumbnail)
            self._captured_quad = None
            self._dist = 0.0
            self._retries = 0
            self._quiet = 0
            self._errored = False
            if forget:
                self._last_fp = None

    def forget_last(self, rearm: bool = True):
        """Forget the slip that was entered last, so the next settled slip - even that very one,
        if it is still there - is read and emitted again (deliberate re-scan)."""
        with self._lock:
            self._last_fp = None
            if rearm and self.state == WAIT_REMOVE:
                self.reset()

    @property
    def last_fingerprint(self) -> sv.Fingerprint | None:
        return self._last_fp

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _drift(a: np.ndarray, b: np.ndarray, frame_shape) -> float:
        diag = float(np.hypot(frame_shape[0], frame_shape[1]))
        return float(np.mean(np.linalg.norm(a - b, axis=1)) / diag)

    @staticmethod
    def _roi_of(quad: np.ndarray | None, shape) -> tuple[int, int, int, int]:
        """The slip's bounding box plus a margin, so a hand or the slip's own edge is noticed early."""
        h, w = shape[:2]
        if quad is None:
            return 0, 0, w, h
        x, y, bw, bh = cv2.boundingRect(np.asarray(quad, np.int32))
        m = int(0.12 * max(bw, bh))
        x0, y0, x1, y1 = max(0, x - m), max(0, y - m), min(w, x + bw + m), min(h, y + bh + m)
        if x1 - x0 < 32 or y1 - y0 < 32:
            return 0, 0, w, h
        return x0, y0, x1, y1

    @staticmethod
    def _thumb(frame: np.ndarray, roi) -> np.ndarray:
        x0, y0, x1, y1 = roi
        gray = cv2.cvtColor(frame[y0:y1, x0:x1], cv2.COLOR_BGR2GRAY)
        return cv2.GaussianBlur(cv2.resize(gray, (_THUMB, _THUMB), interpolation=cv2.INTER_AREA)
                                .astype(np.float32), (0, 0), 1.0)

    def _diff(self, cur: np.ndarray, ref: np.ndarray) -> tuple[float, float]:
        """(fraction of pixels that changed, sensor noise estimate) between two thumbnails, after
        cancelling a global brightness change (flicker, auto exposure) and edge wobble."""
        gain = float(np.clip(ref.mean() / max(float(cur.mean()), 1.0), 0.75, 1.33))
        cur = cur * gain
        d = np.abs(cur - ref)
        noise = 1.4826 * float(np.median(d))
        slack = np.maximum(self._edges(ref), self._edges(cur)) * self.EDGE_SLACK
        changed = np.maximum(d - slack, 0.0) > max(self.PIX_MIN, self.PIX_K * self._noise)
        return float(np.count_nonzero(changed)) / d.size, noise

    @staticmethod
    def _edges(t: np.ndarray) -> np.ndarray:
        """Local gradient magnitude in grey levels per thumbnail pixel."""
        return np.hypot(cv2.Sobel(t, cv2.CV_32F, 1, 0, ksize=3), cv2.Sobel(t, cv2.CV_32F, 0, 1, ksize=3)) / 8.0

    @staticmethod
    def _sharp(frame: np.ndarray, roi) -> float:
        """Laplacian variance over the slip, divided by brightness squared so flicker cancels out."""
        x0, y0, x1, y1 = roi
        gray = cv2.cvtColor(frame[y0:y1, x0:x1], cv2.COLOR_BGR2GRAY)
        return float(cv2.Laplacian(gray, cv2.CV_16S).var()) / max(float(cv2.mean(gray)[0]), 1.0) ** 2

    def _ms(self, t0: float) -> float:
        return round((time.perf_counter() - t0) * 1000.0, 2)

    # ------------------------------------------------------------------ reading
    def _set_reference(self, frame: np.ndarray, quad: np.ndarray | None, fresh: bool = False):
        """Remember what the captured slip looks like, to notice when it is lifted or swapped."""
        if self._thumbs and self._roi is not None and not fresh:
            roi = self._roi
            ref = np.mean(self._thumbs[-4:], axis=0).astype(np.float32)
        else:
            roi = self._roi_of(quad, frame.shape)
            ref = self._thumb(frame, roi)
        self._ref = (roi, ref)
        self._captured_quad = quad
        self.state = WAIT_REMOVE
        self._dist = 0.0
        self._quiet = 0
        self._stable = 0
        self._prev = None
        self._roi = None
        self._thumbs = []
        self._best_frame = None                          # do not hold on to a 6 MB frame

    def _read(self, frame: np.ndarray, quad: np.ndarray | None, manual: bool, tracked: bool = False) -> Update:
        """read_slip on ``frame``; emit the reading unless it is the slip entered last.
        ``tracked``: the slip was just followed by the tracker, so its thumbnails make the reference."""
        timings = {"read_ms": 0.0}
        fresh = manual or not tracked
        t0 = time.perf_counter()
        try:
            reading = sv.read_slip(frame, self.n_rows, self.settings)
        except sv.SlipNotFound as e:
            timings["read_ms"] = self._ms(t0)
            self._set_reference(frame, quad, fresh)
            self._errored = True
            self._retries = 0 if manual else self._retries
            self._wait_msg = "Slip not recognised. Re-place it flat and fully in view."
            return Update(WAIT_REMOVE, f"Slip seen but not recognised ({e}). Re-place it flat and fully in view.",
                          quad, 1.0, error=str(e), timings=timings)
        timings["read_ms"] = self._ms(t0)
        self._errored = False
        t1 = time.perf_counter()
        fp = sv.fingerprint(reading)
        timings["fingerprint_ms"] = self._ms(t1)
        self._set_reference(frame, quad, fresh)
        if not manual and self.dedupe and self._last_fp is not None and sv.same_slip(self._last_fp, fp):
            self._wait_msg = "Same slip still there - lift it off."
            return Update(WAIT_REMOVE, self._wait_msg, quad, 1.0, duplicate=True, timings=timings)
        self._last_fp = fp
        self._wait_msg = "Captured. Lift the slip - the next one is read automatically."
        return Update(WAIT_REMOVE, self._wait_msg, quad, 1.0, reading=reading, timings=timings)

    def capture_now(self, frame: np.ndarray) -> Update:
        """Manual capture button: skip the settling wait and the duplicate check."""
        with self._lock:
            return self._read(frame, sv.find_slip_quad(frame), manual=True)

    # ------------------------------------------------------------------ per frame
    def feed(self, frame: np.ndarray) -> Update:
        t0 = time.perf_counter()
        with self._lock:
            upd = self._feed(frame)
        upd.timings.setdefault("read_ms", 0.0)
        upd.timings["feed_ms"] = self._ms(t0)
        return upd

    def _watch(self, frame: np.ndarray) -> tuple[bool, np.ndarray | None, bool]:
        """WAIT_REMOVE: has the slip been lifted / swapped? Returns (re-armed, quad, quad computed)."""
        roi, ref = self._ref
        cur = self._thumb(frame, roi)
        frac, _ = self._diff(cur, ref)
        quad = self._captured_quad
        if frac <= self.LO:
            self._quiet += 1
            self._dist = max(0.0, self._dist - 0.5)
            return False, quad, False
        self._quiet = 0
        lost = jumped = False
        quad = sv.find_slip_quad(frame)                  # something changed: look at the slip itself
        if quad is None:
            lost = True
        elif self._captured_quad is not None:
            jumped = self._drift(quad, self._captured_quad, frame.shape) > self.JUMP_DRIFT
        if frac > self.STRONG:
            self._dist = max(self._dist, self.rearm_frames)
        elif frac > self.HI or lost or jumped:
            self._dist += 1.0
        else:
            self._dist += self.LO_WEIGHT
        if self._dist >= self.rearm_frames:
            self.reset()
            return True, quad, True
        return False, quad, True

    def _feed(self, frame: np.ndarray) -> Update:
        if frame.shape != self._shape:                   # first frame, or the camera changed resolution
            self.reset()
            self._shape = frame.shape
        if self.paused:
            if self.state == WAIT_REMOVE:                # keep watching, so the next slip is ready on resume
                self._watch(frame)
            return Update(PAUSED, "Finish the current slip first (Enter = write, Esc = discard).",
                          self._captured_quad if self.state == WAIT_REMOVE else None)

        quad, have_quad = None, False
        if self.state == WAIT_REMOVE:
            rearmed, quad, have_quad = self._watch(frame)
            if not rearmed:
                if self._errored and self._retries < self.retry_unreadable and self._quiet >= self.RETRY_EVERY:
                    self._retries += 1                   # an unreadable slip may just have been out of focus
                    return self._read(frame, self._captured_quad, manual=False)
                return Update(WAIT_REMOVE, self._wait_msg, self._captured_quad, 1.0)
        if not have_quad:
            quad = sv.find_slip_quad(frame)
        if quad is None:
            self.reset()
            return Update(IDLE, "Place a slip flat under the camera.", None)

        # --- tracking: is it still, and is the focus converged? ---
        calm = (self._prev is not None and self._roi is not None and
                self._drift(quad, self._prev, frame.shape) < self.CALM_DRIFT and
                self._drift(quad, self._anchor, frame.shape) < self.SLIDE_DRIFT)
        self._prev = quad
        if not calm:
            self._restart(frame, quad)
        else:
            cur = self._thumb(frame, self._roi)
            frac, noise = self._diff(cur, self._thumbs[-1])
            if frac > self.HI:                           # a hand or something else moving over the slip
                self._restart(frame, quad)
            else:
                self._noise = 0.9 * self._noise + 0.1 * float(np.clip(noise, 0.2, 3.0))
                self._thumbs = (self._thumbs + [cur])[-4:]
                self._stable += 1
        if self.wait_focus:
            sharp = self._sharp(frame, self._roi)
            self._sharps = (self._sharps + [sharp])[-3:]
            self._max_sharp = max(self._max_sharp, sharp)
            if sharp >= self._best_sharp:
                self._best_frame, self._best_sharp = frame, sharp

        progress = min(1.0, self._stable / self.stable_frames)
        if self._stable < self.stable_frames:
            self.state = TRACKING
            return Update(TRACKING, "Hold still...", quad, progress)
        if not self.wait_focus:
            return self._read(frame, quad, manual=False)
        if self._focus_t0 is None:
            self._focus_t0 = self.clock()
        self._focus_n += 1
        plateau = len(self._sharps) == 3 and min(self._sharps) >= 0.85 * max(self._sharps)
        capped = self._focus_n > self.focus_cap_frames or self.clock() - self._focus_t0 >= self.focus_cap_s
        if (plateau and self._sharps[-1] >= 0.85 * self._max_sharp) or capped:
            return self._read(self._best_frame if self._best_frame is not None else frame, quad,
                              manual=False, tracked=True)
        self.state = TRACKING
        return Update(TRACKING, "Focusing...", quad, progress)

    def _restart(self, frame: np.ndarray, quad: np.ndarray):
        """The slip moved (or something moved over it): start counting stillness again."""
        self._stable = 0
        self._anchor = quad
        self._best_frame, self._best_sharp, self._max_sharp, self._sharps = None, 0.0, 0.0, []
        self._roi = self._roi_of(quad, frame.shape)
        self._thumbs = [self._thumb(frame, self._roi)]
        self._focus_t0, self._focus_n = None, 0
