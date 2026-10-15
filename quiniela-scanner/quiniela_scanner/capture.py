"""Decides *when* to read a slip, for a stack of slips under a fixed camera.

Read a slip once it sits still (and in focus), enter it, then wait for the operator to lift it
off.  The next slip of the stack lands in the same place, so "removed" cannot be told from the
quad alone: the engine watches for a *disturbance* (a hand, the slip moving, the quad vanishing
or jumping, a frame-difference spike over the slip) and only then arms again.  A slip that
settles is fingerprinted from its pen ink (slip_vision.fingerprint) and compared with the last
few entered, so a slip that was lifted and put back - or never lifted at all - is never entered
twice, while a different slip with identical picks always is.

Two things make that robust.  A hand on the slip (skin-coloured blob) is never read: the engine
stays TRACKING and says "Take your hand off the slip" (after ``hand_timeout_s`` it reads anyway,
rows under the hand flagged OCCLUDED, so the stack cannot stall silently).  And while it waits
for the lift it now and then re-checks that what lies there is still the slip it entered
(``_verify``): the lift of a slip is easy to miss at a low frame rate or while a read is
running, and the next slip, lying in the same place, would never be read.  That re-check needs two
looks that agree and differ from every slip entered lately, and it only counts while the picture is
as sharp, as quiet and as colour-true as when the slip was read (a blur, a burst of noise or a colour
cast make the ink of one slip look like another's).
"""
from __future__ import annotations

import threading
import time
from collections import deque
from dataclasses import dataclass, field

import cv2
import numpy as np

from . import slip_vision as sv

IDLE = "idle"                # nothing in view
TRACKING = "tracking"        # slip seen, waiting for it to settle
WAIT_REMOVE = "wait_remove"  # captured (or seen again); waiting for the slip to be lifted or changed
PAUSED = "paused"            # a scan is waiting for the operator

HAND_MESSAGE = "Take your hand off the slip"

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
    timings: dict = field(default_factory=dict)   # ms: 'read_ms' (read_slip), 'fingerprint_ms', 'verify_ms', 'feed_ms'
    fingerprint: sv.Fingerprint | None = None     # pen-ink signature of ``reading`` (set whenever reading is)


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
    HAND_RECHECK = 3         # still frames between two looks at a slip that is under a hand
    HAND_GAP_S = 1.0         # 'a hand has been on the slip' survives gaps this long
    HAND_MSG_S = 0.4         # a hand this long on the slip (a lift is shorter): say so even if the slip never settles
    HAND_PRE = 2.0           # the cheap look needs this many times Settings.skin_slip to refuse a read (the read itself decides the rest)
    HISTORY = 5              # fingerprints of the last slips entered
    # lift insurance (see _verify): the first look comes this long after a read, then one every POLL_EVERY_S (down to
    # POLL_MIN_S when the camera is slow: its frames are the ones a lift can slip between), ever more rarely
    POLL_FIRST_S = 0.4
    POLL_EVERY_S = 0.6
    POLL_MIN_S = 0.25
    POLL_SOON_S = 0.1        # the second look after one that found a different slip
    POLL_AFTER_DUP_S = 0.15  # ... and the first one after a re-read that found the same slip: the lift may just have gone by
    POLL_CALM = 2            # frames in a row without a real disturbance (a hand, a moving slip) needed before a look
    POLL_AGREE = 0.7         # the two looks at a new slip must agree this well (the same slip, measured twice: ~1.0)
    POLL_DEFER_S = 0.25      # a look is put off this long when the picture is not fit to tell slips apart (see _verify)
    CRISP_KEEP = 0.5         # ... because it is blurred (less than this share of the sharpness it had when the slip was read), noisy,
    TINT_KEEP = 0.1          # ... or its white balance moved by more than this share
    BAD_ROWS = 5             # ... or because its read leaves this many rows unclear (multiple, weak, heavy, hidden): garbage, not a slip

    def __init__(self, n_rows: int = 9, stable_frames: int = 4, settings: sv.Settings | None = None, *,
                 wait_focus: bool = True, focus_cap_frames: int = 12, focus_cap_s: float = 0.4,
                 rearm_frames: int = 2, dedupe: bool = True, retry_unreadable: int = 2, clock=time.monotonic,
                 hand_timeout_s: float = 3.0, verify: bool = True):
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
        self.hand_timeout_s = hand_timeout_s     # a hand this long on a slip: read it anyway, rows under it flagged
        self.verify = verify                     # re-check, while waiting for the lift, that the slip is still the one entered
        self.paused = False
        self._lock = threading.RLock()
        self._fps: deque = deque(maxlen=self.HISTORY)
        self._shape = None
        self._noise = 1.0                        # grey levels: sensor noise of the thumbnails (see _diff)
        self._frame_gap = 1 / 30                 # seconds between the camera's frames (measured, see feed)
        self._last_t = None
        self._busy = False                       # the last feed spent time on a read, which says nothing about the camera
        self.reset()

    def reset(self, forget: bool = False):
        """Back to IDLE (arms immediately). The entered slips' fingerprints are kept unless ``forget``."""
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
            self._calm = 0                       # frames in a row that differ from the reference by less than HI
            self._errored = False
            self._hand_t0 = None                 # when the current streak of 'a hand is on the slip' began
            self._hand_last = 0.0
            self._hand_seen = False              # ... and a real read (not just the cheap look) has confirmed it
            self._holding = False                # the last look found a hand: say so while waiting for the next
            self._base = None                    # (fingerprint, grid, alignment template) of the slip lying there since the last read
            self._base_q = (0.0, 0.0, np.ones(2, np.float32))   # ... and the (sharpness, noise, tint) of its picture (see _fit)
            self._cand = None                    # a fingerprint that differed from _base: waiting for a second look
            self._poll_at = None
            self._polls = 0
            self._extra: dict = {}
            if forget:
                self._fps.clear()

    def forget_last(self, rearm: bool = True):
        """Forget the slips that were entered, so the next settled slip - even that very one,
        if it is still there - is read and emitted again (deliberate re-scan)."""
        with self._lock:
            self._fps.clear()
            if rearm and self.state == WAIT_REMOVE:
                self.reset()

    @property
    def last_fingerprint(self) -> sv.Fingerprint | None:
        """Fingerprint of the slip emitted last (None before the first one / after forget_last)."""
        return self._fps[-1] if self._fps else None

    @property
    def history(self) -> list:
        """Fingerprints of the last HISTORY slips emitted, oldest first."""
        return list(self._fps)

    @property
    def _last_fp(self) -> sv.Fingerprint | None:         # the old name
        return self.last_fingerprint

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

    @staticmethod
    def _quality(frame: np.ndarray, roi) -> tuple[float, float, np.ndarray]:
        """(sharpness, noise, tint) of the slip's picture, for _verify. Sharpness: Laplacian variance of the half-size grey
        crop over brightness squared (the halving keeps sensor noise from hiding a blur: a blur of 1.5 px leaves 0.12-0.22
        of it, sensor noise 0-10 grey levels, gain 0.5-1; the same picture again varies by < 2%). Noise: grey levels, median
        deviation from a 3x3 mean (~0.7 x the sensor's sigma; the flat paper dominates the median, the ink does not).
        Tint: blue/green and red/green of the median colour (white balance; an exposure change leaves it alone)."""
        x0, y0, x1, y1 = roi
        crop = frame[y0:y1, x0:x1]
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        half = cv2.resize(gray, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
        sharp = float(cv2.Laplacian(half, cv2.CV_32F).var()) / max(float(half.mean()), 1.0) ** 2
        dev = np.abs(gray.astype(np.int16) - cv2.blur(gray, (3, 3)))[::2, ::2]
        b, g, r = np.median(cv2.resize(crop, (32, 32), interpolation=cv2.INTER_AREA).reshape(-1, 3), axis=0).astype(np.float32)
        return sharp, 1.4826 * float(np.median(dev)), np.array([b, r], np.float32) / max(float(g), 1.0)

    def _fit(self, frame: np.ndarray) -> bool:
        """Is the picture of the slip that waits for its lift as good as it was when the slip was read? A blur (autofocus
        hunting after a hand went by), a burst of noise or a colour cast (exposure, white balance being readjusted) makes the
        ink of the slip look like another's: no basis to call it a new slip."""
        sharp, noise, tint = self._quality(frame, self._ref[0])
        b_sharp, b_noise, b_tint = self._base_q
        return (sharp >= self.CRISP_KEEP * b_sharp and noise <= 2.0 * b_noise + 3.0
                and bool(np.all(np.abs(tint / np.maximum(b_tint, 1e-3) - 1.0) <= self.TINT_KEEP)))

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
        self._calm = 0
        self._stable = 0
        self._prev = None
        self._roi = None
        self._thumbs = []
        self._best_frame = None                          # do not hold on to a 6 MB frame

    def _hand_streak(self) -> bool:
        """A hand was seen on the slip just now: True while it is still allowed to hold the slip back."""
        now = self.clock()
        if self._hand_t0 is None or now - self._hand_last > self.HAND_GAP_S:
            self._hand_t0, self._hand_seen = now, False
        self._hand_last = now
        return now - self._hand_t0 < self.hand_timeout_s

    def _look_hand(self, frame: np.ndarray, quad: np.ndarray | None) -> bool:
        """~2 ms: is a hand clearly on the slip? Needs no grid; keeps the streak that ends in hand_timeout_s."""
        if quad is None or not self.settings.hand_guard:
            return False
        if sv.hand_fraction(frame, quad) < self.HAND_PRE * self.settings.skin_slip:
            return False
        self._hand_streak()
        return True

    def _hold(self, quad: np.ndarray | None, timings: dict) -> Update:
        """A hand is on the slip: do not read it, look again in a few frames (with a fresh best frame)."""
        self._holding = True
        self._stable = max(0, self.stable_frames - self.HAND_RECHECK)
        self._best_frame, self._best_sharp, self._max_sharp, self._sharps = None, 0.0, 0.0, []
        self._focus_t0, self._focus_n = None, 0
        self.state = TRACKING
        return Update(TRACKING, HAND_MESSAGE, quad, 0.5, timings=timings)

    def _read(self, frame: np.ndarray, quad: np.ndarray | None, manual: bool, tracked: bool = False,
              verified: bool = False, agree: sv.Fingerprint | None = None) -> Update:
        """read_slip on ``frame``; emit the reading unless it is a slip entered lately.
        ``tracked``: the slip was just followed by the tracker, so its thumbnails make the reference.
        ``verified``: _verify found it different from the slip entered (a lift that went unseen); ``agree`` is
        what its first look saw: the reading only counts when it shows the same slip."""
        timings = {"read_ms": 0.0}
        fresh = manual or not tracked
        guard = self.settings.hand_guard and not manual
        looked = False
        if guard and not verified and self._look_hand(frame, quad):    # no point in a 200 ms read under a hand ...
            looked = True
            age = self.clock() - self._hand_t0
            if age < self.hand_timeout_s and (self._hand_seen or age < self.HAND_MSG_S):
                return self._hold(quad, timings)                       # ... unless the cheap look keeps saying so: ask a real read
        t0 = time.perf_counter()
        try:
            reading = sv.read_slip(frame, self.n_rows, self.settings)
        except sv.SlipNotFound as e:
            if looked and self._hand_streak():                         # a hand hides the rules: not a bad slip, wait for the hand
                self._hand_seen = True
                timings["read_ms"] = self._ms(t0)
                return self._hold(quad, timings)
            timings["read_ms"] = self._ms(t0)
            if verified:                                     # the slip entered is still there: a look that fails is no news
                self._poll_at = self.clock() + self._poll_every()
                return Update(WAIT_REMOVE, self._wait_msg, self._captured_quad, 1.0, timings=timings)
            self._set_reference(frame, quad, fresh)
            self._errored = True
            self._base = None
            self._hand_t0, self._hand_seen, self._holding = None, False, False
            self._retries = 0 if manual else self._retries
            self._wait_msg = "Slip not recognised. Re-place it flat and fully in view."
            return Update(WAIT_REMOVE, f"Slip seen but not recognised ({e}). Re-place it flat and fully in view.",
                          quad, 1.0, error=str(e), timings=timings)
        timings["read_ms"] = self._ms(t0)
        if guard and reading.hand:
            if verified:                                     # a hand arrived meanwhile: not a new slip, look again later
                self._poll_at = self.clock() + self._poll_every()
                return Update(WAIT_REMOVE, self._wait_msg, self._captured_quad, 1.0, timings=timings)
            if self._hand_streak():
                self._hand_seen = True
                return self._hold(quad, timings)
        t1 = time.perf_counter()
        fp = sv.fingerprint(reading)
        timings["fingerprint_ms"] = self._ms(t1)
        if agree is not None and not sv.same_ink(agree, fp, self.POLL_AGREE):
            self._poll_at = self.clock() + self._poll_every()      # the picture changed again meanwhile: not settled yet
            return Update(WAIT_REMOVE, self._wait_msg, self._captured_quad, 1.0, timings=timings)
        self._hand_t0, self._hand_seen, self._holding = None, False, False
        if verified and sum(v in (sv.MULTIPLE, sv.WEAK, sv.HEAVY, sv.OCCLUDED) for v in reading.verdicts) >= self.BAD_ROWS:
            self._poll_at = self.clock() + self.POLL_DEFER_S      # a picture gone bad (blur, a colour cast) is no new slip
            return Update(WAIT_REMOVE, self._wait_msg, self._captured_quad, 1.0, timings=timings)
        self._errored = False
        self._set_reference(frame, quad, fresh)
        self._base = (fp, reading.grid, sv.alignment_template(reading))
        self._base_q = self._quality(frame, self._ref[0])
        self._cand = None
        self._polls = 0
        self._poll_at = self.clock() + self.POLL_FIRST_S
        if not manual and self.dedupe and any(sv.same_slip(h, fp) for h in self._fps):
            self._wait_msg = "Same slip still there - lift it off."
            self._poll_at = self.clock() + self.POLL_AFTER_DUP_S
            return Update(WAIT_REMOVE, self._wait_msg, quad, 1.0, duplicate=True, timings=timings)
        self._fps.append(fp)
        self._wait_msg = "Captured. Lift the slip - the next one is read automatically."
        if reading.hand:                                     # only after hand_timeout_s (or a manual capture)
            self._wait_msg = "Captured with a hand on the slip - check the rows marked HAND."
        return Update(WAIT_REMOVE, self._wait_msg, quad, 1.0, reading=reading, timings=timings, fingerprint=fp)

    def capture_now(self, frame: np.ndarray) -> Update:
        """Manual capture button: skip the settling wait, the hand wait and the duplicate check."""
        with self._lock:
            return self._read(frame, sv.find_slip_quad(frame), manual=True)

    # ------------------------------------------------------------------ per frame
    def feed(self, frame: np.ndarray) -> Update:
        t0 = time.perf_counter()
        with self._lock:
            self._extra = {}
            now = self.clock()
            if self._last_t is not None and not self._busy and 0 < now - self._last_t < 0.5:
                self._frame_gap += 0.1 * (now - self._last_t - self._frame_gap)
            self._last_t = now
            upd = self._feed(frame)
            upd.timings.update(self._extra)
            self._busy = upd.timings.get("read_ms", 0.0) > 0 or "verify_ms" in upd.timings
        upd.timings.setdefault("read_ms", 0.0)
        upd.timings["feed_ms"] = self._ms(t0)
        return upd

    def _poll_every(self) -> float:
        """Seconds between two looks at the slip that waits for its lift: shorter for a slower camera (a lift is
        seen in fewer frames), longer and longer the more looks found nothing."""
        base = float(np.clip(self.POLL_EVERY_S * (1 / 30) / max(self._frame_gap, 1e-3), self.POLL_MIN_S, self.POLL_EVERY_S))
        return base * (1 if self._polls <= 5 else 3 if self._polls <= 15 else 8)

    def _watch(self, frame: np.ndarray) -> tuple[bool, np.ndarray | None, bool]:
        """WAIT_REMOVE: has the slip been lifted / swapped? Returns (re-armed, quad, quad computed)."""
        roi, ref = self._ref
        cur = self._thumb(frame, roi)
        frac, _ = self._diff(cur, ref)
        quad = self._captured_quad
        self._calm = self._calm + 1 if frac <= self.HI else 0
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

    def _verify(self, frame: np.ndarray) -> Update | None:
        """WAIT_REMOVE and the picture has been quiet: every now and then check that the slip lying there is
        still the one entered. A lift that fell inside a read, or between two frames of a slow camera, leaves
        the next slip - in the same place - looking just like the last one to the frame differences.
        Two looks in a row that agree with each other but not with the slip entered mean a new slip: the first is
        a quick look (the slip's ink, ~90 ms), the second the full read, which is also compared with the history
        (so a slip that was merely nudged is not entered twice)."""
        if not (self.verify and self.dedupe) or self._base is None or self._errored or self._calm < self.POLL_CALM:
            return None
        now = self.clock()
        if self._poll_at is None or now < self._poll_at:
            return None
        base, grid, template = self._base
        if not self._fit(frame):
            self._cand = None
            self._poll_at = now + self.POLL_DEFER_S
            return None
        if self._cand is not None:                       # the second look
            cand, self._cand = self._cand, None
            quad = sv.find_slip_quad(frame)
            if quad is None:
                return None
            return self._read(frame, quad, manual=False, verified=True, agree=cand)
        t0 = time.perf_counter()
        self._polls += 1
        self._poll_at = now + self._poll_every()
        try:
            fp, skin = sv.quick_fingerprint(frame, grid, self.settings, template)
        except sv.SlipNotFound:                          # nothing to look at / the picture is a mess (_watch notices lifts)
            return None
        self._extra["verify_ms"] = self._ms(t0)
        if (self.settings.hand_guard and skin >= self.settings.skin_slip) or sv.same_ink(base, fp):
            return None                                  # a hand, or the same slip: nothing to do
        self._cand = fp                                  # different from the slip entered: make sure with a second look
        self._poll_at = now + self.POLL_SOON_S
        return None

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
                upd = self._verify(frame)
                if upd is not None:
                    return upd
                return Update(WAIT_REMOVE, self._wait_msg, self._captured_quad, 1.0)
        if not have_quad:
            quad = sv.find_slip_quad(frame)
        if quad is None:
            self.reset()
            return Update(IDLE, "Place a slip flat under the camera.", None)

        # --- a hand on the slip? (also while it is still moving: a big hand makes the outline jump about) ---
        if self._look_hand(frame, quad):
            if self.clock() - self._hand_t0 >= self.hand_timeout_s:      # it has been there for good: read what is there
                return self._read(frame, quad, manual=False)
            self._holding = self._holding or self.clock() - self._hand_t0 >= self.HAND_MSG_S
        else:
            self._holding = False

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
            if self._holding:                            # a hand was on the slip a moment ago: still looking
                return Update(TRACKING, HAND_MESSAGE, quad, 0.5)
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
