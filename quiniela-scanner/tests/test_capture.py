import numpy as np

from quiniela_scanner.capture import CaptureEngine, IDLE, WAIT_REMOVE
from helpers import draw_slip, on_webcam_canvas

PICKS = list("LVLLELLVE")


def _frames():
    slip = on_webcam_canvas(draw_slip(PICKS, rng=np.random.default_rng(1)), scale=0.5, angle=5)
    empty = np.full_like(slip, 235)
    return slip, empty


def test_waits_for_stillness_reads_once_then_needs_removal():
    slip, empty = _frames()
    eng = CaptureEngine(9, stable_frames=5)
    assert eng.feed(empty).state == IDLE
    readings = []
    for _ in range(30):
        u = eng.feed(slip)
        if u.reading is not None:
            readings.append(u.reading)
    assert len(readings) == 1                                   # exactly one capture
    assert readings[0].picks == PICKS
    assert eng.feed(slip).state == WAIT_REMOVE                  # still lying there: no re-capture
    for _ in range(8):
        eng.feed(empty)
    assert eng.state == IDLE
    again = [eng.feed(slip) for _ in range(30)]
    assert sum(u.reading is not None for u in again) == 1       # next slip is read as well


def test_moving_slip_is_not_captured():
    slip, empty = _frames()
    eng = CaptureEngine(9, stable_frames=5)
    for k in range(25):
        shifted = np.roll(slip, k * 40, axis=1)                 # slip keeps sliding sideways
        u = eng.feed(shifted)
        assert u.reading is None


def test_paused_engine_never_captures():
    slip, _ = _frames()
    eng = CaptureEngine(9, stable_frames=3)
    eng.paused = True
    assert all(eng.feed(slip).reading is None for _ in range(15))


def test_manual_capture_skips_the_wait():
    slip, _ = _frames()
    u = CaptureEngine(9).capture_now(slip)
    assert u.reading is not None and u.reading.picks == PICKS
