import numpy as np

from quiniela_scanner.capture import CaptureEngine, IDLE, PAUSED, TRACKING, WAIT_REMOVE
from helpers import draw_slip, draw_slip_named, on_webcam_canvas

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
    assert eng.state == IDLE                                    # lifted: armed again


def test_next_slip_after_removal_is_read_but_the_same_slip_put_back_is_not():
    slip, empty = _frames()
    nxt = on_webcam_canvas(draw_slip_named(list("EEELLVVLV"), "ANA", seed=5), scale=0.5, angle=5)
    eng = CaptureEngine(9, stable_frames=5)
    assert sum(eng.feed(slip).reading is not None for _ in range(30)) == 1
    for _ in range(4):
        eng.feed(empty)
    again = [eng.feed(slip) for _ in range(30)]                 # the very same slip, put back down
    assert not any(u.reading is not None for u in again)
    assert sum(u.duplicate for u in again) == 1 and eng.state == WAIT_REMOVE
    for _ in range(4):
        eng.feed(empty)
    assert sum(eng.feed(nxt).reading is not None for _ in range(30)) == 1     # a different slip is read


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
    assert all(eng.feed(slip).reading is None and eng.feed(slip).state == PAUSED for _ in range(15))


def test_manual_capture_skips_the_wait():
    slip, _ = _frames()
    u = CaptureEngine(9).capture_now(slip)
    assert u.reading is not None and u.reading.picks == PICKS
    assert u.timings["read_ms"] > 0 and u.state == WAIT_REMOVE


def test_manual_capture_of_a_repeat_still_scans():
    slip, _ = _frames()
    eng = CaptureEngine(9)
    assert eng.capture_now(slip).reading is not None
    again = eng.capture_now(slip)
    assert again.reading is not None and not again.duplicate


def test_default_is_four_stable_frames_and_first_read_lands_on_the_fifth_frame():
    slip, _ = _frames()
    eng = CaptureEngine(9)
    assert eng.stable_frames == 4
    states = [eng.feed(slip) for _ in range(8)]
    assert [u.state for u in states[:4]] == [TRACKING] * 4
    assert states[4].reading is not None and states[4].state == WAIT_REMOVE


def test_unreadable_slip_reports_the_error_and_re_arms_on_the_next_disturbance():
    flat = np.full((1080, 1920, 3), 235, np.uint8)
    for i in range(6):
        flat[300 + i * 80:303 + i * 80, 500:1400] = 30
    for x in (500, 800, 1400):
        flat[300:700, x - 1:x + 2] = 30
    empty = np.full_like(flat, 235)
    slip, _ = _frames()
    eng = CaptureEngine(9)
    ups = [eng.feed(flat) for _ in range(12)]
    errs = [u for u in ups if u.error]
    assert len(errs) == 1 and errs[0].reading is None and errs[0].state == WAIT_REMOVE
    for _ in range(4):
        eng.feed(empty)
    assert eng.state == IDLE
    assert sum(eng.feed(slip).reading is not None for _ in range(12)) == 1


def test_the_update_and_engine_contract_for_the_gui():
    """Update.fingerprint rides on every emitted reading (and only then); last_fingerprint follows it; forget_last clears it."""
    from quiniela_scanner import slip_vision as sv
    slip, empty = _frames()
    eng = CaptureEngine(9)
    assert eng.last_fingerprint is None and eng.feed(empty).fingerprint is None
    ups = [eng.feed(slip) for _ in range(8)]
    got = [u for u in ups if u.reading is not None]
    assert len(got) == 1 and isinstance(got[0].fingerprint, sv.Fingerprint) and got[0].fingerprint is eng.last_fingerprint
    assert all(u.fingerprint is None for u in ups if u.reading is None)
    manual = eng.capture_now(slip)
    assert manual.fingerprint is eng.last_fingerprint and manual.fingerprint is not got[0].fingerprint
    eng.forget_last(rearm=False)
    assert eng.last_fingerprint is None and eng.state == WAIT_REMOVE            # rearm=False: still waiting for the lift
    eng.forget_last()
    assert eng.state == IDLE
