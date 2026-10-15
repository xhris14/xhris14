"""A hand on the slip (review findings E1 and E2): it must not make the engine enter the slip a second time, and
rows it covers must never be read as picks or as 'empty' (F1).

The skin guard is validated for false positives on the real photo, coloured inks, colour casts and light/dark
backgrounds; the engine is exercised with the tests' Stage (a fixed 1080p camera over a stack) and a virtual clock."""
from __future__ import annotations

import dataclasses
import os

import cv2
import numpy as np
import pytest

from quiniela_scanner import slip_vision as sv
from quiniela_scanner.capture import HAND_MESSAGE, TRACKING, WAIT_REMOVE, CaptureEngine
from helpers import (Stage, VClock, draw_hand, draw_slip, draw_slip_named, on_webcam_canvas, with_hand)

PICKS = list("LVLLELLVE")
OTHER = list("EEELLVVLV")
THIRD = list("VVVLLLEEE")
REAL_SLIP = os.environ.get("QUINIELA_REAL_SLIP", "")        # a real slip photo, kept out of the repo
FPS = 30


# =============================================================================== the skin detector
def _slip_frame(ink=(170, 60, 25), bg=235, **kw):
    return on_webcam_canvas(draw_slip(PICKS, ink=ink, rng=np.random.default_rng(1)), scale=0.5, angle=4, bg=bg, noise=3, **kw)


@pytest.mark.parametrize("ink", [(170, 60, 25), (20, 20, 20), (40, 40, 200), (30, 120, 240), (30, 60, 110), (150, 40, 120)],
                         ids=["blue", "black", "red", "orange", "brown", "purple"])
@pytest.mark.parametrize("bg", [40, 235])
def test_clean_slips_show_no_skin_whatever_the_ink_and_the_table(ink, bg):
    frame = _slip_frame(ink, bg)
    rd = sv.read_slip(frame)
    assert rd.skin_fraction < 0.001 and not rd.hand and not rd.name_hidden and rd.occluded_rows == []
    assert sv.hand_fraction(frame, sv.find_slip_quad(frame)) < 0.001


@pytest.mark.parametrize("tint", [(1, 1, 1), (0.75, 0.9, 1.0), (0.6, 0.85, 1.0), (0.5, 0.8, 1.0), (0.5, 0.65, 0.95),
                                  (1.0, 0.9, 0.75), (0.9, 1.0, 0.9)])
@pytest.mark.parametrize("gain", [0.5, 1.0, 1.3])
def test_clean_slips_show_no_skin_under_colour_casts_and_exposures(tint, gain):
    """tint is a BGR multiplier: (0.5, 0.8, 1.0) is a strong orange lamp, (1.0, 0.9, 0.75) a cold blue light."""
    frame = np.clip(_slip_frame().astype(np.float32) * np.float32(tint)[None, None] * gain, 0, 255).astype(np.uint8)
    rd = sv.read_slip(frame)
    assert rd.skin_fraction < 0.001 and not rd.hand, rd.skin_fraction
    assert sv.hand_fraction(frame, sv.find_slip_quad(frame)) < 0.001


@pytest.mark.parametrize("paper", [(0.82, 0.78, 1.0), (0.6, 0.95, 1.0), (0.5, 0.75, 1.0), (0.85, 1.0, 0.85)],
                         ids=["pink", "yellow", "orange", "green"])
def test_coloured_paper_is_not_skin(paper):
    slip = np.clip(draw_slip(PICKS, rng=np.random.default_rng(1)).astype(np.float32) * np.float32(paper)[None, None], 0, 255).astype(np.uint8)
    rd = sv.read_slip(on_webcam_canvas(slip, scale=0.5, angle=4, noise=3))
    assert not rd.hand and rd.skin_fraction < 0.001


def test_thick_felt_markers_in_skin_hues_are_not_a_hand():
    """1.1 mm markers in red / orange / brown, every pick circled as well as crossed: still far below 1% skin."""
    for ink in ((40, 40, 200), (30, 120, 240), (30, 60, 110)):
        slip = draw_slip([""] * 9)
        for r, c in enumerate("LEVLEVLEV"):
            cx, cy = {"L": 40, "E": 104, "V": 168}[c], 200 + 80 * r
            cv2.line(slip, (cx - 24, cy - 24), (cx + 24, cy + 24), ink, 11, cv2.LINE_AA)
            cv2.line(slip, (cx + 24, cy - 24), (cx - 24, cy + 24), ink, 11, cv2.LINE_AA)
            cv2.circle(slip, (cx, cy), 34, ink, 11, cv2.LINE_AA)
        rd = sv.read_slip(on_webcam_canvas(slip, scale=0.5, angle=4, noise=3))
        assert rd.skin_fraction < 0.008 and not rd.hand, (ink, rd.skin_fraction)


@pytest.mark.parametrize("tone", [(170, 190, 235), (105, 150, 198), (80, 125, 170), (55, 90, 135), (35, 60, 95)],
                         ids=["fair", "medium", "tan", "brown", "dark"])
def test_a_hand_is_found_for_every_skin_tone(tone):
    clean = _slip_frame()
    g = sv.locate_grid(clean)
    # the hand lies over the lower-left of the slip: found with the clean slip's grid (the hand hides the rules)
    frame = with_hand(clean, 800, 640, 90, 1.0, tone)
    rd = sv.read_slip(frame, grid=g)
    assert rd.hand and rd.skin_fraction > 0.02 and rd.occluded_rows
    assert sv.hand_fraction(frame, sv.find_slip_quad(frame)) > 0.02


def test_a_blue_pen_and_a_hand_in_one_picture_are_told_apart():
    clean = _slip_frame()
    g = sv.locate_grid(clean)
    rd = sv.read_slip(with_hand(clean, 1150, 700, 200, 0.7), grid=g)         # a hand over the lower right: no marks there
    assert rd.hand and rd.picks[:3] == PICKS[:3] and rd.verdicts[:3] == [sv.OK] * 3


# =============================================================================== OCCLUDED rows
def _under_a_hand():
    clean = _slip_frame()
    g = sv.locate_grid(clean)
    return clean, g, sv.read_slip(with_hand(clean, 780, 560, 90, 0.8), grid=g)


def test_rows_under_a_hand_are_occluded_never_picks_and_never_empty():
    clean, g, rd = _under_a_hand()
    truth = sv.read_slip(clean, grid=g)
    assert truth.picks == PICKS and truth.clean
    occ = rd.occluded_rows
    assert len(occ) >= 3
    for r in occ:
        assert rd.verdicts[r] == sv.OCCLUDED and rd.picks[r] == "" and rd.candidates[r] == ()
        assert rd.verdicts[r] not in (sv.OK, sv.BLANK)                      # nothing for the sink to enter, not even F1
    assert not rd.clean and set(occ) <= set(rd.problem_rows)
    for r in range(9):                                                        # rows far from the hand read as before
        if r not in occ and abs(r - occ[0]) > 2 and abs(r - occ[-1]) > 2:
            assert rd.verdicts[r] == sv.OK and rd.picks[r] == PICKS[r]


def test_occluded_rows_are_drawn_in_their_own_colour_and_label():
    clean, g, rd = _under_a_hand()
    colour = sv._VERDICT_COLOUR[sv.OCCLUDED]
    assert colour not in [c for v, c in sv._VERDICT_COLOUR.items() if v != sv.OCCLUDED]       # unlike every other verdict
    img = sv.draw_reading(rd)
    as_weak = sv.draw_reading(dataclasses.replace(rd, verdicts=[sv.WEAK if v == sv.OCCLUDED else v for v in rd.verdicts]))
    mine = (img == np.array(colour, np.uint8)).all(axis=2)
    assert mine.sum() > 500 and not (as_weak == np.array(colour, np.uint8)).all(axis=2).any()
    r = rd.occluded_rows[0]
    rows = np.nonzero(mine.any(axis=1))[0]
    assert g.corners[r, 0, 1] - 6 <= rows.min() and rows.max() <= g.corners[rd.occluded_rows[-1] + 1, 0, 1] + 6   # drawn on those rows only
    assert not (img == as_weak).all()                                        # and the label says HAND, not '??'


def test_a_finger_on_the_left_edge_used_to_turn_hidden_marks_into_f1():
    """E2 on a synthetic slip: a thumb pinning the left edge hides the marks; they used to read BLANK (F1) and the slip
    was auto-eligible (at most 2 blanks); now the covered rows are OCCLUDED and the slip needs a look."""
    clean = _slip_frame()
    g = sv.locate_grid(clean)
    rd = sv.read_slip(with_hand(clean, 700, 470, 90, 0.55, (105, 150, 198)), grid=g)
    assert sv.BLANK not in [rd.verdicts[r] for r in rd.occluded_rows]
    assert rd.occluded_rows and not rd.clean
    assert all(v in (sv.OK, sv.BLANK) for v in rd.verdicts) is False          # never auto-eligible


def test_the_guard_can_be_switched_off():
    clean = _slip_frame()
    g = sv.locate_grid(clean)
    rd = sv.read_slip(with_hand(clean, 780, 560, 90, 0.8), 9, sv.Settings(hand_guard=False), grid=g)
    assert not rd.hand and rd.occluded_rows == [] and rd.skin_fraction == 0.0


# =============================================================================== same_slip is occlusion tolerant
def _fp(slip, **kw):
    kw = {"scale": 0.5, "angle": 4, "seed": 3, "noise": 3, **kw}
    return sv.fingerprint(sv.read_slip(on_webcam_canvas(slip, **kw)))


def test_a_hand_over_the_name_strip_does_not_make_a_new_slip():
    """E1 at the fingerprint level: same marks, the handwriting gone (hidden by a hand, or wiped)."""
    a = _fp(draw_slip_named(PICKS, "CLEMENTE O.", seed=7))
    hidden = _fp(draw_slip_named(PICKS, "", seed=7), angle=-3, seed=5)
    assert hidden.name_blank and not a.name_blank
    assert sv.same_slip(a, hidden) and sv.same_slip(hidden, a)
    other = _fp(draw_slip_named(PICKS, "CLEMENTE O.", seed=9), angle=-3, seed=5)       # same name, marks drawn again
    assert not sv.same_slip(a, other)


def test_rows_hidden_in_one_of_the_photos_are_ignored():
    a_read = sv.read_slip(_slip_frame())
    a = sv.fingerprint(a_read)
    g = a_read.grid
    covered = sv.fingerprint(sv.read_slip(with_hand(_slip_frame(), 780, 560, 90, 0.8), grid=g))
    assert covered.occluded and covered.picks != a.picks
    assert sv.same_slip(a, covered) and sv.same_slip(covered, a)
    # a different slip with a hand over it is still a different slip
    other = _fp(draw_slip_named(OTHER, "ANA", seed=5))
    assert not sv.same_slip(other, covered)


def test_two_marks_only_slips_are_not_matched_on_marks_alone():
    """One or two X marks prove little: below _FP_MIN_ROWS marked rows the handwriting still has to agree."""
    few = [""] * 9
    few[1], few[6] = "L", "V"
    a = _fp(draw_slip_named(few, "CLEMENTE O.", seed=7))
    b = _fp(draw_slip_named(few, "MARIA", seed=7), angle=-3, seed=5)           # the very same two marks, another name
    assert sv.slip_similarity(a, b)[0] > sv.MARKS_STRONG
    assert not sv.same_slip(a, b)


# =============================================================================== the engine and a hand on the slip
def make_stage(seed=2, **kw):
    st = Stage(seed=seed, **kw)
    ids = [st.add(draw_slip_named(p, n, seed=s)) for p, n, s in ((PICKS, "CLEMENTE O.", 1), (OTHER, "ANA", 5), (THIRD, "BETO", 6))]
    return st, ids


def run_clock(eng, clock, *gens):
    """Feed every frame (a list of frames or a generator of (frame, info)) at 30 fps of virtual time: [(t, update)]."""
    out = []
    for gen in gens:
        for item in gen:
            frame = item[0] if isinstance(item, tuple) else item
            clock.t += 1 / FPS
            out.append((round(clock.t, 3), eng.feed(frame)))
    return out


def emitted(ev):
    return [(t, u) for t, u in ev if u.reading is not None]


def cycle(frames, n):
    """n frames made of the (few, noisy) ``frames`` over and over: a slip lying still for as long as needed, cheaply."""
    return [frames[i % len(frames)] for i in range(n)]


def hand_scene(st, A, B, x, y, angle, rest, speed=12, side="left", size=1.7, tone=(110, 150, 205), hold=14):
    """A is read, a hand reaches in from ``side`` to (x, y) and rests ``rest`` frames, leaves; A is then lifted, B read."""
    sx, sy = {"left": (x - 900, y + 60), "right": (x + 900, y + 200), "bottom": (x + 50, y + 900)}[side]
    n = max(2, int(np.hypot(x - sx, y - sy) / speed))
    bank = [f for f, _ in st.hold(A, 8, under=B)]
    base = cycle(bank, n + rest + 2)
    app = [with_hand(base[i], sx + (x - sx) * (i + 1) / n, sy + (y - sy) * (i + 1) / n, angle, size, tone) for i in range(n)]
    app += [with_hand(base[n + i], x, y, angle, size, tone) for i in range(rest)]
    return [[f for f, _ in st.empty(3, under=B)], cycle(bank, hold), app, cycle(bank, 10),
            [f for f, _ in st.lift(A, B, 8)], [f for f, _ in st.hold(B, 14)]]


# placements that used to enter the slip again (old engine, measured: 1-2 extra rows each, some with wrong picks)
SCENES = [("name strip", "bottom", (1000, 700), 0, 12, 6), ("name strip, fast", "bottom", (1000, 700), 0, 25, 4),
          ("marks, thumb", "left", (760, 500), 90, 12, 6), ("marks, fast", "left", (760, 500), 90, 25, 4),
          ("middle", "right", (960, 540), 200, 12, 6)]


@pytest.mark.parametrize("label,side,xy,angle,speed,rest", SCENES, ids=[s[0] for s in SCENES])
def test_a_hand_reaching_for_a_slip_already_entered_does_not_enter_it_again(label, side, xy, angle, speed, rest):
    st, (A, B, _) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    parts = hand_scene(st, A, B, *xy, angle, rest, speed, side)
    ev = run_clock(eng, clock, *parts)
    got = emitted(ev)
    assert [u.reading.picks for _, u in got] == [PICKS, OTHER], [(t, u.reading.picks) for t, u in got]
    assert all(u.reading.clean for _, u in got)
    assert not [u for _, u in ev if u.error]


def test_the_engine_says_so_while_a_hand_rests_on_the_slip_and_reads_nothing():
    st, (A, B, _) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    ev = run_clock(eng, clock, *hand_scene(st, A, B, 1000, 700, 0, 40, 12, "bottom")[:3])      # (the last parts are not needed)
    held = [u for _, u in ev if u.message == HAND_MESSAGE]
    assert held and all(u.state == TRACKING and u.reading is None and not u.duplicate for u in held)
    assert HAND_MESSAGE == "Take your hand off the slip"
    assert len(emitted(ev)) == 1                                              # only A, before the hand arrived


def test_a_new_slip_under_a_hand_waits_for_the_hand_and_is_then_read_clean():
    st, (A, B, C) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    base = [f for f, _ in st.hold(C, 8, under=B)]
    hand = [with_hand(f, 1000, 700, 0, 0.8) for f in base]                    # 1.5 s under a hand: shorter than the time-out
    ev = run_clock(eng, clock, cycle(hand, 45), cycle(base, 15))
    got = emitted(ev)
    assert len(got) == 1 and got[0][1].reading.picks == THIRD and got[0][1].reading.clean
    assert got[0][0] > 1.5                                                    # not before the hand was gone
    assert any(u.message == HAND_MESSAGE for _, u in ev)


def test_after_the_time_out_a_slip_under_a_hand_is_emitted_with_its_covered_rows_flagged():
    """The stack must not stall silently: after hand_timeout_s the reading goes out, rows under the hand OCCLUDED."""
    st, (A, B, C) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock, hand_timeout_s=3.0)
    base = [f for f, _ in st.hold(C, 8, under=B)]
    hand = [with_hand(f, 775, 500, 90, 0.7) for f in base]
    ev = run_clock(eng, clock, cycle(hand, 150), cycle(base, 30))
    got = emitted(ev)
    assert len(got) == 1
    t, u = got[0]
    assert 2.9 < t < 3.6, t                                                   # ~3 s of continuous occlusion
    rd = u.reading
    assert rd.hand and rd.occluded_rows and not rd.clean and u.fingerprint is not None
    assert not all(v in (sv.OK, sv.BLANK) for v in rd.verdicts)                # the window's auto-enter rule: never fast-entered
    assert all(rd.verdicts[r] == sv.OCCLUDED and rd.picks[r] == "" for r in rd.occluded_rows)
    assert [i for i, v in enumerate(rd.verdicts) if v == sv.OK] and u.state == WAIT_REMOVE
    assert sum(1 for _, x in ev if x.message == HAND_MESSAGE) > 10            # it said so all along
    # the hand leaves, the slip is still there: the clean re-read is the same slip, not a second ticket
    assert not [1 for _, x in ev if x.reading is not None and x is not u]
    assert any(x.duplicate for _, x in ev[150:])


def test_a_hand_resting_for_ever_on_an_entered_slip_never_makes_a_second_ticket():
    st, (A, B, _) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock, hand_timeout_s=2.0)
    base = [f for f, _ in st.hold(A, 8, under=B)]
    hand = [with_hand(f, 775, 500, 90, 0.7) for f in base]
    ev = run_clock(eng, clock, cycle(base, 14), cycle(hand, 120), cycle(base, 20),
                   [f for f, _ in st.lift(A, B, 8)], [f for f, _ in st.hold(B, 14)])
    assert [u.reading.picks for _, u in emitted(ev)] == [PICKS, OTHER]
    assert any(u.duplicate for _, u in ev)                                    # the occluded read matched A and was dropped


def test_a_cheap_look_that_cries_hand_for_ever_is_checked_by_a_real_read(monkeypatch):
    """The ~2 ms look at the slip's outline could be fooled (a skin-coloured table in view...): it can hold a read back for
    0.4 s, no longer - then a real read decides, and a slip that shows no skin goes in at once instead of after 3 s."""
    st, (A, B, C) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    monkeypatch.setattr(sv, "hand_fraction", lambda frame, quad, width=480: 0.5)
    ev = run_clock(eng, clock, st.empty(3, under=B), st.hold(C, 40, under=B))
    got = emitted(ev)
    assert len(got) == 1 and got[0][1].reading.picks == THIRD and got[0][1].reading.clean
    assert got[0][0] < 1.4, got[0][0]                                          # 0.1 s of empty table + ~0.5 s, not 3 s


def test_hand_guard_off_reads_under_a_hand_like_before():
    st, (A, B, C) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, 4, sv.Settings(hand_guard=False), clock=clock)
    base = [f for f, _ in st.hold(C, 8, under=B)]
    ev = run_clock(eng, clock, cycle([with_hand(f, 1000, 700, 0, 0.8) for f in base], 40))
    assert len(emitted(ev)) == 1 and not any(u.message == HAND_MESSAGE for _, u in ev)


def test_manual_capture_never_waits_for_the_hand():
    st, (A, B, C) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    frame = with_hand(next(st.hold(C, 1, under=B))[0], 775, 500, 90, 0.7)
    u = eng.capture_now(frame)
    assert u.reading is not None and u.reading.hand and u.reading.occluded_rows and u.fingerprint is not None
    assert u.state == WAIT_REMOVE and not u.duplicate


# =============================================================================== history of fingerprints, and the contract
def test_an_earlier_slip_put_back_is_still_recognised_the_history_is_more_than_the_last_one():
    st, (A, B, C) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    ev = run_clock(eng, clock, st.empty(3, under=B), st.hold(A, 14, under=B), st.lift(A, B, 6), st.hold(B, 14, under=C),
                   st.lift(B, C, 6), st.hold(C, 14), st.lift(C, None, 6), st.empty(4), st.drop(A, None, 5), st.hold(A, 16))
    assert [u.reading.picks for _, u in emitted(ev)] == [PICKS, OTHER, THIRD]     # A is not entered a third time
    assert len([u for _, u in ev if u.duplicate]) >= 1
    assert len(eng.history) == 3 and eng.last_fingerprint is eng.history[-1]


def test_the_history_keeps_the_last_five_and_forget_last_clears_it():
    eng = CaptureEngine(9)
    assert eng.last_fingerprint is None and eng.history == []
    fake = [object() for _ in range(7)]
    for f in fake:
        eng._fps.append(f)
    assert eng.history == fake[-5:] and eng.last_fingerprint is fake[-1]
    eng.forget_last(rearm=False)
    assert eng.last_fingerprint is None and eng.history == []


def test_every_emitted_reading_carries_its_fingerprint():
    st, (A, B, _) = make_stage()
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    ev = run_clock(eng, clock, st.empty(3, under=B), st.hold(A, 14, under=B))
    (t, u), = emitted(ev)
    assert isinstance(u.fingerprint, sv.Fingerprint) and u.fingerprint is eng.last_fingerprint
    assert u.fingerprint.picks == tuple(PICKS)
    assert all(x.fingerprint is None for _, x in ev if x.reading is None)
    manual = eng.capture_now(next(st.hold(B, 1))[0])
    assert manual.fingerprint is not None and manual.fingerprint.picks == tuple(OTHER)
    assert eng.last_fingerprint is manual.fingerprint


# =============================================================================== the real photo
@pytest.fixture(scope="module")
def real():
    if not REAL_SLIP or not os.path.exists(REAL_SLIP):
        pytest.skip("set QUINIELA_REAL_SLIP to a real slip photo to run the real-photo checks")
    return cv2.imread(REAL_SLIP)


def _real_frame(img, angle=4.0, scale=0.5, gain=1.0, noise=0.0, tint=(1, 1, 1), seed=0):
    rng = np.random.default_rng(seed)
    s = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    h, w = s.shape[:2]
    pad = int(0.3 * max(h, w))
    big = cv2.copyMakeBorder(s, pad, pad, pad, pad, cv2.BORDER_REPLICATE)
    hh, ww = big.shape[:2]
    big = cv2.warpAffine(big, cv2.getRotationMatrix2D((ww / 2, hh / 2), angle, 1.0), (ww, hh), borderMode=cv2.BORDER_REPLICATE)
    canvas = np.full((1080, 1920, 3), int(np.median(s[:20, :20])), np.uint8)
    ox, oy = (1920 - ww) // 2, (1080 - hh) // 2
    canvas[max(oy, 0):max(oy, 0) + min(hh, 1080), max(ox, 0):max(ox, 0) + min(ww, 1920)] = \
        big[max(0, -oy):max(0, -oy) + min(hh, 1080), max(0, -ox):max(0, -ox) + min(ww, 1920)]
    f = canvas.astype(np.float32) * np.float32(tint)[None, None] * gain
    if noise:
        f = f + rng.normal(0, noise, f.shape)
    return np.clip(f, 0, 255).astype(np.uint8)


@pytest.mark.parametrize("kw", [dict(), dict(gain=0.6), dict(angle=15), dict(noise=6), dict(tint=(0.75, 0.9, 1.0)),
                                dict(tint=(0.6, 0.85, 1.0), gain=0.6), dict(tint=(0.5, 0.65, 0.95)), dict(tint=(1.0, 0.9, 0.75), gain=1.3)])
def test_the_real_photo_has_no_skin_in_any_light(real, kw):
    frame = _real_frame(real, **kw)
    rd = sv.read_slip(frame)
    assert rd.skin_fraction < 0.001 and not rd.hand and rd.occluded_rows == [] and rd.clean
    assert sv.hand_fraction(frame, sv.find_slip_quad(frame)) < 0.001


def test_a_hand_over_the_real_photos_marks_and_name(real):
    clean = _real_frame(real)
    g = sv.locate_grid(clean)
    truth = sv.read_slip(clean, grid=g)
    assert truth.picks == list("LVLLLLLLE") and truth.clean
    for (x, y, ang, size) in ((775, 560, 90, 0.8), (1000, 735, 0, 0.8), (960, 540, 200, 1.0)):
        rd = sv.read_slip(with_hand(clean, x, y, ang, size), grid=g)
        assert rd.hand
        for r in range(9):                                                   # whatever the hand touched is flagged, never a wrong pick
            assert rd.verdicts[r] == sv.OCCLUDED or (rd.picks[r] == truth.picks[r] and rd.verdicts[r] == sv.OK), (r, rd.verdicts[r])


def test_real_photo_with_a_hand_resting_on_it_is_entered_once(real):
    top = _real_frame(real, seed=1)
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    rng = np.random.default_rng(1)
    bank = [np.rint(rng.normal(0, 3, (top.shape[0], top.shape[1], 3))).astype(np.int16) for _ in range(3)]

    def noisy(f):
        return np.clip(f.astype(np.int16) * np.int16(1) + bank[int(rng.integers(3))], 0, 255).astype(np.uint8)
    seq = [noisy(top) for _ in range(14)]
    for (sx, sy), (gx, gy), ang in (((1860, 740), (960, 540), 200), ((-210, 580), (690, 520), 90)):
        for i in range(40):
            t = (i + 1) / 40
            seq.append(noisy(with_hand(top, sx + (gx - sx) * t, sy + (gy - sy) * t, ang, 1.7)))
        seq += [noisy(with_hand(top, gx, gy, ang, 1.7)) for _ in range(8)]
        seq += [noisy(top) for _ in range(12)]
    ev = run_clock(eng, clock, seq)
    assert len(emitted(ev)) == 1 and emitted(ev)[0][1].reading.picks == list("LVLLLLLLE")


def test_draw_hand_still_draws_a_hand():
    canvas = np.full((300, 300, 3), 235, np.uint8)
    draw_hand(canvas, 150, 150, 0.0, 0.5)
    assert (canvas != 235).any()
