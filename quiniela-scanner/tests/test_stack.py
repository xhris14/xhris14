"""Stacks of slips under a fixed camera: 'is it the same physical slip?' and the capture engine that
uses the answer to read each slip of a stack exactly once, hands-free."""
from __future__ import annotations

import os
import time

import cv2
import numpy as np
import pytest

from quiniela_scanner import slip_vision as sv
from quiniela_scanner.capture import IDLE, PAUSED, TRACKING, WAIT_REMOVE, CaptureEngine
from helpers import Stage, VClock, draw_slip_named, draw_unreadable_form, on_webcam_canvas

PICKS = list("LVLLELLVE")
OTHER = list("EEELLVVLV")
REAL_SLIP = os.environ.get("QUINIELA_REAL_SLIP", "")        # a real slip photo, kept out of the repo


# =============================================================================== fingerprints
def _fp(slip, **kw):
    kw = {"scale": 0.5, "angle": 4, "seed": 3, "noise": 3, **kw}
    return sv.fingerprint(sv.read_slip(on_webcam_canvas(slip, **kw)))


@pytest.fixture(scope="module")
def ref():
    return _fp(draw_slip_named(PICKS, "CLEMENTE O.", seed=7), angle=5, seed=1)


SAME_POSES = [
    dict(scale=0.5, angle=0),
    dict(scale=0.5, angle=12, seed=1),
    dict(scale=0.45, angle=-20, seed=2),
    dict(scale=0.4, angle=5, persp=0.03, seed=3),
    dict(scale=0.5, angle=3, blur=1.0, noise=6, seed=4),
    dict(scale=0.5, angle=4, flip=True),                       # put back upside down
    dict(scale=0.5, angle=8, bg=40, seed=5),                   # dark table
    dict(scale=0.55, angle=-3, persp=0.05, seed=6, noise=8, blur=1.5),
    dict(scale=0.35, angle=30, seed=9, noise=4),
]


@pytest.mark.parametrize("kw", SAME_POSES)
def test_same_slip_rephotographed_is_the_same(ref, kw):
    other = _fp(draw_slip_named(PICKS, "CLEMENTE O.", seed=7), **kw)
    marks, name = sv.slip_similarity(ref, other)
    assert sv.same_slip(ref, other), (marks, name)
    assert marks > 0.95 and name > 0.95


@pytest.mark.parametrize("name", ["JUAN PEREZ", "MARIA", "LUIS G.", "ANA L.", "JOSE", "PEDRO M.", "CARLOS"])
def test_marks_identical_to_the_pixel_are_one_slip_whatever_the_handwriting_says(ref, name):
    """E1: a hand over the name strip changes the handwriting score alone (measured on a real photo: marks 1.00,
    name 0.28), so marks that match as well as one slip's own two photos decide on their own. Two hand-drawn
    slips never match like that (real ones peak at ~0.76, see the jitter tests below)."""
    # draw_slip's default rng gives both slips the very same mark jitter: only the handwriting differs
    other = _fp(draw_slip_named(PICKS, name, seed=7), angle=-3, seed=2)
    marks, name_score = sv.slip_similarity(ref, other)
    assert marks > sv.MARKS_STRONG                              # the marks really are identical ...
    assert name_score < sv.NAME_SAME and sv.same_slip(ref, other)   # ... the handwriting is not, and it does not matter


@pytest.mark.parametrize("seed", range(20, 32))
def test_identical_picks_and_name_but_other_mark_jitter_is_different(ref, seed):
    other = _fp(draw_slip_named(PICKS, "CLEMENTE O.", seed=seed), angle=-3, seed=seed)
    marks, name_score = sv.slip_similarity(ref, other)
    assert name_score > 0.95                                    # same name text, pixel for pixel
    assert marks < sv.MARKS_SAME and not sv.same_slip(ref, other)


@pytest.mark.parametrize("i", range(6))
def test_identical_picks_other_name_other_jitter_is_different(ref, i):
    name = ["JUAN PEREZ", "MARIA", "LUIS G.", "ANA L.", "JOSE", "EL CHAPO"][i]
    other = _fp(draw_slip_named(PICKS, name, seed=50 + i), angle=2, seed=50 + i)
    assert not sv.same_slip(ref, other)


@pytest.mark.parametrize("row", range(9))
def test_a_single_different_mark_is_a_different_slip(ref, row):
    picks = list(PICKS)
    picks[row] = "LEV"[("LEV".index(picks[row]) + 1) % 3]
    other = _fp(draw_slip_named(picks, "CLEMENTE O.", seed=7), angle=1, seed=2)
    assert other.picks != ref.picks and not sv.same_slip(ref, other)


def test_double_marked_rows_keep_their_letters_in_the_fingerprint():
    lv, ev = list(PICKS), list(PICKS)
    lv[2], ev[2] = "LV", "EV"                                   # both rows read as 'ambiguous' ('')
    a = _fp(draw_slip_named(lv, "CLEMENTE O.", seed=7), angle=2, seed=2)
    b = _fp(draw_slip_named(ev, "CLEMENTE O.", seed=7), angle=-5, seed=4)
    a2 = _fp(draw_slip_named(lv, "CLEMENTE O.", seed=7), angle=-5, seed=4)
    assert not sv.same_slip(a, b)
    assert sv.same_slip(a, a2)


def test_blank_name_strip_is_not_enough_to_be_the_same_slip():
    blank = _fp(draw_slip_named(PICKS, "", seed=7), angle=5, seed=1)
    assert blank.name_blank and blank.name_px == 0              # no printed NOMBRE / ULTIMO DIA text leaks in
    again = _fp(draw_slip_named(PICKS, "", seed=7), angle=-8, seed=2, persp=0.03)
    assert sv.same_slip(blank, again)                           # same marks: the same slip
    for seed in (8, 9, 11, 12):                                 # same picks, other jitter: another slip
        other = _fp(draw_slip_named(PICKS, "", seed=seed), angle=-2, seed=seed)
        assert not sv.same_slip(blank, other)
    named = _fp(draw_slip_named(PICKS, "ANA", seed=8), angle=2, seed=2)       # other marks and a name: another slip
    assert not sv.same_slip(blank, named) and not sv.same_slip(named, blank)
    covered = _fp(draw_slip_named(PICKS, "ANA", seed=7), angle=2, seed=2)     # the very same marks, name hidden or not
    assert sv.same_slip(blank, covered) and sv.same_slip(covered, blank)
    other_pick = list(PICKS)
    other_pick[3] = "V"
    assert not sv.same_slip(blank, _fp(draw_slip_named(other_pick, "", seed=7), angle=2, seed=2))


def test_two_completely_blank_slips_cannot_be_told_apart_so_they_are_not_called_the_same():
    a = _fp(draw_slip_named([""] * 9, "", seed=7))
    b = _fp(draw_slip_named([""] * 9, "", seed=8), angle=3)
    assert a.marks_px == 0 and a.name_px == 0 and not sv.same_slip(a, b)


def test_name_strip_holds_handwriting_only(ref):
    assert ref.name_px > 500 and ref.marks_px > 500


def test_fingerprint_and_comparison_are_cheap(ref):
    rd = sv.read_slip(on_webcam_canvas(draw_slip_named(PICKS, "CLEMENTE O.", seed=7), scale=0.5, angle=5))
    sv.fingerprint(rd)
    t = time.perf_counter()
    for _ in range(10):
        fp = sv.fingerprint(rd)
    fp_ms = (time.perf_counter() - t) * 100
    t = time.perf_counter()
    for _ in range(10):
        sv.same_slip(ref, fp)
    cmp_ms = (time.perf_counter() - t) * 100
    assert fp_ms < 40 and cmp_ms < 40


# ----- the real photo (kept outside the repo: set QUINIELA_REAL_SLIP to its path)
def _real_variant(img, angle=0.0, scale=0.5, persp=0.0, gain=1.0, gamma=1.0, grad=0.0, noise=0.0, blur=0.0, seed=0,
                  flip=False):
    rng = np.random.default_rng(seed)
    if flip:
        img = cv2.rotate(img, cv2.ROTATE_180)
    s = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    h, w = s.shape[:2]
    pad = int(0.3 * max(h, w))
    big = cv2.copyMakeBorder(s, pad, pad, pad, pad, cv2.BORDER_REPLICATE)
    hh, ww = big.shape[:2]
    big = cv2.warpAffine(big, cv2.getRotationMatrix2D((ww / 2, hh / 2), angle, 1.0), (ww, hh), borderMode=cv2.BORDER_REPLICATE)
    if persp:
        src = np.float32([[0, 0], [ww, 0], [ww, hh], [0, hh]])
        dst = src + rng.uniform(-persp * min(ww, hh), persp * min(ww, hh), src.shape).astype(np.float32)
        big = cv2.warpPerspective(big, cv2.getPerspectiveTransform(src, dst), (ww, hh), borderMode=cv2.BORDER_REPLICATE)
    canvas = np.full((1080, 1920, 3), int(np.median(s[:20, :20])), np.uint8)
    ox, oy = (1920 - ww) // 2, (1080 - hh) // 2
    x0, y0 = max(0, -ox), max(0, -oy)
    canvas[max(oy, 0):max(oy, 0) + min(hh, 1080), max(ox, 0):max(ox, 0) + min(ww, 1920)] = \
        big[y0:y0 + min(hh, 1080), x0:x0 + min(ww, 1920)]
    f = canvas.astype(np.float32)
    if grad:
        f = f * np.linspace(1 - grad, 1 + grad, 1920, dtype=np.float32)[None, :, None]
    f = 255 * np.power(np.clip(f * gain / 255, 0, 1), gamma)
    if blur:
        f = cv2.GaussianBlur(f, (0, 0), blur)
    if noise:
        f = f + rng.normal(0, noise, f.shape)
    return np.clip(f, 0, 255).astype(np.uint8)


@pytest.fixture(scope="module")
def real():
    if not REAL_SLIP or not os.path.exists(REAL_SLIP):
        pytest.skip("set QUINIELA_REAL_SLIP to a real slip photo to run the real-photo checks")
    img = cv2.imread(REAL_SLIP)
    rd = sv.read_slip(_real_variant(img))
    return img, rd, sv.fingerprint(rd)


REAL_POSES = [
    dict(angle=15, scale=0.55, seed=2), dict(angle=-25, scale=0.45, seed=3), dict(angle=5, scale=0.5, persp=0.03, seed=4),
    dict(angle=3, scale=0.5, noise=6, blur=1.2, seed=5), dict(angle=0, scale=0.5, gain=0.7, seed=6),
    dict(angle=0, scale=0.5, gain=1.2, gamma=1.3, seed=7), dict(angle=8, scale=0.5, grad=0.25, seed=8),
    dict(angle=-10, scale=0.4, persp=0.05, noise=8, seed=9), dict(angle=4, scale=0.5, flip=True, seed=10),
    dict(angle=-6, scale=0.42, persp=0.04, noise=6, seed=13),
]


@pytest.mark.parametrize("kw", REAL_POSES)
def test_real_photo_and_a_transformed_copy_are_the_same_slip(real, kw):
    img, rd, fp = real
    other = sv.fingerprint(sv.read_slip(_real_variant(img, **kw)))
    assert other.picks == fp.picks
    assert sv.same_slip(fp, other), sv.slip_similarity(fp, other)


def _mutated(rd, seed, mark_shift=0, new_name=None):
    """The real slip's ink, but with every row's mark nudged by up to ``mark_shift`` px and/or a new name."""
    rng = np.random.default_rng(seed)
    fin, g = rd.rectified.copy(), rd.grid
    out = fin.copy()
    if mark_shift:
        x0, x1 = int(g.corners[:, 0, 0].mean()), int(g.corners[:, 3, 0].mean())
        for r in range(g.n_rows):
            y0, y1 = int(g.corners[r, 0, 1]) + 4, int(g.corners[r + 1, 0, 1]) - 4
            dx, dy = rng.integers(-mark_shift, mark_shift + 1, 2)
            patch = fin[y0:y1, x0 + 4:x1 - 4]
            out[y0:y1, x0 + 4:x1 - 4] = cv2.warpAffine(patch, np.float32([[1, 0, dx], [0, 1, dy]]),
                                                       (patch.shape[1], patch.shape[0]), borderMode=cv2.BORDER_REPLICATE)
    if new_name:
        nx0, nx1, ny0, ny1 = sv._name_window(fin, g)
        out[ny0:ny1, nx0:nx1] = np.median(fin.reshape(-1, 3), axis=0).astype(np.uint8)
        cv2.putText(out, new_name, (nx0 + 80 + int(rng.integers(-20, 20)), ny1 - 14), cv2.FONT_HERSHEY_SCRIPT_COMPLEX,
                    2.0, (170, 60, 25), 3)
    return sv.fingerprint(sv.SlipReading(picks=rd.picks, verdicts=rd.verdicts, ink=rd.ink, rectified=out,
                                         name_crop=rd.name_crop, title_crop=rd.title_crop, grid=rd.grid,
                                         candidates=rd.candidates))


@pytest.mark.parametrize("seed,shift,name", [(3, 6, None), (4, 7, None), (5, 6, "ELENA G."), (6, 8, "PACO")])
def test_real_ink_with_other_marks_is_a_different_slip(real, seed, shift, name):
    img, rd, fp = real
    assert not sv.same_slip(fp, _mutated(rd, seed, shift, name)), sv.slip_similarity(fp, _mutated(rd, seed, shift, name))


@pytest.mark.parametrize("seed,name", [(1, "JUAN P."), (2, "MARIA")])
def test_real_ink_with_only_the_name_changed_is_still_the_same_slip(real, seed, name):
    """The hand-over-the-name case on the real ink: marks identical, handwriting wiped / replaced."""
    img, rd, fp = real
    other = _mutated(rd, seed, 0, name)
    marks, name_score = sv.slip_similarity(fp, other)
    assert marks > sv.MARKS_STRONG and (name_score is None or name_score < sv.NAME_SAME)
    assert sv.same_slip(fp, other) and sv.same_slip(other, fp)


# =============================================================================== the engine on a stack
SPECS = [  # (picks, name, mark-jitter seed): identical picks with other names (the marks drawn anew each time), one single-mark change
    (PICKS, "CLEMENTE O.", 1), (PICKS, "MARIA G.", 2), (PICKS[:3] + ["V"] + PICKS[4:], "CLEMENTE O.", 3),
    (PICKS, "LUIS PEREZ", 4), (OTHER, "ANA", 5), (OTHER, "JOSE M.", 6),
]


def make_stage(seed=2, specs=SPECS, **kw):
    st = Stage(seed=seed, **kw)
    return st, [st.add(draw_slip_named(p, n, seed=s)) for p, n, s in specs]


def run(eng, *gens):
    """Feed every frame of the generators to the engine; -> [(frame number, info, update)]."""
    out, n = [], 0
    for gen in gens:
        for frame, info in gen:
            out.append((n, info, eng.feed(frame)))
            n += 1
    return out


def reads(ev):
    return [(n, info, u) for n, info, u in ev if u.reading is not None]


def dups(ev):
    return [(n, info, u) for n, info, u in ev if u.duplicate]


@pytest.fixture(scope="module")
def stack_run():
    """Six slips, each held 0.7 s (noise sigma 3, +-3% flicker, +-1.5% focus breathing) then lifted by a hand."""
    st, ids = make_stage()
    eng = CaptureEngine(9)
    gens = [st.empty(4)]
    for k, i in enumerate(ids):
        under = ids[k + 1] if k + 1 < len(ids) else None
        gens += [st.hold(i, 21, under=under), st.lift(i, under, 6)]
    return eng, ids, run(eng, *gens)


def test_every_slip_of_a_stack_is_read_exactly_once_in_order(stack_run):
    eng, ids, ev = stack_run
    got = reads(ev)
    assert [info["top"] for _, info, _ in got] == ids
    assert [u.reading.picks for _, _, u in got] == [p for p, _, _ in SPECS]
    assert not dups(ev) and not [u for _, _, u in ev if u.error]
    assert all(info["phase"] == "hold" for _, info, _ in got)        # never mid-lift, never from the hand


def test_next_slip_is_read_within_stable_frames_plus_three_of_settling(stack_run):
    eng, ids, ev = stack_run
    for n, info, u in reads(ev):
        first = next(m for m, i2, _ in ev if i2["phase"] == "hold" and i2["top"] == info["top"])
        # the slip has been sitting still since the first 'hold' frame (the lift tail already shows it clean)
        assert n - first <= eng.stable_frames + 3, (info["top"], n - first)


def test_engine_re_arms_during_every_lift_and_is_never_stuck(stack_run):
    eng, ids, ev = stack_run
    for k in range(len(ids) - 1):
        lift = [(n, u) for n, info, u in ev if info["phase"] == "lift" and info["top"] == ids[k]]
        assert any(u.state != WAIT_REMOVE for _, u in lift)           # left WAIT_REMOVE within the 6 lift frames
    assert eng.state == IDLE or eng.state == TRACKING


def test_reads_carry_timings(stack_run):
    _, _, ev = stack_run
    for _, _, u in reads(ev):
        assert u.timings["read_ms"] > 0 and u.timings["fingerprint_ms"] >= 0 and u.timings["feed_ms"] >= u.timings["read_ms"]
    assert all("read_ms" in u.timings and "feed_ms" in u.timings for _, _, u in ev)


def test_feed_stays_cheap_on_full_hd_frames(stack_run):
    """Frames that read nothing (a read is ~190 ms, a re-check of the slip waiting for its lift ~90 ms, both reported)."""
    _, _, ev = stack_run
    ms = [u.timings["feed_ms"] for _, _, u in ev if u.reading is None and not u.duplicate and not u.error
          and "verify_ms" not in u.timings]
    assert np.mean(ms) < 30 and np.percentile(ms, 95) < 60, (np.mean(ms), np.percentile(ms, 95))     # measured: ~6 ms, p95 ~20
    checks = [u.timings["verify_ms"] for _, _, u in ev if "verify_ms" in u.timings]
    assert len(checks) <= 2 * len(SPECS) and np.mean(checks) < 200, (len(checks), np.mean(checks))


def test_hand_only_frames_are_never_read():
    st, ids = make_stage()
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(3), st.hand_pass(10), st.empty(3), st.hand_pass(8), st.empty(3))
    assert not reads(ev) and eng.state == IDLE


def test_hand_passing_over_a_captured_slip_does_not_enter_it_again():
    st, ids = make_stage()
    A = ids[0]
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(2), st.hold(A, 12), st.hand_pass(10, top=A), st.hold(A, 14), st.hand_pass(10, top=A), st.hold(A, 12))
    assert [info["top"] for _, info, _ in reads(ev)] == [A]


def test_lifted_and_put_back_is_the_same_slip_and_not_entered_twice():
    st, ids = make_stage()
    A, B = ids[:2]
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(3), st.hold(A, 12), st.lift(A, None, 6), st.empty(4), st.drop(A, None, 5), st.hold(A, 14),
             st.lift(A, B, 6), st.hold(B, 14))
    assert [info["top"] for _, info, _ in reads(ev)] == [A, B]
    d = dups(ev)
    assert len(d) == 1 and "same slip" in d[0][2].message.lower() and d[0][2].state == WAIT_REMOVE
    assert d[0][2].reading is None and d[0][2].error is None


def test_a_slip_that_never_sits_still_is_never_read():
    st, ids = make_stage()
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(3), st.slide(ids[0], 40, 20))
    assert not reads(ev)
    assert {u.state for _, _, u in ev[3:]} <= {TRACKING, IDLE}


def test_still_slip_under_flicker_noise_and_focus_breathing_never_re_arms():
    for kw in (dict(noise=3, flicker=0.03, breath=0.015), dict(noise=6, flicker=0.06, breath=0.03)):
        st, ids = make_stage(seed=5, **kw)
        eng = CaptureEngine(9)
        ev = run(eng, st.empty(2), st.hold(ids[0], 75))
        assert len(reads(ev)) == 1 and not dups(ev) and eng.state == WAIT_REMOVE, kw


def test_unreadable_form_is_reported_and_never_wedges_the_engine():
    st = Stage(seed=4)
    form = st.add(draw_unreadable_form())
    a, b = (st.add(draw_slip_named(PICKS, n, seed=s)) for n, s in (("ANA", 1), ("LUIS", 2)))
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(2), st.hold(form, 20), st.lift(form, None, 6), st.hold(a, 12))
    errs = [u for _, _, u in ev if u.error]
    assert len(errs) == 1 and errs[0].reading is None and errs[0].state == WAIT_REMOVE
    assert [info["top"] for _, info, _ in reads(ev)] == [a]
    # the form is swapped for a slip without the table ever being empty: the stack advances
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(2), st.hold(form, 20, under=a), st.lift(form, a, 6), st.hold(a, 12), st.lift(a, b, 6), st.hold(b, 12))
    assert [info["top"] for _, info, _ in reads(ev)] == [a, b]


def test_blank_slip_is_read_like_any_other_and_does_not_block_the_next():
    st = Stage(seed=6)
    blank = st.add(draw_slip_named([""] * 9, "", seed=1))
    a = st.add(draw_slip_named(PICKS, "ANA", seed=2))
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(2), st.hold(blank, 14, under=a), st.lift(blank, a, 6), st.hold(a, 14))
    got = reads(ev)
    assert [info["top"] for _, info, _ in got] == [blank, a] and got[0][2].reading.picks == [""] * 9


def test_unreadable_slip_that_just_lies_there_is_retried(monkeypatch):
    st, ids = make_stage()
    calls = []
    real_read = sv.read_slip

    def flaky(frame, *a, **k):
        calls.append(1)
        if len(calls) == 1:
            raise sv.SlipNotFound("out of focus")
        return real_read(frame, *a, **k)
    monkeypatch.setattr(sv, "read_slip", flaky)
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(2), st.hold(ids[0], 50))
    assert [bool(u.error) for _, _, u in ev if u.error or u.reading is not None] == [True, False]
    assert len(reads(ev)) == 1


def test_slip_not_found_keeps_the_old_contract(monkeypatch):
    st, ids = make_stage()
    monkeypatch.setattr(sv, "read_slip", lambda *a, **k: (_ for _ in ()).throw(sv.SlipNotFound("no grid")))
    eng = CaptureEngine(9, retry_unreadable=0)
    ev = run(eng, st.empty(2), st.hold(ids[0], 12), st.lift(ids[0], None, 6), st.empty(3))
    errs = [u for _, _, u in ev if u.error]
    assert len(errs) == 1 and errs[0].error == "no grid" and errs[0].reading is None
    assert eng.state == IDLE                                         # re-armed by the lift, not stuck


def test_manual_capture_bypasses_the_duplicate_check_and_forget_last_clears_it():
    st, ids = make_stage()
    A, B = ids[:2]
    eng = CaptureEngine(9)
    frame_a = next(st.hold(A, 1))[0]
    ev = run(eng, st.empty(2), st.hold(A, 12))
    assert len(reads(ev)) == 1
    manual = eng.capture_now(frame_a)                                # the button always scans, even a repeat
    assert manual.reading is not None and not manual.duplicate and manual.state == WAIT_REMOVE
    # lift A and put it back: still the slip entered last (the manual capture counts)
    ev = run(eng, st.lift(A, None, 6), st.empty(3), st.drop(A, None, 5), st.hold(A, 12))
    assert not reads(ev) and len(dups(ev)) == 1
    # a deliberate re-scan: forget, then the slip that is still lying there is read again at once
    eng.forget_last()
    assert eng.last_fingerprint is None and eng.state == IDLE
    ev = run(eng, st.hold(A, 12))
    assert len(reads(ev)) == 1 and not dups(ev)
    # and a different slip afterwards is of course read too
    ev = run(eng, st.lift(A, B, 6), st.hold(B, 12))
    assert [info["top"] for _, info, _ in reads(ev)] == [B]


def test_dedupe_can_be_switched_off():
    st, ids = make_stage()
    A = ids[0]
    eng = CaptureEngine(9, dedupe=False)
    ev = run(eng, st.empty(2), st.hold(A, 12), st.lift(A, None, 6), st.empty(3), st.drop(A, None, 5), st.hold(A, 12))
    assert len(reads(ev)) == 2


def test_paused_engine_reads_nothing_but_keeps_watching_so_the_next_slip_is_ready():
    st, ids = make_stage()
    A, B = ids[:2]
    eng = CaptureEngine(9)
    ev = run(eng, st.empty(2), st.hold(A, 12))
    assert len(reads(ev)) == 1
    eng.paused = True                                                # the card of A is open
    ev = run(eng, st.hold(A, 6, under=B), st.lift(A, B, 6), st.hold(B, 12))
    assert all(u.state == PAUSED and u.reading is None for _, _, u in ev)
    eng.paused = False
    ev = run(eng, st.hold(B, 12))
    assert [info["top"] for _, info, _ in reads(ev)] == [B]          # B was lying there, armed, when resumed


def test_wait_focus_can_be_skipped_and_is_capped():
    st, ids = make_stage(noise=0, flicker=0, breath=0)
    A = ids[0]
    n_fast = next(n for n, _, u in run(CaptureEngine(9, wait_focus=False), st.empty(1), st.hold(A, 20)) if u.reading is not None)
    assert n_fast == 1 + 4                                           # first frame + stable_frames still frames

    # a focus that never converges (sharpness keeps changing by >15% per frame) must not hold the read back
    base = next(st.hold(A, 1))[0]

    def hunting():
        for i in range(60):
            yield cv2.GaussianBlur(base, (0, 0), 0.2 + 1.6 * (i % 2)), dict(phase="hold", top=A, i=i)
    ev = run(CaptureEngine(9, focus_cap_frames=12), hunting())
    n_capped = next(n for n, _, u in ev if u.reading is not None)
    assert 4 < n_capped <= 1 + 4 + 12 + 1
    assert any(u.message == "Focusing..." for _, _, u in ev[:n_capped])


def test_defaults_and_backwards_compatible_constructor():
    eng = CaptureEngine()
    assert eng.stable_frames == 4 and eng.wait_focus and eng.dedupe and eng.n_rows == 9
    old_style = CaptureEngine(9, 7, sv.Settings(min_ink=0.02))
    assert old_style.stable_frames == 7 and old_style.settings.min_ink == 0.02
    assert CaptureEngine(9, stable_frames=1).stable_frames == 3


def test_a_fast_partial_lift_still_re_arms_the_engine():
    """The user lifts each slip so fast that only 4 of the 7 lift frames reach the camera and the rest of the
    picture is just the next slip: one very strong frame must be enough to re-arm (regression)."""
    specs = [(list("LVLLELLVE"), "CLEMENTE O.", 1), (list("EEELLVVLV"), "ANA", 5), (list("VVVLLLEEE"), "BETO", 6),
             (list("LLLEEEVVV"), "CARLOS M.", 7), (list("LEVLEVLEV"), "DOLORES", 8), (list("VELVELVEL"), "PACO 2", 9),
             (list("LLVVEELLV"), "JUAN G.", 10), (list("EVLEVLEVL"), "TONY 5000", 11)]
    st = Stage(seed=2)
    ids = [st.add(draw_slip_named(p, n, seed=s)) for p, n, s in specs]
    frames = [f for f, _ in st.empty(4)]
    for k, i in enumerate(ids):
        under = ids[k + 1] if k + 1 < len(ids) else None
        hold = [f for f, _ in st.hold(i, 8, under=under)]
        lift = [f for f, _ in st.lift(i, under, 7)]
        frames += [hold[n % 8] for n in range(14)] + [lift[j] for j in (1, 3, 5, 6)]
    frames += [f for f, _ in st.empty(10)]
    eng = CaptureEngine(9)
    got = [u.reading for u in map(eng.feed, frames) if u.reading is not None]
    assert [r.picks for r in got] == [list(p) for p, _, _ in specs]


# =============================================================================== SOAK-4: slips missed at low frame rates / after a hand sweep
# A lift that is too short for the camera, or that falls inside a read, leaves the next slip - lying in the same place -
# looking like the last one to the frame differences (0.3-0.8% of the thumbnail against LO = 1%): it was never read, and
# nothing said so. The engine now re-checks, while it waits for the lift, that the slip there is still the one it entered.
FILM_SPECS = [(list("LVLLELLVE"), "CLEMENTE O.", 1), (list("EEELLVVLV"), "ANA", 5), (list("VVVLLLEEE"), "BETO", 6),
              (list("LLLEEEVVV"), "CARLOS M.", 7), (list("LEVLEVLEV"), "DOLORES", 8), (list("VELVELVEL"), "PACO 2", 9),
              (list("LLVVEELLV"), "JUAN G.", 10), (list("EVLEVLEVL"), "TONY 5000", 11), (list("ELVELVELV"), "RITA", 12),
              (list("VLEVLEVLE"), "OMAR", 13)]


class Film:
    """A 30 fps film of a stack: every slip lies still ``hold`` frames (optionally a hand sweeps over it after it was read, and
    rests ``hand_pause`` frames) and is then lifted in ``lift`` frames, revealing the next one."""

    def __init__(self, n=10, lift=4, hold=30, hand_over=(), hand_pause=4, size=None, dim=0.0, noise=0.0, seed=2):
        st = Stage(seed=seed)
        ids = [st.add(draw_slip_named(p, nm, seed=s)) for p, nm, s in FILM_SPECS[:n]]
        self.picks = [tuple(p) for p, _, _ in FILM_SPECS[:n]]
        frames = [f for f, _ in st.empty(4)]
        for k, i in enumerate(ids):
            under = ids[k + 1] if k + 1 < len(ids) else None
            bank = [f for f, _ in st.hold(i, 8, under=under)]
            frames += [bank[j % 8] for j in range(hold)]
            if k in hand_over:
                frames += [f for f, _ in st.hand_pass(6, top=i, under=under)] + [bank[j % 8] for j in range(hand_pause)]
            frames += [f for f, _ in st.lift(i, under, lift)]
        frames += [f for f, _ in st.empty(10)]
        rng = np.random.default_rng(3)
        bank = None
        self.frames = []
        for f in frames:
            if size:
                f = cv2.resize(f, size, interpolation=cv2.INTER_AREA)
            if dim:
                f = cv2.convertScaleAbs(f, alpha=1 - dim)
            if noise:
                if bank is None:                              # a few noise fields, each frame takes a random crop of one
                    bank = [np.rint(rng.normal(0, noise, (f.shape[0] + 32, f.shape[1] + 32, 3))).astype(np.int16) for _ in range(3)]
                oy, ox = rng.integers(0, 32, 2)
                f = np.clip(f.astype(np.int16) + bank[int(rng.integers(3))][oy:oy + f.shape[0], ox:ox + f.shape[1]], 0, 255).astype(np.uint8)
            self.frames.append(f)


def play(film, fps=30, read_s=0.19, verify_s=0.09, frame_s=0.02, **kw):
    """A camera of ``fps`` frames per second (the film at 30 fps, every (30/fps)-th frame): the engine takes the newest frame
    whenever it is free, a read stalls it for ``read_s`` and a re-check for ``verify_s`` (virtual time, so the outcome does not
    depend on the machine). -> (readings, duplicates, errors, number of re-checks)."""
    clock = VClock()
    eng = CaptureEngine(9, clock=clock, **kw)
    stride = 30.0 / fps
    i, readings, dups, errs, checks = 0.0, [], 0, 0, 0
    while int(round(i)) < len(film.frames):
        idx = int(round(i))
        clock.t = idx / 30.0
        u = eng.feed(film.frames[idx])
        cost = frame_s
        if u.reading is not None or u.duplicate or u.error:
            cost += read_s
        if "verify_ms" in u.timings:
            checks += 1
            cost += verify_s
        if u.reading is not None:
            readings.append(tuple(u.reading.picks))
        dups += u.duplicate
        errs += bool(u.error)
        i = max(idx + stride, np.floor((idx / 30.0 + cost) * 30.0 / stride) * stride)
    return readings, dups, errs, checks


def lost_and_repeated(film, readings):
    lost = [k for k, p in enumerate(film.picks) if p not in readings]
    repeated = [k for k, p in enumerate(film.picks) if readings.count(p) > 1]
    stray = [r for r in readings if r not in film.picks]
    return lost, repeated, stray


@pytest.fixture(scope="module")
def film():
    return Film(n=8, lift=4, hold=30)


@pytest.mark.parametrize("fps", [15, 10, 7.5])
def test_every_slip_is_read_exactly_once_whatever_the_camera_speed(film, fps):
    """Lifts of 4 frames (130 ms) at a camera of 15 / 10 / 7.5 fps: the old engine lost 0 / 7 of 12 / 6 of 12 slips here (6 and 29
    of 60 in the full soak at 15 and 10 fps): the next slip, lying in the same place, never looked 'different' to it."""
    readings, dups, errs, checks = play(film, fps)
    lost, repeated, stray = lost_and_repeated(film, readings)
    assert not lost and not repeated and not stray and not errs, (lost, repeated, stray, errs)
    assert dups == 0                                          # and the re-checks raised no false alarm


def test_without_the_re_check_the_slow_camera_loses_slips(film):
    """The scenario is the hard one it claims to be: with verify=False (the old behaviour) slips go missing at 10 fps."""
    readings, dups, errs, checks = play(film, 10, verify=False)
    lost, repeated, stray = lost_and_repeated(film, readings)
    assert len(lost) >= 2 and checks == 0 and not repeated


def test_dim_light_at_15_fps():
    """Gain 0.5 + noise 6 (the old engine lost 11 of 12 slips here, 22 of 60 in the soak: the lift hides in the noise)."""
    f = Film(n=5, lift=4, hold=30, dim=0.5, noise=6)
    lost, repeated, stray = lost_and_repeated(f, play(f, 15)[0])
    assert not lost and not repeated and not stray
    lost_old, *_ = lost_and_repeated(f, play(f, 15, verify=False)[0])
    assert len(lost_old) >= 2


def test_a_hand_sweeping_over_half_of_the_slips_after_they_are_read_costs_none():
    """The sweep re-arms the engine, which re-reads the slip just as the lift happens (the stall hides it): the old engine
    never read the next slip of 5 of 12 (14 of 80 in the soak)."""
    f = Film(n=6, lift=5, hold=30, hand_over=(0, 2, 4), hand_pause=5)
    readings, dups, errs, checks = play(f, 30)
    lost, repeated, stray = lost_and_repeated(f, readings)
    assert not lost and not repeated and not stray and not errs
    assert dups >= 3                                          # the re-reads after the sweeps were recognised, not entered
    lost_old, *_ = lost_and_repeated(f, play(f, 30, verify=False)[0])
    assert len(lost_old) >= 1


def test_1280x720_makes_no_difference():
    f = Film(n=5, lift=4, hold=30, size=(1280, 720))
    lost, repeated, stray = lost_and_repeated(f, play(f, 15)[0])
    assert not lost and not repeated and not stray


def test_the_re_check_costs_one_short_stall_per_slip(film):
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    ms, n_reads = [], 0
    for k, frame in enumerate(film.frames):
        clock.t = k / 30.0
        u = eng.feed(frame)
        if "verify_ms" in u.timings:
            ms.append(u.timings["verify_ms"])
        n_reads += u.reading is not None
    assert n_reads == 8
    per_slip = len(ms) / 8
    assert per_slip <= 3 and sum(ms) / 8 < 250, (per_slip, sum(ms) / 8)         # at most ~0.25 s of re-checking per slip
    assert max(ms) < 250


def test_a_single_odd_look_does_not_make_a_new_slip(monkeypatch):
    """Hysteresis: a quick look that disagrees with the slip entered is only a candidate; the full read that follows must
    agree with it AND differ from the history. Here the quick look is wrong (shows another slip): nothing is entered."""
    st, ids = make_stage()
    A, B = ids[:2]
    other = sv.fingerprint(sv.read_slip(next(st.hold(B, 1))[0]))
    clock = VClock()
    eng = CaptureEngine(9, clock=clock)
    monkeypatch.setattr(sv, "quick_fingerprint", lambda frame, grid, settings=None, template=None: (other, 0.0))
    ev = []
    for k, (frame, info) in enumerate(list(st.empty(3)) + list(st.hold(A, 120))):
        clock.t = k / 30.0
        ev.append((k, info, eng.feed(frame)))
    assert [info["top"] for _, info, _ in reads(ev)] == [A]            # only the first read
    assert len(eng.history) == 1 and not [u for _, _, u in ev if u.error]
    assert eng.state == WAIT_REMOVE
