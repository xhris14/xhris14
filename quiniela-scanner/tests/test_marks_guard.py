"""Marks that must not be entered silently (review findings E3, E6, E7, E8): a crossed-out mark beside a
corrected one, a faint pencil mark, a staple lying in an empty row, one X drawn on the rule between two rows -
and, as important, that ordinary marks of every size and weight still read clean."""
from __future__ import annotations

import os

import cv2
import numpy as np
import pytest

from quiniela_scanner import slip_vision as sv
from helpers import draw_cross, draw_scribble, draw_slip, draw_staple, on_webcam_canvas

REAL_SLIP = os.environ.get("QUINIELA_REAL_SLIP", "")        # a real slip photo, kept out of the repo


def read(img, seed=0, **kw):
    kw = {"scale": 0.5, "angle": 4, "noise": 3, **kw}
    return sv.read_slip(on_webcam_canvas(img, seed=seed, **kw))


def blank_slip(name="ANA"):
    return draw_slip([""] * 9, name=name)


# ===================================================================== E3: crossed out + corrected
@pytest.mark.parametrize("strokes,thick,size,th", [(6, 5, 30, 3), (8, 4, 32, 3), (6, 6, 28, 3), (6, 4, 32, 3)])
def test_crossed_out_mark_beside_a_corrected_one_is_flagged(strokes, thick, size, th):
    """Scribbled-out L (ink 0.28-0.35) and a small new X in V (0.11): the ratio is under multi_ratio and the
    scribble under the old heavy_ink, so the row used to read 'L, ok' and go in silently."""
    img = blank_slip()
    draw_cross(img, 3, 0)
    draw_scribble(img, 3, 0, strokes, thick)
    draw_cross(img, 3, 2, size, th)
    rd = read(img)
    s = sv.Settings()
    ink = rd.ink[3]
    assert ink[2] / ink[0] < s.multi_ratio and ink[0] < 0.40, ink          # the case the old rules let through
    assert rd.verdicts[3] in (sv.MULTIPLE, sv.HEAVY) and not rd.clean and 3 in rd.problem_rows
    if rd.verdicts[3] == sv.MULTIPLE:
        assert rd.picks[3] == ""


def test_the_old_blind_spot_is_closed_by_the_absolute_second_cell_rule():
    img = blank_slip()
    draw_cross(img, 3, 0)
    draw_scribble(img, 3, 0, 6, 5)
    draw_cross(img, 3, 2, 30, 3)
    # the same reading judged by the old rule set (relative ratio only, heavy at 0.40) would be a clean 'L'
    old = sv.Settings(second_ink=1.0, heavy_ink=0.40)
    assert read(img).verdicts[3] != sv.OK
    assert sv.read_slip(on_webcam_canvas(img, scale=0.5, angle=4, noise=3), 9, old).verdicts[3] == sv.OK


@pytest.mark.parametrize("size", [30, 48, 60])
@pytest.mark.parametrize("thick", [2, 4, 6])
def test_ordinary_marks_of_every_size_and_weight_stay_clean(size, thick):
    img = blank_slip()
    for r in range(9):
        draw_cross(img, r, r % 3, size, thick, dx=(r % 3 - 1) * 3, dy=(r % 2) * 4 - 2)
    rd = read(img, seed=size + thick)
    assert rd.picks == ["LEV"[r % 3] for r in range(9)] and rd.clean, (rd.verdicts, rd.ink.max(axis=1))


# ===================================================================== E6: faint pencil / dry pen
@pytest.mark.parametrize("contrast", [0.06, 0.07, 0.09])
def test_a_faint_pencil_mark_is_weak_not_blank(contrast):
    """15-22 grey levels below the paper: the main pass sees ~nothing (blank, entered as F1); the faint pass
    shows it, with the right letter as a guess, and the row goes to the card."""
    v = int(238 * (1 - contrast))
    img = blank_slip()
    rows = (0, 2, 4, 6)
    for r in rows:
        draw_cross(img, r, (r // 2) % 3, 48, 4, ink=(v, v, v))
    rd = read(img)
    for r in rows:
        assert rd.verdicts[r] == sv.WEAK and rd.picks[r] == "LEV"[(r // 2) % 3], (r, rd.verdicts[r], rd.ink[r])
        assert max(rd.ink[r]) < sv.Settings().min_ink          # the main pass really did not see it
    assert [rd.verdicts[r] for r in (1, 3, 5, 7, 8)] == [sv.BLANK] * 5
    assert not rd.clean


@pytest.mark.parametrize("kw", [
    dict(noise=0), dict(noise=6), dict(noise=6, blur=0.8), dict(noise=3, blur=0.8, scale=0.4),
    dict(noise=6, blur=0.8, scale=0.4, angle=-8), dict(noise=6, persp=0.03, seed=2), dict(noise=6, bg=60),
])
def test_blank_slips_stay_blank_under_noise_and_blur(kw):
    rd = read(blank_slip(), **kw)
    assert rd.verdicts == [sv.BLANK] * 9, rd.verdicts


# ===================================================================== E7: staples
@pytest.mark.parametrize("length,angle", [(70, 20), (70, 160), (90, 45), (80, 70)])
@pytest.mark.parametrize("col", [0, 1, 2])
def test_a_long_thin_dark_object_in_an_empty_row_is_not_a_clean_pick(length, angle, col):
    img = blank_slip()
    draw_staple(img, 4, col, length, angle)
    rd = read(img)
    assert rd.verdicts[4] == sv.WEAK and not rd.clean, (rd.verdicts[4], rd.ink[4])
    assert [v for i, v in enumerate(rd.verdicts) if i != 4] == [sv.BLANK] * 8


# ===================================================================== E8: an X on the rule
@pytest.mark.parametrize("size", [50, 60])
def test_one_x_on_the_rule_between_two_rows_is_not_two_clean_picks(size):
    img = blank_slip()
    draw_cross(img, 3, 0, size, 5, dy=40)                    # the rule is 40 px below the row-3 centre
    rd = read(img)
    assert rd.verdicts[3] != sv.OK and rd.verdicts[4] != sv.OK and not rd.clean, (rd.verdicts[3:5], rd.ink[3:5])


@pytest.mark.parametrize("dy", [-14, 0, 14])
def test_an_x_a_bit_off_centre_in_its_row_is_still_clean(dy):
    img = blank_slip()
    draw_cross(img, 3, 0, 60, 5, dy=dy)
    rd = read(img)
    assert rd.verdicts[3] == sv.OK and rd.picks[3] == "L"


# ===================================================================== the real photo must not notice any of it
@pytest.fixture(scope="module")
def real():
    if not REAL_SLIP or not os.path.exists(REAL_SLIP):
        pytest.skip("set QUINIELA_REAL_SLIP to a real slip photo to run the real-photo checks")
    return cv2.imread(REAL_SLIP)


def _canvas(img, scale=0.5, angle=0.0, gain=1.0, noise=0.0, blur=0.0, seed=0):
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
    f = canvas.astype(np.float32) * gain
    if blur:
        f = cv2.GaussianBlur(f, (0, 0), blur)
    if noise:
        f = f + rng.normal(0, noise, f.shape)
    return np.clip(f, 0, 255).astype(np.uint8)


@pytest.mark.parametrize("kw", [dict(), dict(angle=15), dict(gain=0.7), dict(noise=6, blur=1.2), dict(scale=0.35),
                                dict(scale=0.7), dict(gain=1.3, noise=4)])
def test_the_real_photos_nine_x_marks_still_read_clean(real, kw):
    rd = sv.read_slip(_canvas(real, **kw))
    assert rd.picks == list("LVLLLLLLE") and rd.clean, (rd.verdicts, rd.ink.max(axis=1))
    assert 0.10 < rd.ink.max(axis=1).min() and rd.ink.max(axis=1).max() < 0.30      # real X marks: 0.12-0.2 of a cell
