import numpy as np
import pytest

from quiniela_scanner import slip_vision as sv
from helpers import draw_slip, on_webcam_canvas

PICKS = list("LVLLELLVE")


def _read(frame):
    return sv.read_slip(frame)


@pytest.mark.parametrize("kw", [
    dict(scale=0.5, angle=0),
    dict(scale=0.5, angle=12, seed=1),
    dict(scale=0.45, angle=-20, seed=2),
    dict(scale=0.4, angle=5, persp=0.03, seed=3),
    dict(scale=0.5, angle=3, blur=1.0, noise=6, seed=4),
    dict(scale=0.5, angle=4, flip=True),                      # slip upside down
    dict(scale=0.5, angle=8, bg=40, seed=5),                  # dark table
])
def test_reads_all_nine_picks_under_webcam_conditions(kw):
    frame = on_webcam_canvas(draw_slip(PICKS, rng=np.random.default_rng(7)), **kw)
    rd = _read(frame)
    assert rd.picks == PICKS
    assert rd.clean
    assert rd.rectified.shape[1] == 1000


def test_black_pen_works_too():
    frame = on_webcam_canvas(draw_slip(PICKS, ink=(20, 20, 20), rng=np.random.default_rng(3)), scale=0.5, angle=6)
    assert _read(frame).picks == PICKS


def test_blank_and_double_marks_are_flagged_not_guessed():
    picks = ["L", "", "LV", "E", "V", "", "L", "E", "V"]
    rd = _read(on_webcam_canvas(draw_slip(picks, rng=np.random.default_rng(2)), scale=0.5, angle=4))
    assert rd.verdicts[1] == sv.BLANK and rd.verdicts[5] == sv.BLANK
    assert rd.verdicts[2] == sv.MULTIPLE
    assert rd.picks[1] == "" and rd.picks[2] == "" and rd.picks[5] == ""
    assert rd.candidates[2] == ("L", "V")
    assert rd.problem_rows == [1, 2, 5]
    assert not rd.clean
    assert [rd.picks[i] for i in (0, 3, 4, 6, 7, 8)] == ["L", "E", "V", "L", "E", "V"]


def test_name_and_title_crops_are_the_right_strips():
    rd = _read(on_webcam_canvas(draw_slip(PICKS), scale=0.5, angle=2))
    nh, nw = rd.name_crop.shape[:2]
    assert nw == 1000 and 80 < nh < 260
    th, tw = rd.title_crop.shape[:2]
    assert tw == 1000 and th < 200
    # the handwriting strip must actually contain blue-ish ink
    b, g, r = [rd.name_crop[..., i].astype(int) for i in range(3)]
    assert ((b - r) > 60).sum() > 500


@pytest.mark.parametrize("frame", [
    np.full((1080, 1920, 3), 230, np.uint8),
    np.random.default_rng(0).integers(0, 255, (1080, 1920, 3), dtype=np.uint8),
])
def test_empty_frames_find_no_slip(frame):
    assert sv.find_slip_quad(frame) is None
    with pytest.raises(sv.SlipNotFound):
        sv.locate_grid(frame)


def test_unrelated_printed_table_is_rejected():
    import cv2
    img = np.full((1080, 1920, 3), 235, np.uint8)
    for i in range(6):
        cv2.line(img, (500, 300 + i * 80), (1400, 300 + i * 80), (30, 30, 30), 3)
    for x in (500, 800, 1400):
        cv2.line(img, (x, 300), (x, 700), (30, 30, 30), 3)
    with pytest.raises(sv.SlipNotFound):
        sv.locate_grid(img)


def test_occluded_is_a_problem_row_and_is_drawn():
    """E2: the OCCLUDED verdict is part of the vocabulary - never clean, never a pick - and draw_reading shows it."""
    import dataclasses
    rd = _read(on_webcam_canvas(draw_slip(PICKS, rng=np.random.default_rng(7)), scale=0.5, angle=4))
    assert sv.OCCLUDED == "occluded" and rd.clean and rd.occluded_rows == []
    covered = dataclasses.replace(rd, picks=["", ""] + rd.picks[2:], verdicts=[sv.OCCLUDED, sv.OCCLUDED] + rd.verdicts[2:],
                                  candidates=[(), ()] + rd.candidates[2:])
    assert not covered.clean and covered.occluded_rows == [0, 1] and covered.problem_rows == [0, 1]
    img = sv.draw_reading(covered)
    assert img.shape == rd.rectified.shape and not (img == sv.draw_reading(rd)).all()


def test_read_slip_can_reuse_a_grid_and_skin_defaults_are_off_for_a_clean_slip():
    frame = on_webcam_canvas(draw_slip(PICKS, rng=np.random.default_rng(7)), scale=0.5, angle=4)
    g = sv.locate_grid(frame)
    rd = sv.read_slip(frame, grid=g)
    assert rd.picks == PICKS and rd.grid is g
    assert rd.skin_fraction == 0.0 and not rd.hand and not rd.name_hidden
