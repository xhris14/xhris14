"""Computer vision for Liga MX quiniela slips.

Pipeline
--------
1. ``find_slip_quad``  - find the printed table (the slip) in a camera frame.
2. ``locate_grid``     - flatten it, detect the ruled lines and work out where the
                         9 match rows and the L / E / V columns are.
3. ``read_slip``       - measure pen ink inside every L/E/V cell and decide the pick
                         for each match; also crops the handwritten name.

Everything works on plain numpy BGR images so it can be tested without a camera.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace

import cv2
import numpy as np

PICK_LETTERS = ("L", "E", "V")

# Row verdicts
OK = "ok"
BLANK = "blank"          # nothing marked
MULTIPLE = "multiple"    # two columns marked (correction / double mark)
WEAK = "weak"            # a mark was found but it is faint or small
HEAVY = "heavy"          # scribbled over - probably a crossed-out mark
OCCLUDED = "occluded"    # a hand / finger covers the row's mark cells: nothing can be said about the row


class SlipNotFound(Exception):
    """No slip (or no usable grid) could be located in the image."""


# --------------------------------------------------------------------------- #
# Stage 1: where is the slip?
# --------------------------------------------------------------------------- #
def _order_quad(pts: np.ndarray) -> np.ndarray:
    """Order 4 points TL, TR, BR, BL."""
    pts = np.asarray(pts, np.float32).reshape(4, 2)
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).ravel()
    return np.array([pts[np.argmin(s)], pts[np.argmin(d)],
                     pts[np.argmax(s)], pts[np.argmax(d)]], np.float32)


def _dark_mask(gray: np.ndarray, block: int, c: int) -> np.ndarray:
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    return cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                 cv2.THRESH_BINARY_INV, block | 1, c)


def find_slip_quad(frame: np.ndarray, work_dim: int = 900,
                   min_area_frac: float = 0.02) -> np.ndarray | None:
    """Return the rough outline (TL, TR, BR, BL, in ``frame`` pixels) of the slip.

    A black-hat filter keeps only *thin* dark strokes (printed rules, letters, pen
    marks) and ignores shadows, paper edges and the table, so the biggest cluster of
    strokes is the printed table - on any plain surface, light or dark.
    """
    h, w = frame.shape[:2]
    scale = min(1.0, work_dim / max(h, w))
    small = cv2.resize(frame, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) if scale < 1 else frame
    gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
    sh, sw = gray.shape
    ks = max(7, int(round(0.017 * max(sh, sw))) | 1)
    bh = cv2.morphologyEx(gray, cv2.MORPH_BLACKHAT, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ks, ks)))
    bh = cv2.GaussianBlur(bh, (3, 3), 0)
    otsu, _ = cv2.threshold(bh, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    mask = (bh > max(otsu, 16)).astype(np.uint8) * 255
    k = max(5, int(round(0.012 * max(sh, sw))) | 1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (k, k)))
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None
    c = max(cnts, key=cv2.contourArea)
    hull = cv2.convexHull(c)
    area = cv2.contourArea(hull)
    if area < min_area_frac * sh * sw or area > 0.97 * sh * sw:
        return None
    peri = cv2.arcLength(hull, True)
    quad = None
    for eps in np.linspace(0.01, 0.12, 12):
        ap = cv2.approxPolyDP(hull, eps * peri, True)
        if len(ap) == 4:
            quad = ap.reshape(4, 2)
            break
    if quad is None:
        quad = cv2.boxPoints(cv2.minAreaRect(hull))
    quad = _order_quad(quad).astype(np.float32)
    # Grow 2% so the outermost printed border is safely inside the rough warp.
    ctr = quad.mean(axis=0)
    quad = ctr + (quad - ctr) * 1.02
    wq = (np.linalg.norm(quad[1] - quad[0]) + np.linalg.norm(quad[2] - quad[3])) / 2
    hq = (np.linalg.norm(quad[3] - quad[0]) + np.linalg.norm(quad[2] - quad[1])) / 2
    if min(wq, hq) < 0.06 * max(sh, sw) or not (0.4 < wq / max(hq, 1) < 2.5):
        return None
    return quad / scale


# --------------------------------------------------------------------------- #
# Stage 2: ruled lines -> grid
# --------------------------------------------------------------------------- #
@dataclass
class _Line:
    """A straight ruled line. Horizontal: y = k*x + o.  Vertical: x = k*y + o."""
    k: float
    o: float
    lo: float   # extent along the line's own direction (min)
    hi: float   # (max)

    def at(self, t: float) -> float:
        return self.k * t + self.o


def _fit(pts_u: np.ndarray, pts_v: np.ndarray) -> tuple[float, float]:
    """Least-squares v = k*u + o."""
    k, o = np.polyfit(pts_u, pts_v, 1)
    return float(k), float(o)


def _lines_from_mask(mask: np.ndarray, horizontal: bool, min_len: float,
                     merge_tol: float, max_slope: float = 0.4) -> list[_Line]:
    n, lab, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    comps = []
    for i in range(1, n):
        x, y, w, h, _a = stats[i]
        length = w if horizontal else h
        if length < min_len:
            continue
        ys, xs = np.nonzero(lab == i)
        u, v = (xs, ys) if horizontal else (ys, xs)
        k, o = _fit(u, v)
        if abs(k) > max_slope:
            continue
        comps.append((u, v, k, o))
    if not comps:
        return []
    mid = (mask.shape[1] if horizontal else mask.shape[0]) / 2.0
    comps.sort(key=lambda c: c[2] * mid + c[3])
    groups: list[list] = []
    for c in comps:
        pos = c[2] * mid + c[3]
        if groups and abs(pos - (groups[-1][0][2] * mid + groups[-1][0][3])) < merge_tol \
                and abs(c[2] - groups[-1][0][2]) < 0.04:
            groups[-1].append(c)
        else:
            groups.append([c])
    lines = []
    for g in groups:
        u = np.concatenate([c[0] for c in g])
        v = np.concatenate([c[1] for c in g])
        k, o = _fit(u, v)
        lines.append(_Line(k, o, float(u.min()), float(u.max())))
    lines.sort(key=lambda L: L.at(mid))
    return lines


def _line_kernel(length: int, angle_deg: float) -> np.ndarray:
    """1-pixel-thick line structuring element of the given length and angle."""
    r = length // 2
    size = 2 * r + 1
    k = np.zeros((size, size), np.uint8)
    a = np.deg2rad(angle_deg)
    dx, dy = int(round(r * np.cos(a))), int(round(r * np.sin(a)))
    cv2.line(k, (r - dx, r - dy), (r + dx, r + dy), 1, 1)
    return k


def _oriented_open(mask: np.ndarray, length: int, base_deg: float,
                   spread: float = 9.0, step: float = 3.0) -> np.ndarray:
    """Keep only long thin runs within +-spread degrees of ``base_deg``."""
    out = np.zeros_like(mask)
    for a in np.arange(-spread, spread + 0.1, step):
        k = _line_kernel(length, base_deg + a)
        out = cv2.bitwise_or(out, cv2.morphologyEx(mask, cv2.MORPH_OPEN, k))
    return out


def _extract_lines(gray: np.ndarray) -> tuple[list[_Line], list[_Line]]:
    h, w = gray.shape
    th = _dark_mask(gray, 41, 12)
    kx = max(15, int(0.045 * w)) | 1
    ky = max(15, int(0.045 * h)) | 1
    hm = _oriented_open(th, kx, 0.0)
    vm = _oriented_open(th, ky, 90.0)
    hl = _lines_from_mask(hm, True, 0.2 * w, 0.012 * h)
    vl = _lines_from_mask(vm, False, 0.2 * h, 0.012 * w)
    return hl, vl


def _intersect(hl: _Line, vl: _Line) -> tuple[float, float]:
    """Intersection of y = a*x + b with x = c*y + d."""
    a, b, c, d = hl.k, hl.o, vl.k, vl.o
    x = (c * b + d) / (1.0 - c * a)
    return x, a * x + b


@dataclass
class Grid:
    quad: np.ndarray                 # rough slip outline in frame pixels (TL,TR,BR,BL)
    matrix: np.ndarray               # frame -> rectified image
    size: tuple[int, int]            # rectified (width, height)
    corners: np.ndarray              # (rows+1, 4, 2) lattice points, rectified coords
    title_bottom_y: float            # rectified y of the line under the title row
    name_top_y: float                # rectified y where the NOMBRE row starts
    row_h: float
    n_rows: int
    rotated_180: bool = False


def _rect_from_rough(frame: np.ndarray, quad: np.ndarray, width: int = 1000):
    wq = (np.linalg.norm(quad[1] - quad[0]) + np.linalg.norm(quad[2] - quad[3])) / 2
    hq = (np.linalg.norm(quad[3] - quad[0]) + np.linalg.norm(quad[2] - quad[1])) / 2
    height = max(200, int(round(width * hq / max(wq, 1))))
    dst = np.array([[0, 0], [width, 0], [width, height], [0, height]], np.float32)
    m = cv2.getPerspectiveTransform(quad.astype(np.float32), dst)
    return m, (width, height)


def _lattice(hl: list[_Line], vl: list[_Line], w: int, n_rows: int):
    """Work out the L/E/V columns and the n_rows match rows from the raw lines.

    Returns (rows_h, lev_v, title_line, top_border, left_v, right_v); raises
    SlipNotFound if the printed form is not recognised.
    """
    if len(vl) < 4 or len(hl) < n_rows + 2:
        raise SlipNotFound("not enough ruled lines (%d horizontal, %d vertical)" % (len(hl), len(vl)))
    h_est = max(l.hi for l in vl)
    xs = [v.at(h_est / 2) for v in vl]
    # The three narrow, equally wide L / E / V columns: 4 consecutive verticals on the left.
    qi = None
    for i in range(len(vl) - 3):
        gaps = [xs[i + 1] - xs[i], xs[i + 2] - xs[i + 1], xs[i + 3] - xs[i + 2]]
        g = float(np.mean(gaps))
        if xs[i] > 0.3 * w or g < 0.03 * w or g > 0.16 * w:
            continue
        if max(gaps) / max(min(gaps), 1e-6) > 1.35:
            continue
        qi = i
        break
    if qi is None:
        raise SlipNotFound("could not find the three L / E / V columns")
    lev_v = vl[qi:qi + 4]
    # The interior separators end where the last match row ends.
    inner = lev_v[1:]
    y_bot_v = float(np.median([v.hi for v in inner]))
    y_top_v = float(np.median([v.lo for v in inner]))
    x_ref = float(np.mean([v.at((y_top_v + y_bot_v) / 2) for v in lev_v]))
    ys = np.array([l.at(x_ref) for l in hl])
    j_bot = int(np.argmin(np.abs(ys - y_bot_v)))
    if j_bot < n_rows:
        raise SlipNotFound("found only %d rows above the name line, expected %d" % (j_bot, n_rows))
    j_top = j_bot - n_rows
    hts = np.diff(ys[j_top:j_bot + 1])
    med = float(np.median(hts))
    if abs(ys[j_bot] - y_bot_v) > 0.45 * med:
        raise SlipNotFound("cannot locate the bottom of the match rows")
    if hts.min() < 0.6 * med or hts.max() > 1.6 * med:
        raise SlipNotFound("match rows are not evenly spaced")
    rows_h = hl[j_top:j_bot + 1]
    title_line = hl[j_top - 1] if j_top >= 1 else None
    top_border = hl[0] if j_top >= 2 else None
    right_v = vl[-1] if vl[-1] is not lev_v[-1] and vl[-1].at(h_est / 2) > 0.8 * w else None
    return rows_h, lev_v, title_line, top_border, lev_v[0], right_v


def _locate_in_rough(rough: np.ndarray, n_rows: int):
    gray = cv2.cvtColor(rough, cv2.COLOR_BGR2GRAY)
    hl, vl = _extract_lines(gray)
    w = rough.shape[1]
    rows_h, lev_v, title_line, top_border, left_v, right_v = _lattice(hl, vl, w, n_rows)
    return rows_h, lev_v, title_line, top_border, left_v, right_v


def locate_grid(frame: np.ndarray, n_rows: int = 9, quad: np.ndarray | None = None) -> Grid:
    """Find the slip in ``frame`` and the lattice of match rows x L/E/V columns."""
    if quad is None:
        quad = find_slip_quad(frame)
    if quad is None:
        raise SlipNotFound("no slip in view")
    quad = _order_quad(quad)
    errors: list[str] = []
    for rotated in (False, True):
        q = np.roll(quad, 2, axis=0) if rotated else quad     # 180 degrees: swap TL<->BR, TR<->BL
        m1, (w1, h1) = _rect_from_rough(frame, q)
        rough = cv2.warpPerspective(frame, m1, (w1, h1), flags=cv2.INTER_LINEAR)
        try:
            rows_h, lev_v, title_line, top_border, left_v, right_v = _locate_in_rough(rough, n_rows)
        except SlipNotFound as e:
            errors.append(str(e))
            continue
        break
    else:
        raise SlipNotFound(errors[0])

    # Lattice corners (rough-warp coordinates): rows x 4 verticals
    pts = np.array([[_intersect(hr, v) for v in lev_v] for hr in rows_h], np.float32)

    # Second, tight warp so the saved/overlaid image is straight.
    top_line = top_border if top_border is not None else (title_line if title_line is not None else rows_h[0])
    tl = np.array(_intersect(top_line, left_v), np.float32)
    # right edge: the right border if seen, else extend beyond the home|away column
    if right_v is not None:
        tr = np.array(_intersect(top_line, right_v), np.float32)
        br_anchor = np.array(_intersect(rows_h[-1], right_v), np.float32)
        br_prev = np.array(_intersect(rows_h[-2], right_v), np.float32)
    else:
        tr = tl + (pts[0, 3] - pts[0, 0]) * 6.0
        br_anchor = pts[-1, 0] + (pts[-1, 3] - pts[-1, 0]) * 6.0
        br_prev = pts[-2, 0] + (pts[-2, 3] - pts[-2, 0]) * 6.0
    bl_anchor = np.array(_intersect(rows_h[-1], left_v), np.float32)
    bl_prev = np.array(_intersect(rows_h[-2], left_v), np.float32)
    name_rows = 1.7
    bl = bl_anchor + (bl_anchor - bl_prev) * name_rows
    br = br_anchor + (br_anchor - br_prev) * name_rows
    q2 = np.array([tl, tr, br, bl], np.float32)
    wd = 1000
    wq = (np.linalg.norm(tr - tl) + np.linalg.norm(br - bl)) / 2
    hq = (np.linalg.norm(bl - tl) + np.linalg.norm(br - tr)) / 2
    hd = int(round(wd * hq / max(wq, 1)))
    m2 = cv2.getPerspectiveTransform(q2, np.array([[0, 0], [wd, 0], [wd, hd], [0, hd]], np.float32))
    total = m2 @ m1
    corners = cv2.perspectiveTransform(pts.reshape(-1, 1, 2), m2).reshape(pts.shape)
    row_h = float(np.median(np.diff(corners[:, 0, 1])))
    if title_line is not None:
        tb = cv2.perspectiveTransform(
            np.array([[[_intersect(title_line, left_v)[0], _intersect(title_line, left_v)[1]]]], np.float32), m2)
        title_bottom = float(tb[0, 0, 1])
    else:
        title_bottom = 0.0
    return Grid(quad=quad, matrix=total, size=(wd, hd), corners=corners,
                title_bottom_y=title_bottom, name_top_y=float(corners[-1, 0, 1]),
                row_h=row_h, n_rows=n_rows, rotated_180=rotated)


# --------------------------------------------------------------------------- #
# Stage 3: reading the marks
# --------------------------------------------------------------------------- #
@dataclass
class Settings:
    min_ink: float = 0.015          # fraction of a cell that must be inked to count as a mark
    multi_ratio: float = 0.45       # 2nd-best / best above this -> double mark
    heavy_ink: float = 0.30         # fraction of cell inked beyond which it looks scribbled out (a real X: 0.13-0.19)
    weak_ink: float = 0.04          # below this a mark is flagged as faint
    second_ink: float = 0.04        # a 2nd cell this much more inked than the slip's other off-mark cells is a 2nd mark
                                    # whatever the ratio: a corrected mark beside a scribbled-out one has a low ratio,
                                    # while a lone X leaves ~0.002 in the other cells (a strong colour cast: ~0.07)
    faint_ink: float = 0.022        # BLANK rows get a second, more sensitive look: this much ink = a faint mark (WEAK)
    thin_aspect: float = 6.0        # ink this long and thin (staple, pen slash) is not a mark (real X marks: aspect <= 1.8)
    edge_off: float = 0.25          # ink centred this far (in row heights) from its cell's centre lies on a rule (X marks: <= 0.15)
    hand_guard: bool = True         # look for skin: OCCLUDED rows, and tell the capture engine a hand is on the slip
    skin_slip: float = 0.01         # share of the slip covered by skin that means "a hand is on it"
    skin_row: float = 0.01          # share of a row's mark area (+ half a cell around) under skin: row is OCCLUDED


@dataclass
class SlipReading:
    picks: list[str]                        # 'L'/'E'/'V' or '' when blank, ambiguous or hidden
    verdicts: list[str]                     # OK / BLANK / MULTIPLE / WEAK / HEAVY / OCCLUDED per row
    ink: np.ndarray                         # (rows, 3) ink fraction per cell
    rectified: np.ndarray                   # straightened slip, BGR
    name_crop: np.ndarray                   # handwriting strip, BGR
    title_crop: np.ndarray                  # printed title strip, BGR
    grid: Grid
    candidates: list[tuple[str, ...]] = field(default_factory=list)   # letters that carry ink per row
    ink_mask: np.ndarray | None = None      # _ink_mask(rectified), kept so fingerprint() need not recompute it
    skin_fraction: float = 0.0              # share of the rectified slip covered by skin-coloured blobs
    hand: bool = False                      # a hand is on the slip (skin_fraction >= Settings.skin_slip, or it covers a row)
    name_hidden: bool = False               # skin covers the handwriting strip

    @property
    def clean(self) -> bool:
        return all(v == OK for v in self.verdicts)

    @property
    def problem_rows(self) -> list[int]:
        return [i for i, v in enumerate(self.verdicts) if v != OK]

    @property
    def occluded_rows(self) -> list[int]:
        return [i for i, v in enumerate(self.verdicts) if v == OCCLUDED]


def _ink_mask(fin: np.ndarray) -> np.ndarray:
    """Pen ink in the rectified slip: coloured pen OR dark strokes that are not ruled lines."""
    gray = cv2.cvtColor(fin, cv2.COLOR_BGR2GRAY)
    hsv = cv2.cvtColor(fin, cv2.COLOR_BGR2HSV)
    dark = _dark_mask(gray, 41, 14)
    h, w = gray.shape
    hk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(15, int(0.04 * w)), 1)))
    vk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(15, int(0.04 * h)))))
    lines = cv2.dilate(cv2.bitwise_or(hk, vk), cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5)))
    ink_dark = cv2.bitwise_and(dark, cv2.bitwise_not(lines))
    # Coloured ink: saturated and clearly darker than paper.
    paper = cv2.medianBlur(gray, 41)
    coloured = ((hsv[..., 1] > 70) & (hsv[..., 2] > 35) & (gray < paper.astype(np.int16) - 25)).astype(np.uint8) * 255
    ink = cv2.bitwise_or(ink_dark, coloured)
    return cv2.morphologyEx(ink, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))


# ----- hands: skin-coloured blobs on the slip
def skin_mask(img: np.ndarray, kernel: int = 0) -> np.ndarray:
    """Skin-coloured blobs of a BGR image (uint8 0/255): hue red..orange, clearly more saturated than the
    paper (median saturation + 30) yet not as saturated as the strongest inks, opened with a kernel of ~3% of
    the width (a finger is 10-18%) so pen strokes - even 1.6 mm felt-tip circles - and specks vanish.
    Measured: 0 pixels on clean frames under six colour casts x three exposures, with blue/black/red/orange/
    brown ink and thick markers; every finger of the test batch is found. Blind to skin that a colour cast
    turns grey (strong blue light)."""
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    lo, hi = int(np.median(hsv[::4, ::4, 1])) + 31, 225
    if lo > hi:                                          # paper more saturated than skin: cannot tell
        return np.zeros(img.shape[:2], np.uint8)
    m = cv2.bitwise_or(cv2.inRange(hsv, (0, lo, 41), (25, hi, 255)), cv2.inRange(hsv, (170, lo, 41), (179, hi, 255)))
    k = kernel or max(3, int(round(0.031 * img.shape[1])) | 1)
    return cv2.morphologyEx(m, cv2.MORPH_OPEN, np.ones((k, k), np.uint8))


def hand_fraction(frame: np.ndarray, quad: np.ndarray, width: int = 480) -> float:
    """Share of the slip outline ``quad`` (frame pixels) covered by skin: a ~3 ms 'is a hand on it?' that
    needs no grid, so the capture engine can refuse to read (and can say so) before spending ~200 ms."""
    q = _order_quad(quad)
    c = q.mean(axis=0)
    m, size = _rect_from_rough(frame, c + (q - c) * 0.94, width)         # inside the table: no background
    mask = skin_mask(cv2.warpPerspective(frame, m, size, flags=cv2.INTER_LINEAR), 15)  # a finger is >= 48 px wide here
    return float(cv2.countNonZero(mask)) / max(mask.size, 1)


def _skin_report(skin: np.ndarray, g: Grid, s: Settings) -> tuple[list[bool], bool]:
    """(per row: a hand over its mark cells?, a hand over the handwriting strip?) from a skin mask."""
    h, w = skin.shape
    c = g.corners
    rows = []
    for r in range(g.n_rows):
        xs, ys = c[r:r + 2, :, 0], c[r:r + 2, :, 1]
        cw = float(xs.max() - xs.min()) / 3
        x0, x1 = max(0, int(xs.min() - 0.5 * cw)), min(w, int(xs.max() + 0.5 * cw))
        y0, y1 = max(0, int(ys.min() - 0.5 * g.row_h)), min(h, int(ys.max() + 0.5 * g.row_h))
        box = skin[y0:y1, x0:x1]
        rows.append(box.size > 0 and cv2.countNonZero(box) > s.skin_row * box.size)
    y0 = max(0, int(g.name_top_y - 0.5 * g.row_h))
    box = skin[y0:, int(0.2 * w):int(0.975 * w)]
    return rows, box.size > 0 and cv2.countNonZero(box) > s.skin_row * box.size


def _faint_fractions(fin: np.ndarray, g: Grid) -> np.ndarray:
    """(rows, 3) ink share per cell from a more sensitive pass (blur 5x5, adaptive threshold C=6): only run
    for rows that read BLANK, to tell a dry pen / pencil from an empty box. Measured on the real photo and
    synthetic blanks (blur 0.8, noise <= 6): 0 blank cells above 0.02; faint marks the main pass misses: >= 0.0215."""
    h, w = fin.shape[:2]
    c = g.corners
    pad = 46
    x0, x1 = max(0, int(c[..., 0].min()) - pad), min(w, int(c[..., 0].max()) + pad)
    y0, y1 = max(0, int(c[..., 1].min()) - pad), min(h, int(c[..., 1].max()) + pad)
    gray = cv2.cvtColor(fin[y0:y1, x0:x1], cv2.COLOR_BGR2GRAY)
    dark = cv2.adaptiveThreshold(cv2.GaussianBlur(gray, (5, 5), 0), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                 cv2.THRESH_BINARY_INV, 41, 6)
    hk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(15, int(0.04 * w)), 1)))
    vk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(15, int(0.04 * h)))))
    lines = cv2.dilate(cv2.bitwise_or(hk, vk), cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5)))
    ink = cv2.morphologyEx(cv2.bitwise_and(dark, cv2.bitwise_not(lines)), cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
    out = np.zeros((g.n_rows, 3), np.float32)
    for r in range(g.n_rows):
        for k in range(3):
            out[r, k] = _cell_ink(ink, g, r, k, (x0, y0))[0]
    return out


def _cell_ink(ink: np.ndarray, g: Grid, r: int, c: int, origin=(0, 0)):
    """(ink share of cell (r, c), its ink pixels as a crop, the crop's top-left) - the cell polygon is shrunk 6%
    towards its centre so the ruled lines themselves are ignored."""
    poly = np.array([g.corners[r, c], g.corners[r, c + 1], g.corners[r + 1, c + 1], g.corners[r + 1, c]], np.float32)
    ctr = poly.mean(axis=0)
    poly = (ctr + (poly - ctr) * 0.94).astype(np.int32) - np.array(origin, np.int32)
    x0, y0 = np.maximum(poly.min(axis=0), 0)
    x1, y1 = np.minimum(poly.max(axis=0) + 1, (ink.shape[1], ink.shape[0]))
    if x1 <= x0 or y1 <= y0:
        return 0.0, None, (0, 0)
    m = np.zeros((y1 - y0, x1 - x0), np.uint8)
    cv2.fillConvexPoly(m, poly - (x0, y0), 255)
    inked = cv2.bitwise_and(ink[y0:y1, x0:x1], m)
    return cv2.countNonZero(inked) / max(float(cv2.countNonZero(m)), 1.0), inked, (int(x0) + origin[0], int(y0) + origin[1])


def _mark_shape(inked: np.ndarray, origin, ctr_y: float, row_h: float) -> tuple[float, float]:
    """(aspect, off-centre) of the ink of one cell: the aspect ratio of its minimum-area rectangle (a real X: 1-2.2,
    a staple or a slash: > 8) and how far its vertical centroid lies from the cell's centre, in row heights."""
    ys, xs = np.nonzero(inked)
    if ys.size < 30:
        return 1.0, 0.0
    (_, _), (rw, rh), _ = cv2.minAreaRect(np.column_stack([xs, ys]).astype(np.float32))
    return max(rw, rh) / max(min(rw, rh), 1.0), abs(float(ys.mean()) + origin[1] - ctr_y) / row_h


def read_slip(frame: np.ndarray, n_rows: int = 9, settings: Settings | None = None,
              grid: Grid | None = None) -> SlipReading:
    s = settings or Settings()
    g = grid or locate_grid(frame, n_rows=n_rows)
    fin = cv2.warpPerspective(frame, g.matrix, g.size, flags=cv2.INTER_CUBIC)
    ink = _ink_mask(fin)

    skin_fraction, occluded, name_hidden = 0.0, [False] * g.n_rows, False
    if s.hand_guard:
        skin = skin_mask(fin)
        skin_fraction = float(cv2.countNonZero(skin)) / max(skin.size, 1)
        if skin_fraction > 0:
            occluded, name_hidden = _skin_report(skin, g, s)

    fractions = np.zeros((g.n_rows, 3), np.float32)
    cells: dict = {}
    for r in range(g.n_rows):
        for c in range(3):
            fractions[r, c], inked, org = _cell_ink(ink, g, r, c)
            cells[r, c] = (inked, org)

    # what an unmarked cell measures on this photo (~0.00 in even light, up to ~0.07 under a strong colour cast): the
    # typical 2nd-strongest cell of a row - never a mark, except in the odd double-marked row
    floor = float(np.median(np.sort(fractions, axis=1)[:, -2]))
    faint = None
    picks, verdicts, cands = [], [], []
    for r in range(g.n_rows):
        f = fractions[r]
        if occluded[r]:                                   # a hand over the cells: the ink there says nothing
            picks.append("")
            verdicts.append(OCCLUDED)
            cands.append(())
            continue
        order = np.argsort(f)[::-1]
        best, second = float(f[order[0]]), float(f[order[1]])
        marked = tuple(PICK_LETTERS[i] for i in range(3) if f[i] >= s.min_ink)
        if best < s.min_ink:
            if faint is None:
                faint = _faint_fractions(fin, g)
            fr = faint[r]
            if float(fr.max()) >= s.faint_ink:            # a dry pen / pencil, not an empty box: show a guess, but flag it
                marked = tuple(PICK_LETTERS[i] for i in range(3) if fr[i] >= s.faint_ink)
                picks.append(PICK_LETTERS[int(np.argmax(fr))])
                verdicts.append(WEAK)
            else:
                picks.append("")
                verdicts.append(BLANK)
        elif second >= s.min_ink and (second / best > s.multi_ratio or second >= floor + s.second_ink):
            picks.append("")
            verdicts.append(MULTIPLE)
        else:
            c = int(order[0])
            v = HEAVY if best > s.heavy_ink else WEAK if best < s.weak_ink else OK
            if v == OK and cells[r, c][0] is not None:
                ctr_y = float((g.corners[r, c, 1] + g.corners[r + 1, c, 1] + g.corners[r, c + 1, 1] + g.corners[r + 1, c + 1, 1]) / 4)
                aspect, off = _mark_shape(cells[r, c][0], cells[r, c][1], ctr_y, g.row_h)
                if aspect >= s.thin_aspect or off > s.edge_off:   # a staple / slash, or a mark that sits on a rule
                    v = WEAK
            picks.append(PICK_LETTERS[c])
            verdicts.append(v)
        cands.append(marked)

    h_img, w_img = fin.shape[:2]
    y_name0 = int(max(0, g.name_top_y - 0.18 * g.row_h))
    name_crop = fin[y_name0:h_img, :].copy()
    y_title = int(max(g.title_bottom_y, 0.9 * g.row_h)) if g.title_bottom_y > 0 else int(g.row_h)
    title_crop = fin[0:max(10, y_title), :].copy()
    return SlipReading(picks=picks, verdicts=verdicts, ink=fractions, rectified=fin,
                       name_crop=name_crop, title_crop=title_crop, grid=g, candidates=cands, ink_mask=ink,
                       skin_fraction=skin_fraction, hand=skin_fraction >= s.skin_slip or any(occluded), name_hidden=name_hidden)


# --------------------------------------------------------------------------- #
# Stage 4: "is this the same physical slip?"  (stack handling)
# --------------------------------------------------------------------------- #
# Slips of a stack share the identical printed form, so the fingerprint is made of pen ink
# only: the L/E/V mark columns and the handwriting part of the name strip (never the printed
# NOMBRE label, the rules, or the 'ULTIMO DIA...' line under the strip).
_FP_MARKS = (104, 370)          # (w, h) of the mark raster: the mark columns at half the rectified size
_FP_NAME = (388, 47)            # (w, h) of the handwriting raster (same scale)
_FP_SHIFT = 3                   # +-px (raster) of registration slack between two photos
_FP_DILATE = 1                  # px (raster) a stroke may be off and still count as the same stroke
_FP_SPECK = 7                   # connected ink blobs smaller than this (raster px) are noise
_FP_MIN_PX = 40                 # fewer ink pixels than this = nothing written there
MARKS_SAME = 0.90               # min stroke coverage (both ways) for the marks of one slip
NAME_SAME = 0.80                # ... and for its handwriting
MARKS_STRONG = 0.95             # marks this alike are one slip whatever the handwriting says (a hand may cover it)
_FP_MIN_ROWS = 4                # ... provided at least this many rows carry a mark: one or two X prove little


@dataclass(eq=False)
class Fingerprint:
    picks: tuple[str, ...]       # per row: the pick, or the letters that carry ink when there is no single pick
    marks: np.ndarray = field(repr=False)    # binary (uint8) raster of the pen ink in the L/E/V columns
    name: np.ndarray = field(repr=False)     # binary raster of the handwriting
    marks_px: int
    name_px: int
    occluded: tuple[int, ...] = ()           # rows hidden by a hand: they match anything (see same_slip)
    soft: tuple[int, ...] = ()               # rows read WEAK: a faint mark may come out blank on the next photo
    name_hidden: bool = False                # a hand covers the handwriting: it cannot tell slips apart
    row_bands: tuple = field(default=(), repr=False)   # (y0, y1) of every row in the marks raster

    @property
    def name_blank(self) -> bool:
        return self.name_px < _FP_MIN_PX


def _name_window(fin: np.ndarray, g: Grid) -> tuple[int, int, int, int]:
    """(x0, x1, y0, y1) of the handwriting area in the rectified slip: right of the printed
    NOMBRE label, under the last match row, above the strip's bottom rule."""
    h, w = fin.shape[:2]
    y0 = int(round(g.name_top_y + 0.08 * g.row_h))
    y_max = int(round(g.name_top_y + 1.15 * g.row_h))
    y_end = y_max
    top = int(round(g.name_top_y + 0.5 * g.row_h))
    if 0 <= top < h - 8:
        dark = _dark_mask(cv2.cvtColor(fin[top:], cv2.COLOR_BGR2GRAY), 41, 14)
        rules = _oriented_open(dark, max(15, int(0.25 * w)) | 1, 0.0, spread=3.0, step=1.5)
        rows = np.nonzero(rules.sum(axis=1) > 0.25 * 255 * w)[0]
        if rows.size:
            y_rule = top + int(rows.max())
            if g.name_top_y + 0.6 * g.row_h <= y_rule <= g.name_top_y + 1.5 * g.row_h:
                y_end = min(y_max, y_rule - 2)
    return int(0.2 * w), int(0.975 * w), y0, max(y0 + 4, y_end)


def _raster(ink: np.ndarray, x0: int, x1: int, y0: int, y1: int, size: tuple[int, int],
            full_h: int | None = None) -> np.ndarray:
    """Cut ink[y0:y1, x0:x1] (zero padded to full_h rows), shrink to ``size`` -> binary uint8."""
    h, w = ink.shape[:2]
    x0, x1 = max(0, x0), min(w, x1)
    crop = ink[max(0, y0):max(0, min(h, y1)), x0:x1]
    if full_h is not None and crop.shape[0] < full_h:
        crop = np.vstack([crop, np.zeros((full_h - crop.shape[0], crop.shape[1]), np.uint8)])
    if crop.size == 0:
        return np.zeros((size[1], size[0]), np.uint8)
    small = (cv2.resize(crop, size, interpolation=cv2.INTER_AREA) > 80).astype(np.uint8)
    _, lab, stats, _ = cv2.connectedComponentsWithStats(small, connectivity=8)
    keep = stats[:, cv2.CC_STAT_AREA] >= _FP_SPECK          # tiny blobs are sensor noise or dust
    keep[0] = False
    return keep[lab].astype(np.uint8)


def _fingerprint(fin: np.ndarray, g: Grid, ink: np.ndarray, picks: tuple[str, ...] = (), occluded=(), soft=(),
                 name_hidden: bool = False) -> Fingerprint:
    c = g.corners
    inset = 3
    mx0, mx1 = int(round(float(c[:, 0, 0].mean()))) + inset, int(round(float(c[:, 3, 0].mean()))) - inset
    my0, my1 = int(round(float(c[0, 0, 1]))) + inset, int(round(float(c[-1, 0, 1]))) - inset
    marks = _raster(ink, mx0, mx1, my0, my1, _FP_MARKS)
    nx0, nx1, ny0, ny1 = _name_window(fin, g)
    full_h = int(round(1.15 * g.row_h - 0.08 * g.row_h))
    name = _raster(ink, nx0, nx1, ny0, ny1, _FP_NAME, full_h)
    edges = (c[:, 0, 1] - my0) * _FP_MARKS[1] / max(my1 - my0, 1)
    bands = tuple((int(np.floor(edges[r])) - 1, int(np.ceil(edges[r + 1])) + 1) for r in range(g.n_rows))
    return Fingerprint(picks, marks, name, int(marks.sum()), int(name.sum()), tuple(occluded), tuple(soft),
                       name_hidden, bands)


def fingerprint(reading: SlipReading) -> Fingerprint:
    """Pen-ink signature of a slip (see :func:`same_slip`). ~3 ms when ``reading.ink_mask`` is set."""
    ink = reading.ink_mask if reading.ink_mask is not None else _ink_mask(reading.rectified)
    cands = reading.candidates or [()] * len(reading.picks)
    rows = tuple(p or "".join(c) for p, c in zip(reading.picks, cands))   # blank '' / double mark 'LV' stay distinct
    return _fingerprint(reading.rectified, reading.grid, ink, rows,
                        [i for i, v in enumerate(reading.verdicts) if v == OCCLUDED],
                        [i for i, v in enumerate(reading.verdicts) if v == WEAK], reading.name_hidden)


_ALIGN = 4                      # the registration of two photos of the slip works on images this many times smaller


def _align_gray(fin: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(fin, cv2.COLOR_BGR2GRAY)
    small = cv2.resize(gray, (fin.shape[1] // _ALIGN, fin.shape[0] // _ALIGN), interpolation=cv2.INTER_AREA)
    return cv2.GaussianBlur(small, (0, 0), 1.0).astype(np.float32)


def alignment_template(reading: SlipReading) -> np.ndarray:
    """What :func:`quick_fingerprint` registers a later photo against: the rectified slip, small and grey."""
    return _align_gray(reading.rectified)


def quick_fingerprint(frame: np.ndarray, grid: Grid, settings: Settings | None = None,
                      template: np.ndarray | None = None) -> tuple[Fingerprint, float]:
    """(pen-ink signature, skin share) of whatever lies where ``grid`` was found, without the verdict logic and
    without locating the grid again (~90 ms instead of ~190): lets the capture engine keep checking that the slip
    it entered is still the one there. With ``template`` (see :func:`alignment_template`) the photo is first
    registered on the printed lines (focus breathing, a nudge), which an outline of the slip is too coarse for;
    raises SlipNotFound when that fails."""
    s = settings or Settings()
    matrix = grid.matrix
    if template is not None:
        fin = cv2.warpPerspective(frame, matrix, grid.size, flags=cv2.INTER_LINEAR)
        warp = np.eye(2, 3, dtype=np.float32)
        try:
            _, warp = cv2.findTransformECC(template, _align_gray(fin), warp, cv2.MOTION_AFFINE,
                                           (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, 30, 1e-4), None, 5)
        except cv2.error as e:
            raise SlipNotFound("cannot register the slip") from e
        warp[:, 2] *= _ALIGN
        matrix = np.linalg.inv(np.vstack([warp, [0, 0, 1]])) @ matrix       # frame -> rectified, registered
    fin = cv2.warpPerspective(frame, matrix, grid.size, flags=cv2.INTER_CUBIC)
    skin = skin_mask(fin)
    share = float(cv2.countNonZero(skin)) / max(skin.size, 1)
    hidden = _skin_report(skin, grid, s)[1] if share > 0 else False
    return _fingerprint(fin, grid, _ink_mask(fin), name_hidden=hidden), share


def _coverage(a: np.ndarray, b: np.ndarray) -> float:
    """Best (over small shifts) of min(share of a's ink lying on b's strokes, share of b's on a's)."""
    pad = _FP_SHIFT
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2 * _FP_DILATE + 1, 2 * _FP_DILATE + 1))
    ad, bd = cv2.dilate(a, k), cv2.dilate(b, k)

    def padded(m):
        return cv2.copyMakeBorder(m, pad, pad, pad, pad, cv2.BORDER_CONSTANT, value=0).astype(np.float32)
    # c1[d] = sum a[x] * bd[x + d]  (a's ink explained by b);  c2[d] = sum ad[x] * b[x + d]  (b's ink by a)
    c1 = cv2.matchTemplate(padded(bd), a.astype(np.float32), cv2.TM_CCORR)
    c2 = cv2.matchTemplate(padded(b), ad.astype(np.float32), cv2.TM_CCORR)
    best = np.minimum(c1 / max(float(a.sum()), 1.0), c2 / max(float(b.sum()), 1.0))
    return float(best.max())


def slip_similarity(a: Fingerprint, b: Fingerprint) -> tuple[float, float | None]:
    """(marks score, handwriting score) in 0..1; the handwriting score is None when either
    name strip is blank. Measured: one slip photographed again scores >= 0.93 (usually > 0.98),
    another slip with identical picks <= 0.61 on the marks and <= 0.71 on the name."""
    marks = _coverage(a.marks, b.marks) if a.marks_px >= _FP_MIN_PX and b.marks_px >= _FP_MIN_PX else \
        (1.0 if a.marks_px < _FP_MIN_PX and b.marks_px < _FP_MIN_PX else 0.0)
    name = None if (a.name_blank or b.name_blank) else _coverage(a.name, b.name)
    return marks, name


def _without_rows(fp: Fingerprint, rows) -> Fingerprint:
    """``fp`` with the ink of ``rows`` wiped from its marks raster (rows hidden in the other photo)."""
    if not fp.row_bands:
        return fp
    m = fp.marks.copy()
    for r in rows:
        y0, y1 = fp.row_bands[r]
        m[max(0, y0):max(0, y1)] = 0
    return replace(fp, marks=m, marks_px=int(m.sum()))


def _picks_alike(a: Fingerprint, b: Fingerprint, hidden: set) -> bool:
    if len(a.picks) != len(b.picks):
        return False
    soft = set(a.soft) | set(b.soft)
    for i, (x, y) in enumerate(zip(a.picks, b.picks)):
        if x == y or i in hidden or (i in soft and "" in (x, y)):    # a faint mark may or may not be seen again
            continue
        return False
    return True


def same_ink(a: Fingerprint, b: Fingerprint, marks_min: float = MARKS_SAME) -> bool:
    """Do two fingerprints show the same pen ink (picks are not compared)? For re-checking a slip that lies
    where it was read: two empty slips count as alike (nothing to tell them apart by). ``marks_min`` is the
    stroke coverage the marks need (the handwriting needs NAME_SAME, scaled down alike)."""
    hidden = set(a.occluded) | set(b.occluded)
    if hidden:
        a, b = _without_rows(a, hidden), _without_rows(b, hidden)
    if a.marks_px < _FP_MIN_PX and b.marks_px < _FP_MIN_PX and a.name_blank and b.name_blank:
        return True
    marks, name = slip_similarity(a, b)
    if a.name_hidden or b.name_hidden:
        name = None
    return marks >= marks_min and (name is None or name >= NAME_SAME * marks_min / MARKS_SAME)


def same_slip(a: Fingerprint, b: Fingerprint) -> bool:
    """True when two readings are the same physical slip (re-photographed, moved a little,
    relit, partly under a hand), False for a different slip - even one with identical picks.
    Rows a hand covered in either photo are ignored; marks that match as well as one slip's
    two photos do decide on their own (a hand over the name strip must not make a new slip)."""
    hidden = set(a.occluded) | set(b.occluded)
    if not _picks_alike(a, b, hidden):
        return False
    if hidden:
        a, b = _without_rows(a, hidden), _without_rows(b, hidden)
    marks, name = slip_similarity(a, b)
    informative = sum(1 for i, p in enumerate(a.picks) if p and i not in hidden) >= _FP_MIN_ROWS
    if informative and a.marks_px >= 3 * _FP_MIN_PX and b.marks_px >= 3 * _FP_MIN_PX and marks >= MARKS_STRONG:
        return True
    if a.name_hidden or b.name_hidden:                  # no handwriting to compare: the marks have to carry it
        return a.marks_px >= _FP_MIN_PX and b.marks_px >= _FP_MIN_PX and marks >= MARKS_SAME
    if a.name_blank != b.name_blank:                    # someone wrote / nobody wrote a name
        return False
    if a.name_blank and a.marks_px < _FP_MIN_PX and b.marks_px < _FP_MIN_PX:
        return False                                    # nothing written at all: cannot tell
    return marks >= MARKS_SAME and (name is None or name >= NAME_SAME)


# --------------------------------------------------------------------------- #
# Drawing helpers (UI + debugging)
# --------------------------------------------------------------------------- #
_VERDICT_COLOUR = {OK: (60, 180, 60), BLANK: (0, 0, 220), MULTIPLE: (0, 140, 255),
                   WEAK: (0, 200, 255), HEAVY: (0, 140, 255), OCCLUDED: (200, 0, 200)}


def draw_reading(reading: SlipReading) -> np.ndarray:
    """Rectified slip with the detected pick circled in each row (a row under a hand is tinted magenta, labelled HAND)."""
    out = reading.rectified.copy()
    g = reading.grid
    for r in range(g.n_rows):
        colour = _VERDICT_COLOUR[reading.verdicts[r]]
        hidden = reading.verdicts[r] == OCCLUDED
        if hidden:
            poly = np.array([g.corners[r, 0], g.corners[r, 3], g.corners[r + 1, 3], g.corners[r + 1, 0]], np.int32)
            tint = out.copy()
            cv2.fillConvexPoly(tint, poly, colour)
            out = cv2.addWeighted(tint, 0.4, out, 0.6, 0)
            cv2.polylines(out, [poly], True, colour, 3)
        pk = reading.picks[r]
        cols = [PICK_LETTERS.index(pk)] if pk else [PICK_LETTERS.index(x) for x in reading.candidates[r]]
        if not cols and not hidden:
            ctr = (g.corners[r, 0] + g.corners[r + 1, 3]) / 2
            cv2.putText(out, "?", (int(ctr[0]) - 8, int(ctr[1]) + 10), cv2.FONT_HERSHEY_SIMPLEX, 1.0, colour, 3)
        for c in cols:
            ctr = (g.corners[r, c] + g.corners[r, c + 1] + g.corners[r + 1, c] + g.corners[r + 1, c + 1]) / 4
            rad = int(0.36 * min(np.linalg.norm(g.corners[r, c + 1] - g.corners[r, c]), g.row_h))
            cv2.circle(out, (int(ctr[0]), int(ctr[1])), max(rad, 8), colour, 3)
        lab = f"{r + 1}:{pk or ('F1' if reading.verdicts[r] == BLANK else 'HAND' if hidden else '??')}"
        y = int(g.corners[r, 3][1] + 0.62 * g.row_h)
        (tw, th), _ = cv2.getTextSize(lab, cv2.FONT_HERSHEY_SIMPLEX, 0.85, 2)
        x = out.shape[1] - tw - 14                       # right margin, clear of the team names
        cv2.rectangle(out, (x - 6, y - th - 6), (x + tw + 6, y + 8), (255, 255, 255), -1)
        cv2.putText(out, lab, (x, y), cv2.FONT_HERSHEY_SIMPLEX, 0.85, colour, 2)
    return out


def draw_quad(frame: np.ndarray, quad: np.ndarray | None, colour=(0, 255, 0), thick: int = 3) -> np.ndarray:
    out = frame.copy()
    if quad is not None:
        cv2.polylines(out, [np.asarray(quad, np.int32).reshape(-1, 1, 2)], True, colour, thick)
    return out


def sharpness(frame: np.ndarray, quad: np.ndarray | None = None) -> float:
    """Variance of the Laplacian over the slip area (focus quality)."""
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    if quad is not None:
        x, y, w, h = cv2.boundingRect(np.asarray(quad, np.int32))
        x, y = max(x, 0), max(y, 0)
        gray = gray[y:y + h, x:x + w]
    if gray.size == 0:
        return 0.0
    return float(cv2.Laplacian(gray, cv2.CV_64F).var())
