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

from dataclasses import dataclass, field

import cv2
import numpy as np

PICK_LETTERS = ("L", "E", "V")

# Row verdicts
OK = "ok"
BLANK = "blank"          # nothing marked
MULTIPLE = "multiple"    # two columns marked (correction / double mark)
WEAK = "weak"            # a mark was found but it is faint or small
HEAVY = "heavy"          # scribbled over - probably a crossed-out mark


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
    heavy_ink: float = 0.40         # fraction of cell inked beyond which it looks scribbled out
    weak_ink: float = 0.04          # below this a mark is flagged as faint


@dataclass
class SlipReading:
    picks: list[str]                        # 'L'/'E'/'V' or '' when blank or ambiguous
    verdicts: list[str]                     # OK / BLANK / MULTIPLE / WEAK / HEAVY per row
    ink: np.ndarray                         # (rows, 3) ink fraction per cell
    rectified: np.ndarray                   # straightened slip, BGR
    name_crop: np.ndarray                   # handwriting strip, BGR
    title_crop: np.ndarray                  # printed title strip, BGR
    grid: Grid
    candidates: list[tuple[str, ...]] = field(default_factory=list)   # letters that carry ink per row

    @property
    def clean(self) -> bool:
        return all(v == OK for v in self.verdicts)

    @property
    def problem_rows(self) -> list[int]:
        return [i for i, v in enumerate(self.verdicts) if v != OK]


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


def read_slip(frame: np.ndarray, n_rows: int = 9, settings: Settings | None = None,
              grid: Grid | None = None) -> SlipReading:
    s = settings or Settings()
    g = grid or locate_grid(frame, n_rows=n_rows)
    fin = cv2.warpPerspective(frame, g.matrix, g.size, flags=cv2.INTER_CUBIC)
    ink = _ink_mask(fin)

    fractions = np.zeros((g.n_rows, 3), np.float32)
    for r in range(g.n_rows):
        for c in range(3):
            poly = np.array([g.corners[r, c], g.corners[r, c + 1],
                             g.corners[r + 1, c + 1], g.corners[r + 1, c]], np.float32)
            # shrink 6% towards the centre so the ruled lines themselves are ignored
            ctr = poly.mean(axis=0)
            poly = ctr + (poly - ctr) * 0.94
            m = np.zeros(ink.shape, np.uint8)
            cv2.fillConvexPoly(m, poly.astype(np.int32), 255)
            area = float(cv2.countNonZero(m))
            fractions[r, c] = cv2.countNonZero(cv2.bitwise_and(ink, m)) / max(area, 1.0)

    picks, verdicts, cands = [], [], []
    for r in range(g.n_rows):
        f = fractions[r]
        order = np.argsort(f)[::-1]
        best, second = float(f[order[0]]), float(f[order[1]])
        marked = tuple(PICK_LETTERS[i] for i in range(3) if f[i] >= s.min_ink)
        cands.append(marked)
        if best < s.min_ink:
            picks.append("")
            verdicts.append(BLANK)
        elif second >= s.min_ink and second / best > s.multi_ratio:
            picks.append("")
            verdicts.append(MULTIPLE)
        else:
            picks.append(PICK_LETTERS[int(order[0])])
            verdicts.append(HEAVY if best > s.heavy_ink else WEAK if best < s.weak_ink else OK)

    h_img, w_img = fin.shape[:2]
    y_name0 = int(max(0, g.name_top_y - 0.18 * g.row_h))
    name_crop = fin[y_name0:h_img, :].copy()
    y_title = int(max(g.title_bottom_y, 0.9 * g.row_h)) if g.title_bottom_y > 0 else int(g.row_h)
    title_crop = fin[0:max(10, y_title), :].copy()
    return SlipReading(picks=picks, verdicts=verdicts, ink=fractions, rectified=fin,
                       name_crop=name_crop, title_crop=title_crop, grid=g, candidates=cands)


# --------------------------------------------------------------------------- #
# Drawing helpers (UI + debugging)
# --------------------------------------------------------------------------- #
_VERDICT_COLOUR = {OK: (60, 180, 60), BLANK: (0, 0, 220), MULTIPLE: (0, 140, 255),
                   WEAK: (0, 200, 255), HEAVY: (0, 140, 255)}


def draw_reading(reading: SlipReading) -> np.ndarray:
    """Rectified slip with the detected pick circled in each row."""
    out = reading.rectified.copy()
    g = reading.grid
    for r in range(g.n_rows):
        colour = _VERDICT_COLOUR[reading.verdicts[r]]
        pk = reading.picks[r]
        cols = [PICK_LETTERS.index(pk)] if pk else [PICK_LETTERS.index(x) for x in reading.candidates[r]]
        if not cols:
            ctr = (g.corners[r, 0] + g.corners[r + 1, 3]) / 2
            cv2.putText(out, "?", (int(ctr[0]) - 8, int(ctr[1]) + 10), cv2.FONT_HERSHEY_SIMPLEX, 1.0, colour, 3)
        for c in cols:
            ctr = (g.corners[r, c] + g.corners[r, c + 1] + g.corners[r + 1, c] + g.corners[r + 1, c + 1]) / 4
            rad = int(0.36 * min(np.linalg.norm(g.corners[r, c + 1] - g.corners[r, c]), g.row_h))
            cv2.circle(out, (int(ctr[0]), int(ctr[1])), max(rad, 8), colour, 3)
        lab = f"{r + 1}:{pk or ('F1' if reading.verdicts[r] == BLANK else '??')}"
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
