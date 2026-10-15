"""Test fixtures: a pool workbook laid out like the real one, and synthetic slip photos."""
from __future__ import annotations

import cv2
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter

TEAMS = [("NECAXA", "ATLAS"), ("ATLANTE", "PUMAS"), ("TIJUANA", "PUEBLA"), ("CHIVAS", "TIGRES"),
         ("SANTOS", "QUERETARO"), ("LEON", "AMERICA"), ("TOLUCA", "SAN LUIS"),
         ("CRUZ AZUL", "JUAREZ"), ("MONTERREY", "PACHUCA")]


def make_pool_workbook(path, filled=(10, 3, 0), first_tickets=(241, 481, 721), rows=52, suffixes=("RRR", "CH", "")):
    """ENTRY PG1 with three side-by-side blocks:  # | NOMBRE | 1..9 | T   (header on row 4)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ENTRY PG1"
    wb.create_sheet("ENTRY PG2")
    wb.create_sheet("RESULTS")
    wb.create_sheet("SLIPS")
    ws["N3"] = "SOLO DIVERSION"
    start_col = 13                              # column M, like the real sheet
    for b, (first, nfilled, suf) in enumerate(zip(first_tickets, filled, suffixes)):
        c0 = start_col + b * 12                 # '#', NOMBRE, 9 picks, T -> 12 columns
        ws.cell(4, c0, "#")
        ws.cell(4, c0 + 1, "NOMBRE")
        for k in range(9):
            ws.cell(4, c0 + 2 + k, k + 1)
        ws.cell(4, c0 + 11, "T")
        for i in range(rows):
            r = 5 + i
            ws.cell(r, c0, first + i)
            pick_first = get_column_letter(c0 + 2)
            pick_last = get_column_letter(c0 + 10)
            ws.cell(r, c0 + 11, f'=COUNTIF({pick_first}{r}:{pick_last}{r},"V")')
            if i < nfilled:
                ws.cell(r, c0 + 1, f"PLAYER {i + 1} {suf}".strip())
                for k in range(9):
                    ws.cell(r, c0 + 2 + k, "LEV"[(i + k) % 3])
    wb.save(path)
    return path


# ----------------------------------------------------------------------------- slips
def draw_slip(picks, name="CLEMENTE O.", ink=(170, 60, 25), width=1000, rng=None, jitter=True):
    """Draw an upright slip like the real form. picks: list of 9 items, each '' / 'L' / 'E' / 'V' / 'LV' (two marks)."""
    rng = rng or np.random.default_rng(0)
    h = 1050
    img = np.full((h, width, 3), 238, np.uint8)
    blk = (25, 25, 25)
    x = [8, 72, 136, 200, 585, 990]            # vertical rules: border, L|E, E|V, V|team, home|away, border
    y0, ytitle, yhead = 8, 85, 160
    row_h = 80
    ys = [yhead + i * row_h for i in range(10)]    # 10 lines bound the 9 match rows
    ybot = ys[-1] + 90                              # NOMBRE row bottom
    t = 4
    cv2.rectangle(img, (x[0], y0), (x[-1], ybot), blk, t)
    cv2.line(img, (x[0], ytitle), (x[-1], ytitle), blk, t)
    for yy in ys:
        cv2.line(img, (x[0], yy), (x[-1], yy), blk, t)
    for xx in x[1:5]:
        cv2.line(img, (xx, ytitle), (xx, ys[-1]), blk, t)
    cv2.putText(img, "LIGA MX JORNADA 12", (480, 62), cv2.FONT_HERSHEY_DUPLEX, 1.1, blk, 3)
    for lab, xx in zip("LEV", (22, 88, 152)):
        cv2.putText(img, lab, (xx, 125), cv2.FONT_HERSHEY_DUPLEX, 0.9, blk, 3)
    for i, (home, away) in enumerate(TEAMS):
        base = ys[i] + 52
        cv2.putText(img, home, (215, base), cv2.FONT_HERSHEY_DUPLEX, 1.1, blk, 3)
        cv2.putText(img, away, (600, base + 10), cv2.FONT_HERSHEY_DUPLEX, 1.1, blk, 3)
    cv2.putText(img, "NOMBRE", (14, ybot - 28), cv2.FONT_HERSHEY_DUPLEX, 0.95, blk, 3)
    cv2.line(img, (170, ybot - 22), (900, ybot - 22), blk, 2)
    cv2.putText(img, name, (260, ybot - 30), cv2.FONT_HERSHEY_SCRIPT_COMPLEX, 2.2, ink, 4)
    cv2.putText(img, "ULTIMO DIA MIERCOLES EL 14", (190, ybot + 40), cv2.FONT_HERSHEY_DUPLEX, 0.9, blk, 3)
    for i, p in enumerate(picks):
        for letter in p:
            c = "LEV".index(letter)
            cx = (x[c] + x[c + 1]) // 2
            cy = (ys[i] + ys[i + 1]) // 2
            jx, jy = (rng.integers(-8, 9), rng.integers(-8, 9)) if jitter else (0, 0)
            s = 24
            cv2.line(img, (cx - s + jx, cy - s + jy), (cx + s + jx, cy + s + jy), ink, 4)
            cv2.line(img, (cx + s + jx, cy - s + jy), (cx - s + jx, cy + s + jy), ink, 4)
    return img


def on_webcam_canvas(slip, scale=0.5, angle=0.0, persp=0.0, bg=235, blur=0.0, noise=0.0, seed=0, flip=False):
    """Place the slip on a 1920x1080 'webcam' frame with rotation / perspective / blur / noise."""
    rng = np.random.default_rng(seed)
    if flip:
        slip = cv2.rotate(slip, cv2.ROTATE_180)
    s = cv2.resize(slip, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    pad = int(0.3 * max(s.shape[:2]))                       # room to rotate without clipping the slip
    s = cv2.copyMakeBorder(s, pad, pad, pad, pad, cv2.BORDER_CONSTANT, value=(bg,) * 3)
    h, w = s.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    s = cv2.warpAffine(s, m, (w, h), borderValue=(bg,) * 3)
    if persp:
        src = np.float32([[0, 0], [w, 0], [w, h], [0, h]])
        dst = src + rng.uniform(-persp * min(w, h), persp * min(w, h), src.shape).astype(np.float32)
        s = cv2.warpPerspective(s, cv2.getPerspectiveTransform(src, dst), (w, h), borderValue=(bg,) * 3)
    canvas = np.full((1080, 1920, 3), bg, np.uint8)
    ox, oy = (1920 - w) // 2, (1080 - h) // 2
    canvas[oy:oy + h, ox:ox + w] = s
    if blur:
        canvas = cv2.GaussianBlur(canvas, (0, 0), blur)
    if noise:
        canvas = np.clip(canvas + rng.normal(0, noise, canvas.shape), 0, 255).astype(np.uint8)
    return canvas
