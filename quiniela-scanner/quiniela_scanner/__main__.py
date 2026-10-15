"""Command line:

    python -m quiniela_scanner                 open the scanner window (webcam)
    python -m quiniela_scanner --demo slip.jpg try the whole app with a photo instead of a camera
    python -m quiniela_scanner scan slip.jpg   read one photo and print the result (no Excel)
    python -m quiniela_scanner excel           show how the open workbook was understood
    python -m quiniela_scanner cameras         list webcams

scanner.log (next to config.json) gets one line per entered slip with its timings.
"""
from __future__ import annotations

import argparse
import os
import sys

from .config import Config


def _write_image(path: str, img) -> bool:
    """cv2.imwrite returns False (no exception) for a path it cannot write and cannot do non-ANSI paths on Windows."""
    import cv2
    ext = os.path.splitext(path)[1] or ".png"
    try:
        ok, buf = cv2.imencode(ext, img)
        if ok:
            buf.tofile(path)
        return bool(ok)
    except (OSError, cv2.error):
        return False


def cmd_scan(args, cfg: Config) -> int:
    import cv2
    from . import slip_vision as sv
    img = cv2.imread(args.image)
    if img is None:
        print(f"cannot read {args.image}")
        return 2
    try:
        rd = sv.read_slip(img, cfg.matches_per_slip, sv.Settings(cfg.min_ink, cfg.multi_ratio))
    except sv.SlipNotFound as e:
        print(f"slip not recognised: {e}")
        return 1
    for i, (p, v, ink) in enumerate(zip(rd.picks, rd.verdicts, rd.ink), 1):
        print(f"  match {i}: {p or '-':1s}  {v:8s}  ink L/E/V = " + " / ".join(f"{x:.3f}" for x in ink))
    print("picks:", " ".join(p or "-" for p in rd.picks), "| clean" if rd.clean else "| NEEDS CHECK")
    if args.out:
        root, ext = os.path.splitext(args.out)
        name_out = f"{root}_name{ext or '.png'}"              # r.jpg -> r_name.jpg (also for ./dir.x/r.jpg)
        for path, img in ((args.out, sv.draw_reading(rd)), (name_out, rd.name_crop)):
            if not _write_image(path, img):
                print(f"cannot write {path}")
                return 2
            print("wrote", path)
    if args.name:
        from .name_reader import make_reader
        reader, note = make_reader(True, cfg.claude_model, cfg.claude_effort)
        if reader is None:
            print("name:", note)
        else:
            res = reader.read(rd.name_crop, rd.title_crop)
            print("name:", res.name or "(none)", "| jornada:", res.jornada, "| legible:", res.legible, res.error or "")
    return 0


def cmd_excel(args, cfg: Config) -> int:
    from .excel_sink import ExcelError, ExcelSink, open_backend
    try:
        be = open_backend(args.backend or cfg.backend, cfg.workbook, args.path or cfg.workbook_path)
        sink = ExcelSink(be, cfg.sheets, cfg.name_header)
    except Exception as e:                                # (not only ExcelError: this is the first check of a new setup)
        print("Excel problem:", e if isinstance(e, ExcelError) else f"{e.__class__.__name__}: {e}")
        return 1
    print("workbook:", be.label)
    for info in sink.block_infos():
        nxt = info.first_free.describe() if info.first_free else "FULL"
        print(" -", info.label, "\n     next free:", nxt)
    for w in sink.layout_warnings:                        # the same lines the window shows in its amber bar
        print("WARNING:", w)
    return 0


def cmd_cameras(args, cfg: Config) -> int:
    from .camera import list_cameras
    found = list_cameras(backend=cfg.camera_backend)
    print("cameras found:", found or "none")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="quiniela_scanner", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--demo", metavar="IMAGE", help="loop a photo as a fake camera")
    ap.add_argument("--xlsx", metavar="PATH", help="edit this .xlsx on disk instead of live Excel")
    ap.add_argument("--camera", type=int, help="camera index (default from config.json)")
    sub = ap.add_subparsers(dest="cmd")
    sp = sub.add_parser("scan", help="read one photo")
    sp.add_argument("image")
    sp.add_argument("--out", help="save the annotated, straightened slip here")
    sp.add_argument("--name", action="store_true", help="also read the handwritten name with Claude")
    sp = sub.add_parser("excel", help="show the detected ticket blocks")
    sp.add_argument("--backend", choices=["auto", "com", "file"])
    sp.add_argument("--path")
    sub.add_parser("cameras", help="list webcams")
    args = ap.parse_args(argv)

    from .config import load_dotenv
    try:
        env_note = load_dotenv()
    except Exception as e:                                # a bad .env must never stop the program from starting
        env_note = f".env could not be read ({e.__class__.__name__}) - the API key was not loaded."
    cfg = Config.load()                                   # an old config.json is migrated here (see config.py)
    if env_note:
        cfg.notes.append(env_note)
    if args.cmd:
        for note in cfg.notes:
            print("NOTE:", note, file=sys.stderr)
    if args.camera is not None:
        cfg.camera_index = args.camera
    if args.xlsx:
        cfg.override(backend="file", workbook_path=args.xlsx)        # this run only: never saved to config.json

    if args.cmd == "scan":
        return cmd_scan(args, cfg)
    if args.cmd == "excel":
        return cmd_excel(args, cfg)
    if args.cmd == "cameras":
        return cmd_cameras(args, cfg)

    from . import gui
    source = None
    if args.demo:
        from .camera import ImageSource
        source = ImageSource(args.demo, demo=True)
    gui.run(cfg, source)
    return 0


if __name__ == "__main__":
    sys.exit(main())
