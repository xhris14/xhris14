"""Command line:

    python -m quiniela_scanner                 open the scanner window (webcam)
    python -m quiniela_scanner --demo slip.jpg try the whole app with a photo instead of a camera
    python -m quiniela_scanner scan slip.jpg   read one photo and print the result (no Excel)
    python -m quiniela_scanner excel           show how the open workbook was understood
    python -m quiniela_scanner cameras         list webcams
"""
from __future__ import annotations

import argparse
import sys

from .config import Config


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
        cv2.imwrite(args.out, sv.draw_reading(rd))
        cv2.imwrite(args.out.replace(".", "_name.", 1), rd.name_crop)
        print("wrote", args.out)
    if args.name:
        from .name_reader import make_reader
        reader, note = make_reader(True, cfg.claude_model)
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
    except ExcelError as e:
        print("Excel problem:", e)
        return 1
    print("workbook:", be.label)
    for info in sink.block_infos():
        nxt = info.first_free.describe() if info.first_free else "FULL"
        print(" -", info.label, "\n     next free:", nxt)
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
    load_dotenv()
    cfg = Config.load()
    if args.camera is not None:
        cfg.camera_index = args.camera
    if args.xlsx:
        cfg.backend, cfg.workbook_path = "file", args.xlsx

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
