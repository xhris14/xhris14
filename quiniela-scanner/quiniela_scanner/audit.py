"""Keeps a permanent record of every slip that was entered: a straightened photo with the
detected picks circled, plus one CSV row. Handy for settling any "that's not what I marked"."""
from __future__ import annotations

import csv
import re
from datetime import datetime
from pathlib import Path

import cv2

from . import slip_vision as sv


class AuditLog:
    FIELDS = ["timestamp", "workbook", "sheet", "row", "ticket", "name", "picks", "status", "image"]

    def __init__(self, base: Path):
        self.base = Path(base)
        self.csv_path = self.base / "log.csv"

    def record(self, reading: sv.SlipReading, picks: list, name: str, workbook: str, sheet: str,
               row: int, ticket, status: str) -> str:
        now = datetime.now()
        day = self.base / now.strftime("%Y-%m-%d")
        try:
            day.mkdir(parents=True, exist_ok=True)
            safe = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")[:30] or "slip"
            img_path = day / f"{now.strftime('%H%M%S')}_T{ticket if ticket is not None else 'x'}_{safe}.jpg"
            cv2.imwrite(str(img_path), sv.draw_reading(reading), [cv2.IMWRITE_JPEG_QUALITY, 85])
            new = not self.csv_path.exists()
            with open(self.csv_path, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if new:
                    w.writerow(self.FIELDS)
                w.writerow([now.isoformat(timespec="seconds"), workbook, sheet, row,
                            ticket if ticket is not None else "", name,
                            " ".join(p or "-" for p in picks), status, str(img_path)])
            return str(img_path)
        except OSError:
            return ""
