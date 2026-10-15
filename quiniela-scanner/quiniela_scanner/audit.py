"""Keeps a permanent record of every slip that was entered: a straightened photo with the
detected picks circled, plus one CSV row. Handy for settling any "that's not what I marked".

Also ``ScanLog`` (scanner.log): one line per entered slip with its timings - what to send when the
app feels slow or something went wrong on the real PC."""
from __future__ import annotations

import csv
import re
import threading
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


class ScanLog:
    """Append-only text log, safe to call from any thread:  ``2026-10-15 12:00:01  ENTRY  ticket=251 read_ms=190 ...``"""

    MAX_BYTES = 2_000_000                  # older lines move to scanner.log.1 once the file gets this big

    def __init__(self, path: Path):
        self.path = Path(path)
        self._lock = threading.Lock()

    @staticmethod
    def _fmt(v) -> str:
        if isinstance(v, float):
            v = round(v, 1)
        s = str(v)
        return f'"{s}"' if (not s or " " in s or '"' in s) else s

    def line(self, event: str, /, **fields) -> None:
        text = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {event}  " + \
               " ".join(f"{k}={self._fmt(v)}" for k, v in fields.items() if v is not None)
        with self._lock:
            try:
                if self.path.exists() and self.path.stat().st_size > self.MAX_BYTES:
                    self.path.replace(self.path.with_name(self.path.name + ".1"))
                with open(self.path, "a", encoding="utf-8") as f:
                    f.write(text.replace("\n", " ") + "\n")
            except OSError:
                pass
