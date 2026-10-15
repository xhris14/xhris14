"""Keeps a permanent record of every slip that was entered: a straightened photo with the
detected picks circled, plus one CSV row. Handy for settling any "that's not what I marked".

Also ``ScanLog`` (scanner.log): one line per entered slip with its timings - what to send when the
app feels slow or something went wrong on the real PC."""
from __future__ import annotations

import csv
import re
import shutil
import threading
from datetime import datetime
from pathlib import Path

import cv2

from . import slip_vision as sv


class AuditLog:
    """One JPEG + one CSV row per slip. ``record`` never raises: it returns the picture path and leaves a sentence in
    ``last_error`` ("" = all fine) for the caller to log and show - a lost audit row must not go unnoticed."""

    FIELDS = ["timestamp", "workbook", "sheet", "row", "ticket", "name", "picks", "status", "image"]
    LOW_DISK_MB = 1000                       # warn below this much free space (a slip picture is ~200 KB)
    DISK_CHECK_EVERY = 50

    def __init__(self, base: Path):
        self.base = Path(base)
        self.csv_path = self.base / "log.csv"
        self.last_error = ""
        self._fallback: Path | None = None   # log-<time>.csv, used while log.csv cannot be appended to
        self._n = 0

    def record(self, reading: sv.SlipReading, picks: list, name: str, workbook: str, sheet: str,
               row: int, ticket, status: str) -> str:
        now = datetime.now()
        errors: list = []
        img_path = ""
        day = self.base / now.strftime("%Y-%m-%d")
        try:
            day.mkdir(parents=True, exist_ok=True)
            safe = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")[:30] or "slip"
            path = day / f"{now.strftime('%H%M%S')}_T{ticket if ticket is not None else 'x'}_{safe}.jpg"
            ok, buf = cv2.imencode(".jpg", sv.draw_reading(reading), [cv2.IMWRITE_JPEG_QUALITY, 85])
            if not ok:
                raise OSError("could not encode the picture")
            path.write_bytes(buf.tobytes())      # (cv2.imwrite fails silently, and cannot write non-ANSI paths on Windows)
            img_path = str(path)
        except (OSError, cv2.error) as e:
            errors.append(f"picture not saved ({e.__class__.__name__}: {e})")
        row_values = [now.isoformat(timespec="seconds"), workbook, sheet, row, ticket if ticket is not None else "",
                      name, " ".join(p or "-" for p in picks), status, img_path]
        target = self._fallback or self.csv_path
        try:
            self._append(target, row_values)
        except OSError as e:                     # on Windows: log.csv open in Excel is locked (PermissionError)
            self._fallback = self.base / f"log-{now.strftime('%Y%m%d-%H%M%S')}.csv"
            try:
                self._append(self._fallback, row_values)
                errors.append(f"log.csv cannot be written ({e.__class__.__name__}: close it in Excel) - "
                              f"this and the next rows go to {self._fallback.name}")
            except OSError as e2:
                self._fallback = None
                errors.append(f"audit row not saved ({e2.__class__.__name__}: {e2})")
        self._n += 1
        if self._n % self.DISK_CHECK_EVERY == 1:
            try:
                free = shutil.disk_usage(self.base).free / 1e6
                if free < self.LOW_DISK_MB:
                    errors.append(f"only {free:.0f} MB free on the disk that holds scans/")
            except OSError:
                pass
        self.last_error = "; ".join(errors)
        return img_path

    def _append(self, path: Path, values: list) -> None:
        self.base.mkdir(parents=True, exist_ok=True)
        new = not path.exists()
        with open(path, "a", newline="", encoding="utf-8-sig" if new else "utf-8") as f:   # BOM: Excel shows accents right
            w = csv.writer(f)
            if new:
                w.writerow(self.FIELDS)
            w.writerow(values)


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
