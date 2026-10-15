"""Getting scanned slips into the Excel pool sheet.

The sheet is laid out as side-by-side *ticket blocks*:

    #  | NOMBRE | 1 | 2 | ... | 9 | T        (header row)
    241| RECA 30| V | V | ... | L | 5        (one row per ticket)

``discover_blocks`` finds every block from its header (so nothing is hard-coded to a
column letter), ``ExcelSink`` picks the next free ticket and writes name + 9 picks.
T (and anything else to the right) is left alone - it is normally a formula.

Two backends:
  * ComBackend      - drives the Excel window you already have open (Windows, pywin32).
  * OpenpyxlBackend - edits an .xlsx file on disk (workbook must be closed in Excel).
"""
from __future__ import annotations

import fnmatch
import re
import time
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

N_PICKS = 9


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def norm(v):
    """Cell value -> None | int | float | str (whole floats become ints)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


_FORMULA_OR_NUMBERISH = re.compile(r"^[=+\-@]|^[\d\s\-/.:,]+$")


def needs_text_prefix(s: str) -> bool:
    """True if Excel would turn this string into a formula / number / date."""
    return bool(_FORMULA_OR_NUMBERISH.match(s))


class ExcelError(Exception):
    pass


class SlotTaken(ExcelError):
    pass


# --------------------------------------------------------------------------- #
# Snapshot + layout discovery
# --------------------------------------------------------------------------- #
@dataclass
class Snapshot:
    top: int                       # 1-based row of values[0]
    left: int                      # 1-based column of values[0][0]
    values: list

    def get(self, row: int, col: int):
        r, c = row - self.top, col - self.left
        if 0 <= r < len(self.values) and 0 <= c < len(self.values[r]):
            return norm(self.values[r][c])
        return None

    @property
    def last_row(self) -> int:
        return self.top + len(self.values) - 1

    @property
    def last_col(self) -> int:
        return self.left + (max(len(r) for r in self.values) if self.values else 0) - 1


@dataclass
class Block:
    sheet: str
    header_row: int
    name_col: int
    ticket_col: int | None
    rows: list

    @property
    def pick_col(self) -> int:
        return self.name_col + 1

    @property
    def key(self) -> str:
        return f"{self.sheet}!{col_letter(self.name_col)}{self.header_row}"


@dataclass
class Slot:
    block: Block
    row: int
    ticket: int | None

    def describe(self) -> str:
        b = self.block
        t = f"ticket {self.ticket}" if self.ticket is not None else "next row"
        return f"{t}  ({b.sheet} row {self.row}, name in column {col_letter(b.name_col)})"


@dataclass
class BlockInfo:
    block: Block
    free: int
    total: int
    first_free: Slot | None
    suffix: str                    # dominant trailing word of existing names (e.g. "RRR")
    label: str


def discover_blocks(sheet: str, snap: Snapshot, name_header: str = "NOMBRE",
                    rows_without_tickets: int = 52, max_header_row: int = 500) -> list:
    want = name_header.strip().upper()
    blocks = []
    last_hdr = min(snap.last_row, max_header_row)
    for r in range(snap.top, last_hdr + 1):
        for c in range(snap.left, snap.last_col + 1):
            v = snap.get(r, c)
            if not isinstance(v, str) or v.upper() != want:
                continue
            if any(str(snap.get(r, c + 1 + k)) != str(k + 1) for k in range(N_PICKS)):
                continue
            tc = c - 1 if str(snap.get(r, c - 1) or "").upper() in ("#", "NO", "NO.", "TICKET", "FOLIO") else None
            rows = []
            if tc is not None:
                rr = r + 1
                while isinstance(snap.get(rr, tc), int):
                    rows.append(rr)
                    rr += 1
            else:
                rows = list(range(r + 1, r + 1 + rows_without_tickets))
            if rows:
                blocks.append(Block(sheet, r, c, tc, rows))
    return blocks


def slot_is_free(snap: Snapshot, block: Block, row: int) -> bool:
    return all(snap.get(row, block.name_col + k) is None for k in range(N_PICKS + 1))


def _dominant_suffix(names: list) -> str:
    toks = [n.split()[-1] for n in names if isinstance(n, str) and len(n.split()) > 1]
    if len(toks) < 3:
        return ""
    word, cnt = Counter(toks).most_common(1)[0]
    return word if cnt / len(names) >= 0.5 and not word.isdigit() else ""


# --------------------------------------------------------------------------- #
# Backends
# --------------------------------------------------------------------------- #
class OpenpyxlBackend:
    """Edit an .xlsx on disk. Fine for testing / when Excel is closed; the workbook must NOT be
    open in Excel while writing. openpyxl may drop pictures/charts, so prefer ComBackend."""

    def __init__(self, path: str):
        import openpyxl
        self._openpyxl = openpyxl
        self.path = str(path)
        self.label = Path(path).name + " (file)"
        self.workbook_name = Path(path).name
        self._load()

    def _load(self):
        self._wb = self._openpyxl.load_workbook(self.path)
        self._wbv = self._openpyxl.load_workbook(self.path, data_only=True)

    def sheet_names(self):
        return list(self._wb.sheetnames)

    def snapshot(self, sheet: str, max_rows: int = 600) -> Snapshot:
        self._load()
        ws = self._wbv[sheet]
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows),
                                              max_col=ws.max_column, values_only=True)]
        return Snapshot(1, 1, rows or [[None]])

    def write_row(self, sheet: str, row: int, col: int, values: list):
        self._load()
        ws = self._wb[sheet]
        for i, v in enumerate(values):
            cell = ws.cell(row=row, column=col + i)
            cell.value = v
            if isinstance(v, str):
                cell.data_type = "s"          # never let a name like "=1" become a formula
        self._wb.save(self.path)

    def save(self):
        pass


class ComBackend:
    """Drive the Excel instance that is already running (Windows + pywin32)."""

    BUSY = {-2147418111, -2147417846, -2146777998}   # call rejected / busy / "Excel is busy"

    def __init__(self, workbook_hint: str = "", app=None):
        if app is None:
            try:
                import pythoncom
                import win32com.client
            except ImportError as e:
                raise ExcelError("pywin32 is not installed (pip install pywin32)") from e
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            try:
                app = win32com.client.GetActiveObject("Excel.Application")
            except Exception as e:
                raise ExcelError("Excel is not running - open your jornada workbook first") from e
        self.app = app
        wb = self._pick_workbook(workbook_hint)
        self.workbook_name = str(wb.Name)
        self.label = f"{self.workbook_name} (Excel, live)"

    def _pick_workbook(self, hint: str):
        count = self.app.Workbooks.Count
        if count == 0:
            raise ExcelError("Excel is running but no workbook is open")
        if not hint:
            return self.app.ActiveWorkbook
        for i in range(1, count + 1):
            wb = self.app.Workbooks(i)
            if hint.lower() in str(wb.Name).lower():
                return wb
        raise ExcelError(f"no open workbook matches '{hint}'")

    @property
    def wb(self):
        return self.app.Workbooks(self.workbook_name)

    def _retry(self, fn):
        last = None
        for _ in range(10):
            try:
                return fn()
            except Exception as e:                      # pywintypes.com_error
                hr = getattr(e, "hresult", None)
                if hr in self.BUSY or (hr is None and "busy" in str(e).lower()):
                    last = e
                    time.sleep(0.4)
                    continue
                raise ExcelError(f"Excel error: {e}") from e
        raise ExcelError("Excel is busy (is a cell being edited?). Press Esc in Excel and retry.") from last

    def sheet_names(self):
        def go():
            ws = self.wb.Worksheets
            return [str(ws.Item(i).Name) for i in range(1, ws.Count + 1)]
        return self._retry(go)

    def snapshot(self, sheet: str, max_rows: int = 600) -> Snapshot:
        def go():
            ws = self.wb.Worksheets(sheet)
            ur = ws.UsedRange
            top, left = int(ur.Row), int(ur.Column)
            nrows = min(int(ur.Rows.Count), max_rows)
            ncols = int(ur.Columns.Count)
            rng = ws.Range(ws.Cells(top, left), ws.Cells(top + nrows - 1, left + ncols - 1))
            v = rng.Value
            if v is None:
                vals = [[None]]
            elif not isinstance(v, tuple):
                vals = [[v]]
            else:
                vals = [list(r) for r in v]
            return Snapshot(top, left, vals)
        return self._retry(go)

    def write_row(self, sheet: str, row: int, col: int, values: list):
        def go():
            ws = self.wb.Worksheets(sheet)
            vals = [("'" + v) if isinstance(v, str) and needs_text_prefix(v) else v for v in values]
            ws.Range(ws.Cells(row, col), ws.Cells(row, col + len(vals) - 1)).Value = (tuple(vals),)
        self._retry(go)

    def save(self):
        self._retry(lambda: self.wb.Save())


def open_backend(kind: str, workbook: str = "", path: str = ""):
    """kind: auto | com | file"""
    import sys
    if kind == "file" or (kind == "auto" and sys.platform != "win32"):
        if not path:
            raise ExcelError("backend 'file' needs workbook_path in config.json")
        return OpenpyxlBackend(path)
    return ComBackend(workbook)


# --------------------------------------------------------------------------- #
# The sink
# --------------------------------------------------------------------------- #
@dataclass
class Receipt:
    slot: Slot
    sheet: str
    row: int
    col: int
    previous: list
    written: list
    ticket: int | None


class ExcelSink:
    def __init__(self, backend, sheet_patterns=("ENTRY*",), name_header="NOMBRE",
                 uppercase: bool = True, save_after_write: bool = False):
        self.backend = backend
        self.patterns = list(sheet_patterns) or ["*"]
        self.name_header = name_header
        self.uppercase = uppercase
        self.save_after_write = save_after_write
        self.snaps: dict = {}
        self.blocks: list = []
        self.refresh()

    # -- reading the layout ------------------------------------------------
    def refresh(self):
        self.snaps, self.blocks = {}, []
        names = [n for n in self.backend.sheet_names()
                 if any(fnmatch.fnmatch(n.upper(), p.upper()) for p in self.patterns)]
        if not names:
            raise ExcelError(f"no sheet matches {self.patterns}; sheets are: {self.backend.sheet_names()}")
        for n in names:
            snap = self.backend.snapshot(n)
            self.snaps[n] = snap
            self.blocks += discover_blocks(n, snap, self.name_header)
        if not self.blocks:
            raise ExcelError(f"no '{self.name_header} | 1 2 3 ... 9' header found on {names}")
        return self.blocks

    def _ticket(self, block: Block, row: int):
        if block.ticket_col is None:
            return None
        t = self.snaps[block.sheet].get(row, block.ticket_col)
        return t if isinstance(t, int) else None

    def block_infos(self) -> list:
        infos = []
        for b in self.blocks:
            snap = self.snaps[b.sheet]
            free = [r for r in b.rows if slot_is_free(snap, b, r)]
            tickets = [t for t in (self._ticket(b, r) for r in b.rows) if t is not None]
            names = [snap.get(r, b.name_col) for r in b.rows]
            names = [n for n in names if n]
            rng = f"tickets {min(tickets)}-{max(tickets)}" if tickets else f"rows {b.rows[0]}-{b.rows[-1]}"
            first = Slot(b, free[0], self._ticket(b, free[0])) if free else None
            suffix = _dominant_suffix(names)
            label = f"{b.sheet}  {col_letter(b.name_col)}:  {rng}  -  {len(free)} free"
            if suffix:
                label += f"  [{suffix}]"
            infos.append(BlockInfo(b, len(free), len(b.rows), first, suffix, label))
        infos.sort(key=lambda i: (i.block.ticket_col is None,
                                  self._ticket(i.block, i.block.rows[0]) or 0, i.block.name_col))
        return infos

    def default_block(self):
        infos = self.block_infos()
        return next((i for i in infos if i.free), infos[0] if infos else None)

    def next_slot(self, block: Block | None = None) -> Slot | None:
        for info in self.block_infos():
            if block is not None and info.block.key != block.key:
                continue
            if info.first_free:
                return info.first_free
        return None

    def slot_for_ticket(self, ticket: int) -> Slot | None:
        for b in self.blocks:
            for r in b.rows:
                if self._ticket(b, r) == ticket:
                    return Slot(b, r, ticket)
        return None

    def is_free(self, slot: Slot) -> bool:
        return slot_is_free(self.snaps[slot.block.sheet], slot.block, slot.row)

    # -- writing -----------------------------------------------------------
    def format_name(self, name: str, suffix: str = "") -> str:
        name = " ".join(name.split())
        if self.uppercase:
            name = name.upper()
        suffix = suffix.strip()
        if suffix and not name.upper().endswith(suffix.upper()):
            name = f"{name} {suffix}"
        return name

    def write(self, slot: Slot, name: str, picks: list, allow_overwrite: bool = False) -> Receipt:
        b = slot.block
        # Always re-read right before writing so manual edits are respected.
        self.snaps[b.sheet] = self.backend.snapshot(b.sheet)
        snap = self.snaps[b.sheet]
        if not allow_overwrite and not slot_is_free(snap, b, slot.row):
            raise SlotTaken(f"{slot.describe()} already has data")
        previous = [snap.get(slot.row, b.name_col + k) for k in range(N_PICKS + 1)]
        letters = [p if p in ("L", "E", "V") else None for p in picks]
        letters += [None] * (N_PICKS - len(letters))
        values = [name] + letters[:N_PICKS]
        self.backend.write_row(b.sheet, slot.row, b.name_col, values)
        if self.save_after_write:
            self.backend.save()
        self.snaps[b.sheet] = self.backend.snapshot(b.sheet)
        return Receipt(slot, b.sheet, slot.row, b.name_col, previous, values, slot.ticket)

    def undo(self, r: Receipt):
        self.backend.write_row(r.sheet, r.row, r.col, list(r.previous))
        if self.save_after_write:
            self.backend.save()
        self.snaps[r.sheet] = self.backend.snapshot(r.sheet)
