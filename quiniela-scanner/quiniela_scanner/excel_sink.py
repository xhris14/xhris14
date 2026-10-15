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

Speed: the full-sheet snapshot is only a CACHE for choosing the next free ticket. It is taken once (and on
refresh()); a write then costs one small read (is the slot still free and still that ticket?) plus one write.
The name can be filled in later than the picks with update_name().
"""
from __future__ import annotations

import fnmatch
import re
import time
from collections import Counter
from copy import copy
from dataclasses import dataclass, field
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


def format_name(name: str, tag: str = "", position: str = "after", sep: str = " ",
                uppercase: bool = True) -> str:
    """Tidy the name and add the group tag before or after it.

    The tag is skipped only if the slip already carries it as a separate word at that end
    (so a tag of "E" is still added to "JOSE" or "ELENA")."""
    name = " ".join(name.split())
    tag = tag.strip()
    if uppercase:
        name, tag = name.upper(), tag.upper()
    if not tag or not name:
        return name
    words = name.split(" ")
    edge = words[0] if position == "before" else words[-1]
    if edge.upper() == tag.upper():
        return name
    return f"{tag}{sep}{name}" if position == "before" else f"{name}{sep}{tag}"


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

    def blank(self, row: int, col: int, n: int) -> bool:
        """True if the n cells from (row, col) are all empty (same rule as norm(), without building values)."""
        r, c = row - self.top, col - self.left
        if not 0 <= r < len(self.values):
            return True
        for v in self.values[r][max(c, 0):max(c + n, 0)]:
            if v is not None and not (isinstance(v, str) and not v.strip()):
                return False
        return True

    def set(self, row: int, col: int, value):
        """Update one cell of the cache in place (grows the grid if the cell is past its edge)."""
        r, c = row - self.top, col - self.left
        if r < 0 or c < 0:
            return
        while len(self.values) <= r:
            self.values.append([])
        line = self.values[r]
        if len(line) <= c:
            line.extend([None] * (c + 1 - len(line)))
        line[c] = value


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
    return snap.blank(row, block.name_col, N_PICKS + 1)


def _dominant_suffix(names: list) -> str:
    toks = [n.split()[-1] for n in names if isinstance(n, str) and len(n.split()) > 1]
    if len(toks) < 3:
        return ""
    word, cnt = Counter(toks).most_common(1)[0]
    return word if cnt / len(names) >= 0.5 and not word.isdigit() else ""


# --------------------------------------------------------------------------- #
# Backends
# --------------------------------------------------------------------------- #
def _a1(row: int, col: int, n: int = 1) -> str:
    """'N12' or, for n > 1 cells to the right, 'N12:W12'."""
    first = f"{col_letter(col)}{row}"
    return first if n == 1 else f"{first}:{col_letter(col + n - 1)}{row}"


def _a1_areas(row: int, col: int, offsets: list) -> str:
    """Multi-area address of the cells col+offset on one row: 'P5,R5:T5'."""
    runs, start, prev = [], None, None
    for off in sorted(offsets):
        if start is None:
            start = prev = off
        elif off == prev + 1:
            prev = off
        else:
            runs.append((start, prev))
            start = prev = off
    if start is not None:
        runs.append((start, prev))
    return ",".join(_a1(row, col + a, b - a + 1) for a, b in runs)


class OpenpyxlBackend:
    """Edit an .xlsx on disk. Fine for testing / when Excel is closed; the workbook must NOT be
    open in Excel while writing. openpyxl may drop pictures/charts, so prefer ComBackend."""

    def __init__(self, path: str):
        import openpyxl
        self._openpyxl = openpyxl
        self.path = str(path)
        self.label = Path(path).name + " (file)"
        self.workbook_name = Path(path).name
        self.calls = Counter()                 # public-method call counts, for tests / diagnostics
        self._load()

    def _load(self):
        self._load_formulas()
        self._load_values()

    def _load_formulas(self):
        self._wb = self._openpyxl.load_workbook(self.path)

    def _load_values(self):
        self._wbv = self._openpyxl.load_workbook(self.path, data_only=True)

    def sheet_names(self):
        self.calls["sheet_names"] += 1
        self._load_formulas()
        return list(self._wb.sheetnames)

    def snapshot(self, sheet: str, max_rows: int = 600) -> Snapshot:
        self.calls["snapshot"] += 1
        self._load_values()
        ws = self._wbv[sheet]
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows),
                                              max_col=ws.max_column, values_only=True)]
        return Snapshot(1, 1, rows or [[None]])

    def read_row(self, sheet: str, row: int, col: int, n: int) -> list:
        """n raw cell values of one row, starting at (row, col)."""
        self.calls["read_row"] += 1
        self._load_values()
        ws = self._wbv[sheet]
        return [ws.cell(row=row, column=col + i).value for i in range(n)]

    def font_sizes(self, sheet: str, row: int, col: int, n: int) -> list:
        self.calls["font_sizes"] += 1
        self._load_formulas()
        ws = self._wb[sheet]
        return [ws.cell(row=row, column=col + i).font.sz for i in range(n)]

    def write_row(self, sheet: str, row: int, col: int, values: list, font_sizes: dict | None = None):
        self.calls["write_row"] += 1
        self._load_formulas()
        ws = self._wb[sheet]
        for i, v in enumerate(values):
            cell = ws.cell(row=row, column=col + i)
            cell.value = v
            if isinstance(v, str):
                cell.data_type = "s"          # never let a name like "=1" become a formula
        for off, size in (font_sizes or {}).items():
            cell = ws.cell(row=row, column=col + off)
            f = copy(cell.font)
            f.sz = size
            cell.font = f
        self._wb.save(self.path)

    def save(self):
        self.calls["save"] += 1


class ComBackend:
    """Drive the Excel instance that is already running (Windows + pywin32)."""

    BUSY = {-2147418111, -2147417846, -2146777998}   # call rejected / busy / "Excel is busy"
    RETRY_TRIES = 6                                  # COM runs on the Tk thread: ~1.2 s at most, then the app
    RETRY_DELAY = 0.2                                # shows the problem instead of freezing

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
        self.calls = Counter()                 # public-method call counts, for tests / diagnostics
        self._sheets: dict = {}                # sheet name -> Worksheet COM object (saves 2 round trips per call)
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

    def _ws(self, sheet: str):
        ws = self._sheets.get(sheet)
        if ws is None:
            ws = self._sheets[sheet] = self.wb.Worksheets(sheet)
        return ws

    def _retry(self, fn):
        last = None
        dropped = False
        for _ in range(self.RETRY_TRIES):
            try:
                return fn()
            except Exception as e:                      # pywintypes.com_error
                hr = getattr(e, "hresult", None)
                if hr in self.BUSY or (hr is None and "busy" in str(e).lower()):
                    last = e
                    time.sleep(self.RETRY_DELAY)
                    continue
                if self._sheets and not dropped:        # a cached sheet may be dead (workbook reopened): resolve again
                    self._sheets.clear()
                    dropped = True
                    continue
                raise ExcelError(f"Excel error: {e}") from e
        raise ExcelError("Excel is busy (is a cell being edited?). Press Esc in Excel and retry.") from last

    def _screen_updating(self, on: bool):
        try:
            self.app.ScreenUpdating = on
        except Exception:
            pass

    def sheet_names(self):
        self.calls["sheet_names"] += 1

        def go():
            ws = self.wb.Worksheets
            return [str(ws.Item(i).Name) for i in range(1, ws.Count + 1)]
        return self._retry(go)

    def snapshot(self, sheet: str, max_rows: int = 600) -> Snapshot:
        self.calls["snapshot"] += 1
        self._sheets.pop(sheet, None)                   # a (re)connect re-resolves the sheet object

        def go():
            ws = self._ws(sheet)
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

    def read_row(self, sheet: str, row: int, col: int, n: int) -> list:
        """n raw cell values of one row, starting at (row, col), with ONE Range read."""
        self.calls["read_row"] += 1

        def go():
            v = self._ws(sheet).Range(_a1(row, col, n)).Value
            if isinstance(v, tuple):                    # n > 1 -> ((a, b, ...),)
                v = list(v[0]) if v and isinstance(v[0], tuple) else list(v)
            else:                                       # n == 1 -> the bare value
                v = [v]
            return (v + [None] * n)[:n]
        return self._retry(go)

    def font_sizes(self, sheet: str, row: int, col: int, n: int) -> list:
        self.calls["font_sizes"] += 1

        def go():
            ws = self._ws(sheet)
            size = ws.Range(_a1(row, col, n)).Font.Size        # Excel answers Null when the sizes differ
            if size:
                return [size] * n
            return [ws.Range(_a1(row, col + i)).Font.Size for i in range(n)]
        return self._retry(go)

    def write_row(self, sheet: str, row: int, col: int, values: list, font_sizes: dict | None = None):
        self.calls["write_row"] += 1

        def go():
            ws = self._ws(sheet)
            vals = [("'" + v) if isinstance(v, str) and needs_text_prefix(v) else v for v in values]
            by_size: dict = {}
            for off, size in sorted((font_sizes or {}).items()):
                by_size.setdefault(size, []).append(off)
            self._screen_updating(False)
            try:
                rng = ws.Range(_a1(row, col, len(vals)))
                if len(vals) == 1:
                    rng.Value = vals[0]
                else:
                    rng.Value = (tuple(vals),)
                for size, offs in by_size.items():              # one multi-area range per distinct size
                    ws.Range(_a1_areas(row, col, offs)).Font.Size = size
            finally:
                self._screen_updating(True)
        self._retry(go)

    def save(self):
        self.calls["save"] += 1
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
    col: int                       # first column written = the NAME column
    previous: list                 # name + 9 picks as they were before (normalised), for undo
    written: list                  # name + 9 picks as written; written[0] is None while the name is still empty
    ticket: int | None
    prev_sizes: dict = field(default_factory=dict)     # font sizes to put back on undo: {offset: size}
    name_col: int = 0              # same as col; kept as its own field for callers that think in "name column"

    def __post_init__(self):
        if not self.name_col:
            self.name_col = self.col


MAX_PROBES = 25                    # free-looking slots tried (one small read each) before a full refresh


class ExcelSink:
    def __init__(self, backend, sheet_patterns=("ENTRY*",), name_header="NOMBRE",
                 uppercase: bool = True, save_after_write: bool = False,
                 blank_marker: str = "F1", blank_font_size: int = 9):
        self.backend = backend
        self.blank_marker = blank_marker
        self.blank_font_size = blank_font_size
        self.patterns = list(sheet_patterns) or ["*"]
        self.name_header = name_header
        self.uppercase = uppercase
        self.save_after_write = save_after_write
        self.snaps: dict = {}              # CACHE of each ENTRY sheet; kept current in place by our own writes
        self.blocks: list = []
        self.refresh_count = 0             # full re-reads so far (construction counts as 1)
        self.refresh()

    # -- reading the layout ------------------------------------------------
    def refresh(self):
        """Re-read every ENTRY sheet in full. Slow over COM: only at start, on reconnect, or when a
        verification read finds the cache stale (write_row_fast does that by itself)."""
        names = [n for n in self.backend.sheet_names()
                 if any(fnmatch.fnmatch(n.upper(), p.upper()) for p in self.patterns)]
        if not names:
            raise ExcelError(f"no sheet matches {self.patterns}; sheets are: {self.backend.sheet_names()}")
        snaps, blocks = {}, []
        for n in names:
            snap = self.backend.snapshot(n)
            snaps[n] = snap
            blocks += discover_blocks(n, snap, self.name_header)
        if not blocks:
            raise ExcelError(f"no '{self.name_header} | 1 2 3 ... 9' header found on {names}")
        self.snaps, self.blocks = snaps, blocks            # swapped only once everything was read
        self.refresh_count += 1
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

    # -- cache helpers (no COM) ---------------------------------------------
    def _first_free_row(self, block: Block, skip=()):
        snap = self.snaps[block.sheet]
        return next((r for r in block.rows if r not in skip and slot_is_free(snap, block, r)), None)

    def _resolve_block(self, old: Block | None):
        """The Block object of the current cache that stands for `old` (None = the default block).
        A refresh rebuilds the Blocks, and rows inserted above a header change its key."""
        if old is None:
            ordered = sorted(self.blocks, key=lambda b: (b.ticket_col is None,
                                                         self._ticket(b, b.rows[0]) or 0, b.name_col))
            return next((b for b in ordered if self._first_free_row(b) is not None),
                        ordered[0] if ordered else None)
        same = [b for b in self.blocks if b.sheet == old.sheet and b.name_col == old.name_col]
        return next((b for b in same if b.key == old.key),
                    min(same, key=lambda b: abs(b.header_row - old.header_row), default=None))

    def _remember(self, sheet: str, row: int, col: int, values: list):
        snap = self.snaps[sheet]
        for k, v in enumerate(values):
            snap.set(row, col + k, v)

    # -- writing -----------------------------------------------------------
    def format_name(self, name: str, tag: str = "", position: str = "after", sep: str = " ") -> str:
        return format_name(name, tag, position, sep, self.uppercase)

    def _row_values(self, name: str, picks) -> tuple:
        """-> ([name, p1..p9], {offset: small font size}). A box left empty becomes the marker (default F1)."""
        fill = self.blank_marker or None
        letters = [p if p in ("L", "E", "V") else fill for p in picks]
        letters += [fill] * (N_PICKS - len(letters))
        values = [(name or "").strip() or None] + letters[:N_PICKS]
        small = {k: self.blank_font_size for k, v in enumerate(values) if k >= 1 and v is not None and v == fill}
        return values, small

    def _probe(self, slot: Slot) -> tuple:
        """ONE small read of [ticket?, name, 9 picks] -> (status, the 10 live cells), status being
        'stale' (the ticket number there is not the one we cached: rows were inserted/deleted),
        'taken' (some target cell has data) or 'ok'. The cache learns what was read."""
        b = slot.block
        lead = 1 if b.ticket_col is not None else 0
        live = [norm(v) for v in self.backend.read_row(b.sheet, slot.row, b.name_col - lead, lead + N_PICKS + 1)]
        if lead and live[0] != slot.ticket:
            return "stale", live[lead:]
        cells = live[lead:]
        self._remember(b.sheet, slot.row, b.name_col, cells)
        return ("taken" if any(v is not None for v in cells) else "ok"), cells

    def _commit(self, slot: Slot, previous: list, values: list, small: dict) -> Receipt:
        b = slot.block
        prev_sizes = {}
        if small:                                       # old sizes only where the marker goes, in one read
            lo, hi = min(small), max(small)
            sizes = self.backend.font_sizes(b.sheet, slot.row, b.name_col + lo, hi - lo + 1)
            prev_sizes = {k: sizes[k - lo] for k in small if sizes[k - lo]}
        self.backend.write_row(b.sheet, slot.row, b.name_col, values, small)
        if self.save_after_write:
            self.backend.save()
        self._remember(b.sheet, slot.row, b.name_col, values)
        return Receipt(slot, b.sheet, slot.row, b.name_col, list(previous), list(values), slot.ticket, prev_sizes)

    def write_row_fast(self, block: Block | None, picks, name: str = "", ticket: int | None = None,
                       allow_overwrite: bool = False) -> Receipt:
        """Write name + 9 picks into the next free ticket of `block` (None = default block), or into the
        explicit `ticket`. Typical cost: one small read + one write, no full-sheet read; a font read/write
        is added only for boxes left empty. `name` may be '' (the cell stays empty; fill it in later with
        update_name). SlotTaken only for an explicit ticket that already has data (unless allow_overwrite);
        a free-looking slot that turns out to be taken is skipped. allow_overwrite does not apply to "next free"."""
        values, small = self._row_values(name, picks)
        block = getattr(block, "block", block)             # a BlockInfo works too
        if ticket is not None:
            return self._write_ticket(block, ticket, values, small, allow_overwrite)
        return self._write_next_free(block, values, small)

    def _write_next_free(self, block, values, small) -> Receipt:
        stale = False
        for attempt in (0, 1):
            b = self._resolve_block(block)
            skip = set()
            stale = False
            for _ in range(MAX_PROBES):
                row = self._first_free_row(b, skip) if b else None
                if row is None:
                    break
                slot = Slot(b, row, self._ticket(b, row))
                status, live = self._probe(slot)
                if status == "ok":
                    return self._commit(slot, live, values, small)
                if status == "stale":
                    stale = True
                    break
                skip.add(row)                           # taken by hand; the cache now knows
            if attempt == 0:
                self.refresh()
        if stale:
            raise ExcelError("Excel changed while writing (ticket numbers moved twice); "
                             "nothing was written - try again")
        if b is None:
            raise ExcelError("no free ticket in block: the block is gone from the sheet")
        first, last = b.rows[0], b.rows[-1]
        t = [self._ticket(b, first), self._ticket(b, last)]
        span = f"tickets {t[0]}-{t[1]}" if None not in t else f"rows {first}-{last}"
        raise ExcelError(f"no free ticket in block {b.sheet} column {col_letter(b.name_col)} ({span}): "
                         f"every row already has data")

    def _write_ticket(self, block, ticket, values, small, allow_overwrite) -> Receipt:
        for attempt in (0, 1):
            slot = self._find_ticket(ticket, block)
            if slot is not None:
                status, live = self._probe(slot)
                if status == "ok" or (status == "taken" and allow_overwrite):
                    return self._commit(slot, live, values, small)
                if status == "taken":
                    raise SlotTaken(f"{slot.describe()} already has data")
            if attempt == 0:
                self.refresh()
        raise ExcelError(f"ticket {ticket} is not in the sheet (or its row keeps moving)")

    def _find_ticket(self, ticket: int, block: Block | None = None) -> Slot | None:
        first = self._resolve_block(block) if block is not None else None
        for b in ([first] if first else []) + [b for b in self.blocks if b is not first]:
            for r in b.rows:
                if self._ticket(b, r) == ticket:
                    return Slot(b, r, ticket)
        return None

    def write(self, slot: Slot, name: str, picks: list, allow_overwrite: bool = False) -> Receipt:
        """Backwards-compatible: write into this slot (verified live, SlotTaken if it has data)."""
        values, small = self._row_values(name, picks)
        if slot.ticket is not None:
            return self._write_ticket(slot.block, slot.ticket, values, small, allow_overwrite)
        status, live = self._probe(slot)                 # a block without ticket numbers: nothing to relocate by
        if status == "taken" and not allow_overwrite:
            raise SlotTaken(f"{slot.describe()} already has data")
        return self._commit(slot, live, values, small)

    # -- changing a ticket after it was written -----------------------------
    def _read_receipt_row(self, r: Receipt, n: int) -> list:
        """The n cells from the name column of the receipt's row, after checking (same read) that the
        row still carries the receipt's ticket; if rows moved, refresh once and follow the ticket."""
        b = r.slot.block
        lead = 1 if b.ticket_col is not None and r.ticket is not None else 0
        if lead + n == 0:
            return []
        for attempt in (0, 1):
            live = [norm(v) for v in self.backend.read_row(r.sheet, r.row, r.col - lead, lead + n)]
            if not lead or live[0] == r.ticket:
                return live[lead:]
            if attempt == 0:
                self.refresh()
                slot = self._find_ticket(r.ticket, b)
                if slot is None:
                    break
                r.slot, r.sheet, r.row = slot, slot.block.sheet, slot.row
                r.col = r.name_col = slot.block.name_col
                b = slot.block
        raise ExcelError(f"ticket {r.ticket} moved or was removed in Excel; nothing was changed")

    def update_name(self, r: Receipt, name: str, expect=None) -> None:
        """Fill in / replace ONLY the name cell of a ticket written earlier (one write plus one small read).
        The cell must still be empty or hold `expect` (default: what we wrote there), else SlotTaken."""
        want = norm(r.written[0] if expect is None else expect)
        cur = self._read_receipt_row(r, 1)[0]
        if cur is not None and cur != want:
            raise SlotTaken(f"{r.slot.describe()}: the name cell now holds {cur!r}; left as is")
        value = (name or "").strip() or None
        self.backend.write_row(r.sheet, r.row, r.col, [value], None)
        if self.save_after_write:
            self.backend.save()
        self._remember(r.sheet, r.row, r.col, [value])
        r.written[0] = value

    def undo(self, r: Receipt):
        self._read_receipt_row(r, 0)                    # the row must still be that ticket (a read of its number only)
        self.backend.write_row(r.sheet, r.row, r.col, list(r.previous), r.prev_sizes or None)
        if self.save_after_write:
            self.backend.save()
        self._remember(r.sheet, r.row, r.col, list(r.previous))
