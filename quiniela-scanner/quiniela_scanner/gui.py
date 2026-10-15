"""Desktop window (tkinter).

Stack workflow ("fast mode"): the engine reads a slip, the picks go into Excel at once, and the handwritten
name is read in the background and filled into the same row a few seconds later. The engine is never paused,
so the next slip of the stack is picked up as soon as the previous one is lifted. Slips that need a human
(unclear marks, no name reader, auto-enter off) open the classic card; the engine keeps reading meanwhile and
every slip that arrives while a card is open waits in ``backlog`` (loud: beep + alert) and is entered, in
arrival order, as soon as the card is closed.

Threading rules: tkinter and Excel (COM) are touched from the Tk thread only. Worker threads (camera, name
reads, disk writes) hand their results over through ``self.events``.
"""
from __future__ import annotations

import collections
import os
import queue
import re
import statistics
import sys
import threading
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from tkinter import font as tkfont, messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageTk

from . import slip_vision as sv
from . import sound
from .audit import AuditLog, ScanLog
from .camera import CameraError, FrameGrabber, WebcamSource
from .capture import HAND_MESSAGE, WAIT_REMOVE, CaptureEngine, Update
from .config import Config, app_dir
from .excel_sink import BlockFull, ExcelBusy, ExcelError, ExcelSink, Receipt, SlotTaken, format_name, open_backend
from .name_reader import NameResult, make_reader

FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
GOOD, WARN, BAD = "#1b7f3b", "#b36b00", "#c62828"
NOTE_BG, NOTE_FG = "#ffe0a3", "#5c3a00"          # the amber strip: things to know that do not stop the stack
VERDICT_TEXT = {sv.MULTIPLE: "two marks", sv.WEAK: "faint mark", sv.HEAVY: "scribbled?",
                sv.OCCLUDED: "finger over this row"}

PREVIEW_W = 520             # the live picture is shrunk to this before anything is drawn on it
PREVIEW_FPS = 15
TICK_MS = 25
BREAK_S = 60                # a gap between two entries longer than this is a break, not part of the average
KEEP_CROPS = 20             # name pictures kept (as arrays) for the most recent entries
CLOSE_WAIT_S = 10.0         # closing the window: wait this long for names that are still being read
CAPTURE_DEBOUNCE_S = 1.5    # 'Capture now' ignores clicks this soon after the last one
READER_FAILS = 3            # consecutive service failures (key, credit, line, rate limit) = the name reader is down
READER_HANG_S = 10.0        # READER_FAILS requests silent for this long (or half name_timeout_s) and none answered: down
TICK_BUDGET_S = 0.15        # one tick works through the waiting events for at most this long, so Tk keeps painting
BUSY_RETRY_S = 1.0          # Excel was busy: ask again this often ...
NAME_BUSY_S = 60.0          # ... for this long before a finished name becomes a red 'NAME NOT WRITTEN' row
LOW_TICKETS = 5             # a block with this few tickets left is shown in amber ...
LOW_BEEP = (3, 0)           # ... with one beep when it gets down to each of these
FPS_WARN = 20.0             # the camera delivers fewer frames per second than this (for FPS_CHECKS checks in a row): warn
FPS_OK = 24.0               # ... and the warning goes away above this
FPS_CHECKS = 3
GAP_FACTOR = 2.5            # a slip that comes this many times later than the usual pace: 'a slip may have been skipped'
GAP_MIN = 4                 # ... judged once this many gaps are known

_JORNADA_RE = re.compile(r"jornada[\s_\-.#:]*(?:n(?:o|º|°|um|úm|ro)?\.?[\s_\-.#:]*)?(\d+)", re.I)


def jornada_of(workbook_name: str) -> str:
    """'JORNADA 10 PC.xlsx', 'Jornada_10', 'JORNADA No. 10', 'JORNADA #10' -> '10'; '' when there is no number."""
    m = _JORNADA_RE.search(workbook_name or "")
    return m.group(1) if m else ""


def _to_photo(img_bgr: np.ndarray, width: int | None = None, height: int | None = None) -> ImageTk.PhotoImage:
    """Shrink (never enlarge) to fit width/height and wrap for tkinter."""
    h, w = img_bgr.shape[:2]
    s = min([1.0] + ([width / w] if width else []) + ([height / h] if height else []))
    if s != 1.0:
        img_bgr = cv2.resize(img_bgr, None, fx=s, fy=s, interpolation=cv2.INTER_AREA)
    return ImageTk.PhotoImage(Image.fromarray(cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)))


def _shrink(img: np.ndarray, width: int) -> np.ndarray:
    h, w = img.shape[:2]
    if w <= width:
        return img
    return cv2.resize(img, None, fx=width / w, fy=width / w, interpolation=cv2.INTER_AREA)


def _jpeg(img: np.ndarray) -> bytes | None:
    """A name picture as JPEG bytes (~10 KB instead of ~190 KB as an array): review items live as long as the session."""
    try:
        ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, 90])
        return buf.tobytes() if ok else None
    except cv2.error:
        return None


@dataclass(frozen=True)
class TagSnap:
    """The group-tag settings at the moment a slip was scanned: F2 may be pressed before its name arrives."""
    tag: str = ""
    position: str = "after"
    sep: str = " "
    upper: bool = True

    def format(self, name: str) -> str:
        return format_name(name, self.tag, self.position, self.sep, self.upper)


@dataclass
class Pending:
    """A slip on the classic card, waiting for the operator."""
    id: int
    reading: sv.SlipReading
    touched: bool = False                 # operator edited something -> no auto-enter, keep their name
    edited: set = field(default_factory=set)   # rows whose pick the operator chose by hand
    hold: bool = False                    # Excel refused it in fast mode: only a human may enter it now
    name_done: bool = False
    name_legible: bool = False
    slip_jornada: str | None = None
    auto_deadline: float | None = None
    timings: dict = field(default_factory=dict)
    fp: object = field(default=None, repr=False)      # the engine's fingerprint of this slip (Update.fingerprint)


@dataclass
class Job:
    """One slip that is in Excel. In fast mode its name is still being read when it is created."""
    id: int
    iid: str
    sink: ExcelSink = field(repr=False)
    receipt: Receipt = field(repr=False)
    picks: list                           # as shown: the marker for empty boxes
    tag: TagSnap
    expected: str                         # jornada the workbook should be, as of the scan
    workbook: str
    t_scan: float
    t_enter: float
    reading: sv.SlipReading | None = field(default=None, repr=False)   # dropped once the job is settled
    name_crop: np.ndarray | None = field(default=None, repr=False)
    crop_jpg: bytes | None = field(default=None, repr=False)           # name_crop of a review row, as JPEG
    read_ms: float = 0.0
    excel_ms: float = 0.0
    name_ms: float | None = None
    state: str = "reading"                # reading | ok | review | wrong | undone
    text: str = "ENTERED - reading the name..."
    reason: str = ""
    guess: str = ""                       # the name we proposed, marker stripped
    written: str | None = None            # what the name cell holds because of us
    jornada: str | None = None
    manual: bool = False
    t_submit: float = 0.0
    t_started: float = 0.0                # when a pool worker actually began the request (0 = still queued)
    future: object = field(default=None, repr=False)
    feed_ms: float = 0.0
    fp_ms: float = 0.0
    ui_ms: float = 0.0                    # Tk-thread time spent entering it
    lag_ms: float = 0.0                   # how long the scan waited in the queue for the Tk thread
    cycle_s: float | None = None          # time since the previous entry
    logged: bool = False
    fp: object = field(default=None, repr=False)        # the engine's fingerprint of the slip (see App.forget_if_current)
    name_hidden: bool = False             # a hand covered the name strip when the slip was read: the name needs a look
    counted: bool = False                 # its name result was already counted for the reader-outage logic
    writing: bool = False                 # the name is read; waiting for a busy Excel to take it (see name_busy)
    busy_since: float = 0.0               # ... since (monotonic)
    name_answer: tuple | None = field(default=None, repr=False)     # (NameResult, ms) of that name

    def crop(self) -> np.ndarray | None:
        """The name picture, if it is still kept (as an array for recent rows, as JPEG for review rows)."""
        if self.name_crop is not None:
            return self.name_crop
        if self.crop_jpg is not None:
            return cv2.imdecode(np.frombuffer(self.crop_jpg, np.uint8), cv2.IMREAD_COLOR)
        return None

    @property
    def name_done(self) -> bool:
        return self.state != "reading"

    @property
    def review(self) -> bool:
        return self.state in ("review", "wrong")

    @property
    def ticket(self):
        return self.receipt.ticket


class App(tk.Tk):
    def __init__(self, cfg: Config, source=None, backend=None):
        super().__init__()
        self.cfg = cfg
        self.title("Quiniela Scanner")
        sw, sh, dpi = self._screen()
        self.compact = sh / dpi < 900                # small screen (1366x768, 150% DPI on 1080p ...): denser layout
        self.preview_w = 360 if self.compact else PREVIEW_W
        self.slip_box = (200, 215) if self.compact else (310, 325)       # max size of the slip picture on the card
        self.rev_w = 300 if self.compact else 460                       # ... and of the picture in the review panel
        self.tree_rows = 3 if self.compact else 4
        self.minsize(min(1100, sw - 20), min(520, sh - self._margin_h(dpi)))
        base = app_dir()
        self.audit = AuditLog(base / "scans")
        self.scanlog = ScanLog(base / "scanner.log")
        self.events: queue.Queue = queue.Queue()
        self._lock = threading.Lock()
        self._latest: tuple | None = None            # (frame id, frame, Update) of the newest processed frame
        self._stop = threading.Event()
        self.source = None
        self.grabber: FrameGrabber | None = None
        self._cam_thread: threading.Thread | None = None
        self._fixed_source = source
        self._fixed_backend = backend
        # a demo loop replays one slip for ever: dedupe would (rightly) read it only once
        self.engine = CaptureEngine(cfg.matches_per_slip, cfg.stable_frames,
                                    sv.Settings(min_ink=cfg.min_ink, multi_ratio=cfg.multi_ratio),
                                    dedupe=not getattr(source, "demo", False))
        self.engine.paused = False                   # a card never pauses the engine: later slips queue in self.backlog
        self.pending: Pending | None = None
        self.marker = (cfg.blank_marker or "-")      # what an empty box turns into ("F1")
        self._scan_id = 0
        self.sink: ExcelSink | None = None
        self.backend = None
        self.entries: list[Job] = []                 # every slip entered this session, oldest first
        self.iid_jobs: dict[str, Job] = {}
        self.jobs: dict[int, Job] = {}
        self._job_id = 0
        self.entry_times: list[float] = []
        self.backlog: list = []                      # (Update, t_event, TagSnap) that arrived while a card was open
        self._draining = False
        self.card_job: Job | None = None             # the entered slip the card is showing (read-only)
        self.rev_job: Job | None = None              # the history row selected for review
        self.infos: list = []
        self.sel_block = None             # the Block the operator selected (followed through sink.resolve_block)
        self.sel_key: str | None = None
        self.ticket_custom = False        # operator typed a ticket number
        self._ticket_prog = False         # ticket_var is being set by the program, not typed
        self.breaker = ""                 # message while auto-enter is switched off by the circuit breaker
        self._alerts: dict = {}           # kind -> message of every alert that is up (insertion ordered)
        self._entered_msg = ""
        self._lift_msg = ""               # 'Undone / discarded - lift the slip and place it again': shown while it lies there
        self._dup_latched = False         # the engine dropped a repeat read: keep its 'lift it off' message up
        self.dupes_skipped = 0
        self.callback_errors = collections.deque(maxlen=50)      # short strings, never exception objects
        self.callback_error_count = 0
        self._seen_errors: set = set()
        self.closing = False
        self._close_started = False
        self._capturing = False
        self._capture_not_before = 0.0
        self._busy = 0                    # name workers running a request right now (any kind)
        self._busy_lock = threading.Lock()
        self._rd_fail = 0                 # consecutive service failures of the name reader
        self._rd_down = False
        self._reader_breaker = ""
        self._last_name_done = time.perf_counter()     # when any name request last finished (pool progress)
        self._last_name_ok = self._last_name_done
        self._queue_warn = False
        self._excel_hold = 0.0            # monotonic: Excel was busy a moment ago, do not ask it for names before this
        self._name_retry: list = []       # jobs whose finished name waits for a busy Excel
        self._block_gone = False          # the selected block was not found any more after a re-read: write nothing
        self._fail_banner = ""            # banner for the card of a slip Excel refused (set by enter_fast)
        self._connected_wb = None         # name of the workbook the last successful connect was to
        self._block_confirm = ""          # label of the block the stack was moved to after a full one, until a human confirms
        self._low_warned: set = set()     # (block, tickets left) the low-stock beep was given for
        self.low_note = ""                # 'Only 3 tickets left in this block' (strip, status line)
        self._notices: dict = {}          # amber strip: kind -> message
        self._dismissed: dict = {}        # kind -> the message the operator dismissed
        self._shown_layout: tuple = ()
        self._fps = 0.0                   # camera frames per second as measured, 0 = not known yet
        self._fps_bad = 0
        self._next_fps = 0.0
        self.slips_read = 0               # slips the camera read this session (entered or on a card)
        self._read_times: list = []       # perf_counter of each of them, for the 'slip skipped' check
        self.reader, self.reader_note = make_reader(cfg.read_names_with_claude, cfg.claude_model, cfg.claude_effort)
        self.pool = ThreadPoolExecutor(max_workers=max(1, int(cfg.name_workers)), thread_name_prefix="name")
        self.io = ThreadPoolExecutor(max_workers=1, thread_name_prefix="io")   # audit images + scanner.log
        self._photos: dict = {}
        self._prev_photo = None
        self._prev_fid = -1
        self._next_preview = 0.0
        self._next_watch = 0.0
        self._focus_after = None
        self._last_vision_err = ""
        self._last_cb_err = ("", 0.0)

        self._build()
        self.bind_all("<Return>", self._on_return)
        self.bind_all("<Escape>", self._on_escape)
        self.bind_all("<F2>", lambda e: self.tag_on.set(not self.tag_on.get()))
        self.bind("<Configure>", self._on_resize)
        self.bind_all("<MouseWheel>", self._on_wheel)
        self.bind_all("<Button-4>", self._on_wheel)
        self.bind_all("<Button-5>", self._on_wheel)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.log("START", fast=cfg.auto_enter_delay_s <= 0, workers=cfg.name_workers, stable=cfg.stable_frames,
                 reader=type(self.reader).__name__ if self.reader else "none",
                 reset=",".join(getattr(cfg, "migrated", [])) or None)
        if getattr(cfg, "migrated", None):
            self.hint_var.set("Settings from an older version were reset to the new defaults: " +
                              ", ".join(cfg.migrated) + ".")
        if getattr(cfg, "notes", None):              # config.json / .env problems found while starting
            self.alert("  ".join(cfg.notes), "config", beep=False)
        self.connect_excel()
        self.update_idletasks()                      # size to what the widgets need (high-DPI fonts grow), within the screen
        self._sync_mid()
        self.update_idletasks()
        room = 0 if self.compact else 140            # room for the alert strip (on a small screen the middle scrolls instead)
        self.geometry(f"{min(max(1240, self.winfo_reqwidth()), sw - 20)}x"
                      f"{min(max(840 if not self.compact else 0, self.winfo_reqheight() + room), sh - self._margin_h(dpi))}")
        self.start_camera()
        self._tick_id = self.after(TICK_MS, self.tick)

    # ------------------------------------------------------------------ UI
    def _screen(self) -> tuple:
        """(width, height, scale) of the screen; scale = size of a pixel relative to 96 dpi (Windows display scaling)."""
        return (self.winfo_screenwidth(), self.winfo_screenheight(),
                max(0.5, float(self.tk.call("tk", "scaling")) / (96 / 72)))

    @staticmethod
    def _margin_h(dpi: float) -> int:
        """Height of the screen that is not ours: title bar + taskbar (+ a little air)."""
        return int(90 * dpi)

    def _build(self):
        pad = dict(padx=6, pady=3) if not self.compact else dict(padx=6, pady=1)
        style = ttk.Style(self)
        style.configure("Big.TButton", font=(FONT, 12, "bold"), padding=8)
        style.configure("Card.TLabelframe.Label", font=(FONT, 11, "bold"))
        try:                                 # list rows follow the font (high DPI): the default 20 px clips the text at 150 %
            style.configure("Treeview", rowheight=max(20, tkfont.nametofont("TkDefaultFont").metrics("linespace") + 4))
        except tk.TclError:
            pass

        # ----- toolbar: camera / excel
        bar = ttk.Frame(self)
        bar.pack(fill="x", **pad)
        self._bar = bar
        ttk.Label(bar, text="Camera #").pack(side="left")
        self.cam_var = tk.IntVar(value=self.cfg.camera_index)
        ttk.Spinbox(bar, from_=0, to=9, width=3, textvariable=self.cam_var).pack(side="left", padx=4)
        ttk.Button(bar, text="Restart camera", command=self.start_camera).pack(side="left", padx=2)
        ttk.Button(bar, text="Camera settings", command=self.camera_settings).pack(side="left", padx=2)
        self.af_var = tk.BooleanVar(value=self.cfg.autofocus)
        ttk.Checkbutton(bar, text="Autofocus", variable=self.af_var, command=self.apply_focus).pack(side="left", padx=8)
        ttk.Label(bar, text="Focus").pack(side="left")
        self.focus_var = tk.IntVar(value=int(self.cfg.manual_focus))
        self.focus_scale = tk.Scale(bar, from_=0, to=255, resolution=5, orient="horizontal", length=120,
                                    showvalue=True, variable=self.focus_var, command=self.on_focus_move,
                                    highlightthickness=0, borderwidth=0, takefocus=0)
        self.focus_scale.pack(side="left", padx=4)
        self._focus_state()
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        self.excel_var = tk.StringVar(value="Excel: not connected")
        self.excel_lbl = ttk.Label(bar, textvariable=self.excel_var, font=(FONT, 10, "bold"))
        self.excel_lbl.pack(side="left")
        ttk.Button(bar, text="Connect / reconnect Excel", command=self.connect_excel).pack(side="left", padx=8)

        # ----- loud message strip (Excel trouble, circuit breaker); only shown while there is something to say
        self.alert_var = tk.StringVar()
        self.alert_frame = tk.Frame(self, bg=BAD)
        self.alert_lbl = tk.Label(self.alert_frame, textvariable=self.alert_var, bg=BAD, fg="white", justify="left",
                                  font=(FONT, 11 if self.compact else 13, "bold"),
                                  wraplength=max(500, min(1050, self.winfo_screenwidth() - 220)))
        self.alert_lbl.pack(side="left", padx=10, pady=4)
        tk.Button(self.alert_frame, text="Dismiss", command=self.clear_alert).pack(side="right", padx=10)
        # ----- the amber strip under it: things to know that do not stop the stack (slow camera, a skipped slip?,
        # Excel warnings, an odd sheet layout)
        self.notice_var = tk.StringVar()
        self.notice_frame = tk.Frame(self, bg=NOTE_BG)
        self.notice_lbl = tk.Label(self.notice_frame, textvariable=self.notice_var, bg=NOTE_BG, fg=NOTE_FG, justify="left",
                                   font=(FONT, 10 if self.compact else 11, "bold"),
                                   wraplength=max(500, min(1050, self.winfo_screenwidth() - 220)))
        self.notice_lbl.pack(side="left", padx=10, pady=3)
        tk.Button(self.notice_frame, text="Dismiss", command=self.dismiss_notices).pack(side="right", padx=10)

        # ----- target row
        tgt = ttk.LabelFrame(self, text="Enter into")
        tgt.pack(fill="x", **pad)
        self._tgt = tgt
        self.block_lbl = ttk.Label(tgt, text="Block:")
        self.block_lbl.grid(row=0, column=0, **pad)
        self._block_fg = style.lookup("TLabel", "foreground") or "black"
        self.block_var = tk.StringVar()
        self.block_cb = ttk.Combobox(tgt, textvariable=self.block_var, state="readonly", width=58)
        self.block_cb.grid(row=0, column=1, **pad)
        self.block_cb.bind("<<ComboboxSelected>>", lambda e: self.on_block_change())
        ttk.Label(tgt, text="Ticket #:").grid(row=0, column=2, **pad)
        self.ticket_var = tk.StringVar()
        self.ticket_entry = ticket_entry = ttk.Entry(tgt, textvariable=self.ticket_var, width=7, font=(FONT, 12, "bold"))
        ticket_entry.grid(row=0, column=3, **pad)
        self.ticket_var.trace_add("write", lambda *a: self.on_ticket_edit())
        self.auto_var = tk.BooleanVar(value=self.cfg.auto_enter)
        ttk.Checkbutton(tgt, text="Auto-enter clean scans", variable=self.auto_var).grid(row=0, column=4, padx=14)
        self.auto_var.trace_add("write", lambda *a: self.on_auto_change())

        # group tag: letters like "RRR" or "E" added before/after every name, switched on/off as you go
        ttk.Label(tgt, text="Group tag:").grid(row=1, column=0, **pad)
        tagrow = ttk.Frame(tgt)
        tagrow.grid(row=1, column=1, columnspan=4, sticky="w", **pad)
        self.tag_var = tk.StringVar(value=self.cfg.name_tag)
        self.tag_pos = tk.StringVar(value=self.cfg.name_tag_position if self.cfg.name_tag_position in
                                    ("before", "after") else "after")
        self.tag_on = tk.BooleanVar(value=False)          # always starts OFF so nobody is tagged by accident
        self.tag_entry = ttk.Entry(tagrow, textvariable=self.tag_var, width=8, font=(FONT, 12, "bold"))
        self.tag_entry.pack(side="left", padx=(0, 8))
        ttk.Radiobutton(tagrow, text="Before name", value="before", variable=self.tag_pos).pack(side="left", padx=2)
        ttk.Radiobutton(tagrow, text="After name", value="after", variable=self.tag_pos).pack(side="left", padx=2)
        ttk.Checkbutton(tagrow, text="Add tag  (F2)", variable=self.tag_on).pack(side="left", padx=14)
        self.tag_state = tk.StringVar()
        self.tag_state_lbl = ttk.Label(tagrow, textvariable=self.tag_state, font=(FONT, 11, "bold"))
        self.tag_state_lbl.pack(side="left", padx=6)
        for v in (self.tag_var, self.tag_pos, self.tag_on):
            v.trace_add("write", lambda *a: self.on_tag_change())
        self.jornada_var = tk.StringVar()                  # whether the wrong-jornada guard is on (see expected_jornada)
        self.jornada_lbl = ttk.Label(tagrow, textvariable=self.jornada_var, font=(FONT, 10, "bold"))
        self.jornada_lbl.pack(side="left", padx=10)

        # ----- main: preview | card. On a screen that is too small for it the middle scrolls; the strip and the
        # history below it are packed first (side="bottom") so they always keep their room.
        self._mid = ttk.Frame(self)
        self._mid_canvas = tk.Canvas(self._mid, highlightthickness=0, borderwidth=0, takefocus=0,
                                     background=style.lookup("TFrame", "background") or self.cget("background"))
        self._mid_sb = ttk.Scrollbar(self._mid, orient="vertical", command=self._mid_canvas.yview)
        self._mid_canvas.configure(yscrollcommand=self._mid_sb.set)
        self._mid_canvas.pack(side="left", fill="both", expand=True)
        main = ttk.Frame(self._mid_canvas)
        self._mid_win = self._mid_canvas.create_window(0, 0, window=main, anchor="nw")
        self._main = main
        main.bind("<Configure>", lambda e: self._sync_mid())
        self._mid_canvas.bind("<Configure>", lambda e: self._sync_mid())
        left = ttk.Frame(main)
        left.pack(side="left", fill="y", padx=4)
        self.preview = ttk.Label(left)
        self.preview.pack()
        self._prev_photo = ImageTk.PhotoImage(Image.new("RGB", (self.preview_w, round(self.preview_w * 9 / 16)), (236, 236, 236)))
        self.preview.configure(image=self._prev_photo)      # (a placeholder: the window is sized for the real picture)
        self.status_var = tk.StringVar(value="Starting camera...")
        ttk.Label(left, textvariable=self.status_var, font=(FONT, 13, "bold"), wraplength=self.preview_w).pack(
            anchor="w", pady=2 if self.compact else 4)
        prog = ttk.Frame(left)
        prog.pack(anchor="w", fill="x", pady=(0, 2 if self.compact else 4))
        self.progress = ttk.Progressbar(prog, maximum=1.0, length=self.preview_w - 120)
        self.progress.pack(side="left", fill="x", expand=True)
        ttk.Button(prog, text="Capture now", command=self.capture_now).pack(side="left", padx=(8, 0))
        self.last_var = tk.StringVar()
        ttk.Label(left, textvariable=self.last_var, font=(FONT, 11), foreground=GOOD,
                  wraplength=self.preview_w).pack(anchor="w")
        self.hint_var = tk.StringVar(value=self.reader_note)
        ttk.Label(left, textvariable=self.hint_var, foreground="#555", wraplength=self.preview_w).pack(anchor="w")
        rev = ttk.LabelFrame(left, text="Selected entry  (click a row of the list below)")
        rev.pack(anchor="w", fill="x", pady=(4 if self.compact else 10, 0))
        self.rev_img = ttk.Label(rev)
        self.rev_img.grid(row=0, column=0, columnspan=3, sticky="w", padx=4, pady=2)
        self.rev_info = tk.StringVar()
        ttk.Label(rev, textvariable=self.rev_info, wraplength=self.preview_w - 20, foreground=BAD).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=4)
        self.rev_name = tk.StringVar()
        self.rev_entry = ttk.Entry(rev, textvariable=self.rev_name, font=(FONT, 13, "bold"), width=16 if self.compact else 20)
        self.rev_entry.grid(row=2, column=0, padx=4, pady=3)
        self.rev_entry.bind("<Return>", lambda e: (self.save_review_name(), "break")[1])
        self.rev_entry.bind("<Escape>", lambda e: (self.focus_set(), "break")[1])    # Esc leaves the entry, never the card
        self.rev_save = ttk.Button(rev, text="Save name", command=self.save_review_name)
        self.rev_save.grid(row=2, column=1, padx=2)
        self.rev_undo = ttk.Button(rev, text="Undo this entry", command=self.undo_selected)
        self.rev_undo.grid(row=2, column=2, padx=2)

        card = ttk.LabelFrame(main, text="Scanned slip", style="Card.TLabelframe")
        card.pack(side="left", fill="both", expand=True, padx=8)
        self.banner_var = tk.StringVar()
        self.banner = tk.Label(card, textvariable=self.banner_var, font=(FONT, 14, "bold"), anchor="w", padx=8)
        self.banner.grid(row=0, column=0, columnspan=3, sticky="we", padx=6, pady=(2, 4))
        self._banner_look = (self.banner.cget("bg"), self.banner.cget("fg"))
        self.slip_img = ttk.Label(card)
        self.slip_img.grid(row=1, column=0, rowspan=5, sticky="n", **pad)
        self.name_img = ttk.Label(card)
        self.name_img.grid(row=1, column=1, columnspan=2, sticky="w", **pad)
        ttk.Label(card, text="Name:", font=(FONT, 11, "bold")).grid(row=2, column=1, sticky="e")
        self.name_var = tk.StringVar()
        self.name_var.trace_add("write", lambda *a: self.update_final())
        self.name_entry = ttk.Entry(card, textvariable=self.name_var, font=(FONT, 16, "bold"), width=17)
        self.name_entry.grid(row=2, column=2, sticky="w", **pad)
        self.name_entry.bind("<Key>", self.on_name_key)
        self.name_note = tk.StringVar()
        ttk.Label(card, textvariable=self.name_note, wraplength=300).grid(row=3, column=1, columnspan=2, sticky="w", padx=6)
        self.final_var = tk.StringVar()
        ttk.Label(card, textvariable=self.final_var, font=(FONT, 11, "bold"), foreground="#0b5394",
                  wraplength=300).grid(row=4, column=1, columnspan=2, sticky="w", padx=6, pady=2)
        self.jor_note = tk.StringVar()
        self.jor_lbl = ttk.Label(card, textvariable=self.jor_note, wraplength=300)
        self.jor_lbl.grid(row=5, column=1, columnspan=2, sticky="w", padx=6)

        grid = ttk.Frame(card)
        grid.grid(row=6, column=0, columnspan=3, sticky="w", **pad)
        self.pick_vars, self.pick_lbls, self.pick_radios = [], [], []
        for i in range(self.cfg.matches_per_slip):
            col = (i // 5) * 6
            row = i % 5
            lbl = ttk.Label(grid, text=f"{i + 1}", width=3, font=(FONT, 11, "bold"))
            lbl.grid(row=row, column=col, padx=(10, 2))
            v = tk.StringVar(value="")
            self.pick_vars.append(v)
            radios = []
            for j, (txt, val) in enumerate((("L", "L"), ("E", "E"), ("V", "V"), (self.marker, ""))):
                rb = ttk.Radiobutton(grid, text=txt, value=val, variable=v, command=lambda k=i: self.pick_edited(k))
                rb.grid(row=row, column=col + 1 + j, padx=2)
                radios.append(rb)
            self.pick_radios.append(radios)
            note = ttk.Label(grid, text="", width=11, foreground=WARN)
            note.grid(row=row, column=col + 5)
            self.pick_lbls.append((lbl, note))

        btns = ttk.Frame(card)
        btns.grid(row=7, column=0, columnspan=3, sticky="w", **pad)
        self.write_btn = ttk.Button(btns, text="Write to Excel  (Enter)", style="Big.TButton", command=self.commit)
        self.write_btn.pack(side="left", padx=4)
        self.discard_btn = ttk.Button(btns, text="Discard  (Esc)", command=self.discard)
        self.discard_btn.pack(side="left", padx=4)
        self.auto_var_txt = tk.StringVar()
        ttk.Label(card, textvariable=self.auto_var_txt, font=(FONT, 11, "bold"), foreground=GOOD).grid(
            row=8, column=0, columnspan=3, sticky="w", padx=8)

        # ----- stats strip
        self.stats_var = tk.StringVar()
        self.stats_lbl = ttk.Label(self, textvariable=self.stats_var, font=(FONT, 11, "bold"),
                                   wraplength=max(500, self.winfo_screenwidth() - 60))

        # ----- history (red rows = needs a look) + the review panel for the selected row
        hist = ttk.LabelFrame(self, text="Entered this session  (red = check the name)")
        hist.pack(side="bottom", fill="x", **pad)             # (packed first: the strip goes above it, the middle gets the rest)
        self.stats_lbl.pack(side="bottom", fill="x", padx=10, pady=(2, 0))
        self._mid.pack(fill="both", expand=True, **pad)
        cols = ("time", "ticket", "name", "picks", "status")
        self.tree = ttk.Treeview(hist, columns=cols, show="headings", height=self.tree_rows)
        for c, w in zip(cols, (80, 60, 210, 210, 400)):
            self.tree.heading(c, text=c.title())
            self.tree.column(c, width=w, anchor="w")
        self.tree.tag_configure("undone", foreground="#999")
        self.tree.tag_configure("review", foreground=BAD)
        self.tree.tag_configure("wrong", foreground=BAD, font=(FONT, 10, "bold"))
        self.tree.tag_configure("reading", foreground="#0b5394")
        self.tree.pack(side="left", fill="x", expand=True)
        scroll = ttk.Scrollbar(hist, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_tree_select())
        side = ttk.Frame(hist)
        side.pack(side="left", padx=8, fill="y")
        self.undo_btn = ttk.Button(side, text="Undo last", command=self.undo)
        self.undo_btn.pack(anchor="w", pady=2)
        self.review_btn = ttk.Button(side, text="Next to review", command=self.select_next_review)
        self.review_btn.pack(anchor="w", pady=2)
        self.clear_card()
        self.on_tag_change()
        self.on_tree_select()
        self.update_stats()

    def _on_resize(self, e):
        """The long strips wrap at the width of the window (not of the screen), so nothing runs off its right edge."""
        if e.widget is self and e.width > 300:
            self.stats_lbl.configure(wraplength=e.width - 40)
            self.alert_lbl.configure(wraplength=max(400, min(1050, e.width - 200)))
            self.notice_lbl.configure(wraplength=max(400, min(1050, e.width - 200)))

    def _sync_mid(self):
        """Keep the scrolling middle in step: as tall as its content wants, a scrollbar only when it is squeezed."""
        c, main = self._mid_canvas, self._main
        rw, rh = main.winfo_reqwidth(), main.winfo_reqheight()
        c.configure(width=rw, height=rh, scrollregion=(0, 0, max(rw, c.winfo_width()), rh))
        c.itemconfigure(self._mid_win, width=max(rw, c.winfo_width()))
        need = c.winfo_height() > 1 and rh > c.winfo_height() + 1
        if need and not self._mid_sb.winfo_ismapped():
            self._mid_sb.pack(side="right", fill="y")
        elif not need and self._mid_sb.winfo_ismapped():
            self._mid_sb.pack_forget()
            c.yview_moveto(0)

    def scroll_mid(self, bottom: bool):
        """On a small screen the middle scrolls: show its bottom (the review panel) or its top (the card, the banner)."""
        if self._mid_sb.winfo_ismapped():
            self.update_idletasks()                       # (the content may just have grown: take its new size first)
            self._sync_mid()
            self._mid_canvas.yview_moveto(1.0 if bottom else 0.0)

    def _on_wheel(self, e):
        if not self._mid_sb.winfo_ismapped():
            return
        w = self.winfo_containing(e.x_root, e.y_root)
        if w is None or not str(w).startswith(str(self._mid_canvas)):
            return
        up = e.num == 4 or getattr(e, "delta", 0) > 0
        self._mid_canvas.yview_scroll(-2 if up else 2, "units")

    def _key_context(self, e) -> str:
        """Where did Enter / Esc come from? 'card' = the window itself or a widget with no use for the key (it may act
        on the card); 'entry' = an entry / drop-down / spinbox of this window, which keeps the key for itself;
        'other' = another window (a dialog, the drop-down list of a combobox)."""
        w = e.widget
        try:
            if w.winfo_toplevel() is not self:
                return "other"
            if any(str(w) == str(o) for o in (self.ticket_entry, self.tag_entry, self.rev_entry, self.block_cb)):
                return "entry"
            return "entry" if w.winfo_class() in ("TSpinbox", "Spinbox", "Scale") else "card"
        except (AttributeError, tk.TclError):
            return "other"

    def _on_return(self, e):
        ctx = self._key_context(e)
        if ctx == "card":
            self.commit()
        elif ctx == "entry":
            self.focus_set()                      # Enter confirms what was typed; the next Enter writes the card
            return "break"

    def _on_escape(self, e):
        ctx = self._key_context(e)
        if ctx == "card":
            self.escape()
        elif ctx == "entry":
            self.focus_set()                      # Esc leaves the entry; it never discards the card
            return "break"

    # ----------------------------------------------------------- camera
    def start_camera(self):
        self._stop_camera()
        try:
            if self._fixed_source is not None:
                self.source = self._fixed_source
            else:
                self.source = WebcamSource(self.cam_var.get(), self.cfg.frame_width, self.cfg.frame_height,
                                           self.cfg.camera_backend, self.af_var.get(), self.cfg.manual_focus)
        except CameraError as e:
            self.source = None
            self.status_var.set(str(e))
            self.log("CAMERA_ERROR", err=str(e))
            return
        self.cfg.camera_index = self.cam_var.get()
        self.engine.dedupe = not getattr(self.source, "demo", False)
        self.engine.reset()
        with self._lock:
            self._latest = None
        self._prev_fid = -1                 # the new grabber numbers its frames from 1 again
        self._fps, self._fps_bad, self._next_fps = 0.0, 0, 0.0
        self.clear_notice("camera")
        self._stop = threading.Event()
        self.grabber = FrameGrabber(self.source)
        self._cam_thread = threading.Thread(target=self._camera_loop, args=(self._stop, self.grabber),
                                            name="vision", daemon=True)
        self._cam_thread.start()
        self.status_var.set(f"Using {self.source.label}. Place a slip flat under the camera.")
        self.log("CAMERA", source=self.source.label, fps_reported=getattr(self.source, "fps_reported", None) or None)

    def _stop_camera(self):
        self._stop.set()
        if self.grabber is not None:
            self.grabber.stop()
        if self._cam_thread is not None and self._cam_thread is not threading.current_thread():
            self._cam_thread.join(2.0)
        self.grabber = self._cam_thread = None
        if self.source is not None:
            try:
                self.source.close()
            except Exception:
                pass
        self.source = None

    def apply_focus(self):
        self._focus_state()
        if self.source is not None:
            try:
                self.source.set_autofocus(self.af_var.get(), int(self.focus_var.get()))
            except Exception as e:
                self.status_var.set(f"Could not change the focus mode: {e}")

    def _focus_state(self):
        self.focus_scale.configure(state="disabled" if self.af_var.get() else "normal")

    def on_focus_move(self, _value=None):
        """Slider moved: apply the new focus once it stops moving (every step would hit the driver)."""
        if self._focus_after is not None:
            self.after_cancel(self._focus_after)
        self._focus_after = self.after(120, self.apply_manual_focus)

    def apply_manual_focus(self):
        self._focus_after = None
        value = int(round(float(self.focus_var.get()) / 5) * 5)
        self.cfg.manual_focus = value
        if self.af_var.get() or self.source is None:
            return
        try:
            self.source.set_autofocus(False, value)
        except Exception as e:
            self.status_var.set(f"Could not set the focus: {e}")

    def camera_settings(self):
        if self.source is not None:
            self.source.open_driver_settings()

    def _camera_loop(self, stop: threading.Event, grabber: FrameGrabber):
        last = 0
        while not stop.is_set():
            fid, frame = grabber.get(last, 0.25)
            if frame is None:
                continue
            last = fid
            try:
                upd = self.engine.feed(frame)
            except Exception as e:            # never let a vision bug kill the camera thread
                upd = Update("idle", f"vision error: {e}")
                if str(e) != self._last_vision_err:
                    self._last_vision_err = str(e)
                    self.log("VISION_ERROR", err=str(e))
            with self._lock:
                self._latest = (fid, frame, upd)
            if upd.reading is not None or upd.error or upd.duplicate:
                self.events.put(("scan", upd, time.perf_counter()))

    def capture_now(self):
        with self._lock:
            latest = self._latest
        if latest is None:
            return
        if self.pending is not None:
            self.status_var.set("Finish the slip on the card first (Enter = write, Esc = discard).")
            return
        now = time.monotonic()
        if self._capturing or now < self._capture_not_before:      # a double click must not enter the slip twice
            return
        self._capturing = True
        self._capture_not_before = now + CAPTURE_DEBOUNCE_S
        frame = latest[1]

        def run():                            # a slip read takes ~200 ms: keep it off the Tk thread
            try:
                upd = self.engine.capture_now(frame)
            except Exception as e:
                upd = Update("idle", f"capture failed: {e}", error=str(e))
            finally:
                self._capturing = False
            self.events.put(("scan", upd, time.perf_counter()))
        threading.Thread(target=run, name="capture-now", daemon=True).start()

    # ------------------------------------------------------------ Excel
    def connect_excel(self):
        """(Re)connect to the workbook. Whatever goes wrong - not just ExcelError - ends up in the red Excel strip:
        the window must always open. Reconnecting to the same workbook keeps the block that was selected (it is followed
        through rows that moved); a different workbook starts from the first block with room."""
        try:
            self.backend = self._fixed_backend or open_backend(self.cfg.backend, self.cfg.workbook, self.cfg.workbook_path)
            self.sink = ExcelSink(self.backend, self.cfg.sheets, self.cfg.name_header,
                                  self.cfg.uppercase_names, self.cfg.save_after_write,
                                  self.cfg.blank_marker, self.cfg.blank_font_size)
            wb = getattr(self.backend, "workbook_name", "")
            same = self.sel_block is not None and wb == self._connected_wb
            self._connected_wb = wb
            self._shown_layout = ()                       # (a fresh connect shows the layout warnings again)
            self.refresh_blocks(select_default=not same)
        except Exception as e:
            msg = str(e) if isinstance(e, ExcelError) else f"{e.__class__.__name__}: {e}"
            self.sink = None
            self.excel_var.set(f"Excel: NOT CONNECTED - {msg}")
            self.excel_lbl.configure(foreground=BAD)
            self.block_cb["values"] = []
            self.update_jornada_note()
            self.log("EXCEL_NOT_CONNECTED", err=msg)
            return
        self.excel_var.set(f"Excel: {self.backend.label}")
        self.excel_lbl.configure(foreground=GOOD)
        self.update_jornada_note()
        self.log("EXCEL", workbook=getattr(self.backend, "workbook_name", ""), blocks=len(self.sink.blocks),
                 jornada=self.expected_jornada() or "off", layout_warnings=len(self.sink.layout_warnings) or None)

    def update_jornada_note(self):
        exp = self.expected_jornada() if self.sink is not None else ""
        if exp:
            self.jornada_var.set(f"Jornada check: {exp}")
            self.jornada_lbl.configure(foreground=GOOD)
        elif self.sink is not None:
            self.jornada_var.set("Jornada check OFF - no 'jornada <n>' in the workbook name")
            self.jornada_lbl.configure(foreground=WARN)
        else:
            self.jornada_var.set("")

    def refresh_blocks(self, select_default=False, announce=False):
        """Re-list the blocks (from the sink's cache) and keep the selected one selected. The Blocks are rebuilt by every
        full re-read, and rows inserted above a header change its key, so the selection is followed through
        ExcelSink.resolve_block. A block that is really gone is NOT replaced by another one (that would send the stack
        to another group): one alert, and nothing is written until the operator picks a block."""
        if self.sink is None:
            return
        self.infos = self.sink.block_infos()
        self.block_cb["values"] = [i.label for i in self.infos]
        self.show_layout_warnings()
        if select_default or self.sel_block is None:
            chosen = self.sink.default_block()
        else:
            old = self.sel_block
            rb = self.sink.resolve_block(old)
            if rb is not None and None not in (rb.first_ticket, old.first_ticket) and rb.first_ticket != old.first_ticket:
                rb = None                                 # another block moved into its place: other tickets, maybe another group
            chosen = next((i for i in self.infos if i.block is rb), None) if rb is not None else None
            if chosen is None:
                self._block_gone = True
                self.alert("The block you selected is no longer in the sheet (rows or columns were moved or deleted in "
                           "Excel). Nothing is written until you pick a block in the Block list.", "block")
                self.block_lbl.configure(foreground=BAD)
                return
            self._block_gone = False
            self.clear_alert("block")
        if chosen:
            self.follow_info(chosen, announce)

    def on_ticket_edit(self):
        """Only a real edit of the Ticket box (typing, paste, cut) makes the ticket the operator's own choice."""
        if not self._ticket_prog:
            self.ticket_custom = True

    def set_ticket_text(self, text: str):
        self._ticket_prog = True
        try:
            self.ticket_var.set(text)
        finally:
            self._ticket_prog = False

    def follow_info(self, info, announce: bool = False):
        """``info`` is the selected block as of now: refresh the label, the Block box and the suggested ticket. A ticket
        the operator typed stays (ticket_done() ends it once it has been used)."""
        self.sel_block, self.sel_key = info.block, info.block.key
        self.block_var.set(info.label)
        if not self.ticket_custom:
            slot = info.first_free
            self.set_ticket_text(str(slot.ticket) if slot and slot.ticket is not None else "")
        self.check_stock(info, announce)

    def ticket_done(self):
        """A write (or an undo) used up the ticket box: it goes back to the suggestion for the next slip."""
        self.ticket_custom = False

    def check_stock(self, info, announce: bool = False):
        """Few tickets left in the block (<= LOW_TICKETS): amber 'Block:' label and a line in the strip, and when
        ``announce`` (an entry just used one up) one beep at each of LOW_BEEP so that nobody finds out at the slip that
        does not fit."""
        free, b = info.free, info.block
        key = (b.sheet, b.name_col, b.first_ticket)
        if free > LOW_TICKETS:
            self._low_warned -= {w for w in self._low_warned if w[0] == key}
            self.low_note = ""
            self.block_lbl.configure(foreground=self._block_fg)
            return
        self.low_note = ("This block is now FULL - choose the next block before the next slip." if free == 0 else
                         f"Only {free} ticket{'s' if free != 1 else ''} left in this block.")
        self.block_lbl.configure(foreground=BAD if free == 0 else WARN)
        if announce and free in LOW_BEEP and (key, free) not in self._low_warned:
            self._low_warned.add((key, free))
            self.beep(True)
            self.log("LOW_STOCK", block=b.key, free=free)

    def select_info(self, info):
        """The operator (or the program, after a full block) picks a block: its suggestion starts from scratch."""
        self.ticket_custom = False
        self._block_gone = False
        self.clear_alert("block")
        self.sel_block, self.sel_key = info.block, info.block.key
        self.block_var.set(info.label)
        slot = info.first_free
        self.set_ticket_text(str(slot.ticket) if slot and slot.ticket is not None else "")
        self.ticket_custom = False
        self.check_stock(info)

    def selected_info(self):
        return next((i for i in self.infos if i.block is self.sel_block), None) or \
            next((i for i in self.infos if i.block.key == self.sel_key), None)

    def on_block_change(self):
        info = next((i for i in self.infos if i.label == self.block_var.get()), None)
        if info:
            self._block_confirm = ""                  # the operator chose: nothing is left to confirm
            self.clear_alert("fullblock")
            self.select_info(info)

    def expected_jornada(self) -> str:
        if self.cfg.expected_jornada:
            return self.cfg.expected_jornada
        return jornada_of(getattr(self.backend, "workbook_name", ""))

    def target_slot(self):
        if self.sink is None:
            return None
        t = self.ticket_var.get().strip()
        if self.ticket_custom and t.isdigit():
            return self.sink.slot_for_ticket(int(t))
        info = self.selected_info()
        return info.first_free if info else None

    def target(self):
        """(block, explicit ticket or None) for the next write_row_fast. The block is the selected Block object and
        never None while there is a selection (None would mean 'the first block with room', i.e. another group)."""
        t = self.ticket_var.get().strip()
        return self.sel_block, (int(t) if self.ticket_custom and t.isdigit() else None)

    def show_layout_warnings(self):
        """What the sink found odd in the sheet (a block cut short by a text ticket, repeated tickets, formulas that show
        nothing ...): once per change, in the amber strip and in scanner.log."""
        w = tuple(self.sink.layout_warnings) if self.sink is not None else ()
        if w == self._shown_layout:
            return
        self._shown_layout = w
        if not w:
            self.clear_notice("layout")
            return
        shown = " | ".join(w[:4]) + (f" | ... and {len(w) - 4} more" if len(w) > 4 else "")
        self.notice(f"EXCEL LAYOUT - check the sheet: {shown}", "layout")
        for line in w:
            self.log("LAYOUT_WARNING", msg=line)

    # ------------------------------------------------------ messages / sounds / log
    def log(self, event: str, /, **fields):
        """One line in scanner.log (written on the io thread: no disk waits on the Tk thread)."""
        try:
            self.io.submit(self.scanlog.line, event, **fields)
        except RuntimeError:                  # already shut down
            self.scanlog.line(event, **fields)

    def flush_io(self, timeout: float = 5.0):
        """Wait until every queued audit picture / log line is on disk."""
        try:
            self.io.submit(lambda: None).result(timeout)
        except Exception:
            pass

    def beep(self, attention: bool):
        if self.cfg.sounds:
            (sound.attention if attention else sound.ok)()

    def alert(self, msg: str, kind: str = "excel", beep: bool = True):
        """Red strip under the toolbar. One message per ``kind`` (a new one of the same kind replaces it); the strip
        shows every kind that is up, so a backlog or reader alert never hides an Excel one."""
        changed = self._alerts.get(kind) != msg
        self._alerts[kind] = msg
        self._show_alerts()
        if changed or kind == "excel":                    # (every refused slip is an event worth a line; repeats of the others are not)
            self.log("ALERT", kind=kind, msg=msg)
        if beep:
            self.beep(True)

    def _show_alerts(self):
        self.alert_var.set("\n".join(self._alerts.values()))
        if self._alerts and not self.alert_frame.winfo_ismapped():
            self.alert_frame.pack(fill="x", padx=6, pady=3,
                                  before=self.notice_frame if self.notice_frame.winfo_ismapped() else self._tgt)
        elif not self._alerts:
            self.alert_frame.pack_forget()

    def clear_alert(self, kind: str | None = None):
        """One kind, or (the Dismiss button) all of them."""
        if kind is None:
            self._alerts.clear()
        elif self._alerts.pop(kind, None) is None:
            return
        self._show_alerts()

    def notice(self, msg: str, kind: str):
        """Amber strip: something the operator should know that does not stop the stack. One message per ``kind``;
        a message the operator dismissed stays away until it changes (or its condition ended: clear_notice)."""
        if self._dismissed.get(kind) == msg:
            return
        if self._notices.get(kind) != msg:
            self.log("NOTICE", kind=kind, msg=msg)
        self._notices[kind] = msg
        self._show_notices()

    def clear_notice(self, kind: str):
        self._dismissed.pop(kind, None)
        if self._notices.pop(kind, None) is not None:
            self._show_notices()

    def dismiss_notices(self):
        self._dismissed.update(self._notices)
        self._notices.clear()
        self._show_notices()

    def _show_notices(self):
        self.notice_var.set("\n".join(self._notices.values()))
        if self._notices and not self.notice_frame.winfo_ismapped():
            self.notice_frame.pack(fill="x", padx=6, pady=3, before=self._tgt)
        elif not self._notices:
            self.notice_frame.pack_forget()

    def set_banner(self, text: str, colour: str | None = None):
        self.banner_var.set(text)
        self.banner.configure(bg=colour or self._banner_look[0], fg="white" if colour else self._banner_look[1])

    def report_callback_exception(self, exc, val, tb):
        """An exception inside a Tk callback must not go unnoticed (or stop the tick loop). Only a short text is kept:
        an exception object would pin its traceback, and with it every frame (6 MB each) of the failing call."""
        import traceback
        key, now = f"{exc.__name__}: {val}", time.monotonic()
        self.callback_errors.append(key[:200])
        self.callback_error_count += 1
        if key == self._last_cb_err[0] and now - self._last_cb_err[1] < 5:
            return                                        # the same error again and again (a 25 ms loop): say it once
        self._last_cb_err = (key, now)
        where = traceback.format_tb(tb)[-1].strip().replace("\n", " ") if tb else None
        self.log("CALLBACK_ERROR", err=key, where=where)
        try:
            traceback.print_exception(exc, val, tb)
        except Exception:
            pass
        if key not in self._seen_errors and len(self._seen_errors) < 20:     # first time: the operator should know
            self._seen_errors.add(key)
            try:
                self.alert(f"INTERNAL ERROR - {key[:160]}. The app keeps running; if it happens again, send scanner.log.",
                           "internal", beep=False)
            except Exception:
                pass

    # --------------------------------------------------------- scans
    def set_card_enabled(self, on: bool):
        st = "normal" if on else "disabled"
        for radios in self.pick_radios:
            for rb in radios:
                rb.configure(state=st)
        self.write_btn.configure(state=st)
        self.discard_btn.configure(state=st)
        self.name_entry.configure(state="normal" if on else "readonly")

    def clear_card(self):
        self.pending = None
        self.card_job = None
        self.set_card_enabled(True)
        self.set_banner("")
        self.name_var.set("")
        self.name_note.set("")
        self.jor_note.set("")
        self.auto_var_txt.set("")
        self.final_var.set("")
        for v, (lbl, note) in zip(self.pick_vars, self.pick_lbls):
            v.set("")
            note.configure(text="")
            lbl.configure(foreground="black")
        blank = np.full((300, 330, 3), 245, np.uint8)
        cv2.putText(blank, "Waiting for a slip...", (40, 150), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (150, 150, 150), 2)
        blank = _shrink(blank, self.slip_box[0])
        self._photos["slip"] = _to_photo(blank)
        self.slip_img.configure(image=self._photos["slip"])
        self.name_img.configure(image="")

    def show_reading(self, rd: sv.SlipReading):
        """Pick buttons, straightened slip and name picture for a freshly read slip. The picks (the data) come first:
        a drawing problem must not leave a card with the wrong marks on it."""
        for i, (v, (lbl, note)) in enumerate(zip(self.pick_vars, self.pick_lbls)):
            v.set(rd.picks[i])
            verdict = rd.verdicts[i]
            lbl.configure(foreground="black" if verdict == sv.OK else WARN if verdict == sv.BLANK else BAD)
            note.configure(text=f"empty = {self.marker}" if verdict == sv.BLANK else VERDICT_TEXT.get(verdict, ""))
        try:
            self._photos["slip"] = _to_photo(sv.draw_reading(rd), width=self.slip_box[0], height=self.slip_box[1])
            self.slip_img.configure(image=self._photos["slip"])
            self._photos["name"] = _to_photo(rd.name_crop, width=self.slip_box[0])
            self.name_img.configure(image=self._photos["name"])
        except Exception as e:
            self.log("UI_ERROR", where="show_reading", err=repr(e))

    def on_scan(self, upd: Update, t_event: float | None = None):
        if upd.reading is None:
            if upd.duplicate:
                self.on_duplicate(upd)
                return
            self.status_var.set(upd.message)
            if upd.error:                                   # the operator has to re-place the slip
                self.log("UNREADABLE", err=upd.error)
                self.beep(True)
            return
        self.note_slip(t_event)
        if self.pending is not None or self.backlog:        # a card is open (or older slips still wait): keep the order
            self.backlog.append((upd, t_event, self.tag_snapshot()))
            self.backlog_changed(grew=True)
            self.drain_backlog()                            # (does nothing while the card is open)
            return
        self._safe_handle(upd, t_event, None)

    def _safe_handle(self, upd: Update, t_event, snap):
        n0 = len(self.entries)
        try:
            self._handle_scan(upd, t_event, snap)
        except Exception as e:                              # the next slip must still be handled
            self.report_callback_exception(type(e), e, e.__traceback__)
            if len(self.entries) == n0 and self.pending is None:         # neither entered nor on a card
                self.forget_if_current(upd.fingerprint)       # (so that lifting the slip and putting it back reads it again)
                self.alert("A slip could not be processed (internal error). Check the last ticket in Excel, then lift "
                           "the slip and put it back (or press Capture now).", "scanfail")

    def _handle_scan(self, upd: Update, t_event, snap=None):
        t0 = time.perf_counter()
        lag = (t0 - t_event) * 1000 if t_event else 0.0
        self._fail_banner = ""
        if self.fast_ok(upd.reading):
            if self.enter_fast(upd, t0, lag, snap):
                return
            self.open_card(upd, hold=True)                   # Excel said no: the operator takes this slip
            self.set_banner(self._fail_banner or "NOT ENTERED - fix Excel, then press Enter", BAD)
            return
        self.open_card(upd)
        if self._block_confirm:
            self.set_banner("CHECK THIS SLIP - NEW BLOCK, see the red bar", WARN)

    def note_slip(self, t_event):
        """A slip was read: count it, and say so when it comes much later than the pace of the stack. A slip the camera
        never saw leaves no other trace (no beep, no row), so the pace is the only clue there is."""
        now = t_event or time.perf_counter()
        self.slips_read += 1
        ts = self._read_times
        if ts:
            gap = now - ts[-1]
            gaps = [b - a for a, b in zip(ts, ts[1:]) if b - a <= BREAK_S]
            if len(gaps) >= GAP_MIN and gap <= BREAK_S:
                med = statistics.median(gaps)
                if gap > GAP_FACTOR * med:
                    self.notice(f"A SLIP MAY HAVE BEEN SKIPPED - {gap:.0f} s passed since the previous slip (usually "
                                f"{med:.1f} s). Check the stack against the tickets just entered; Capture now reads "
                                "the slip that lies there.", "skipped")
                    self.log("SLIP_GAP", gap_s=round(gap, 1), median_s=round(med, 1))
        ts.append(now)
        del ts[:-12]
        self.update_stats()

    def forget_if_current(self, fp) -> bool:
        """The slip with this fingerprint did NOT go into Excel (card discarded, entry undone, processing failed): the
        engine must read it again when it is lifted and put back, instead of dropping it as 'Same slip still there'.
        Only when it is the slip the engine remembers last - the fingerprint of an older entry is not the slip that
        lies under the camera, and forgetting it would let that one be entered twice. rearm=False: a slip that is
        still lying there is not read again until it has been lifted."""
        if fp is not None and self.engine.last_fingerprint is fp:
            self.engine.forget_last(rearm=False)
            return True
        return False

    def lift_message(self, text: str):
        """What the status line says while the slip that was just undone / discarded still lies under the camera."""
        self._entered_msg, self._lift_msg = "", text
        self.status_var.set(text)

    def drain_backlog(self):
        """Enter every slip that waited behind a card, in arrival order, until one needs a card again."""
        if self._draining:
            return
        self._draining = True
        try:
            while self.backlog and self.pending is None:
                upd, t_event, snap = self.backlog.pop(0)
                self._safe_handle(upd, t_event, snap)
        finally:
            self._draining = False
        self.backlog_changed()

    def next_backlog(self):
        """A card was closed: take the slips that were waiting behind it."""
        self.drain_backlog()

    def backlog_changed(self, grew: bool = False):
        """Slips waiting behind a card are tickets that are not in Excel yet: say so loudly, and beep when the first one
        queues up."""
        n = len(self.backlog)
        if n:
            self.alert(f"{n} slip{' is' if n == 1 else 's are'} WAITING behind the card - Enter (write) or Esc (discard), "
                       f"then {'it is' if n == 1 else 'they are'} entered in order.", "backlog", beep=grew and n == 1)
        else:
            self.clear_alert("backlog")
        self.update_stats()

    def on_duplicate(self, upd: Update):
        """The engine dropped a re-read of the slip that was entered last: count it, log it, keep the message up."""
        self.dupes_skipped += 1
        self._dup_latched = True
        last = next((j.ticket for j in reversed(self.entries) if j.state != "undone"), None)
        self.log("DUPLICATE_SKIPPED", n=self.dupes_skipped, last_ticket=last)
        if self.pending is None:
            self.status_var.set(upd.message)
        self.update_stats()

    def fast_ok(self, rd: sv.SlipReading) -> bool:
        """Can this slip go straight into Excel with its name to follow?"""
        return (self.auto_var.get() and self.cfg.auto_enter_delay_s <= 0 and self.reader is not None
                and self.sink is not None and not self._block_confirm and self.auto_eligible(rd))

    def tag_snapshot(self) -> TagSnap:
        return TagSnap(self.active_tag(), self.tag_pos.get(), self.cfg.tag_separator, self.cfg.uppercase_names)

    def enter_fast(self, upd: Update, t0: float, lag_ms: float, snap: TagSnap | None = None) -> bool:
        """Picks into Excel now, name later. False = Excel refused (the alert is up; the caller opens the card)."""
        rd = upd.reading
        snap, exp = snap or self.tag_snapshot(), self.expected_jornada()   # taken NOW: F2 may change before the name is back
        block, ticket = self.target()
        if self._block_gone:                                           # (no Excel call: the sink would pick a look-alike)
            self._fail_banner = "NOT ENTERED - the selected block is gone from the sheet: pick a block, then press Enter"
            self.alert("NOT ENTERED - the block you selected is no longer in the sheet. Pick a block in the Block "
                       "list, then press Enter.", "block")
            return False
        t1 = time.perf_counter()
        try:
            receipt = self.sink.write_row_fast(block, rd.picks, name="", ticket=ticket)
        except BlockFull as e:                                         # not Excel's fault: the block is used up
            self.block_full(e)
            return False
        except ExcelBusy as e:                                         # nothing was written: the card keeps the slip safe
            self._excel_hold = time.monotonic() + BUSY_RETRY_S
            self._fail_banner = "NOT ENTERED - Excel is busy: press Esc in Excel (or close its dialog), then press Enter"
            self.alert(f"NOT ENTERED - {e}", "excel")
            return False
        except Exception as e:                                         # ExcelError / SlotTaken / anything unexpected
            self.alert(f"NOT ENTERED - {e}", "excel")
            return False
        excel_ms = (time.perf_counter() - t1) * 1000
        self.beep(False)                                               # after the write, so a beep means "it is in"
        shown = [p or self.marker for p in rd.picks]
        job = self.new_job(receipt, rd, shown, snap, exp, upd, excel_ms, t0)
        # The row is in Excel: ask for its name before any drawing, so that a UI error cannot leave a row without a request.
        try:
            job.future = self.pool.submit(self._name_worker, self.reader, "job", job.id, rd.name_crop, rd.title_crop)
        except RuntimeError:                                           # pool already shut down (closing)
            job.future = None
        job.t_submit = time.perf_counter()
        try:
            self.receipt_notice(receipt)
            self.show_entered(job)
            self.clear_alert("excel")
            self.clear_alert("fullblock")
            self.ticket_done()
            self.refresh_blocks(announce=True)
            self._entered_msg = f"Entered ticket {receipt.ticket}. Lift the slip - the next one is read automatically." + \
                (f" {self.low_note}" if self.low_note else "")
            self._lift_msg = ""
            self._dup_latched = False
            self.last_var.set(f"Last entered: ticket {receipt.ticket}   {' '.join(shown)}   (name follows)")
            self.update_stats()
        except Exception as e:                                         # the row is in Excel: never fall back to the card
            self.log("UI_ERROR", where="enter_fast", ticket=receipt.ticket, err=repr(e))
        job.ui_ms = (time.perf_counter() - t0) * 1000
        job.lag_ms = lag_ms
        return True

    def block_full(self, e: BlockFull):
        """The selected block has no free ticket. The slip waits on a held card; the next block that has room (same
        group first: same trailing word in its names, same sheet) is pre-selected and named in the red bar, but nothing
        is written there until a human presses Enter (or picks a block) - another block can be another group's."""
        self.infos = self.sink.block_infos()
        full = e.block
        sfx = next((i.suffix for i in self.infos if i.block is full), "")
        room = [(n, i) for n, i in enumerate(self.infos) if i.free and i.block is not full]
        nxt = min(room, key=lambda c: (bool(sfx) and c[1].suffix != sfx, full is not None and c[1].block.sheet != full.sheet,
                                       c[0]))[1] if room else None
        self._fail_banner = "BLOCK FULL - choose another block in the Block list, then press Enter"
        if nxt is None:
            msg = (f"BLOCK FULL - NOT ENTERED: {e}. No other block has a free ticket: add rows in Excel (or open the "
                   "next workbook and press Connect), then press Enter.")
        else:
            self.select_info(nxt)
            self._block_confirm = nxt.label
            msg = (f"BLOCK FULL - NOT ENTERED: {e}. Press Enter to write this slip to: {nxt.label} - or choose "
                   "another block in the Block list first (a block can belong to another group).")
        self.alert(msg, "fullblock")
        self.log("BLOCK_FULL", block=full.key if full else None, next=nxt.block.key if nxt else None)

    def receipt_notice(self, receipt: Receipt):
        """Excel took the values but a best-effort step failed (font size, Save ...): the row is in, so this is no
        'NOT ENTERED'. Say what happened, in the amber strip."""
        if receipt.warnings:
            t = receipt.ticket if receipt.ticket is not None else f"row {receipt.row}"
            self.notice(f"Entered ticket {t}, but " + "; ".join(receipt.warnings), "warn")

    def new_job(self, receipt: Receipt, rd: sv.SlipReading, shown: list, snap: TagSnap, exp: str, upd: Update,
                excel_ms: float, t0: float, manual: bool = False) -> Job:
        self._job_id += 1
        now = time.perf_counter()
        job = Job(self._job_id, str(self._job_id), self.sink, receipt, shown, snap, exp,
                  getattr(self.backend, "workbook_name", ""), t0, now, reading=rd,
                  name_crop=_shrink(rd.name_crop, 640), read_ms=upd.timings.get("read_ms", 0.0),
                  excel_ms=excel_ms, manual=manual, t_submit=now)      # (t_submit: check_jobs times a job from here)
        job.fp = upd.fingerprint
        job.name_hidden = bool(getattr(rd, "name_hidden", False))
        job.feed_ms = upd.timings.get("feed_ms", 0.0)
        job.fp_ms = upd.timings.get("fingerprint_ms", 0.0)
        job.cycle_s = (job.t_enter - self.entry_times[-1]) if self.entry_times else None
        self.entries.append(job)
        self.jobs[job.id] = job
        self.iid_jobs[job.iid] = job
        self.entry_times.append(job.t_enter)
        if len(self.entries) > KEEP_CROPS:                  # only recent entries, rows still being read and review rows
            old = self.entries[-KEEP_CROPS - 1]             # keep their picture (review rows as JPEG, see settle)
            if old.state != "reading" and not old.review:
                old.name_crop = None
        t = receipt.ticket if receipt.ticket is not None else f"row {receipt.row}"
        try:
            self.tree.insert("", 0, iid=job.iid, values=(time.strftime("%H:%M:%S"), t, "(reading...)", " ".join(shown),
                                                         job.text), tags=("reading",))
        except tk.TclError as e:                            # the job exists and is tracked even without its list row
            self.log("UI_ERROR", where="new_job", err=repr(e))
        return job

    def show_entered(self, job: Job):
        """The card shows the slip that just went in, read-only, until the next one arrives."""
        self.pending = None
        self.card_job = job                  # first: update_final() stays quiet while an entered slip is shown
        self.set_card_enabled(True)
        self.show_reading(job.reading)
        self.set_card_enabled(False)
        self.set_banner(f"ENTERED - ticket {job.ticket}", GOOD)
        tag = job.tag.tag
        self.name_var.set("")
        self.name_note.set("Reading the handwriting..." + (f"  (tag \"{tag}\" {job.tag.position} the name)" if tag else ""))
        self.final_var.set("")
        self.jor_note.set("")
        self.auto_var_txt.set("")
        self.status_var.set(f"Entered ticket {job.ticket}.")

    def open_card(self, upd: Update, hold: bool = False):
        """Classic flow: the slip waits on the card for the operator (the engine pauses meanwhile)."""
        rd = upd.reading
        self._scan_id += 1
        self.card_job = None
        self.set_card_enabled(True)
        self.pending = Pending(self._scan_id, rd, hold=hold, timings=dict(upd.timings), fp=upd.fingerprint)
        self._lift_msg = ""
        self.scroll_mid(False)
        self.show_reading(rd)
        self.set_banner("CHECK THIS SLIP", WARN)
        self.name_var.set("")
        self.jor_note.set("")
        self.auto_var_txt.set("")
        if self.reader is not None:
            self.name_note.set("Reading handwriting...")
            try:
                self.pool.submit(self._name_worker, self.reader, "card", self._scan_id, rd.name_crop, rd.title_crop)
            except RuntimeError:
                pass
        else:
            self.pending.name_done = True
            self.name_note.set(self.reader_note)
        blanks = [i + 1 for i, v in enumerate(rd.verdicts) if v == sv.BLANK]
        flagged = [i + 1 for i, v in enumerate(rd.verdicts) if v not in (sv.OK, sv.BLANK)]
        msg = "Slip captured."
        if blanks:
            msg += f" Empty box in match {', '.join(map(str, blanks))} -> {self.marker}."
        if flagged:
            msg = "Check row(s) " + ", ".join(map(str, flagged)) + " (marked red)." + \
                  (f" Empty: {', '.join(map(str, blanks))} -> {self.marker}." if blanks else "")
        hidden = [i + 1 for i in rd.occluded_rows] if hasattr(rd, "occluded_rows") else []
        if hidden:
            msg += f" A finger or hand covers row(s) {', '.join(map(str, hidden))}: set them by hand, or lift the slip and place it again."
        self.status_var.set(msg)
        self.name_entry.focus_set()
        if not hold:
            self.beep(True)                                          # the line stops here until the operator acts

    # --------------------------------------------------------- names
    def _name_worker(self, reader, kind: str, key: int, name_crop, title_crop):
        """Runs in a pool thread: no tkinter, no Excel in here - the result goes back through the queue."""
        t0 = time.perf_counter()
        job = self.jobs.get(key) if kind == "job" else None
        if job is not None:
            job.t_started = t0                            # (a plain float store: safe from this thread)
        with self._busy_lock:
            self._busy += 1
        try:
            try:
                res = reader.read(name_crop, title_crop)
            except Exception as e:                        # (read() must not raise: a reader that does is itself broken)
                res = NameResult(error=f"name reading failed ({e.__class__.__name__})", systemic=True)
            self.events.put(("name", kind, key, res, (time.perf_counter() - t0) * 1000))
        finally:
            with self._busy_lock:
                self._busy -= 1

    def workers_busy(self) -> bool:
        """Is any name request (job or card) still running in a pool thread? Such a thread would keep the process alive."""
        with self._busy_lock:
            return self._busy > 0

    def note_reader_result(self, res: NameResult):
        """Is the name reader itself working? READER_FAILS service failures in a row (key rejected, no credit, offline,
        rate limit, no answer) raise ONE alert and switch Auto-enter off, so the operator types the names while the
        slip is still in hand; the first read that works again clears it. A slip-specific failure (declined, unreadable
        reply) neither counts nor resets."""
        if res.error:
            if res.systemic:
                self._rd_fail += 1
                if self._rd_fail >= READER_FAILS and not self._rd_down:
                    self.reader_down(res.error)
        else:
            self._last_name_ok = time.perf_counter()
            if self._rd_down:
                self.reader_recovered()
            self._rd_fail = 0

    def reader_down(self, reason: str):
        self._rd_down = True
        self._rd_fail = max(self._rd_fail, READER_FAILS)
        msg = (f"NAME READING IS NOT WORKING ({reason}). The picks are still entered, but the names come out as '?'. "
               "Auto-enter is OFF: each slip waits on the card for you to type the name.")
        self.breaker = self._reader_breaker = msg
        if self.auto_var.get():
            self.auto_var.set(False)
        self.alert(msg, "reader")
        self.log("READER_DOWN", reason=reason)
        self.update_stats()

    def reader_recovered(self):
        self._rd_down = False
        self.clear_alert("reader")
        if self.breaker and self.breaker == self._reader_breaker:      # we switched it off, and nobody has touched it since
            self.breaker = self._reader_breaker = ""
            self.auto_var.set(True)
        self.hint_var.set("Name reading works again.")
        self.touch()                                      # a card that is open now waits for Enter, it is not auto-written
        self.log("READER_OK")
        self.update_stats()

    def on_name(self, kind: str, key: int, res: NameResult, ms: float):
        self._last_name_done = time.perf_counter()
        if kind == "job":
            job = self.jobs.get(key)
            if job is not None and time.monotonic() < self._excel_hold and not job.name_done and not job.writing:
                job.name_answer = (res, ms, True)             # Excel was busy a moment ago: do not ask it again for
                if not job.counted:                           # every name that arrives, write them one after the other
                    job.counted = True
                    self.note_reader_result(res)
                self.name_busy(job)
            elif job is not None:
                self.apply_name(job, res, ms)
        else:
            self.on_card_name(key, res)

    def on_card_name(self, scan_id: int, res: NameResult):
        try:
            self._show_card_name(scan_id, res)
        finally:
            self.note_reader_result(res)                  # (after the name is filled in: a recovery touches the card)

    def _show_card_name(self, scan_id: int, res: NameResult):
        p = self.pending
        if p is None or p.id != scan_id:
            return
        p.name_done = True
        p.slip_jornada = res.jornada
        hidden = bool(getattr(p.reading, "name_hidden", False))          # a hand was over the name strip
        p.name_legible = res.legible and not hidden
        if res.error:
            self.name_note.set(f"{res.error}. Type the name.")
        else:
            if not p.touched and not self.name_var.get().strip():
                self.name_var.set(res.name)
            self.name_note.set("A hand covered the name when the slip was read - check it against the slip." if hidden else
                               "Read by Claude - check it against the handwriting." if res.legible else
                               "Claude was unsure - please check the name.")
        exp = self.expected_jornada()
        if res.jornada and exp and res.jornada != exp:
            self.jor_note.set(f"Slip says JORNADA {res.jornada} but the workbook is JORNADA {exp}!")
            self.jor_lbl.configure(foreground=BAD)
        elif res.jornada:
            self.jor_note.set(f"Slip: JORNADA {res.jornada}")
            self.jor_lbl.configure(foreground=GOOD)

    def apply_name(self, job: Job, res: NameResult, ms: float, *, beep: bool = True):
        """The name of an entered slip is back (possibly out of order, seconds after the next slip went in)."""
        if job.name_done or job.writing:                  # undone meanwhile, already timed out, or waiting for Excel
            return
        job.name_answer = (res, ms, beep)
        self.finish_name(job)

    def finish_name(self, job: Job, final: bool = False):
        """Write the name of ``job`` (its answer is in ``name_answer``) into its row and settle the row. A busy Excel does
        not lose the answer: the job stays 'reading' with ``writing`` set and the write is retried (retry_names) for up
        to NAME_BUSY_S; ``final`` = no more waiting (the window is closing)."""
        res, ms, beep = job.name_answer
        if not job.counted:
            job.counted = True
            self.note_reader_result(res)
        job.name_ms = ms
        job.jornada = res.jornada
        marker = self.cfg.unsure_marker
        if res.error or not res.name.strip():
            job.reason = res.error or "no name found"
            job.guess = ""
            written, state, text = (marker.strip() or "?"), "review", f"REVIEW - name not read ({job.reason})"
        else:
            job.guess = job.tag.format(res.name)
            if res.legible and not job.name_hidden:
                written, state, text = job.guess, "ok", "ok"
            elif res.legible:                             # clear handwriting, but a hand was over the strip when it was read
                job.reason = "a hand covered the name when the slip was read"
                written, state, text = job.guess + marker, "review", "REVIEW - a hand was over the name"
            else:
                job.reason = "Claude was unsure of the name"
                written, state, text = job.guess + marker, "review", "REVIEW - unsure of the name"
        try:
            warnings = job.sink.update_name(job.receipt, written)
            job.written = written
            if warnings:
                self.notice(f"Ticket {job.ticket}: the name is in, but " + "; ".join(warnings), "warn")
        except ExcelBusy as e:
            if not final and not (job.writing and time.monotonic() - job.busy_since >= NAME_BUSY_S):
                self.name_busy(job)
                return
            job.reason = str(e)
            state, text = "review", "NAME NOT WRITTEN - type it"
        except ExcelError as e:                           # includes SlotTaken: someone typed in that cell
            job.reason = str(e)
            state, text = "review", "NAME NOT WRITTEN - type it"
        if job.writing:
            job.writing = False
            if job in self._name_retry:
                self._name_retry.remove(job)
            self.excel_busy_notice()
        exp = job.expected
        if res.jornada and exp and res.jornada != exp:
            state = "wrong"
            text = f"WRONG JORNADA - slip says {res.jornada}, workbook is {exp}" + \
                   (" (name not written)" if job.written is None else "")
            job.reason = text
            self.trip_breaker(f"WRONG JORNADA - ticket {job.ticket} says JORNADA {res.jornada} but the workbook is "
                              f"JORNADA {exp}. Auto-enter is OFF: check that slip (Undo this entry), then tick "
                              "Auto-enter again.")
        job.state, job.text = state, text
        self.refresh_row(job)
        if self.card_job is job:
            self.show_card_name(job)
        if self.rev_job is job:
            self.on_tree_select()
        if state != "ok" and beep and self._rd_fail < READER_FAILS:    # (once the reader is down: one alert, no beep per slip)
            self.beep(True)
        self.settle(job)
        self.update_stats()

    def name_busy(self, job: Job):
        """Excel is busy: keep the finished name and write it a little later. The row stays 'reading' (so the name
        time-out leaves it alone) and says why."""
        now = time.monotonic()
        if not job.writing:
            job.writing, job.busy_since = True, now
            self._name_retry.append(job)
        self._excel_hold = now + BUSY_RETRY_S
        job.text = "writing name (Excel busy)..."
        self.refresh_row(job)
        self.excel_busy_notice()
        self.update_stats()

    def retry_names(self):
        """Tick stage: names waiting for a busy Excel, one attempt per BUSY_RETRY_S (a busy Excel costs a second of waiting,
        not a second per queued name)."""
        while self._name_retry and time.monotonic() >= self._excel_hold:
            job = self._name_retry[0]
            if job.state != "reading" or not job.writing or job.name_answer is None:    # undone / settled meanwhile
                job.writing = False
                self._name_retry.pop(0)
                continue
            self.finish_name(job)
            if job.writing:                                # still busy: the hold has been renewed, try again later
                break
        self.excel_busy_notice()

    def excel_busy_notice(self):
        n = sum(1 for j in self._name_retry if j.writing and j.state == "reading")
        if n:
            self.notice(f"EXCEL IS BUSY - {n} name{'s' if n != 1 else ''} waiting to be written. Press Esc in Excel (leave "
                        "the cell you are editing, close its dialog); they are written by themselves.", "excelbusy")
        else:
            self.clear_notice("excelbusy")

    def show_card_name(self, job: Job):
        self.name_var.set(job.written or "")
        self.final_var.set("")
        if job.state == "ok":
            self.name_note.set("Name read by Claude.")
            self.set_banner(f"ENTERED - ticket {job.ticket}", GOOD)
        else:
            self.name_note.set(job.text)
            self.set_banner(f"ENTERED - ticket {job.ticket} - {job.text}", BAD)
        if job.jornada:
            self.jor_note.set(f"Slip: JORNADA {job.jornada}")
            self.jor_lbl.configure(foreground=BAD if job.state == "wrong" else GOOD)

    def trip_breaker(self, msg: str):
        """A slip for the wrong jornada went in: stop automatic entry until the operator takes over again."""
        self.breaker = msg
        self.auto_var.set(False)
        self.alert(msg, "breaker", beep=False)       # apply_name beeps
        self.log("BREAKER", msg=msg)

    def on_auto_change(self):
        if not self.auto_var.get():
            return
        if self.breaker:
            self.breaker = self._reader_breaker = ""
            self.clear_alert("breaker")
        if self._rd_down:                                 # the operator takes over again: count reader failures anew
            self._rd_down, self._rd_fail = False, 0
            self.clear_alert("reader")
            self.update_stats()

    def check_jobs(self, now: float):
        """Give up on names that never came back (a hung request must not leave a hole in the sheet).
        The clock runs from the moment a worker starts the request - a name that merely waits its turn in a busy pool
        is not late - with a generous cap on the waiting itself in case every worker is stuck (and only while no
        request at all has finished for a whole timeout: a pool that is making progress is just slow)."""
        limit = self.cfg.name_timeout_s
        if limit <= 0:
            return
        queued_wait, hung = 0.0, []
        for job in self.entries:
            if job.state != "reading" or job.writing:         # (a finished name that waits for Excel is not late)
                continue
            sub = job.t_submit or job.t_enter                 # (a job that never got its request is timed from its creation)
            if job.t_started:
                since = now - job.t_started
                if since > limit:
                    self.apply_name(job, NameResult(error=f"no answer after {limit:g} s", systemic=True), (now - sub) * 1000)
                    if job.future is not None:
                        job.future.cancel()
                elif since > max(READER_HANG_S, limit / 2):
                    hung.append(job)
            elif job.future is None:                          # the Excel row exists but no name request was ever made
                if now - sub > limit:
                    self.apply_name(job, NameResult(error="name request never started"), (now - sub) * 1000)
            else:
                queued_wait = max(queued_wait, now - sub)
                if now - sub > 4 * limit and now - self._last_name_done > limit:
                    self.apply_name(job, NameResult(error="no free name reader", systemic=True), (now - sub) * 1000)
                    job.future.cancel()
        # several requests silent for a long time and nothing answered since they started: the reader is down, say so now
        # instead of waiting for every one of them to hit the timeout
        if len(hung) >= READER_FAILS and not self._rd_down and self._last_name_ok < min(j.t_started for j in hung):
            self.reader_down("Claude is not answering")
        warn = queued_wait > 10
        if warn != self._queue_warn:                          # (update_stats only runs on events: refresh when this flips)
            self._queue_warn = warn
            self.update_stats()

    # ------------------------------------------------------- finishing a job
    def settle(self, job: Job):
        """The job reached its end state: write the audit picture + the scanner.log line, free the big arrays."""
        if job.logged:
            return
        job.logged = True
        rd = job.reading
        if rd is not None and job.state != "undone":
            try:
                self.io.submit(self._audit_task, rd, job.picks, job.written or "", job.workbook, job.receipt.sheet,
                               job.receipt.row, job.ticket, job.text if job.manual else job.state)
            except RuntimeError:
                pass
        self.log("ENTRY", ticket=job.ticket, read_ms=job.read_ms, excel_ms=job.excel_ms, name_ms=job.name_ms,
                 status=job.state, name=job.written, reason=job.reason or None, fp_ms=job.fp_ms, feed_ms=job.feed_ms,
                 ui_ms=job.ui_ms, lag_ms=job.lag_ms, cycle_s=job.cycle_s, manual=job.manual or None)
        job.reading = None
        if job.review:                                    # a row the operator still has to fix keeps its picture, as a
            if job.name_crop is not None:                 # small JPEG (the session may be long and every row can fail)
                job.crop_jpg = _jpeg(job.name_crop) or job.crop_jpg
                if job.crop_jpg is not None:
                    job.name_crop = None
        elif len(self.entries) - job.id >= KEEP_CROPS:
            job.name_crop = None

    def _audit_task(self, rd, picks, name, workbook, sheet, row, ticket, status):
        """io thread: write the audit picture + CSV row; a failure goes to scanner.log and (once) to the operator."""
        self.audit.record(rd, picks, name, workbook, sheet, row, ticket, status)
        err = self.audit.last_error
        if err:
            self.scanlog.line("AUDIT_ERROR", ticket=ticket, err=err)
            self.events.put(("audit", err))

    def on_audit(self, err: str):
        self.alert(f"AUDIT LOG PROBLEM - {err}. The slips are still entered in Excel; see scanner.log.", "audit", beep=False)

    # ----------------------------------------------------- tag / name helpers
    def active_tag(self) -> str:
        return self.tag_var.get().strip() if self.tag_on.get() else ""

    def final_name(self, name: str) -> str:
        return format_name(name, self.active_tag(), self.tag_pos.get(), self.cfg.tag_separator,
                           self.cfg.uppercase_names)

    def update_final(self):
        if self.card_job is not None:
            return
        name = self.name_var.get().strip()
        if not name:
            self.final_var.set("")
            return
        self.final_var.set("Will be written as:  " + self.final_name(name))

    def on_tag_change(self):
        tag, on = self.tag_var.get().strip(), self.tag_on.get()
        if on and tag:
            where = "before" if self.tag_pos.get() == "before" else "after"
            self.tag_state.set(f"TAG ON: \"{tag.upper() if self.cfg.uppercase_names else tag}\" {where} name")
            self.tag_state_lbl.configure(foreground=GOOD)
        elif on:
            self.tag_state.set("TAG ON but empty - type the letters")
            self.tag_state_lbl.configure(foreground=WARN)
        else:
            self.tag_state.set("tag off")
            self.tag_state_lbl.configure(foreground="#777")
        self.update_final()
        if self.pending is not None:
            self.touch()                       # changing the tag mid-slip: confirm with Enter, don't auto-write

    def on_name_key(self, e):
        if e.keysym in ("Return", "Escape", "Tab") or e.keysym.startswith(("Shift", "Control", "Alt", "Caps")):
            return
        self.touch()

    def touch(self):
        if self.pending:
            self.pending.touched = True
            self.pending.auto_deadline = None
            self.auto_var_txt.set("")

    def pick_edited(self, i: int):
        self.touch()
        if self.pending is not None:
            self.pending.edited.add(i)
        lbl, note = self.pick_lbls[i]
        lbl.configure(foreground="black")
        note.configure(text="")

    def auto_eligible(self, rd: sv.SlipReading) -> bool:
        """Every row is clear - or simply empty (F1), as long as not too many are empty."""
        if any(v not in (sv.OK, sv.BLANK) for v in rd.verdicts):
            return False
        blanks = sum(v == sv.BLANK for v in rd.verdicts)
        return blanks == 0 or (self.cfg.auto_enter_blanks and blanks <= self.cfg.auto_enter_max_blank)

    def maybe_auto(self):
        """Old hands-on mode (auto_enter_delay_s > 0): wait for the name, count down, then write."""
        p = self.pending
        if p is None or p.touched or p.hold or not self.auto_var.get() or self.sink is None:
            return
        exp = self.expected_jornada()
        jornada_bad = bool(p.slip_jornada and exp and p.slip_jornada != exp)
        ok = (self.auto_eligible(p.reading) and p.name_done and self.name_var.get().strip() and not jornada_bad
              and (p.name_legible or self.reader is None) and self.target_slot() is not None)
        if not ok:
            p.auto_deadline = None
            self.auto_var_txt.set("")
            return
        now = time.time()
        if p.auto_deadline is None:
            p.auto_deadline = now + self.cfg.auto_enter_delay_s
        left = p.auto_deadline - now
        if left <= 0:
            self.commit(auto=True)
        else:
            slot = self.target_slot()
            self.auto_var_txt.set(f"Entering ticket {slot.ticket} in {left:.0f}s ...  "
                                  "(Esc to cancel, or edit anything to take over)")

    # ------------------------------------------------------- commit etc
    def commit(self, auto: bool = False):
        """Classic card: write the slip on the card (name as shown) into Excel."""
        p = self.pending
        if p is None:
            return
        if self.sink is None:
            self.alert("Excel is not connected. Open the workbook, then press 'Connect / reconnect Excel'.")
            return
        name = self.name_var.get().strip()
        if not name:
            self.name_entry.focus_set()
            self.status_var.set("Type the name first.")
            return
        picks = [v.get() for v in self.pick_vars]
        unresolved = [i + 1 for i, v in enumerate(p.reading.verdicts)      # two marks / faint / scribbled, still on 'empty'
                      if v not in (sv.OK, sv.BLANK) and not picks[i] and i not in p.edited]
        if unresolved and not messagebox.askyesno(
                "Unresolved row", f"Row {', '.join(map(str, unresolved))} of the slip could not be read and still says "
                f"{self.marker} (empty).\n\nWrite {self.marker} anyway?"):
            return
        shown = [x or self.marker for x in picks]
        final_name = self.final_name(name)
        snap, exp = self.tag_snapshot(), self.expected_jornada()
        block, ticket = self.target()
        if self._block_gone:
            self.alert("The block you selected is no longer in the sheet. Pick a block in the Block list first.", "block")
            return
        t0 = time.perf_counter()
        try:
            try:
                receipt = self.sink.write_row_fast(block, picks, final_name, ticket)
            except SlotTaken as e:
                if not messagebox.askyesno("Ticket already used", f"{e}\n\nOverwrite it?"):
                    return
                receipt = self.sink.write_row_fast(block, picks, final_name, ticket, allow_overwrite=True)
        except BlockFull as e:
            self.cancel_auto()
            self.block_full(e)
            self.set_banner(self._fail_banner, BAD)
            return
        except ExcelError as e:
            self.cancel_auto()
            if isinstance(e, ExcelBusy):
                self._excel_hold = time.monotonic() + BUSY_RETRY_S
                self.set_banner("NOT ENTERED - Excel is busy: press Esc in Excel (or close its dialog), then press Enter", BAD)
            self.alert(str(e))
            return
        excel_ms = (time.perf_counter() - t0) * 1000
        status = "auto" if auto else ("checked" if p.touched or p.hold or not self.auto_eligible(p.reading) else "ok")
        upd = Update(WAIT_REMOVE, "", timings=p.timings, fingerprint=p.fp)
        job = self.new_job(receipt, p.reading, shown, snap, exp, upd, excel_ms, t0, manual=True)
        job.written, job.guess, job.state, job.text = final_name, final_name, "ok", status
        self.refresh_row(job)
        self.settle(job)
        self.beep(False)
        self.clear_alert("excel")
        self.clear_alert("fullblock")
        self._block_confirm = ""                          # (a human pressed Enter: the block it went to is confirmed)
        self.receipt_notice(receipt)
        self.last_var.set(f"Last entered: ticket {receipt.ticket}  {final_name}   {' '.join(shown)}")
        self._entered_msg = ""
        self.clear_card()
        self.ticket_done()
        self.refresh_blocks(announce=True)
        self.update_stats()
        self.next_backlog()

    def cancel_auto(self):
        if self.pending:
            self.pending.touched = True
            self.pending.auto_deadline = None
            self.auto_var_txt.set("")

    def discard(self):
        """Throw the slip on the card away (it is in no row of Excel - a refused slip included): the engine forgets it,
        so lifting it and putting it back reads it again instead of dropping it as a repeat."""
        p = self.pending
        if p is not None:
            self.clear_card()
            self.status_var.set("Slip discarded.")
            if self.forget_if_current(p.fp):
                self.lift_message("Slip discarded - lift it and place it again to rescan it.")
            self.next_backlog()

    def escape(self):
        p = self.pending
        if p is not None and p.auto_deadline is not None:
            self.cancel_auto()
            self.status_var.set("Auto-enter cancelled - edit, then press Enter to write.")
        else:
            self.discard()

    # ------------------------------------------------------- history / review
    def refresh_row(self, job: Job):
        if not self.tree.exists(job.iid):
            return
        vals = list(self.tree.item(job.iid, "values"))
        vals[2] = job.written if job.written is not None else ("(reading...)" if job.state == "reading" else "")
        vals[4] = job.text
        tag = {"reading": "reading", "review": "review", "wrong": "wrong", "undone": "undone"}.get(job.state, "")
        self.tree.item(job.iid, values=vals, tags=(tag,) if tag else ())

    def select_next_review(self):
        """Select the oldest history row that still needs a look."""
        for job in self.entries:
            if job.review:
                self.tree.selection_set(job.iid)
                self.tree.see(job.iid)
                self.on_tree_select()
                self.scroll_mid(True)                     # (small screen: bring the review panel into view)
                return
        self.rev_info.set("Nothing to review.")

    def on_tree_select(self):
        sel = self.tree.selection()
        job = self.iid_jobs.get(sel[0]) if sel else None
        self.rev_job = job
        usable = job is not None and job.state != "undone"
        if job is None:
            self.rev_info.set("Select a row to see its name picture.")
        else:
            self.rev_info.set(f"Ticket {job.ticket}: {job.text}")
        crop = job.crop() if usable else None
        if crop is not None:
            self._photos["rev"] = _to_photo(crop, width=self.rev_w)
            self.rev_img.configure(image=self._photos["rev"])
        else:
            self.rev_img.configure(image="")
        self.rev_name.set((job.guess if job.review else (job.written or "")) if usable else "")
        can_edit = usable and job.name_done
        self.rev_entry.configure(state="normal" if can_edit else "disabled")
        self.rev_save.configure(state="normal" if can_edit else "disabled")
        self.rev_undo.configure(state="normal" if usable else "disabled")

    def save_review_name(self):
        job = self.rev_job
        if job is None or job.state == "undone" or not job.name_done:
            return
        text = job.tag.format(self.rev_name.get())
        if not text:
            self.status_var.set("Type the name first.")
            self.rev_entry.focus_set()
            return
        try:
            job.sink.update_name(job.receipt, text, expect=job.written)
        except ExcelError as e:                           # SlotTaken: the cell was changed in Excel meanwhile
            self.alert(f"Ticket {job.ticket}: {e}")
            return
        job.written, job.guess, job.reason = text, text, ""
        job.state, job.text = "ok", "ok (name fixed by hand)"
        self.refresh_row(job)
        if self.card_job is job:
            self.show_card_name(job)
        self.on_tree_select()
        self.log("NAME_FIXED", ticket=job.ticket, name=text)
        self.update_stats()

    def undo(self):
        job = next((j for j in reversed(self.entries) if j.state != "undone"), None)
        if job is not None:
            self.undo_job(job)

    def undo_selected(self):
        if self.rev_job is not None:
            self.undo_job(self.rev_job)

    def undo_job(self, job: Job):
        """Clear the ticket from Excel again - also while its name is still being read (the late result is ignored)."""
        if job.state == "undone":
            return
        who = job.written or job.guess or "name still reading"
        if not messagebox.askyesno("Undo", f"Clear ticket {job.ticket} ({who}) from Excel?"):
            return
        if job.state == "undone":                         # (the dialog keeps the event loop running)
            return
        try:
            warnings = job.sink.undo(job.receipt)
        except ExcelError as e:
            self.alert(f"Undo failed: {e}")
            return
        if warnings:
            self.notice(f"Ticket {job.ticket} was cleared, but " + "; ".join(warnings), "warn")
        was_reading = job.state == "reading"
        self._entered_msg = ""                            # (the status bar must not go on saying 'Entered ticket N')
        job.writing = False
        job.state, job.text, job.written = "undone", "UNDONE", None
        if job.future is not None:
            job.future.cancel()
        self.refresh_row(job)
        if self.card_job is job:
            self.card_job = None
            self.clear_card()
        self.last_var.set(f"Undone: ticket {job.ticket}")
        if self.forget_if_current(job.fp):                # the slip lying there is not in Excel any more
            self.lift_message("Undone - lift the slip and place it again to scan it.")
        self.settle(job)
        self.log("UNDO", ticket=job.ticket, while_reading=was_reading or None)
        if self.rev_job is job:
            self.on_tree_select()
        job.reading = job.name_crop = job.crop_jpg = None
        self.ticket_done()
        self.refresh_blocks()
        self.update_stats()

    # ------------------------------------------------------------ stats
    def avg_cycle(self) -> float | None:
        t = self.entry_times[-11:]
        gaps = [b - a for a, b in zip(t, t[1:]) if b - a <= BREAK_S]
        return sum(gaps) / len(gaps) if gaps else None

    def update_stats(self):
        live = [j for j in self.entries if j.state != "undone"]
        parts = [f"Entered: {len(live)}", f"slips read: {self.slips_read}"]
        avg = self.avg_cycle()
        parts.append(f"{avg:.1f} s per slip (last 10)" if avg is not None else "- s per slip")
        if self.entries:
            j = self.entries[-1]
            name = "-" if j.manual else (f"{j.name_ms / 1000:.1f} s" if j.name_ms is not None else "reading...")
            parts.append(f"last slip: read {j.read_ms:.0f} ms, Excel {j.excel_ms:.0f} ms, name {name}")
        waiting = sum(j.state == "reading" and not j.writing for j in self.entries)
        if waiting:
            parts.append(f"names being read: {waiting}")
        busy = sum(j.state == "reading" and j.writing for j in self.entries)
        if busy:
            parts.append(f"names waiting for Excel: {busy}")
        if self._queue_warn:
            parts.append("names are queueing: Claude is slower than the stack - raise name_workers")
        if self.dupes_skipped:
            parts.append(f"Skipped repeats: {self.dupes_skipped}")
        if self._fps:
            parts.append(f"camera {self._fps:.0f} fps")
        if self.low_note:
            parts.append(self.low_note)
        if self.backlog:
            parts.append(f"slips waiting behind the card: {len(self.backlog)}")
        if self._rd_down:
            parts.append("NAME READER DOWN")
        review = sum(j.review for j in self.entries)
        if review:
            parts.append(f"TO REVIEW: {review}")
        self.stats_var.set("     |     ".join(parts))
        self.stats_lbl.configure(foreground=BAD if (review or self.backlog or self._rd_down) else
                                 WARN if (self._queue_warn or self._fps_bad >= FPS_CHECKS or busy) else "#222")

    # --------------------------------------------------------- main loop
    def tick(self):
        try:
            self._tick()
        except Exception as e:                # the loop must survive anything
            self.report_callback_exception(type(e), e, e.__traceback__)
        if not self.closing:
            self._tick_id = self.after(TICK_MS, self.tick)

    def _stage(self, fn, *args):
        """One stage of the tick: a bug in it must not stop the stages after it (the name time-outs, the backlog ...)."""
        try:
            fn(*args)
        except Exception as e:
            self.report_callback_exception(type(e), e, e.__traceback__)

    def _events(self):
        """Handle what the worker threads handed over - for at most TICK_BUDGET_S: one slow Excel call per event (a busy
        Excel costs ~1.2 s each) must not add up to seconds inside one Tk callback, where nothing is painted and
        Windows calls the window 'not responding'. The rest waits for the next tick."""
        end = time.perf_counter() + TICK_BUDGET_S
        for _ in range(50):
            try:
                ev = self.events.get_nowait()
            except queue.Empty:
                break
            try:
                if ev[0] == "scan":
                    self.on_scan(ev[1], ev[2] if len(ev) > 2 else None)
                elif ev[0] == "name":
                    self.on_name(ev[1], ev[2], ev[3], ev[4])
                elif ev[0] == "audit":
                    self.on_audit(ev[1])
            except Exception as e:            # one bad event must not drop the ones behind it
                self.report_callback_exception(type(e), e, e.__traceback__)
            if time.perf_counter() > end:
                break

    def check_camera(self, now: float):
        """About once a second: how many frames per second does the camera really deliver? Too few (a C920s in dim light
        drops to 10-15 on its own) and the lift of a slip can fall between two pictures: say so, until it is fixed."""
        if now < self._next_fps or self.grabber is None:
            return
        self._next_fps = now + 1.0
        fps = self.grabber.fps()
        if fps <= 0:
            return
        shown = round(fps) != round(self._fps)
        self._fps = fps
        if fps < FPS_WARN:
            self._fps_bad += 1
        elif fps >= FPS_OK:
            self._fps_bad = 0
            self.clear_notice("camera")
        if self._fps_bad >= FPS_CHECKS:
            self.notice(f"CAMERA IS SLOW (under {FPS_WARN:.0f} pictures per second): slips may be missed. Add light, or turn "
                        "off Low Light Compensation / RightLight in Camera settings.", "camera")
            shown = True
        if shown:
            self.update_stats()

    def _tick(self):
        now = time.perf_counter()
        self._stage(self._events)
        self._stage(self.retry_names)
        self._stage(self.check_camera, now)
        if now >= self._next_watch:
            self._next_watch = now + 0.25
            self._stage(self.check_jobs, now)
        if self.backlog and self.pending is None:         # safety net: slips never wait without a card in front of them
            self._stage(self.drain_backlog)
        if now >= self._next_preview:
            self._stage(self.render_preview, now)
        if self.pending is not None:
            self._stage(self.maybe_auto)

    def render_preview(self, now: float):
        """~15 fps, only when there is a new frame, and the full-size picture is never copied or converted."""
        with self._lock:
            latest = self._latest
        if self.grabber is not None and self.grabber.age() > 2.0:
            self.status_var.set("No picture from the camera - is it unplugged, or used by another app?")
            return
        if latest is None or latest[0] == self._prev_fid:
            return
        fid, frame, upd = latest
        self._prev_fid = fid
        self._next_preview = now + 1.0 / PREVIEW_FPS
        h, w = frame.shape[:2]
        s = self.preview_w / w
        small = cv2.resize(frame, (self.preview_w, max(1, round(h * s))), interpolation=cv2.INTER_AREA)
        quad = None if upd.quad is None else np.asarray(upd.quad, np.float32) * s
        vis = sv.draw_quad(small, quad, (0, 0, 230) if upd.message == HAND_MESSAGE else
                           (0, 200, 0) if upd.progress >= 1 else (0, 200, 255), 2)
        img = Image.fromarray(cv2.cvtColor(vis, cv2.COLOR_BGR2RGB))
        if self._prev_photo is not None and (self._prev_photo.width(), self._prev_photo.height()) == img.size:
            self._prev_photo.paste(img)
        else:
            self._prev_photo = ImageTk.PhotoImage(img)
            self.preview.configure(image=self._prev_photo)
        if self.pending is None:
            msg = upd.message
            if upd.state != WAIT_REMOVE:
                self._entered_msg = self._lift_msg = ""
                self._dup_latched = False
            elif not upd.duplicate and not upd.error and not self._dup_latched:
                msg = self._entered_msg or self._lift_msg or msg    # (a dropped repeat keeps the engine's 'lift it off' text)
            if self.status_var.get() != msg:
                self.status_var.set(msg)
        if self.progress["value"] != upd.progress:
            self.progress["value"] = upd.progress

    def names_in_flight(self) -> int:
        return sum(1 for j in self.entries if j.state == "reading" and j.future is not None and not j.future.done())

    def shutdown(self):
        """Stop every thread we started. Safe to call twice."""
        if self.closing:
            return
        self.closing = True
        for after_id in (getattr(self, "_tick_id", None), self._focus_after):
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
        self._stop_camera()
        for job in self.entries:                          # entries still waiting for their name get a log line too
            if job.state == "reading":
                self.log("ENTRY", ticket=job.ticket, read_ms=job.read_ms, excel_ms=job.excel_ms, status="closed-before-name")
        self.pool.shutdown(wait=False, cancel_futures=True)
        self.io.shutdown(wait=True)

    def destroy(self):
        self.shutdown()
        super().destroy()

    def finish_names(self, wait: float):
        """Closing: pump Tk until the names that are still being read have come back (at most ``wait`` s), then give up on
        the rest the way a timeout does - '?' in the cell, a review row, the audit picture and the CSV row."""
        def pending_names():
            return [j for j in self.entries if j.state == "reading" and j.future is not None]
        end = time.perf_counter() + wait
        while pending_names() and time.perf_counter() < end:
            self.status_var.set(f"Closing - waiting for {len(pending_names())} name(s) still being read ...")
            try:
                self.update()                             # lets tick() deliver the 'name' events (apply_name -> settle)
            except tk.TclError:
                return
            time.sleep(0.02)
        for job in list(self.entries):
            if job.state == "reading":
                try:
                    if job.writing and job.name_answer is not None:      # the name is here, Excel was busy: last try
                        self.finish_name(job, final=True)
                    else:
                        self.apply_name(job, NameResult(error="app closed before the name came back"), 0.0, beep=False)
                except Exception as e:
                    self.log("CLOSE_ERROR", ticket=job.ticket, err=repr(e))

    def on_close(self):
        if self._close_started:
            return
        unentered = (1 if self.pending is not None else 0) + len(self.backlog)
        if unentered:
            self._close_started = True                    # (the dialog keeps the event loop running: no second close)
            try:
                go = messagebox.askyesno("Close", f"{unentered} slip(s) are NOT in Excel yet "
                                         f"({1 if self.pending is not None else 0} on the card, {len(self.backlog)} "
                                         "waiting behind it).\n\nClose anyway and lose them?")
            finally:
                self._close_started = False
            if not go:
                return
            self.log("CLOSE_DROPPED", card=1 if self.pending is not None else 0, backlog=len(self.backlog))
        self._close_started = True
        try:
            self._stop_camera()                           # no new slips while the names are collected
            self.finish_names(CLOSE_WAIT_S)
            if not unentered and (self.pending is not None or self.backlog):     # a slip already on its way needed a card
                self.log("CLOSE_DROPPED", card=1 if self.pending is not None else 0, backlog=len(self.backlog), late=1)
            self.flush_io()                               # the audit pictures / CSV rows / log lines of those names
        except Exception as e:
            self.log("CLOSE_ERROR", err=repr(e))
        finally:
            try:
                self.cfg.name_tag = self.tag_var.get().strip()
                self.cfg.name_tag_position = self.tag_pos.get()
                self.cfg.auto_enter = self.auto_var.get() or bool(self.breaker)   # the breaker's "off" is for this session only
                self.cfg.autofocus = self.af_var.get()
                self.cfg.manual_focus = int(self.focus_var.get())
                self.cfg.save()
            finally:
                self.destroy()


def run(cfg: Config, source=None, backend=None):
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App(cfg, source, backend)
    app.mainloop()
    if app.workers_busy() or app.names_in_flight():
        # a name request that is still hanging (even one the app already gave up on) would keep this process alive for
        # up to a minute: pool threads are non-daemon and joined at exit. Everything is saved and flushed by now.
        sys.stdout.flush()
        sys.stderr.flush()
        os._exit(0)
