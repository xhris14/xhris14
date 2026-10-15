"""Desktop window (tkinter).

Stack workflow ("fast mode"): the engine reads a slip, the picks go into Excel at once, and the handwritten
name is read in the background and filled into the same row a few seconds later. The engine is never paused,
so the next slip of the stack is picked up as soon as the previous one is lifted. Slips that need a human
(unclear marks, no name reader, auto-enter off) open the classic card and pause the engine until Enter / Esc.

Threading rules: tkinter and Excel (COM) are touched from the Tk thread only. Worker threads (camera, name
reads, disk writes) hand their results over through ``self.events``.
"""
from __future__ import annotations

import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from tkinter import messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageTk

from . import slip_vision as sv
from . import sound
from .audit import AuditLog, ScanLog
from .camera import CameraError, FrameGrabber, WebcamSource
from .capture import WAIT_REMOVE, CaptureEngine, Update
from .config import Config, app_dir
from .excel_sink import ExcelError, ExcelSink, Receipt, SlotTaken, format_name, open_backend
from .name_reader import NameResult, make_reader

FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
GOOD, WARN, BAD = "#1b7f3b", "#b36b00", "#c62828"
VERDICT_TEXT = {sv.MULTIPLE: "two marks", sv.WEAK: "faint mark", sv.HEAVY: "scribbled?"}

PREVIEW_W = 520             # the live picture is shrunk to this before anything is drawn on it
PREVIEW_FPS = 15
TICK_MS = 25
BREAK_S = 60                # a gap between two entries longer than this is a break, not part of the average
KEEP_CROPS = 20             # name pictures kept for the most recent entries (review items are always kept)


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
    hold: bool = False                    # Excel refused it in fast mode: only a human may enter it now
    name_done: bool = False
    name_legible: bool = False
    slip_jornada: str | None = None
    auto_deadline: float | None = None
    timings: dict = field(default_factory=dict)


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
        self.minsize(1180, 780)
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
        self.backlog: list = []                      # readings that arrived while a card was open (never dropped)
        self.card_job: Job | None = None             # the entered slip the card is showing (read-only)
        self.rev_job: Job | None = None              # the history row selected for review
        self.infos: list = []
        self.sel_key: str | None = None
        self.ticket_custom = False        # operator typed a ticket number
        self.breaker = ""                 # message while auto-enter is switched off by the circuit breaker
        self._alert_kind = ""
        self._entered_msg = ""
        self.callback_errors: list = []
        self.closing = False
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
        self.bind_all("<Return>", lambda e: self.commit())
        self.bind_all("<Escape>", lambda e: self.escape())
        self.bind_all("<F2>", lambda e: self.tag_on.set(not self.tag_on.get()))
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.log("START", fast=cfg.auto_enter_delay_s <= 0, workers=cfg.name_workers, stable=cfg.stable_frames,
                 reader=type(self.reader).__name__ if self.reader else "none",
                 reset=",".join(getattr(cfg, "migrated", [])) or None)
        if getattr(cfg, "migrated", None):
            self.hint_var.set("Settings from an older version were reset to the new defaults: " +
                              ", ".join(cfg.migrated) + ".")
        self.update_idletasks()                      # size to what the widgets need (high-DPI fonts grow), within the screen
        self.geometry(f"{min(max(1240, self.winfo_reqwidth()), self.winfo_screenwidth() - 20)}x"
                      f"{min(max(840, self.winfo_reqheight() + 70), self.winfo_screenheight() - 60)}")   # +70: room for the alert strip
        self.connect_excel()
        self.start_camera()
        self._tick_id = self.after(TICK_MS, self.tick)

    # ------------------------------------------------------------------ UI
    def _build(self):
        pad = dict(padx=6, pady=3)
        style = ttk.Style(self)
        style.configure("Big.TButton", font=(FONT, 12, "bold"), padding=8)
        style.configure("Card.TLabelframe.Label", font=(FONT, 11, "bold"))

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
        tk.Label(self.alert_frame, textvariable=self.alert_var, bg=BAD, fg="white", font=(FONT, 13, "bold"),
                 wraplength=1050, justify="left").pack(side="left", padx=10, pady=6)
        tk.Button(self.alert_frame, text="Dismiss", command=self.clear_alert).pack(side="right", padx=10)

        # ----- target row
        tgt = ttk.LabelFrame(self, text="Enter into")
        tgt.pack(fill="x", **pad)
        self._tgt = tgt
        ttk.Label(tgt, text="Block:").grid(row=0, column=0, **pad)
        self.block_var = tk.StringVar()
        self.block_cb = ttk.Combobox(tgt, textvariable=self.block_var, state="readonly", width=58)
        self.block_cb.grid(row=0, column=1, **pad)
        self.block_cb.bind("<<ComboboxSelected>>", lambda e: self.on_block_change())
        ttk.Label(tgt, text="Ticket #:").grid(row=0, column=2, **pad)
        self.ticket_var = tk.StringVar()
        ticket_entry = ttk.Entry(tgt, textvariable=self.ticket_var, width=7, font=(FONT, 12, "bold"))
        ticket_entry.grid(row=0, column=3, **pad)
        ticket_entry.bind("<Key>", self.on_ticket_key)
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
        ttk.Entry(tagrow, textvariable=self.tag_var, width=8, font=(FONT, 12, "bold")).pack(side="left", padx=(0, 8))
        ttk.Radiobutton(tagrow, text="Before name", value="before", variable=self.tag_pos).pack(side="left", padx=2)
        ttk.Radiobutton(tagrow, text="After name", value="after", variable=self.tag_pos).pack(side="left", padx=2)
        ttk.Checkbutton(tagrow, text="Add tag  (F2)", variable=self.tag_on).pack(side="left", padx=14)
        self.tag_state = tk.StringVar()
        self.tag_state_lbl = ttk.Label(tagrow, textvariable=self.tag_state, font=(FONT, 11, "bold"))
        self.tag_state_lbl.pack(side="left", padx=6)
        for v in (self.tag_var, self.tag_pos, self.tag_on):
            v.trace_add("write", lambda *a: self.on_tag_change())

        # ----- main: preview | card
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, **pad)
        left = ttk.Frame(main)
        left.pack(side="left", fill="y", padx=4)
        self.preview = ttk.Label(left)
        self.preview.pack()
        self.status_var = tk.StringVar(value="Starting camera...")
        ttk.Label(left, textvariable=self.status_var, font=(FONT, 13, "bold"), wraplength=520).pack(anchor="w", pady=4)
        prog = ttk.Frame(left)
        prog.pack(anchor="w", fill="x", pady=(0, 4))
        self.progress = ttk.Progressbar(prog, maximum=1.0, length=400)
        self.progress.pack(side="left", fill="x", expand=True)
        ttk.Button(prog, text="Capture now", command=self.capture_now).pack(side="left", padx=(8, 0))
        self.last_var = tk.StringVar()
        ttk.Label(left, textvariable=self.last_var, font=(FONT, 11), foreground=GOOD, wraplength=520).pack(anchor="w")
        self.hint_var = tk.StringVar(value=self.reader_note)
        ttk.Label(left, textvariable=self.hint_var, foreground="#555", wraplength=520).pack(anchor="w")
        rev = ttk.LabelFrame(left, text="Selected entry  (click a row of the list below)")
        rev.pack(anchor="w", fill="x", pady=(10, 0))
        self.rev_img = ttk.Label(rev)
        self.rev_img.grid(row=0, column=0, columnspan=3, sticky="w", padx=4, pady=2)
        self.rev_info = tk.StringVar()
        ttk.Label(rev, textvariable=self.rev_info, wraplength=500, foreground=BAD).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=4)
        self.rev_name = tk.StringVar()
        self.rev_entry = ttk.Entry(rev, textvariable=self.rev_name, font=(FONT, 13, "bold"), width=20)
        self.rev_entry.grid(row=2, column=0, padx=4, pady=3)
        self.rev_entry.bind("<Return>", lambda e: (self.save_review_name(), "break")[1])
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
        self.stats_lbl = ttk.Label(self, textvariable=self.stats_var, font=(FONT, 11, "bold"))
        self.stats_lbl.pack(fill="x", padx=10, pady=(2, 0))

        # ----- history (red rows = needs a look) + the review panel for the selected row
        hist = ttk.LabelFrame(self, text="Entered this session  (red = check the name)")
        hist.pack(fill="x", **pad)
        cols = ("time", "ticket", "name", "picks", "status")
        self.tree = ttk.Treeview(hist, columns=cols, show="headings", height=4)
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
        ttk.Button(side, text="Undo last", command=self.undo).pack(anchor="w", pady=2)
        ttk.Button(side, text="Next to review", command=self.select_next_review).pack(anchor="w", pady=2)
        self.clear_card()
        self.on_tag_change()
        self.on_tree_select()
        self.update_stats()

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
        self._stop = threading.Event()
        self.grabber = FrameGrabber(self.source)
        self._cam_thread = threading.Thread(target=self._camera_loop, args=(self._stop, self.grabber),
                                            name="vision", daemon=True)
        self._cam_thread.start()
        self.status_var.set(f"Using {self.source.label}. Place a slip flat under the camera.")
        self.log("CAMERA", source=self.source.label)

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
            self.engine.paused = self.pending is not None
            try:
                upd = self.engine.feed(frame)
            except Exception as e:            # never let a vision bug kill the camera thread
                upd = Update("idle", f"vision error: {e}")
                if str(e) != self._last_vision_err:
                    self._last_vision_err = str(e)
                    self.log("VISION_ERROR", err=str(e))
            with self._lock:
                self._latest = (fid, frame, upd)
            if upd.reading is not None or upd.error:
                self.events.put(("scan", upd, time.perf_counter()))

    def capture_now(self):
        with self._lock:
            latest = self._latest
        if latest is None or self.pending is not None:
            return
        frame = latest[1]

        def run():                            # a slip read takes ~200 ms: keep it off the Tk thread
            try:
                upd = self.engine.capture_now(frame)
            except Exception as e:
                upd = Update("idle", f"capture failed: {e}", error=str(e))
            self.events.put(("scan", upd, time.perf_counter()))
        threading.Thread(target=run, name="capture-now", daemon=True).start()

    # ------------------------------------------------------------ Excel
    def connect_excel(self):
        try:
            self.backend = self._fixed_backend or open_backend(self.cfg.backend, self.cfg.workbook, self.cfg.workbook_path)
            self.sink = ExcelSink(self.backend, self.cfg.sheets, self.cfg.name_header,
                                  self.cfg.uppercase_names, self.cfg.save_after_write,
                                  self.cfg.blank_marker, self.cfg.blank_font_size)
        except ExcelError as e:
            self.sink = None
            self.excel_var.set(f"Excel: NOT CONNECTED - {e}")
            self.excel_lbl.configure(foreground=BAD)
            self.block_cb["values"] = []
            self.log("EXCEL_NOT_CONNECTED", err=str(e))
            return
        self.excel_var.set(f"Excel: {self.backend.label}")
        self.excel_lbl.configure(foreground=GOOD)
        self.log("EXCEL", workbook=getattr(self.backend, "workbook_name", ""), blocks=len(self.sink.blocks))
        self.refresh_blocks(select_default=True)

    def refresh_blocks(self, select_default=False):
        if self.sink is None:
            return
        self.infos = self.sink.block_infos()
        self.block_cb["values"] = [i.label for i in self.infos]
        chosen = None if select_default else next((i for i in self.infos if i.block.key == self.sel_key), None)
        if chosen is None:
            chosen = self.sink.default_block()
        if chosen:
            self.select_info(chosen)

    def on_ticket_key(self, e):
        if e.keysym not in ("Return", "Escape", "Tab"):
            self.ticket_custom = True

    def select_info(self, info):
        self.ticket_custom = False
        self.sel_key = info.block.key
        self.block_var.set(info.label)
        slot = info.first_free
        self.ticket_var.set(str(slot.ticket) if slot and slot.ticket is not None else "")

    def selected_info(self):
        return next((i for i in self.infos if i.block.key == self.sel_key), None)

    def on_block_change(self):
        info = next((i for i in self.infos if i.label == self.block_var.get()), None)
        if info:
            self.select_info(info)

    def expected_jornada(self) -> str:
        if self.cfg.expected_jornada:
            return self.cfg.expected_jornada
        m = re.search(r"jornada\s*(\d+)", getattr(self.backend, "workbook_name", ""), re.I)
        return m.group(1) if m else ""

    def target_slot(self):
        if self.sink is None:
            return None
        t = self.ticket_var.get().strip()
        if self.ticket_custom and t.isdigit():
            return self.sink.slot_for_ticket(int(t))
        info = self.selected_info()
        return info.first_free if info else None

    def target(self):
        """(block, explicit ticket or None) for the next write_row_fast."""
        info = self.selected_info()
        t = self.ticket_var.get().strip()
        return (info.block if info else None), (int(t) if self.ticket_custom and t.isdigit() else None)

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
        self.alert_var.set(msg)
        self._alert_kind = kind
        if not self.alert_frame.winfo_ismapped():
            self.alert_frame.pack(fill="x", padx=6, pady=3, before=self._tgt)
        self.log("ALERT", kind=kind, msg=msg)
        if beep:
            self.beep(True)

    def clear_alert(self, kind: str | None = None):
        if kind is not None and kind != self._alert_kind:
            return
        self._alert_kind = ""
        self.alert_var.set("")
        self.alert_frame.pack_forget()

    def set_banner(self, text: str, colour: str | None = None):
        self.banner_var.set(text)
        self.banner.configure(bg=colour or self._banner_look[0], fg="white" if colour else self._banner_look[1])

    def report_callback_exception(self, exc, val, tb):
        """An exception inside a Tk callback must not go unnoticed (or stop the tick loop)."""
        import traceback
        self.callback_errors.append(val)
        key, now = f"{exc.__name__}: {val}", time.monotonic()
        if key == self._last_cb_err[0] and now - self._last_cb_err[1] < 5:
            return                                        # the same error again and again (a 25 ms loop): say it once
        self._last_cb_err = (key, now)
        where = traceback.format_tb(tb)[-1].strip().replace("\n", " ") if tb else None
        self.log("CALLBACK_ERROR", err=key, where=where)
        try:
            traceback.print_exception(exc, val, tb)
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
        self._photos["slip"] = _to_photo(blank)
        self.slip_img.configure(image=self._photos["slip"])
        self.name_img.configure(image="")

    def show_reading(self, rd: sv.SlipReading):
        """Straightened slip, name picture and pick buttons for a freshly read slip."""
        self._photos["slip"] = _to_photo(sv.draw_reading(rd), width=310, height=325)
        self.slip_img.configure(image=self._photos["slip"])
        self._photos["name"] = _to_photo(rd.name_crop, width=310)
        self.name_img.configure(image=self._photos["name"])
        for i, (v, (lbl, note)) in enumerate(zip(self.pick_vars, self.pick_lbls)):
            v.set(rd.picks[i])
            verdict = rd.verdicts[i]
            lbl.configure(foreground="black" if verdict == sv.OK else WARN if verdict == sv.BLANK else BAD)
            note.configure(text=f"empty = {self.marker}" if verdict == sv.BLANK else VERDICT_TEXT.get(verdict, ""))

    def on_scan(self, upd: Update, t_event: float | None = None):
        if upd.reading is None:
            self.status_var.set(upd.message)
            if upd.error:                                   # the operator has to re-place the slip
                self.log("UNREADABLE", err=upd.error)
                self.beep(True)
            return
        if self.pending is not None:                        # e.g. Excel was busy for seconds and two scans queued up
            self.backlog.append((upd, t_event))
            self.update_stats()
            return
        t0 = time.perf_counter()
        lag = (t0 - t_event) * 1000 if t_event else 0.0
        if self.fast_ok(upd.reading):
            if self.enter_fast(upd, t0, lag):
                return
            self.open_card(upd, hold=True)                   # Excel said no: the operator takes this slip
            self.set_banner("NOT ENTERED - fix Excel, then press Enter", BAD)
            return
        self.open_card(upd)

    def next_backlog(self):
        """A card was closed: take the slip that was waiting behind it, if any."""
        if self.backlog and self.pending is None:
            upd, t_event = self.backlog.pop(0)
            self.on_scan(upd, t_event)
            self.update_stats()

    def fast_ok(self, rd: sv.SlipReading) -> bool:
        """Can this slip go straight into Excel with its name to follow?"""
        return (self.auto_var.get() and self.cfg.auto_enter_delay_s <= 0 and self.reader is not None
                and self.sink is not None and self.auto_eligible(rd))

    def tag_snapshot(self) -> TagSnap:
        return TagSnap(self.active_tag(), self.tag_pos.get(), self.cfg.tag_separator, self.cfg.uppercase_names)

    def enter_fast(self, upd: Update, t0: float, lag_ms: float) -> bool:
        """Picks into Excel now, name later. False = Excel refused (the alert is up; the caller opens the card)."""
        rd = upd.reading
        snap, exp = self.tag_snapshot(), self.expected_jornada()      # taken NOW: F2 may change before the name is back
        block, ticket = self.target()
        t1 = time.perf_counter()
        try:
            receipt = self.sink.write_row_fast(block, rd.picks, name="", ticket=ticket)
        except Exception as e:                                         # ExcelError / SlotTaken / anything unexpected
            self.alert(f"NOT ENTERED - {e}", "excel")
            return False
        excel_ms = (time.perf_counter() - t1) * 1000
        self.beep(False)                                               # after the write, so a beep means "it is in"
        shown = [p or self.marker for p in rd.picks]
        job = self.new_job(receipt, rd, shown, snap, exp, upd, excel_ms, t0)
        self.show_entered(job)
        self.clear_alert("excel")
        self.refresh_blocks()
        job.t_submit = time.perf_counter()
        try:
            job.future = self.pool.submit(self._name_worker, self.reader, "job", job.id, rd.name_crop, rd.title_crop)
        except RuntimeError:                                           # pool already shut down (closing)
            job.future = None
        self._entered_msg = f"Entered ticket {receipt.ticket}. Lift the slip - the next one is read automatically."
        self.last_var.set(f"Last entered: ticket {receipt.ticket}   {' '.join(shown)}   (name follows)")
        self.update_stats()
        job.ui_ms = (time.perf_counter() - t0) * 1000
        job.lag_ms = lag_ms
        return True

    def new_job(self, receipt: Receipt, rd: sv.SlipReading, shown: list, snap: TagSnap, exp: str, upd: Update,
                excel_ms: float, t0: float, manual: bool = False) -> Job:
        self._job_id += 1
        job = Job(self._job_id, str(self._job_id), self.sink, receipt, shown, snap, exp,
                  getattr(self.backend, "workbook_name", ""), t0, time.perf_counter(), reading=rd,
                  name_crop=_shrink(rd.name_crop, 640), read_ms=upd.timings.get("read_ms", 0.0),
                  excel_ms=excel_ms, manual=manual)
        job.feed_ms = upd.timings.get("feed_ms", 0.0)
        job.fp_ms = upd.timings.get("fingerprint_ms", 0.0)
        job.cycle_s = (job.t_enter - self.entry_times[-1]) if self.entry_times else None
        self.entries.append(job)
        self.jobs[job.id] = job
        self.iid_jobs[job.iid] = job
        self.entry_times.append(job.t_enter)
        if len(self.entries) > KEEP_CROPS:                  # only recent entries (and review items) keep their picture
            old = self.entries[-KEEP_CROPS - 1]
            if not old.review:
                old.name_crop = None
        t = receipt.ticket if receipt.ticket is not None else f"row {receipt.row}"
        self.tree.insert("", 0, iid=job.iid, values=(time.strftime("%H:%M:%S"), t, "(reading...)", " ".join(shown),
                                                     job.text), tags=("reading",))
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
        self.pending = Pending(self._scan_id, rd, hold=hold, timings=dict(upd.timings))
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
        try:
            res = reader.read(name_crop, title_crop)
        except Exception as e:
            res = NameResult(error=f"name reading failed ({e.__class__.__name__})")
        self.events.put(("name", kind, key, res, (time.perf_counter() - t0) * 1000))

    def on_name(self, kind: str, key: int, res: NameResult, ms: float):
        if kind == "job":
            job = self.jobs.get(key)
            if job is not None:
                self.apply_name(job, res, ms)
        else:
            self.on_card_name(key, res)

    def on_card_name(self, scan_id: int, res: NameResult):
        p = self.pending
        if p is None or p.id != scan_id:
            return
        p.name_done = True
        p.slip_jornada = res.jornada
        p.name_legible = res.legible
        if res.error:
            self.name_note.set(f"{res.error}. Type the name.")
        else:
            if not p.touched and not self.name_var.get().strip():
                self.name_var.set(res.name)
            self.name_note.set("Read by Claude - check it against the handwriting." if res.legible else
                               "Claude was unsure - please check the name.")
        exp = self.expected_jornada()
        if res.jornada and exp and res.jornada != exp:
            self.jor_note.set(f"Slip says JORNADA {res.jornada} but the workbook is JORNADA {exp}!")
            self.jor_lbl.configure(foreground=BAD)
        elif res.jornada:
            self.jor_note.set(f"Slip: JORNADA {res.jornada}")
            self.jor_lbl.configure(foreground=GOOD)

    def apply_name(self, job: Job, res: NameResult, ms: float):
        """The name of an entered slip is back (possibly out of order, seconds after the next slip went in)."""
        if job.name_done:                                 # undone meanwhile, or already timed out: ignore the late one
            return
        job.name_ms = ms
        job.jornada = res.jornada
        marker = self.cfg.unsure_marker
        if res.error or not res.name.strip():
            job.reason = res.error or "no name found"
            job.guess = ""
            written, state, text = (marker.strip() or "?"), "review", f"REVIEW - name not read ({job.reason})"
        else:
            job.guess = job.tag.format(res.name)
            if res.legible:
                written, state, text = job.guess, "ok", "ok"
            else:
                job.reason = "Claude was unsure of the name"
                written, state, text = job.guess + marker, "review", "REVIEW - unsure of the name"
        try:
            job.sink.update_name(job.receipt, written)
            job.written = written
        except ExcelError as e:                           # includes SlotTaken: someone typed in that cell
            job.reason = str(e)
            state, text = "review", "NAME NOT WRITTEN - type it"
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
        if state != "ok":
            self.beep(True)
        self.settle(job)
        self.update_stats()

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
        if self.auto_var.get() and self.breaker:
            self.breaker = ""
            self.clear_alert("breaker")

    def check_jobs(self, now: float):
        """Give up on names that never came back (a hung request must not leave a hole in the sheet).
        The clock runs from the moment a worker starts the request - a name that merely waits its turn in a busy pool
        is not late - with a generous cap on the waiting itself in case every worker is stuck."""
        limit = self.cfg.name_timeout_s
        if limit <= 0:
            return
        for job in self.entries:
            if job.state != "reading" or not job.t_submit:
                continue
            since = now - job.t_started if job.t_started else None
            if (since is not None and since > limit) or (since is None and now - job.t_submit > 4 * limit):
                self.apply_name(job, NameResult(error=f"no answer after {limit:g} s" if since is not None else
                                                "no free name reader"), (now - job.t_submit) * 1000)
                if job.future is not None:
                    job.future.cancel()

    # ------------------------------------------------------- finishing a job
    def settle(self, job: Job):
        """The job reached its end state: write the audit picture + the scanner.log line, free the big arrays."""
        if job.logged:
            return
        job.logged = True
        rd = job.reading
        if rd is not None and job.state != "undone":
            try:
                self.io.submit(self.audit.record, rd, job.picks, job.written or "", job.workbook, job.receipt.sheet,
                               job.receipt.row, job.ticket, job.text if job.manual else job.state)
            except RuntimeError:
                pass
        self.log("ENTRY", ticket=job.ticket, read_ms=job.read_ms, excel_ms=job.excel_ms, name_ms=job.name_ms,
                 status=job.state, name=job.written, reason=job.reason or None, fp_ms=job.fp_ms, feed_ms=job.feed_ms,
                 ui_ms=job.ui_ms, lag_ms=job.lag_ms, cycle_s=job.cycle_s, manual=job.manual or None)
        job.reading = None
        if not job.review and len(self.entries) - job.id >= KEEP_CROPS:
            job.name_crop = None

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
        shown = [x or self.marker for x in picks]
        final_name = self.final_name(name)
        snap, exp = self.tag_snapshot(), self.expected_jornada()
        block, ticket = self.target()
        t0 = time.perf_counter()
        try:
            try:
                receipt = self.sink.write_row_fast(block, picks, final_name, ticket)
            except SlotTaken as e:
                if not messagebox.askyesno("Ticket already used", f"{e}\n\nOverwrite it?"):
                    return
                receipt = self.sink.write_row_fast(block, picks, final_name, ticket, allow_overwrite=True)
        except ExcelError as e:
            self.cancel_auto()
            self.alert(str(e))
            return
        excel_ms = (time.perf_counter() - t0) * 1000
        status = "auto" if auto else ("checked" if p.touched or p.hold or not self.auto_eligible(p.reading) else "ok")
        upd = Update(WAIT_REMOVE, "", timings=p.timings)
        job = self.new_job(receipt, p.reading, shown, snap, exp, upd, excel_ms, t0, manual=True)
        job.written, job.guess, job.state, job.text = final_name, final_name, "ok", status
        self.refresh_row(job)
        self.settle(job)
        self.beep(False)
        self.clear_alert("excel")
        self.last_var.set(f"Last entered: ticket {receipt.ticket}  {final_name}   {' '.join(shown)}")
        self._entered_msg = ""
        self.clear_card()
        self.refresh_blocks()
        self.update_stats()
        self.next_backlog()

    def cancel_auto(self):
        if self.pending:
            self.pending.touched = True
            self.pending.auto_deadline = None
            self.auto_var_txt.set("")

    def discard(self):
        if self.pending is not None:
            self.clear_card()
            self.status_var.set("Slip discarded.")
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
        if usable and job.name_crop is not None:
            self._photos["rev"] = _to_photo(job.name_crop, width=460)
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
            job.sink.undo(job.receipt)
        except ExcelError as e:
            self.alert(f"Undo failed: {e}")
            return
        was_reading = job.state == "reading"
        job.state, job.text, job.written = "undone", "UNDONE", None
        if job.future is not None:
            job.future.cancel()
        self.refresh_row(job)
        if self.card_job is job:
            self.card_job = None
            self.clear_card()
        self.last_var.set(f"Undone: ticket {job.ticket}")
        self.settle(job)
        self.log("UNDO", ticket=job.ticket, while_reading=was_reading or None)
        if self.rev_job is job:
            self.on_tree_select()
        job.reading = job.name_crop = None
        self.refresh_blocks()
        self.update_stats()

    # ------------------------------------------------------------ stats
    def avg_cycle(self) -> float | None:
        t = self.entry_times[-11:]
        gaps = [b - a for a, b in zip(t, t[1:]) if b - a <= BREAK_S]
        return sum(gaps) / len(gaps) if gaps else None

    def update_stats(self):
        live = [j for j in self.entries if j.state != "undone"]
        parts = [f"Entered: {len(live)}"]
        avg = self.avg_cycle()
        parts.append(f"{avg:.1f} s per slip (last 10)" if avg is not None else "- s per slip")
        if self.entries:
            j = self.entries[-1]
            name = "-" if j.manual else (f"{j.name_ms / 1000:.1f} s" if j.name_ms is not None else "reading...")
            parts.append(f"last slip: read {j.read_ms:.0f} ms, Excel {j.excel_ms:.0f} ms, name {name}")
        waiting = sum(j.state == "reading" for j in self.entries)
        if waiting:
            parts.append(f"names being read: {waiting}")
        if self.backlog:
            parts.append(f"slips waiting behind the card: {len(self.backlog)}")
        review = sum(j.review for j in self.entries)
        if review:
            parts.append(f"TO REVIEW: {review}")
        self.stats_var.set("     |     ".join(parts))
        self.stats_lbl.configure(foreground=BAD if review else "#222")

    # --------------------------------------------------------- main loop
    def tick(self):
        try:
            self._tick()
        except Exception as e:                # the loop must survive anything
            self.report_callback_exception(type(e), e, e.__traceback__)
        if not self.closing:
            self._tick_id = self.after(TICK_MS, self.tick)

    def _tick(self):
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
            except Exception as e:            # one bad event must not drop the ones behind it
                self.report_callback_exception(type(e), e, e.__traceback__)
        now = time.perf_counter()
        if now >= self._next_preview:
            self.render_preview(now)
        if now >= self._next_watch:
            self._next_watch = now + 0.25
            self.check_jobs(now)
        if self.pending is not None:
            self.maybe_auto()

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
        s = PREVIEW_W / w
        small = cv2.resize(frame, (PREVIEW_W, max(1, round(h * s))), interpolation=cv2.INTER_AREA)
        quad = None if upd.quad is None else np.asarray(upd.quad, np.float32) * s
        vis = sv.draw_quad(small, quad, (0, 200, 0) if upd.progress >= 1 else (0, 200, 255), 2)
        img = Image.fromarray(cv2.cvtColor(vis, cv2.COLOR_BGR2RGB))
        if self._prev_photo is not None and (self._prev_photo.width(), self._prev_photo.height()) == img.size:
            self._prev_photo.paste(img)
        else:
            self._prev_photo = ImageTk.PhotoImage(img)
            self.preview.configure(image=self._prev_photo)
        if self.pending is None:
            msg = upd.message
            if self._entered_msg and upd.state == WAIT_REMOVE and not upd.duplicate and not upd.error:
                msg = self._entered_msg
            elif upd.state != WAIT_REMOVE:
                self._entered_msg = ""
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

    def on_close(self):
        self.cfg.name_tag = self.tag_var.get().strip()
        self.cfg.name_tag_position = self.tag_pos.get()
        self.cfg.auto_enter = self.auto_var.get() or bool(self.breaker)   # the breaker's "off" is for this session only
        self.cfg.autofocus = self.af_var.get()
        self.cfg.manual_focus = int(self.focus_var.get())
        self.cfg.save()
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
    if app.names_in_flight():
        # a name request that is still hanging would keep this process alive (non-daemon pool threads)
        sys.stdout.flush()
        sys.stderr.flush()
        os._exit(0)
