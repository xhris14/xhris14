"""Desktop window (tkinter)."""
from __future__ import annotations

import queue
import re
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageTk

from . import slip_vision as sv
from .audit import AuditLog
from .camera import CameraError, WebcamSource
from .capture import CaptureEngine, Update
from .config import Config, app_dir
from .excel_sink import ExcelError, ExcelSink, SlotTaken, format_name, open_backend
from .name_reader import NameResult, make_reader

FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
GOOD, WARN, BAD = "#1b7f3b", "#b36b00", "#c62828"
VERDICT_TEXT = {sv.BLANK: "no mark", sv.MULTIPLE: "two marks", sv.WEAK: "faint mark", sv.HEAVY: "scribbled?"}


def _to_photo(img_bgr: np.ndarray, width: int | None = None, height: int | None = None) -> ImageTk.PhotoImage:
    """Shrink (never enlarge) to fit width/height and wrap for tkinter."""
    h, w = img_bgr.shape[:2]
    s = min([1.0] + ([width / w] if width else []) + ([height / h] if height else []))
    if s != 1.0:
        img_bgr = cv2.resize(img_bgr, None, fx=s, fy=s, interpolation=cv2.INTER_AREA)
    return ImageTk.PhotoImage(Image.fromarray(cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)))


@dataclass
class Pending:
    id: int
    reading: sv.SlipReading
    touched: bool = False                 # operator edited something -> no auto-enter
    name_done: bool = False
    name_legible: bool = False
    slip_jornada: str | None = None
    auto_deadline: float | None = None


class App(tk.Tk):
    def __init__(self, cfg: Config, source=None, backend=None):
        super().__init__()
        self.cfg = cfg
        self.title("Quiniela Scanner")
        self.geometry("1200x840")
        self.minsize(1100, 720)
        self.audit = AuditLog(app_dir() / "scans")
        self.events: queue.Queue = queue.Queue()
        self._lock = threading.Lock()
        self._latest: tuple | None = None
        self._stop = threading.Event()
        self.source = None
        self._fixed_source = source
        self._fixed_backend = backend
        self.engine = CaptureEngine(cfg.matches_per_slip, cfg.stable_frames,
                                    sv.Settings(min_ink=cfg.min_ink, multi_ratio=cfg.multi_ratio))
        self.pending: Pending | None = None
        self._scan_id = 0
        self.sink: ExcelSink | None = None
        self.backend = None
        self.receipts: list = []          # (tree iid, Receipt)
        self._row_id = 0
        self.infos: list = []
        self.sel_key: str | None = None
        self.ticket_custom = False        # operator typed a ticket number
        self.reader, self.reader_note = make_reader(cfg.read_names_with_claude, cfg.claude_model)
        self._photos: dict = {}

        self._build()
        self.bind_all("<Return>", lambda e: self.commit())
        self.bind_all("<Escape>", lambda e: self.escape())
        self.bind_all("<F2>", lambda e: self.tag_on.set(not self.tag_on.get()))
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.connect_excel()
        self.start_camera()
        self.after(33, self.tick)

    # ------------------------------------------------------------------ UI
    def _build(self):
        pad = dict(padx=6, pady=3)
        style = ttk.Style(self)
        style.configure("Big.TButton", font=(FONT, 12, "bold"), padding=8)
        style.configure("Card.TLabelframe.Label", font=(FONT, 11, "bold"))

        # ----- toolbar: camera / excel
        bar = ttk.Frame(self)
        bar.pack(fill="x", **pad)
        ttk.Label(bar, text="Camera #").pack(side="left")
        self.cam_var = tk.IntVar(value=self.cfg.camera_index)
        ttk.Spinbox(bar, from_=0, to=9, width=3, textvariable=self.cam_var).pack(side="left", padx=4)
        ttk.Button(bar, text="Restart camera", command=self.start_camera).pack(side="left", padx=2)
        ttk.Button(bar, text="Camera settings", command=self.camera_settings).pack(side="left", padx=2)
        self.af_var = tk.BooleanVar(value=self.cfg.autofocus)
        ttk.Checkbutton(bar, text="Autofocus", variable=self.af_var, command=self.apply_focus).pack(side="left", padx=8)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        self.excel_var = tk.StringVar(value="Excel: not connected")
        self.excel_lbl = ttk.Label(bar, textvariable=self.excel_var, font=(FONT, 10, "bold"))
        self.excel_lbl.pack(side="left")
        ttk.Button(bar, text="Connect / reconnect Excel", command=self.connect_excel).pack(side="left", padx=8)

        # ----- target row
        tgt = ttk.LabelFrame(self, text="Enter into")
        tgt.pack(fill="x", **pad)
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
        self.progress = ttk.Progressbar(left, maximum=1.0, length=520)
        self.progress.pack(anchor="w")
        ttk.Button(left, text="Capture now", command=self.capture_now).pack(anchor="w", pady=6)
        self.last_var = tk.StringVar()
        ttk.Label(left, textvariable=self.last_var, font=(FONT, 11), foreground=GOOD, wraplength=520).pack(anchor="w")
        self.hint_var = tk.StringVar(value=self.reader_note)
        ttk.Label(left, textvariable=self.hint_var, foreground="#555", wraplength=520).pack(anchor="w")

        card = ttk.LabelFrame(main, text="Scanned slip", style="Card.TLabelframe")
        card.pack(side="left", fill="both", expand=True, padx=8)
        self.slip_img = ttk.Label(card)
        self.slip_img.grid(row=0, column=0, rowspan=5, sticky="n", **pad)
        self.name_img = ttk.Label(card)
        self.name_img.grid(row=0, column=1, columnspan=2, sticky="w", **pad)
        ttk.Label(card, text="Name:", font=(FONT, 11, "bold")).grid(row=1, column=1, sticky="e")
        self.name_var = tk.StringVar()
        self.name_var.trace_add("write", lambda *a: self.update_final())
        self.name_entry = ttk.Entry(card, textvariable=self.name_var, font=(FONT, 16, "bold"), width=17)
        self.name_entry.grid(row=1, column=2, sticky="w", **pad)
        self.name_entry.bind("<Key>", self.on_name_key)
        self.name_note = tk.StringVar()
        ttk.Label(card, textvariable=self.name_note, wraplength=300).grid(row=2, column=1, columnspan=2, sticky="w", padx=6)
        self.final_var = tk.StringVar()
        ttk.Label(card, textvariable=self.final_var, font=(FONT, 11, "bold"), foreground="#0b5394",
                  wraplength=300).grid(row=3, column=1, columnspan=2, sticky="w", padx=6, pady=2)
        self.jor_note = tk.StringVar()
        self.jor_lbl = ttk.Label(card, textvariable=self.jor_note, wraplength=300)
        self.jor_lbl.grid(row=4, column=1, columnspan=2, sticky="w", padx=6)

        grid = ttk.Frame(card)
        grid.grid(row=5, column=0, columnspan=3, sticky="w", **pad)
        self.pick_vars, self.pick_lbls = [], []
        for i in range(self.cfg.matches_per_slip):
            col = (i // 5) * 6
            row = i % 5
            lbl = ttk.Label(grid, text=f"{i + 1}", width=3, font=(FONT, 11, "bold"))
            lbl.grid(row=row, column=col, padx=(10, 2))
            v = tk.StringVar(value="")
            self.pick_vars.append(v)
            self.pick_lbls.append(lbl)
            for j, (txt, val) in enumerate((("L", "L"), ("E", "E"), ("V", "V"), ("-", ""))):
                ttk.Radiobutton(grid, text=txt, value=val, variable=v, command=lambda k=i: self.pick_edited(k)
                                ).grid(row=row, column=col + 1 + j, padx=2)
            note = ttk.Label(grid, text="", width=11, foreground=WARN)
            note.grid(row=row, column=col + 5)
            self.pick_lbls[i] = (lbl, note)

        btns = ttk.Frame(card)
        btns.grid(row=6, column=0, columnspan=3, sticky="w", **pad)
        self.write_btn = ttk.Button(btns, text="Write to Excel  (Enter)", style="Big.TButton", command=self.commit)
        self.write_btn.pack(side="left", padx=4)
        ttk.Button(btns, text="Discard  (Esc)", command=self.discard).pack(side="left", padx=4)
        self.auto_var_txt = tk.StringVar()
        ttk.Label(card, textvariable=self.auto_var_txt, font=(FONT, 11, "bold"), foreground=GOOD).grid(
            row=7, column=0, columnspan=3, sticky="w", padx=8)

        # ----- history
        hist = ttk.LabelFrame(self, text="Entered this session")
        hist.pack(fill="x", **pad)
        cols = ("time", "ticket", "name", "picks", "status")
        self.tree = ttk.Treeview(hist, columns=cols, show="headings", height=5)
        for c, w in zip(cols, (70, 70, 260, 260, 220)):
            self.tree.heading(c, text=c.title())
            self.tree.column(c, width=w, anchor="w")
        self.tree.tag_configure("undone", foreground="#999")
        self.tree.pack(side="left", fill="x", expand=True)
        ttk.Button(hist, text="Undo last", command=self.undo).pack(side="left", padx=8)
        self.clear_card()
        self.on_tag_change()

    # ----------------------------------------------------------- camera
    def start_camera(self):
        self._stop.set()
        time.sleep(0.05)
        if self.source is not None:
            try:
                self.source.close()
            except Exception:
                pass
        self.source = None
        try:
            if self._fixed_source is not None:
                self.source = self._fixed_source
            else:
                self.source = WebcamSource(self.cam_var.get(), self.cfg.frame_width, self.cfg.frame_height,
                                           self.cfg.camera_backend, self.af_var.get(), self.cfg.manual_focus)
        except CameraError as e:
            self.status_var.set(str(e))
            return
        self.cfg.camera_index = self.cam_var.get()
        self._stop = threading.Event()
        self.engine.reset()
        threading.Thread(target=self._camera_loop, args=(self._stop, self.source), daemon=True).start()
        self.status_var.set(f"Using {self.source.label}. Place a slip flat under the camera.")

    def apply_focus(self):
        if self.source is not None:
            self.source.set_autofocus(self.af_var.get(), self.cfg.manual_focus)

    def camera_settings(self):
        if self.source is not None:
            self.source.open_driver_settings()

    def _camera_loop(self, stop: threading.Event, source):
        while not stop.is_set():
            frame = source.read()
            if frame is None:
                time.sleep(0.05)
                continue
            self.engine.paused = self.pending is not None
            try:
                upd = self.engine.feed(frame)
            except Exception as e:            # never let a vision bug kill the camera thread
                upd = Update("idle", f"vision error: {e}")
            with self._lock:
                self._latest = (frame, upd)
            if upd.reading is not None or upd.error:
                self.events.put(("scan", upd))

    def capture_now(self):
        with self._lock:
            latest = self._latest
        if latest is None or self.pending is not None:
            return
        self.events.put(("scan", self.engine.capture_now(latest[0])))

    # ------------------------------------------------------------ Excel
    def connect_excel(self):
        try:
            self.backend = self._fixed_backend or open_backend(self.cfg.backend, self.cfg.workbook, self.cfg.workbook_path)
            self.sink = ExcelSink(self.backend, self.cfg.sheets, self.cfg.name_header,
                                  self.cfg.uppercase_names, self.cfg.save_after_write)
        except ExcelError as e:
            self.sink = None
            self.excel_var.set(f"Excel: NOT CONNECTED - {e}")
            self.excel_lbl.configure(foreground=BAD)
            self.block_cb["values"] = []
            return
        self.excel_var.set(f"Excel: {self.backend.label}")
        self.excel_lbl.configure(foreground=GOOD)
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

    # --------------------------------------------------------- scans
    def clear_card(self):
        self.pending = None
        self.name_var.set("")
        self.name_note.set("")
        self.jor_note.set("")
        self.auto_var_txt.set("")
        for v, (lbl, note) in zip(self.pick_vars, self.pick_lbls):
            v.set("")
            note.configure(text="")
            lbl.configure(foreground="black")
        blank = np.full((300, 330, 3), 245, np.uint8)
        cv2.putText(blank, "Waiting for a slip...", (40, 150), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (150, 150, 150), 2)
        self._photos["slip"] = _to_photo(blank)
        self.slip_img.configure(image=self._photos["slip"])
        self.name_img.configure(image="")

    def on_scan(self, upd: Update):
        if upd.reading is None:
            self.status_var.set(upd.message)
            return
        if self.pending is not None:
            return
        rd = upd.reading
        self._scan_id += 1
        self.pending = Pending(self._scan_id, rd)
        self._photos["slip"] = _to_photo(sv.draw_reading(rd), width=330, height=350)
        self.slip_img.configure(image=self._photos["slip"])
        self._photos["name"] = _to_photo(rd.name_crop, width=330)
        self.name_img.configure(image=self._photos["name"])
        self.name_var.set("")
        for i, (v, (lbl, note)) in enumerate(zip(self.pick_vars, self.pick_lbls)):
            v.set(rd.picks[i])
            bad = rd.verdicts[i] != sv.OK
            lbl.configure(foreground=BAD if bad else "black")
            note.configure(text=VERDICT_TEXT.get(rd.verdicts[i], ""))
        if self.reader is not None:
            self.name_note.set("Reading handwriting...")
            threading.Thread(target=self._read_name, args=(self._scan_id, rd), daemon=True).start()
        else:
            self.pending.name_done = True
            self.name_note.set(self.reader_note)
            self.name_entry.focus_set()
        flagged = rd.problem_rows
        self.status_var.set("Slip captured." if not flagged else
                            "Check row(s) " + ", ".join(str(i + 1) for i in flagged) + " (marked red).")
        self.name_entry.focus_set()

    def _read_name(self, scan_id: int, rd: sv.SlipReading):
        try:
            res = self.reader.read(rd.name_crop, rd.title_crop)
        except Exception as e:
            res = NameResult(error=str(e))
        self.events.put(("name", scan_id, res))

    def on_name(self, scan_id: int, res: NameResult):
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

    def active_tag(self) -> str:
        return self.tag_var.get().strip() if self.tag_on.get() else ""

    def final_name(self, name: str) -> str:
        return format_name(name, self.active_tag(), self.tag_pos.get(), self.cfg.tag_separator,
                           self.cfg.uppercase_names)

    def update_final(self):
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

    def maybe_auto(self):
        p = self.pending
        if p is None or p.touched or not self.auto_var.get() or self.sink is None:
            return
        exp = self.expected_jornada()
        jornada_bad = bool(p.slip_jornada and exp and p.slip_jornada != exp)
        ok = (p.reading.clean and p.name_done and self.name_var.get().strip() and not jornada_bad
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
        p = self.pending
        if p is None:
            return
        if self.sink is None:
            messagebox.showerror("Excel", "Excel is not connected. Open the workbook, then press "
                                          "'Connect / reconnect Excel'.")
            return
        name = self.name_var.get().strip()
        if not name:
            self.name_entry.focus_set()
            self.status_var.set("Type the name first.")
            return
        picks = [v.get() for v in self.pick_vars]
        blanks = [str(i + 1) for i, x in enumerate(picks) if not x]
        if blanks and not auto and not messagebox.askyesno(
                "Missing picks", "No pick for match " + ", ".join(blanks) + ".\nEnter the slip anyway "
                "(those cells stay empty)?"):
            return
        try:                                   # pick up anything typed into Excel by hand since the last scan
            self.sink.refresh()
        except ExcelError as e:
            self.cancel_auto()
            messagebox.showerror("Excel", str(e))
            return
        self.infos = self.sink.block_infos()
        slot = self.target_slot()
        if slot is None:
            self.cancel_auto()
            messagebox.showerror("Ticket", "That ticket number is not in the selected sheet, "
                                           "or the block has no free row left.")
            return
        final_name = self.final_name(name)
        try:
            try:
                receipt = self.sink.write(slot, final_name, picks)
            except SlotTaken:
                if not messagebox.askyesno("Ticket already used", f"{slot.describe()} already has data.\n\nOverwrite it?"):
                    return
                receipt = self.sink.write(slot, final_name, picks, allow_overwrite=True)
        except ExcelError as e:
            self.cancel_auto()
            messagebox.showerror("Excel", str(e))
            return
        status = "auto" if auto else ("checked" if p.touched or not p.reading.clean else "ok")
        self.audit.record(p.reading, picks, final_name, getattr(self.backend, "workbook_name", ""),
                          receipt.sheet, receipt.row, receipt.ticket, status)
        self._row_id += 1
        iid = str(self._row_id)
        self.receipts.append((iid, receipt))
        self.tree.insert("", 0, iid=iid, values=(
            time.strftime("%H:%M:%S"), receipt.ticket if receipt.ticket is not None else f"row {receipt.row}",
            final_name, " ".join(x or "-" for x in picks), status))
        self.last_var.set(f"Last entered: ticket {receipt.ticket}  {final_name}   {' '.join(x or '-' for x in picks)}")
        self.clear_card()
        self.refresh_blocks()

    def cancel_auto(self):
        if self.pending:
            self.pending.touched = True
            self.pending.auto_deadline = None
            self.auto_var_txt.set("")

    def discard(self):
        if self.pending is not None:
            self.clear_card()
            self.status_var.set("Slip discarded.")

    def escape(self):
        p = self.pending
        if p is not None and p.auto_deadline is not None:
            self.cancel_auto()
            self.status_var.set("Auto-enter cancelled - edit, then press Enter to write.")
        else:
            self.discard()

    def undo(self):
        if not self.receipts or self.sink is None:
            return
        iid, r = self.receipts[-1]
        if not messagebox.askyesno("Undo", f"Clear ticket {r.ticket} ({r.written[0]}) from Excel?"):
            return
        try:
            self.sink.undo(r)
        except ExcelError as e:
            messagebox.showerror("Excel", str(e))
            return
        self.receipts.pop()
        self.tree.item(iid, tags=("undone",))
        vals = list(self.tree.item(iid, "values"))
        vals[4] = "UNDONE"
        self.tree.item(iid, values=vals)
        self.last_var.set(f"Undone: ticket {r.ticket}")
        self.refresh_blocks()

    # --------------------------------------------------------- main loop
    def tick(self):
        try:
            while True:
                ev = self.events.get_nowait()
                if ev[0] == "scan":
                    self.on_scan(ev[1])
                elif ev[0] == "name":
                    self.on_name(ev[1], ev[2])
        except queue.Empty:
            pass
        with self._lock:
            latest = self._latest
        if latest is not None:
            frame, upd = latest
            vis = sv.draw_quad(frame, upd.quad, (0, 200, 0) if upd.progress >= 1 else (0, 200, 255), 6)
            self._photos["prev"] = _to_photo(vis, width=520, height=390)
            self.preview.configure(image=self._photos["prev"])
            if self.pending is None:
                self.status_var.set(upd.message)
            self.progress["value"] = upd.progress
        self.maybe_auto()
        self.after(33, self.tick)

    def destroy(self):
        self._stop.set()                  # stop the camera thread before the window goes away
        super().destroy()

    def on_close(self):
        self.cfg.name_tag = self.tag_var.get().strip()
        self.cfg.name_tag_position = self.tag_pos.get()
        self.cfg.auto_enter = self.auto_var.get()
        self.cfg.autofocus = self.af_var.get()
        self.cfg.save()
        try:
            if self.source is not None:
                self.source.close()
        except Exception:
            pass
        self.destroy()


def run(cfg: Config, source=None, backend=None):
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App(cfg, source, backend).mainloop()
