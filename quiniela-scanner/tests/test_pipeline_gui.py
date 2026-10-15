"""The stack workflow end to end, through the real window (the GUI tests need a display: xvfb-run on Linux).

Slips arrive from a scripted 30 fps "camera" (a hand lifts a slip, the next one lies beneath), the Excel side is an
openpyxl workbook held in memory that insists on being called from the Tk thread, and the name reader is a fake
whose answers take random seconds and arrive out of order. Smaller tests inject readings straight into the window
to pin down the rules around the pipeline (tag snapshot, review list, circuit breaker, manual card, undo ...).
At the end: config migration, the name reader's failure paths, the frame grabber, sounds and the scanner.log."""
from __future__ import annotations

import base64
import copy
import json
import os
import sys
import threading
import time
from collections import Counter
from dataclasses import asdict

import cv2
import numpy as np
import pytest

tk = pytest.importorskip("tkinter")

from quiniela_scanner import config as config_mod
from quiniela_scanner import gui, name_reader, sound
from quiniela_scanner import slip_vision as sv
from quiniela_scanner.audit import ScanLog
from quiniela_scanner.camera import FrameGrabber
from quiniela_scanner.capture import WAIT_REMOVE, Update
from quiniela_scanner.config import CONFIG_VERSION, Config
from quiniela_scanner.excel_sink import ExcelError, OpenpyxlBackend
from quiniela_scanner.name_reader import ClaudeNameReader, NameResult
from helpers import Stage, draw_slip_named, make_pool_workbook, on_webcam_canvas

NAME_COL = 14                         # block 1 of make_pool_workbook: NOMBRE in N, picks O..W, first free ticket 251 = row 15
FIRST_ROW = 15
NO_DISPLAY = sys.platform != "win32" and not os.environ.get("DISPLAY")

STACK = [  # (picks, name, mark-jitter seed): eight different slips
    (list("LVLLELLVE"), "CLEMENTE O.", 1), (list("EEELLVVLV"), "ANA", 5), (list("VVVLLLEEE"), "BETO", 6),
    (list("LLLEEEVVV"), "CARLOS M.", 7), (list("LEVLEVLEV"), "DOLORES", 8), (list("VELVELVEL"), "PACO 2", 9),
    (list("LLVVEELLV"), "JUAN G.", 10), (list("EVLEVLEVL"), "TONY 5000", 11),
]
NAME_OF = {tuple(p): n for p, n, _ in STACK}
RESULTS: dict = {}                    # numbers worth reporting, filled in by the tests


# ======================================================================================= fakes
class MemBackend(OpenpyxlBackend):
    """OpenpyxlBackend that keeps the workbook in memory (no disk round trip per call), can add a COM-like
    latency to every call, and notes any call that does not come from the Tk thread (COM must not)."""

    def __init__(self, path, latency=0.0):
        self.latency = latency
        self.stamps: list = []                 # (time, kind, row, values) of every write
        self.off_thread: list = []
        self.fail_writes = 0                   # next N write_row calls raise ExcelError
        super().__init__(path)

    def _load_formulas(self):
        if not hasattr(self, "_wb"):
            super()._load_formulas()

    def _load_values(self):
        if not hasattr(self, "_wbv"):
            super()._load_values()

    def _touch(self, kind):
        if threading.current_thread() is not threading.main_thread():
            self.off_thread.append((kind, threading.current_thread().name))
        if self.latency:
            time.sleep(self.latency)

    def sheet_names(self):
        self._touch("sheet_names")
        return super().sheet_names()

    def snapshot(self, sheet, max_rows=600):
        self._touch("snapshot")
        return super().snapshot(sheet, max_rows)

    def read_row(self, sheet, row, col, n):
        self._touch("read_row")
        return super().read_row(sheet, row, col, n)

    def font_sizes(self, sheet, row, col, n):
        self._touch("font_sizes")
        return super().font_sizes(sheet, row, col, n)

    def write_row(self, sheet, row, col, values, font_sizes=None):
        self._touch("write_row")
        self.calls["write_row"] += 1
        if self.fail_writes > 0:
            self.fail_writes -= 1
            raise ExcelError("Excel is busy (is a cell being edited?). Press Esc in Excel and retry.")
        ws, wsv = self._wb[sheet], self._wbv[sheet]
        for i, v in enumerate(values):
            for w in (ws, wsv):
                cell = w.cell(row=row, column=col + i)
                cell.value = v
                if isinstance(v, str):
                    cell.data_type = "s"
        for off, size in (font_sizes or {}).items():
            cell = ws.cell(row=row, column=col + off)
            f = copy.copy(cell.font)
            f.sz = size
            cell.font = f
        self.stamps.append((time.perf_counter(), "write_row", row, list(values)))

    def row(self, row, sheet="ENTRY PG1"):
        """name + 9 picks of one sheet row"""
        return [self._wbv[sheet].cell(row=row, column=NAME_COL + k).value for k in range(10)]

    def type_by_hand(self, row, col, value, sheet="ENTRY PG1"):
        for wb in (self._wb, self._wbv):
            wb[sheet].cell(row=row, column=col).value = value


class Spec:
    def __init__(self, answer, delay=0.0, gate=None):
        self.answer, self.delay, self.gate = answer, delay, gate


class ScriptedReader:
    """Stands in for Claude. The window hands over the very crop object of a reading, so the answer for a slip is
    looked up by that object's identity - the reader sees nothing but the picture, like the real one."""

    def __init__(self):
        self.specs: dict = {}
        self.keep: list = []
        self.started: list = []                # keys, in the order the calls began
        self.finished: list = []               # keys, in the order the answers came back
        self.threads: set = set()
        self._lock = threading.Lock()

    def expect(self, crop, answer, delay=0.0, gate=None, key=None):
        self.specs[id(crop)] = (key if key is not None else len(self.specs), Spec(answer, delay, gate))
        self.keep.append(crop)                 # so the id cannot be reused

    def read(self, name_crop, title_crop=None):
        key, spec = self.specs.get(id(name_crop), (None, Spec(NameResult(error="unexpected picture"))))
        with self._lock:
            self.started.append(key)
            self.threads.add(threading.current_thread().name)
        if spec.gate is not None:
            assert spec.gate.wait(30), "the test never released this name"
        if spec.delay:
            time.sleep(spec.delay)
        with self._lock:
            self.finished.append(key)
        if isinstance(spec.answer, BaseException):
            raise spec.answer
        return spec.answer


class NullSource:
    """A camera that sees an empty table (for tests that inject readings themselves)."""
    label = "empty table"
    demo = False

    def __init__(self):
        self.focus_calls: list = []
        self.frame = np.full((1080, 1920, 3), 235, np.uint8)

    def read(self):
        time.sleep(1 / 30)
        return self.frame

    def set_autofocus(self, on, focus=30):
        self.focus_calls.append((on, focus))

    def open_driver_settings(self):
        pass

    def close(self):
        pass


class Parts:
    """The rendered pieces of a stack scene, stored as JPEG like the camera's MJPEG: 4 empty frames, then per
    slip 8 different 'still' frames (noise, flicker, focus breathing) and a 7-frame lift by a hand that reveals
    the next slip, then 10 empty frames. Film() strings them together at any pace."""

    def __init__(self, specs=STACK, lift=7, seed=2):
        enc = lambda f: cv2.imencode(".jpg", f, [cv2.IMWRITE_JPEG_QUALITY, 92])[1].tobytes()
        st = Stage(seed=seed)
        ids = [st.add(draw_slip_named(p, n, seed=s)) for p, n, s in specs]
        self.start = [enc(f) for f, _ in st.empty(4)]
        self.end = [enc(f) for f, _ in st.empty(10)]
        self.holds, self.lifts = [], []
        for k, i in enumerate(ids):
            under = ids[k + 1] if k + 1 < len(ids) else None
            self.holds.append([enc(f) for f, _ in st.hold(i, 8, under=under)])
            self.lifts.append([enc(f) for f, _ in st.lift(i, under, lift)])


class Film:
    def __init__(self, parts, hold=26, lift_idx=None):
        self.frames = list(parts.start)
        self.hold_start: dict = {}
        for k in range(len(parts.holds)):
            self.hold_start[k] = len(self.frames)
            self.frames += [parts.holds[k][n % 8] for n in range(hold)]
            self.frames += [parts.lifts[k][i] for i in (lift_idx or range(len(parts.lifts[k])))]
        self.frames += parts.end

    def __len__(self):
        return len(self.frames)


class StackSource:
    """A 30 fps camera playing a Film once; after the end it keeps showing the empty table."""
    label = "scripted stack"
    demo = False

    def __init__(self, film, fps=30.0):
        self.film, self.fps, self.n, self.t0 = film, fps, 0, None
        self.served: dict = {}                 # frame number -> time it was served

    def read(self):
        if self.t0 is None:
            self.t0 = time.perf_counter()
        wait = self.t0 + self.n / self.fps - time.perf_counter()
        if wait > 0:
            time.sleep(wait)
        i = min(self.n, len(self.film) - 1)
        self.served[self.n] = time.perf_counter()
        self.n += 1
        return cv2.imdecode(np.frombuffer(self.film.frames[i], np.uint8), cv2.IMREAD_COLOR)

    def set_autofocus(self, *a, **k):
        pass

    def open_driver_settings(self):
        pass

    def close(self):
        pass


# ======================================================================================= fixtures / helpers
@pytest.fixture
def env(tmp_path, monkeypatch):
    if NO_DISPLAY:
        pytest.skip("no display")
    monkeypatch.setattr(gui, "app_dir", lambda: tmp_path)
    monkeypatch.setattr(config_mod, "app_dir", lambda: tmp_path)
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(gui.messagebox, "showerror", lambda *a, **k: (_ for _ in ()).throw(AssertionError(a)))
    beeps = Counter()
    monkeypatch.setattr(sound, "ok", lambda: beeps.update(["ok"]))
    monkeypatch.setattr(sound, "attention", lambda: beeps.update(["attention"]))
    return tmp_path, beeps


@pytest.fixture
def mk(env):
    """mk(source, reader, wbname, latency, **cfg) -> (app, backend). Every app is destroyed afterwards, and no Tk
    callback may have raised."""
    tmp, _ = env
    apps = []

    def make(source=None, reader=None, wbname="JORNADA 12 PC.xlsx", latency=0.0, **cfgkw):
        xlsx = make_pool_workbook(tmp / wbname)
        cfg = Config()
        cfg.read_names_with_claude = False
        for k, v in cfgkw.items():
            setattr(cfg, k, v)
        backend = MemBackend(str(xlsx), latency)
        app = gui.App(cfg, source=source or NullSource(), backend=backend)
        app.reader = reader
        apps.append(app)
        return app, backend
    yield make
    for app in apps:
        errors = list(app.callback_errors)
        try:
            app.destroy()
        except tk.TclError:
            pass
        assert not errors, errors


def pump(app, cond, timeout=15.0, step=0.01):
    end = time.time() + timeout
    while time.time() < end:
        app.update()
        if cond():
            return True
        time.sleep(step)
    return False


def settle(app, seconds=0.3):
    pump(app, lambda: False, seconds)


def read_of(picks, name="ANA", seed=3, angle=4):
    return sv.read_slip(on_webcam_canvas(draw_slip_named(picks, name, seed=seed), scale=0.5, angle=angle))


@pytest.fixture(scope="module")
def readings():
    """Real vision readings of the eight stack slips (made once). Index = the slip's number in STACK."""
    return [read_of(p, n, s) for p, n, s in STACK]


@pytest.fixture(scope="module")
def flagged():
    """Row 2 empty, row 3 a double mark: needs a human."""
    return read_of(["L", "", "LV", "E", "L", "L", "V", "E", "E"], "ANA", 21)


def inject(app, reader, rd, answer=None, delay=0.0, gate=None, key=None):
    """Hand the window a fresh reading as if the engine had just emitted it; `answer` is what the reader will say."""
    rd = copy.copy(rd)
    rd.name_crop = rd.name_crop.copy()
    if reader is not None and answer is not None:
        reader.expect(rd.name_crop, answer, delay, gate, key)
    app.events.put(("scan", Update(WAIT_REMOVE, "Captured.", reading=rd, timings={"read_ms": 180.0, "feed_ms": 200.0}),
                    time.perf_counter()))
    return rd


def good(name, jornada="12"):
    return NameResult(name, jornada, True)


def row_of(backend, n):
    return backend.row(FIRST_ROW + n)


def tags_of(app, job):
    return app.tree.item(job.iid, "tags")


def entered(app, n):
    return pump(app, lambda: len(app.entries) == n)


def named(app, *jobs):
    return pump(app, lambda: all(j.state != "reading" for j in jobs))


# ======================================================================================= the stack
@pytest.fixture(scope="module")
def parts():
    return Parts()


def run_stack(env, mk, film, with_latency=True):
    tmp, beeps = env
    rng = np.random.default_rng(7)
    latency = {tuple(p): float(rng.uniform(0.3, 2.5)) for p, _, _ in STACK}        # random seconds per name, seeded
    reader = ScriptedReader()
    source = StackSource(film)
    app, backend = mk(source, reader, latency=0.004)                               # ~4 ms per COM-like call
    paused: list = []
    feed = app.engine.feed

    def tap(frame):                                      # runs in the camera thread
        paused.append(app.engine.paused)
        upd = feed(frame)
        if upd.reading is not None:
            p = tuple(upd.reading.picks)
            reader.expect(upd.reading.name_crop, good(NAME_OF[p]) if p in NAME_OF else NameResult(error="?"),
                          delay=latency.get(p, 1.0) if with_latency else 0.0, key=p)
        return upd
    app.engine.feed = tap
    t0 = time.perf_counter()
    assert pump(app, lambda: len(app.entries) == 8, 60), f"only {len(app.entries)} slips entered"
    t_entered = time.perf_counter() - t0
    assert pump(app, lambda: all(j.state != "reading" for j in app.entries), 20), "a name never arrived"
    settle(app, 0.4)
    return app, backend, reader, source, paused, beeps, t_entered


def check_stack(app, backend, reader, paused, beeps):
    # exactly 8 rows, in ticket order, right picks, every name in ITS row, nothing else touched
    for k, (picks, name, _) in enumerate(STACK):
        assert row_of(backend, k) == [name] + picks, f"row of slip {k}"
    assert backend.row(FIRST_ROW + 8)[0] is None
    assert [j.ticket for j in app.entries] == list(range(251, 259))
    assert [j.state for j in app.entries] == ["ok"] * 8
    # never paused, no full re-read per slip, COM and Tk only on the Tk thread
    assert not any(paused) and len(paused) > 100
    assert backend.calls["snapshot"] == 2 and app.sink.refresh_count == 1           # the two ENTRY sheets, read once
    assert backend.off_thread == [] and reader.threads and "MainThread" not in reader.threads
    assert backend.calls["write_row"] == 16                                          # picks, then the name, per slip
    assert beeps["ok"] == 8 and beeps["attention"] == 0
    assert app.alert_var.get() == "" and not app.breaker and not app.callback_errors


def test_stack_of_eight_is_entered_hands_free_and_every_name_lands_in_its_own_row(env, mk, parts):
    film = Film(parts, hold=26)                                        # 33 frames = 1.1 s per slip: a relaxed human
    app, backend, reader, source, paused, beeps, t_entered = run_stack(env, mk, film)
    check_stack(app, backend, reader, paused, beeps)
    assert reader.finished != [tuple(p) for p, _, _ in STACK], "the random latencies never reordered the answers"

    picks_at = [t for t, kind, row, vals in backend.stamps if vals[0] is None]
    assert len(picks_at) == 8
    lat = [picks_at[k] - source.served[film.hold_start[k]] for k in range(8)]
    cadence = (picks_at[-1] - picks_at[0]) / 7
    script_slip_s = (len(film) - 14) / 8 / source.fps
    RESULTS["stack"] = dict(script_s_per_slip=round(script_slip_s, 2), app_s_per_slip=round(cadence, 2),
                            all_8_entered_after_s=round(t_entered, 1), still_to_picks_in_excel_s=[round(x, 2) for x in lat],
                            tk_ms_per_slip=[round(j.ui_ms) for j in app.entries],
                            excel_ms=[round(j.excel_ms, 1) for j in app.entries], read_ms=[round(j.read_ms) for j in app.entries])
    print("\nSTACK", json.dumps(RESULTS["stack"]))
    assert max(lat) < 1.5 and float(np.median(lat)) < 1.0, lat          # picks are in Excel well under a second
    assert cadence < script_slip_s + 0.1                                # the app keeps the camera's pace, it does not set it
    assert max(j.ui_ms for j in app.entries) < 400
    assert "Entered: 8" in app.stats_var.get() and "s per slip" in app.stats_var.get()
    app.flush_io()
    lines = (env[0] / "scanner.log").read_text().splitlines()
    assert sum("  ENTRY  " in ln and "status=ok" in ln and "name_ms=" in ln and "excel_ms=" in ln for ln in lines) == 8


def test_a_much_faster_hand_is_still_one_entry_per_slip(env, mk, parts):
    """19 frames (0.63 s) per slip: three times faster than the 2 s target, with a real 7-frame lift. (A lift squeezed
    into 2 frames would be a slip leaving the picture in 66 ms - and, with the next slip right under it, no longer
    shows as a change: that is the engine's documented blind spot, not something the window can fix.)"""
    film = Film(parts, hold=12)
    app, backend, reader, source, paused, beeps, t_entered = run_stack(env, mk, film, with_latency=False)
    check_stack(app, backend, reader, paused, beeps)
    picks_at = [t for t, kind, row, vals in backend.stamps if vals[0] is None]
    RESULTS["fast_stack"] = dict(app_s_per_slip=round((picks_at[-1] - picks_at[0]) / 7, 2),
                                 script_s_per_slip=round((len(film) - 14) / 8 / source.fps, 2))
    print("\nFAST STACK", json.dumps(RESULTS["fast_stack"]))


REAL_SLIP = os.environ.get("QUINIELA_REAL_SLIP", "")           # a real slip photo, kept out of the repo


@pytest.mark.skipif(not REAL_SLIP, reason="set QUINIELA_REAL_SLIP to a real slip photo")
def test_the_real_photo_goes_through_the_whole_pipeline(env, mk):
    from quiniela_scanner.camera import ImageSource
    reader = ScriptedReader()
    app, backend = mk(ImageSource(REAL_SLIP, demo=True, hold=60, gap=40), reader)
    feed = app.engine.feed

    def tap(frame):
        upd = feed(frame)
        if upd.reading is not None:
            reader.expect(upd.reading.name_crop, good("CLEMENTE O."), delay=0.3)
        return upd
    app.engine.feed = tap
    assert pump(app, lambda: len(app.entries) >= 1, 30)
    assert named(app, app.entries[0])
    assert row_of(backend, 0) == ["CLEMENTE O."] + list("LVLLLLLLE")


# ======================================================================================= rules around the pipeline
def test_picks_go_in_at_once_and_the_name_follows(env, mk, readings):
    tmp, beeps = env
    reader, gate = ScriptedReader(), threading.Event()
    app, backend = mk(reader=reader)
    inject(app, reader, readings[0], good("ana"), gate=gate)
    assert entered(app, 1)
    job = app.entries[0]
    assert row_of(backend, 0) == [None] + STACK[0][0] and job.state == "reading"      # picks in, name cell still empty
    assert app.pending is None and app.card_job is job and "ENTERED - ticket 251" in app.banner_var.get()
    assert app.tree.item(job.iid, "values")[2] == "(reading...)" and "reading" in tags_of(app, job)
    assert beeps["ok"] == 1
    assert app.ticket_var.get() == "252"                                              # the next slip goes to the next ticket
    gate.set()
    assert named(app, job)
    assert row_of(backend, 0)[0] == "ANA" and job.state == "ok" and app.name_var.get() == "ANA"
    assert app.tree.item(job.iid, "values")[2] == "ANA" and "TO REVIEW" not in app.stats_var.get()


def test_tag_is_snapshotted_when_the_slip_is_scanned_not_when_its_name_arrives(env, mk, readings):
    reader, g0, g1 = ScriptedReader(), threading.Event(), threading.Event()
    app, backend = mk(reader=reader)
    app.tag_var.set("RRR")
    app.tag_on.set(True)                                              # tag ON when slip 0 is scanned
    inject(app, reader, readings[0], good("ANA"), gate=g0)
    assert entered(app, 1)
    app.tag_on.set(False)                                             # F2 before slip 0's name is back ...
    inject(app, reader, readings[1], good("BETO"), gate=g1)           # ... so slip 1 is scanned with the tag OFF
    assert entered(app, 2)
    app.tag_on.set(True)
    app.tag_pos.set("before")                                         # ON again and moved, still before either name is back
    g1.set()                                                          # slip 1's name comes back first
    assert named(app, app.entries[1])
    g0.set()
    assert named(app, app.entries[0])
    assert row_of(backend, 0)[0] == "ANA RRR" and row_of(backend, 1)[0] == "BETO"
    assert not app.entries[0].review and not app.entries[1].review


def test_names_that_come_back_in_reverse_order_still_land_in_their_own_rows(env, mk, readings):
    reader = ScriptedReader()
    gates = [threading.Event() for _ in range(4)]
    app, backend = mk(reader=reader, name_workers=4)                  # (with 3 workers the 4th name would queue behind them)
    names = ["ANA", "BETO", "CARLOS", "DOLORES"]
    for k in range(4):
        inject(app, reader, readings[k], good(names[k]), gate=gates[k])
    assert entered(app, 4)
    assert [row_of(backend, k)[1:] for k in range(4)] == [STACK[k][0] for k in range(4)]
    for k in (3, 1, 2, 0):
        gates[k].set()
        assert named(app, app.entries[k])
    assert [row_of(backend, k)[0] for k in range(4)] == names


def test_unsure_and_failed_names_get_the_marker_and_the_review_list_fixes_them(env, mk, readings):
    tmp, beeps = env
    reader = ScriptedReader()
    app, backend = mk(reader=reader)
    answers = [NameResult("MAYBE", "12", False), NameResult(error="no internet"), good("CARLOS"), RuntimeError("boom")]
    for rd, a in zip(readings, answers):
        inject(app, reader, rd, a, delay=0.05)
    assert entered(app, 4)
    assert named(app, *app.entries)
    assert [row_of(backend, k)[0] for k in range(4)] == ["MAYBE ?", "?", "CARLOS", "?"]
    j0, j1, j2, j3 = app.entries
    assert [j.state for j in app.entries] == ["review", "review", "ok", "review"]
    assert "unsure" in j0.text and "no internet" in j1.text and "RuntimeError" in j3.text
    assert [("review" in tags_of(app, j)) for j in app.entries] == [True, True, False, True]
    assert beeps["attention"] >= 3 and "TO REVIEW: 3" in app.stats_var.get()

    app.select_next_review()                                          # the oldest red row
    assert app.rev_job is j0 and app.rev_name.get() == "MAYBE"        # prefilled with the guess, marker stripped
    assert str(app.rev_img.cget("image")) != "" and app.rev_save.instate(["!disabled"])
    app.rev_name.set("maria g.")
    app.save_review_name()
    assert row_of(backend, 0)[0] == "MARIA G." and j0.state == "ok" and "review" not in tags_of(app, j0)

    app.tree.selection_set(j1.iid)
    settle(app, 0.05)
    assert app.rev_job is j1 and app.rev_name.get() == ""             # nothing to guess from
    app.rev_name.set("jose")
    app.rev_entry.focus_force()
    settle(app, 0.05)
    app.rev_entry.event_generate("<Return>")                          # Enter inside the entry saves too
    settle(app, 0.05)
    assert row_of(backend, 1)[0] == "JOSE" and j1.state == "ok"
    assert "TO REVIEW: 1" in app.stats_var.get()
    app.flush_io()
    log = (tmp / "scanner.log").read_text()
    assert "status=review" in log and "NAME_FIXED" in log


def test_a_name_cell_typed_by_hand_is_never_overwritten(env, mk, readings):
    reader, gate = ScriptedReader(), threading.Event()
    app, backend = mk(reader=reader)
    inject(app, reader, readings[0], good("ANA"), gate=gate)
    assert entered(app, 1)
    job = app.entries[0]
    backend.type_by_hand(FIRST_ROW, NAME_COL, "TYPED BY HAND")
    gate.set()
    assert named(app, job)
    assert job.state == "review" and job.text == "NAME NOT WRITTEN - type it"
    assert row_of(backend, 0)[0] == "TYPED BY HAND" and "review" in tags_of(app, job)
    app.select_next_review()
    assert app.rev_name.get() == "ANA"                                # the guess is offered
    app.save_review_name()
    assert row_of(backend, 0)[0] == "TYPED BY HAND" and job.state == "review"      # still refused ...
    assert "ticket 251" in app.alert_var.get()                                      # ... loudly
    backend.type_by_hand(FIRST_ROW, NAME_COL, None)                   # the operator clears the cell
    app.save_review_name()
    assert row_of(backend, 0)[0] == "ANA" and job.state == "ok"


def test_wrong_jornada_marks_the_row_trips_the_breaker_and_can_be_undone(env, mk, readings):
    tmp, beeps = env
    reader = ScriptedReader()
    app, backend = mk(reader=reader, wbname="JORNADA 10 PC.xlsx")
    inject(app, reader, readings[0], good("ANA", jornada="12"))       # the slip says 12, the workbook is 10
    assert entered(app, 1)
    job = app.entries[0]
    assert named(app, job)
    assert job.state == "wrong" and "WRONG JORNADA" in job.text and "wrong" in tags_of(app, job)
    assert row_of(backend, 0) == ["ANA"] + STACK[0][0]                # it did go in (the operator decides)
    assert app.auto_var.get() is False and app.breaker and "WRONG JORNADA" in app.alert_var.get()
    assert beeps["attention"] >= 1 and app.alert_frame.winfo_manager() == "pack"

    inject(app, reader, readings[1], good("BETO", jornada="10"))      # the next slip is clean ... but auto-enter is off
    assert pump(app, lambda: app.pending is not None)
    assert len(app.entries) == 1 and "CHECK THIS SLIP" in app.banner_var.get()
    assert pump(app, lambda: app.engine.paused)                       # the classic card pauses the engine
    assert pump(app, lambda: app.name_var.get() == "BETO")            # (its name is still read for the card)
    app.commit()
    assert row_of(backend, 1)[0] == "BETO" and app.pending is None and len(app.entries) == 2

    app.tree.selection_set(job.iid)                                   # the wrong one: take it out again
    settle(app, 0.05)
    assert app.rev_job is job and app.rev_undo.instate(["!disabled"])
    app.undo_selected()
    assert row_of(backend, 0) == [None] * 10 and job.state == "undone"

    app.auto_var.set(True)                                            # operator takes responsibility again
    assert not app.breaker and app.alert_var.get() == ""
    inject(app, reader, readings[2], good("CARLOS", jornada="10"))
    assert entered(app, 3)                                            # fast mode again
    assert named(app, app.entries[2]) and app.entries[2].state == "ok"
    assert row_of(backend, 0)[0] == "CARLOS"                          # the freed ticket is reused


def test_breaker_is_not_saved_as_auto_enter_off(env, mk, readings):
    tmp, _ = env
    reader = ScriptedReader()
    app, backend = mk(reader=reader, wbname="JORNADA 10 PC.xlsx")
    inject(app, reader, readings[0], good("ANA", jornada="12"))
    assert entered(app, 1) and named(app, app.entries[0])
    assert app.auto_var.get() is False
    app.tag_var.set("RRR")
    app.on_close()
    saved = json.loads((tmp / "config.json").read_text())
    assert saved["config_version"] == CONFIG_VERSION and saved["name_tag"] == "RRR" and "auto_enter" not in saved


def test_flagged_slip_uses_the_classic_card_and_never_refreshes_excel(env, mk, flagged):
    tmp, beeps = env
    reader = ScriptedReader()
    app, backend = mk(reader=reader)
    inject(app, reader, flagged, good("ANA"))
    assert pump(app, lambda: app.pending is not None)
    assert len(app.entries) == 0 and "CHECK THIS SLIP" in app.banner_var.get() and beeps["attention"] == 1
    assert pump(app, lambda: app.engine.paused) and app.write_btn.instate(["!disabled"])
    assert pump(app, lambda: app.name_var.get() == "ANA")
    assert app.pick_vars[1].get() == "" and app.pick_vars[2].get() == ""
    app.pick_vars[1].set("E")
    app.pick_edited(1)
    app.pick_vars[2].set("V")
    app.pick_edited(2)
    app.commit()
    assert row_of(backend, 0) == ["ANA", "L", "E", "V", "E", "L", "L", "V", "E", "E"]
    assert app.pending is None and app.entries[0].state == "ok" and app.entries[0].text == "checked"
    assert backend.calls["snapshot"] == 2 and app.sink.refresh_count == 1       # no full re-read for the commit
    assert pump(app, lambda: not app.engine.paused)


def test_a_slip_that_arrives_while_a_card_is_open_is_not_lost(env, mk, readings, flagged):
    reader = ScriptedReader()
    app, backend = mk(reader=reader)
    inject(app, reader, flagged, good("ANA"))
    assert pump(app, lambda: app.pending is not None)
    inject(app, reader, readings[1], good("BETO"))                    # e.g. queued up while Excel was busy
    assert pump(app, lambda: len(app.backlog) == 1) and "waiting behind the card: 1" in app.stats_var.get()
    app.pick_vars[1].set("E")
    app.pick_edited(1)
    app.pick_vars[2].set("V")
    app.pick_edited(2)
    app.name_var.set("ana")
    app.commit()                                                      # closing the card takes the waiting slip
    assert len(app.entries) == 2 and not app.backlog
    assert named(app, app.entries[1])
    assert row_of(backend, 0)[0] == "ANA" and row_of(backend, 1) == ["BETO"] + STACK[1][0]


def test_no_name_reader_means_the_card_and_typing_the_name(env, mk, readings):
    app, backend = mk(reader=None)                                    # no API key
    inject(app, None, readings[0])
    assert pump(app, lambda: app.pending is not None)
    assert app.name_note.get() == app.reader_note and len(app.entries) == 0
    app.name_var.set("clemente o.")
    app.commit()
    assert row_of(backend, 0) == ["CLEMENTE O."] + STACK[0][0] and app.pending is None


def test_auto_enter_off_means_the_card_even_for_a_clean_slip(env, mk, readings):
    reader = ScriptedReader()
    app, backend = mk(reader=reader, auto_enter=False)
    inject(app, reader, readings[0], good("ANA"))
    assert pump(app, lambda: app.pending is not None) and len(app.entries) == 0
    assert pump(app, lambda: app.name_var.get() == "ANA")
    app.commit()
    assert row_of(backend, 0)[0] == "ANA"


def test_a_countdown_longer_than_zero_keeps_the_old_hands_on_mode(env, mk, readings):
    reader = ScriptedReader()
    app, backend = mk(reader=reader, auto_enter_delay_s=0.4)
    inject(app, reader, readings[0], good("ANA"), delay=0.1)
    assert pump(app, lambda: app.pending is not None) and len(app.entries) == 0
    assert pump(app, lambda: "Entering ticket 251" in app.auto_var_txt.get(), 5)
    assert pump(app, lambda: len(app.entries) == 1, 5)
    assert row_of(backend, 0)[0] == "ANA" and app.entries[0].text == "auto"


def test_excel_trouble_in_fast_mode_is_loud_and_falls_back_to_the_card(env, mk, readings):
    tmp, beeps = env
    reader = ScriptedReader()
    app, backend = mk(reader=reader)
    backend.fail_writes = 1
    inject(app, reader, readings[0], good("ANA"))
    assert pump(app, lambda: app.pending is not None)
    assert "NOT ENTERED" in app.alert_var.get() and "busy" in app.alert_var.get() and beeps["attention"] >= 1
    assert "NOT ENTERED" in app.banner_var.get() and len(app.entries) == 0 and row_of(backend, 0) == [None] * 10
    assert pump(app, lambda: app.engine.paused)
    assert pump(app, lambda: app.name_var.get() == "ANA")
    app.commit()                                                      # Excel is fine again: the operator enters it
    assert row_of(backend, 0) == ["ANA"] + STACK[0][0] and app.alert_var.get() == ""


def test_undo_of_an_entry_whose_name_is_still_being_read(env, mk, readings):
    tmp, _ = env
    reader, gate = ScriptedReader(), threading.Event()
    app, backend = mk(reader=reader)
    inject(app, reader, readings[0], good("ANA"), gate=gate)
    assert entered(app, 1)
    job = app.entries[0]
    assert row_of(backend, 0)[1:] == STACK[0][0] and job.state == "reading"
    app.undo()                                                        # "Undo last" while Claude is still reading
    assert row_of(backend, 0) == [None] * 10 and job.state == "undone"
    assert app.tree.item(job.iid, "values")[4] == "UNDONE" and app.ticket_var.get() == "251"
    gate.set()                                                        # the late answer arrives ...
    assert pump(app, lambda: len(reader.finished) == 1)
    settle(app, 0.3)
    assert row_of(backend, 0) == [None] * 10 and job.state == "undone"   # ... and is ignored
    assert backend.calls["write_row"] == 2                            # picks, undo: no stray name write
    inject(app, reader, readings[1], good("BETO"))                    # the freed ticket is reused
    assert entered(app, 2) and named(app, app.entries[1])
    assert row_of(backend, 0)[0] == "BETO"
    app.flush_io()
    log = (tmp / "scanner.log").read_text()
    assert "status=undone" in log and "UNDO" in log and "while_reading=True" in log


def test_a_name_that_never_comes_back_is_given_up_on(env, mk, readings):
    reader, gate = ScriptedReader(), threading.Event()
    app, backend = mk(reader=reader, name_timeout_s=0.3)
    inject(app, reader, readings[0], good("ANA"), gate=gate)
    assert entered(app, 1)
    job = app.entries[0]
    assert named(app, job)
    assert job.state == "review" and "no answer" in job.text and row_of(backend, 0)[0] == "?"
    gate.set()
    settle(app, 0.4)
    assert row_of(backend, 0)[0] == "?" and job.state == "review"     # the late answer changes nothing


def test_a_name_waiting_its_turn_in_a_busy_pool_is_not_late(env, mk, readings):
    reader = ScriptedReader()
    app, backend = mk(reader=reader, name_workers=1, name_timeout_s=0.6)
    inject(app, reader, readings[0], good("ANA"), delay=0.35)
    inject(app, reader, readings[1], good("BETO"), delay=0.35)        # starts only when ANA is done: 0.7 s after it was entered
    assert entered(app, 2) and named(app, *app.entries)
    assert [j.state for j in app.entries] == ["ok", "ok"]
    assert [row_of(backend, k)[0] for k in range(2)] == ["ANA", "BETO"]


def test_with_every_reader_stuck_the_queued_names_are_finally_given_up_on(env, mk, readings):
    reader, gate = ScriptedReader(), threading.Event()
    app, backend = mk(reader=reader, name_workers=1, name_timeout_s=0.3)
    inject(app, reader, readings[0], good("ANA"), gate=gate)          # the only worker hangs on this one
    inject(app, reader, readings[1], good("BETO"))                    # ... so this one never even starts
    assert entered(app, 2) and named(app, *app.entries)
    assert "no answer" in app.entries[0].text and "no free name reader" in app.entries[1].reason
    assert [row_of(backend, k)[0] for k in range(2)] == ["?", "?"]
    gate.set()


def test_sounds_can_be_switched_off(env, mk, readings):
    tmp, beeps = env
    reader = ScriptedReader()
    app, backend = mk(reader=reader, sounds=False)
    inject(app, reader, readings[0], NameResult(error="no internet"))
    assert entered(app, 1) and named(app, app.entries[0])
    assert sum(beeps.values()) == 0


def test_capture_now_goes_through_the_same_pipeline(env, mk, readings):
    reader = ScriptedReader()
    app, backend = mk(reader=reader)
    rd = copy.copy(readings[0])
    rd.name_crop = rd.name_crop.copy()
    reader.expect(rd.name_crop, good("ANA"))
    app.engine.capture_now = lambda frame: Update(WAIT_REMOVE, "Captured.", reading=rd, timings={"read_ms": 150.0})
    assert pump(app, lambda: app._latest is not None)
    app.capture_now()
    assert entered(app, 1) and named(app, app.entries[0])
    assert row_of(backend, 0) == ["ANA"] + STACK[0][0]


def test_unreadable_slip_is_reported_and_beeps(env, mk):
    tmp, beeps = env
    app, backend = mk()
    app.on_scan(Update(WAIT_REMOVE, "Slip seen but not recognised (x).", error="x"), time.perf_counter())
    assert beeps["attention"] == 1 and "not recognised" in app.status_var.get() and not app.entries
    app.flush_io()
    assert "UNREADABLE" in (tmp / "scanner.log").read_text()


# ======================================================================================= window behaviour
def test_focus_slider_is_live_only_without_autofocus_and_remembered(env, mk):
    src = NullSource()
    app, _ = mk(src)
    assert app.af_var.get() is True and str(app.focus_scale.cget("state")) == "disabled"
    assert (app.focus_scale.cget("from"), app.focus_scale.cget("to"), float(app.focus_scale.cget("resolution"))) == (0, 255, 5)
    app.af_var.set(False)
    app.apply_focus()
    assert str(app.focus_scale.cget("state")) == "normal" and src.focus_calls[-1] == (False, 30)
    app.focus_scale.set(125)                                          # dragging the slider ...
    assert pump(app, lambda: src.focus_calls[-1] == (False, 125), 3)  # ... reaches the camera, debounced
    assert app.cfg.manual_focus == 125
    app.focus_var.set(133)                                            # not a multiple of 5: snapped
    app.apply_manual_focus()
    assert src.focus_calls[-1] == (False, 135)
    app.af_var.set(True)
    app.apply_focus()
    assert str(app.focus_scale.cget("state")) == "disabled" and src.focus_calls[-1][0] is True
    app.apply_manual_focus()
    assert src.focus_calls[-1][0] is True                             # autofocus on: the slider never touches the camera


def test_slider_value_is_saved_with_the_config(env, mk):
    tmp, _ = env
    app, _ = mk(autofocus=False, manual_focus=60)
    assert app.focus_var.get() == 60 and str(app.focus_scale.cget("state")) == "normal"
    app.focus_var.set(90)
    app.on_close()
    saved = json.loads((tmp / "config.json").read_text())
    assert saved["manual_focus"] == 90 and saved["autofocus"] is False and saved["config_version"] == CONFIG_VERSION


def test_preview_is_shrunk_before_drawing_and_skips_frames_it_already_showed(env, mk, monkeypatch):
    seen = []
    real = sv.draw_quad
    monkeypatch.setattr(gui.sv, "draw_quad", lambda frame, quad, *a, **k: (seen.append(frame.shape), real(frame, quad, *a, **k))[1])
    app, _ = mk()
    assert pump(app, lambda: len(seen) >= 3, 5)
    assert all(s[1] == gui.PREVIEW_W and s[0] < 400 for s in seen), seen[:3]     # never a 1080p frame
    frame = np.full((1080, 1920, 3), 200, np.uint8)
    upd = Update("tracking", "Hold still...", np.array([[400, 300], [1500, 300], [1500, 900], [400, 900]], np.float32), 0.5)
    app._stop_camera()                                                # nothing else may touch the preview from here on
    n0 = len(seen)
    with app._lock:
        app._latest = (10_000, frame, upd)
    app._next_preview = 0.0
    app.render_preview(time.perf_counter())
    app.render_preview(time.perf_counter() + 1)
    app.render_preview(time.perf_counter() + 2)                       # same frame id: nothing more to draw
    assert len(seen) - n0 == 1
    t0 = time.perf_counter()
    for i in range(60):
        with app._lock:
            app._latest = (20_000 + i, frame, upd)
        app.render_preview(time.perf_counter() + 10 * i)
    ms = (time.perf_counter() - t0) / 60 * 1000
    RESULTS["preview_ms_per_frame"] = round(ms, 2)
    print("\nPREVIEW ms/frame", round(ms, 2))
    assert ms < 25


def test_demo_loop_turns_dedupe_off(env, mk):
    class Demo(NullSource):
        demo = True
    app, _ = mk(Demo())
    assert app.engine.dedupe is False


def test_a_real_camera_leaves_dedupe_on(env, mk):
    app, _ = mk(NullSource())
    assert app.engine.dedupe is True


def test_settings_reset_by_the_migration_are_announced(env, mk):
    cfg_keys = ["stable_frames", "auto_enter"]
    xlsx = make_pool_workbook(env[0] / "JORNADA 12 PC.xlsx")
    cfg = Config()
    cfg.migrated = cfg_keys
    cfg.read_names_with_claude = False
    app = gui.App(cfg, source=NullSource(), backend=MemBackend(str(xlsx)))
    try:
        assert "reset" in app.hint_var.get() and "stable_frames" in app.hint_var.get()
    finally:
        app.destroy()


# ======================================================================================= config
OLD_ONLY = {"config_version", "name_workers", "unsure_marker", "sounds", "claude_effort", "name_timeout_s"}


def old_style_config(**over):
    """What the previous build wrote: every setting, none of the new ones, the old defaults for the three that matter."""
    data = {k: v for k, v in asdict(Config()).items() if k not in OLD_ONLY}
    data.update(stable_frames=7, auto_enter_delay_s=3.0, auto_enter=False)
    data.update(over)
    return data


def test_old_config_json_is_migrated_to_the_new_defaults(tmp_path, monkeypatch):
    monkeypatch.setattr(config_mod, "app_dir", lambda: tmp_path)        # the app folder: Config.load() / save() use it
    path = Config.path()
    assert path == tmp_path / "config.json"
    path.write_text(json.dumps(old_style_config(camera_index=2, name_tag="RRR", autofocus=False, manual_focus=45,
                                                claude_model="my-model", blank_marker="X1"), indent=2))
    cfg = Config.load()
    assert (cfg.stable_frames, cfg.auto_enter_delay_s, cfg.auto_enter) == (4, 0.0, True)       # the new defaults apply
    assert cfg.migrated == ["stable_frames", "auto_enter_delay_s", "auto_enter"]
    assert (cfg.camera_index, cfg.name_tag, cfg.autofocus, cfg.manual_focus) == (2, "RRR", False, 45)   # the rest is kept
    assert (cfg.claude_model, cfg.blank_marker, cfg.config_version) == ("my-model", "X1", CONFIG_VERSION)
    assert (cfg.name_workers, cfg.unsure_marker, cfg.sounds) == (6, " ?", True)
    cfg.save()
    saved = json.loads(path.read_text())
    assert saved == {"config_version": CONFIG_VERSION, "camera_index": 2, "autofocus": False, "manual_focus": 45,
                     "name_tag": "RRR", "claude_model": "my-model", "blank_marker": "X1"}
    again = Config.load()
    assert again.migrated == [] and asdict(again) == asdict(cfg)


def test_missing_or_older_version_both_drop_the_three_settings(tmp_path):
    path = tmp_path / "config.json"
    for version in ({}, {"config_version": 1}, {"config_version": "2"}, {"config_version": None}):
        path.write_text(json.dumps({**old_style_config(), **version}))
        cfg = Config.load(path)
        assert (cfg.stable_frames, cfg.auto_enter) == (4, True), version


def test_a_current_config_keeps_what_the_user_chose(tmp_path):
    path = tmp_path / "config.json"
    path.write_text(json.dumps({"config_version": CONFIG_VERSION, "auto_enter": False, "stable_frames": 6,
                                "auto_enter_delay_s": 2.0, "name_workers": 5, "unsure_marker": " ??", "sounds": False}))
    cfg = Config.load(path)
    assert (cfg.auto_enter, cfg.stable_frames, cfg.auto_enter_delay_s) == (False, 6, 2.0)
    assert (cfg.name_workers, cfg.unsure_marker, cfg.sounds, cfg.migrated) == (5, " ??", False, [])
    path.write_text(json.dumps({"config_version": CONFIG_VERSION + 1, "stable_frames": 9}))
    assert Config.load(path).stable_frames == 9                       # written by a newer build: left alone


def test_defaults_and_save_write_only_what_differs(tmp_path):
    cfg = Config()
    assert (cfg.stable_frames, cfg.auto_enter_delay_s, cfg.auto_enter, cfg.name_workers, cfg.unsure_marker,
            cfg.sounds, cfg.config_version, cfg.claude_model) == (4, 0.0, True, 3, " ?", True, 2, "claude-opus-5-5")
    path = tmp_path / "config.json"
    cfg.save(path)
    assert json.loads(path.read_text()) == {"config_version": CONFIG_VERSION}
    cfg.stable_frames = 5
    cfg.save(path)
    assert json.loads(path.read_text()) == {"config_version": CONFIG_VERSION, "stable_frames": 5}
    assert not list(tmp_path.glob("*.tmp"))


def test_command_line_overrides_are_not_saved(tmp_path):
    path = tmp_path / "config.json"
    path.write_text(json.dumps({"config_version": CONFIG_VERSION, "backend": "com", "name_tag": "RRR"}))
    cfg = Config.load(path)
    cfg.override(backend="file", workbook_path="/tmp/demo.xlsx")        # run.bat --demo slip.jpg --xlsx /tmp/demo.xlsx
    assert (cfg.backend, cfg.workbook_path) == ("file", "/tmp/demo.xlsx")
    cfg.name_tag = "E"                                                  # a real change still is saved
    cfg.save(path)
    assert json.loads(path.read_text()) == {"config_version": CONFIG_VERSION, "backend": "com", "name_tag": "E"}


def test_bad_config_files_fall_back_to_defaults(tmp_path):
    path = tmp_path / "config.json"
    for text in ("not json at all", "[1, 2]", ""):
        path.write_text(text)
        assert asdict(Config.load(path)) == asdict(Config())
    path.write_text(json.dumps({"config_version": CONFIG_VERSION, "stable_frames": "lots", "sounds": "yes",
                                "auto_enter_delay_s": 1, "camera_index": True, "bogus": 1}))
    cfg = Config.load(path)
    assert (cfg.stable_frames, cfg.sounds, cfg.camera_index) == (4, True, 0)       # wrong types are ignored
    assert cfg.auto_enter_delay_s == 1.0 and isinstance(cfg.auto_enter_delay_s, float)


# ======================================================================================= name reader
class FakeAnthropic:
    """Same exception names and hierarchy as the SDK, so the reader's except-order is exercised."""
    class APIError(Exception):
        pass

    class APIConnectionError(APIError):
        pass

    class APITimeoutError(APIConnectionError):
        pass

    class APIStatusError(APIError):
        status_code = 500

    class AuthenticationError(APIStatusError):
        status_code = 401

    class RateLimitError(APIStatusError):
        status_code = 429


def fake_reader(create, effort="low"):
    rd = ClaudeNameReader.__new__(ClaudeNameReader)
    rd._anthropic, rd.model, rd.effort = FakeAnthropic, "m", effort
    rd.client = type("C", (), {"messages": type("M", (), {"create": staticmethod(create)})()})()
    return rd


class _Resp:
    def __init__(self, text, stop="end_turn"):
        self.stop_reason = stop
        self.content = [type("B", (), {"type": "text", "text": text})()]


@pytest.mark.parametrize("exc,expect", [
    (FakeAnthropic.APITimeoutError("t"), "too long"), (FakeAnthropic.APIConnectionError("c"), "no internet"),
    (FakeAnthropic.AuthenticationError("a"), "key rejected"), (FakeAnthropic.RateLimitError("r"), "rate limit"),
    (FakeAnthropic.APIStatusError("s"), "error 500"), (FakeAnthropic.APIError("e"), "APIError"),
    (RuntimeError("anything else"), "RuntimeError"), (MemoryError("x"), "MemoryError"),
])
def test_every_failure_comes_back_as_a_result_with_an_error(exc, expect):
    def create(**kw):
        raise exc
    res = fake_reader(create).read(np.zeros((100, 400, 3), np.uint8), np.zeros((50, 400, 3), np.uint8))
    assert isinstance(res, NameResult) and expect in res.error and not res.name


def test_odd_replies_and_bad_pictures_are_errors_too():
    r = fake_reader(lambda **kw: _Resp("", stop="max_tokens")).read(np.zeros((10, 40, 3), np.uint8))
    assert "tokens" in r.error
    assert fake_reader(lambda **kw: _Resp("no json here")).read(np.zeros((10, 40, 3), np.uint8)).error
    assert fake_reader(lambda **kw: _Resp("[1, 2]")).read(np.zeros((10, 40, 3), np.uint8)).error
    r = fake_reader(lambda **kw: _Resp('{"name": "A"}')).read(np.zeros((10, 40, 3), np.uint8), None)
    assert r.error is None and r.name == "A"
    assert fake_reader(lambda **kw: _Resp("{}")).read(None).error is None           # (an empty reply parses to an empty name)
    assert fake_reader(lambda **kw: _Resp("{}")).read(object()).error               # not even a picture: still no exception


def test_crops_are_shrunk_to_800_px_jpegs_and_the_request_keeps_its_shape():
    sent = {}

    def create(**kw):
        sent.update(kw)
        return _Resp('{"name": "ANA", "jornada": "12", "legible": true}')
    big = np.random.default_rng(0).integers(0, 255, (300, 2000, 3)).astype(np.uint8)
    res = fake_reader(create, effort="medium").read(big, big[:100])
    assert res.name == "ANA"
    assert sent["max_tokens"] == 4000 and sent["output_config"] == {"effort": "medium"}      # thinking tokens need the room
    assert "thinking" not in sent and "temperature" not in sent
    content = sent["messages"][0]["content"]
    assert [c["type"] for c in content] == ["image", "image", "text"]
    for c in content[:2]:
        assert c["source"]["media_type"] == "image/jpeg"
        img = cv2.imdecode(np.frombuffer(base64.b64decode(c["source"]["data"]), np.uint8), cv2.IMREAD_COLOR)
        assert img.shape[1] == name_reader.MAX_IMAGE_W == 800


def test_legible_false_as_text_is_not_true_and_names_are_capped():
    assert name_reader.parse_reply('{"name": "A", "legible": "false"}').legible is False
    assert name_reader.parse_reply('{"name": "A", "legible": "true"}').legible is True
    assert len(name_reader.parse_reply('{"name": "' + "X" * 500 + '"}').name) == name_reader.MAX_NAME_CHARS


def test_real_sdk_timeout_error_is_a_connection_error_so_the_order_of_the_except_clauses_matters():
    anthropic = pytest.importorskip("anthropic")
    assert issubclass(anthropic.APITimeoutError, anthropic.APIConnectionError)
    assert issubclass(anthropic.AuthenticationError, anthropic.APIStatusError)


def test_make_reader_sets_a_background_friendly_timeout_and_one_retry(monkeypatch):
    pytest.importorskip("anthropic")
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-test")
    reader, note = name_reader.make_reader(True, "some-model", "medium")
    assert isinstance(reader, ClaudeNameReader) and reader.model == "some-model" and reader.effort == "medium"
    assert reader.client.max_retries == 1
    t = reader.client.timeout
    assert t.connect == 5.0 and t.read == 25.0
    assert name_reader.make_reader(False, "m")[0] is None
    monkeypatch.delenv("ANTHROPIC_API_KEY")
    monkeypatch.delenv("ANTHROPIC_AUTH_TOKEN", raising=False)
    assert name_reader.make_reader(True, "m")[0] is None


# ======================================================================================= camera, sound, log
def test_frame_grabber_hands_out_only_newer_frames_and_survives_a_failing_camera():
    class Cam:
        def __init__(self):
            self.n = 0

        def read(self):
            time.sleep(0.005)
            self.n += 1
            if self.n % 7 == 0:
                raise RuntimeError("driver hiccup")
            if self.n % 5 == 0:
                return None
            return np.full((4, 4, 3), self.n % 250, np.uint8)
    g = FrameGrabber(Cam())
    try:
        seen, last = [], 0
        for _ in range(12):
            fid, frame = g.get(last, 0.5)
            assert frame is not None and fid > last
            seen.append(fid)
            last = fid
            time.sleep(0.03)                                          # a slow consumer: frames in between are dropped
        assert seen == sorted(set(seen)) and seen[-1] - seen[0] > len(seen)
        assert g.get(last + 10**6, 0.05) == (last + 10**6, None) and g.age() < 1.0
    finally:
        g.stop()
    assert not g._thread.is_alive()


def test_sound_is_silent_off_windows_and_never_raises(monkeypatch):
    sound.ok()
    sound.attention()                                                 # (a no-op on this machine)
    played = []
    fake = type("W", (), {"MB_ICONASTERISK": 1, "MB_ICONEXCLAMATION": 2, "MessageBeep": staticmethod(played.append)})
    monkeypatch.setattr(sound, "_backend", lambda: fake)
    sound.ok()
    sound.attention()
    assert played == [1, 2]

    def broken():
        raise OSError("no sound card")
    monkeypatch.setattr(sound, "_backend", broken)
    sound.ok()
    sound.attention()


def test_scanner_log_lines_and_rotation(tmp_path):
    log = ScanLog(tmp_path / "scanner.log")
    log.line("ENTRY", ticket=251, read_ms=190.4, excel_ms=12.0, name_ms=None, status="ok", name="CLEMENTE O.")
    text = (tmp_path / "scanner.log").read_text()
    assert text.split("  ")[1] == "ENTRY" and "ticket=251" in text and "read_ms=190.4" in text
    assert 'name="CLEMENTE O."' in text and "name_ms" not in text and text.endswith("\n")
    ScanLog.MAX_BYTES = 50
    try:
        log.line("ENTRY", ticket=252, status="ok")
        log.line("ENTRY", ticket=253, status="ok")
    finally:
        ScanLog.MAX_BYTES = 2_000_000
    assert (tmp_path / "scanner.log.1").exists() and "ticket=253" in (tmp_path / "scanner.log").read_text()
