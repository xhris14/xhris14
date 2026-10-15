"""Regression tests for the adversarial review of the Excel layer (findings F1-F12, SOAK-5, SOAK-8).

Each test fails on the code before the fix. They run against the instrumented fake Excel object model of
test_excel_fast.py (whose com_error stand-in has the real pywintypes shape) and, for the file backend, real files.
"""
import os
import re
import sys
import types
import zipfile
from collections import Counter

import openpyxl
import pytest

import quiniela_scanner.excel_sink as xs
from quiniela_scanner.excel_sink import (BlockFull, ComBackend, ExcelBusy, ExcelError, ExcelSink,
                                         OpenpyxlBackend, Receipt, SlotTaken, Snapshot, discover_blocks, is_ticket,
                                         needs_text_prefix)
from helpers import make_pool_workbook
from test_excel_fast import (E_FORMAT_DENIED, NAME_COL, TICKET_COL, VBA_E_IGNORE, FakeApp, FakeComError,
                             FakeRange, PyWinComError, Watch, excel_would_convert, fresh, picks_of,
                             pool_app)

REF = -2146826265                      # #REF!  as pywin32 hands an error cell over


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    sleeps = []
    monkeypatch.setattr(xs.time, "sleep", sleeps.append)
    return sleeps


def raiser(make):
    calls = []

    def go(*a, **k):
        calls.append(1)
        raise make()
    go.calls = calls
    return go


# ------------------------------------------------------------------ F1 / SOAK-5: busy in every shape
@pytest.mark.parametrize("make", [
    lambda: FakeComError(-2147418111, "call rejected"),                              # transport level
    lambda: PyWinComError(VBA_E_IGNORE),                                              # cell in edit mode: excepinfo[5]
    lambda: PyWinComError(-2147417846, hresult=-2147352567),                          # RETRYLATER inside DISP_E_EXCEPTION
], ids=["call-rejected", "edit-mode", "retry-later"])
def test_busy_is_recognised_in_every_com_error_shape(tmp_path, no_sleep, make):
    sink, be, app, ws = fresh(tmp_path)
    ws.check_busy = (boom := raiser(make))
    with pytest.raises(ExcelBusy, match="Press Esc") as ei:
        sink.write_row_fast(None, picks_of(0), name="X")
    assert len(boom.calls) == be.RETRY_TRIES and len(no_sleep) == be.RETRY_TRIES
    assert isinstance(ei.value, ExcelError) and (15, NAME_COL) not in ws.cells and app.screen is True
    rc = None
    ws.check_busy = lambda: None
    rc = sink.write_row_fast(None, picks_of(0), name="")                               # Excel free again
    ws.check_busy = (boom2 := raiser(make))
    no_sleep.clear()
    with pytest.raises(ExcelBusy):
        sink.update_name(rc, "LATE")
    assert len(boom2.calls) == be.RETRY_TRIES


def test_an_unrelated_excel_error_is_not_retried_as_busy(tmp_path, no_sleep):
    sink, be, app, ws = fresh(tmp_path)
    ws.check_busy = (boom := raiser(lambda: PyWinComError(E_FORMAT_DENIED)))
    with pytest.raises(ExcelError, match="Excel error") as ei:
        sink.write_row_fast(None, picks_of(0), name="X")
    assert not isinstance(ei.value, ExcelBusy) and len(boom.calls) == 2 and no_sleep == []   # 1 try + 1 after re-resolving


def test_excel_busy_and_block_full_are_part_of_the_error_hierarchy():
    assert issubclass(ExcelBusy, ExcelError) and issubclass(BlockFull, ExcelError)
    assert not issubclass(ExcelBusy, SlotTaken) and BlockFull("x").block is None
    assert Receipt.__dataclass_fields__["warnings"].default_factory is list


# ------------------------------------------------------------------ F2: the write is atomic from the caller's view
def test_font_size_failing_after_the_values_landed_does_not_raise_and_does_not_duplicate(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.protect_formats = True                                       # entry cells unlocked, 'Format cells' not allowed
    picks = ["L", "", "E", "V", "V", "V", "V", "V", "V"]             # one empty box -> F1 + font size 9
    rc = sink.write_row_fast(None, picks, name="CLEMENTE")
    assert ws.row(15) == ["CLEMENTE", "L", "F1", "E", "V", "V", "V", "V", "V", "V"]
    assert rc.ticket == 251 and rc.row == 15 and any("font size" in w for w in rc.warnings)
    assert not sink.is_free(rc.slot) and app.screen is True
    nxt = sink.write_row_fast(None, picks_of(1), name="B")           # the next slip moves on, nothing is entered twice
    assert nxt.ticket == 252 and nxt.warnings == []
    assert [k for k, v in ws.cells.items() if v == "CLEMENTE"] == [(15, NAME_COL)]
    sink.update_name(rc, "CLEMENTE O.")
    sink.undo(rc)                                                    # the undo's own font reset fails too: a warning
    assert all(ws.cells.get((15, NAME_COL + k)) is None for k in range(10))
    assert any("font size" in w for w in rc.warnings)


def test_a_failing_save_after_the_write_is_a_warning_not_an_error(tmp_path):
    app, ws = pool_app(tmp_path)
    be = ComBackend(app=app)
    sink = ExcelSink(be, save_after_write=True)
    app.Workbooks.books[0].save_error = PyWinComError(0x800A03EC - (1 << 32))      # read-only workbook
    rc = sink.write_row_fast(None, picks_of(0), name="")
    assert ws.cells[(15, NAME_COL + 1)] == "L" and rc.warnings and "could not Save" in rc.warnings[0]
    warnings = sink.update_name(rc, "LATE")
    assert ws.cells[(15, NAME_COL)] == "LATE" and warnings and rc.warnings[-1] == warnings[-1]
    assert sink.write_row_fast(None, picks_of(1), name="").ticket == 252


def test_values_that_landed_although_the_call_raised_count_as_written(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    orig = FakeRange.Value

    def set_then_fail(self, v):
        orig.fset(self, v)
        raise PyWinComError(E_FORMAT_DENIED)                         # Excel stored it, then reported an error

    FakeRange.Value = orig.setter(set_then_fail)
    try:
        rc = sink.write_row_fast(None, picks_of(0), name="A")
    finally:
        FakeRange.Value = orig
    assert rc.ticket == 251 and rc.warnings and "holds the values" in rc.warnings[0]
    assert ws.row(15) == ["A"] + picks_of(0)
    assert sink.write_row_fast(None, picks_of(1), name="B").ticket == 252


def test_a_write_that_did_not_land_still_raises_and_leaves_the_slot_free(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.check_busy = raiser(lambda: PyWinComError(E_FORMAT_DENIED))   # a protected cell: nothing is stored
    with pytest.raises(ExcelError) as ei:
        sink.write_row_fast(None, picks_of(0), name="A")
    assert not isinstance(ei.value, ExcelBusy) and "may already hold" not in str(ei.value)
    assert (15, NAME_COL) not in ws.cells and sink.is_free(sink.slot_for_ticket(251))
    ws.check_busy = lambda: None
    assert sink.write_row_fast(None, picks_of(0), name="A").ticket == 251


def test_a_failed_write_that_cannot_be_verified_says_the_row_may_hold_the_slip(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    state = {"armed": False}
    orig_get, orig_set = FakeRange.Value.fget, FakeRange.Value.fset

    def get(self):
        if state["armed"]:
            raise FakeComError(-2147418111, "busy")
        return orig_get(self)

    def set_(self, v):
        state["armed"] = True
        raise PyWinComError(E_FORMAT_DENIED)

    FakeRange.Value = property(get, set_)
    try:
        with pytest.raises(ExcelError, match=r"row 15 may already hold this data") as ei:
            sink.write_row_fast(None, picks_of(0), name="A")
    finally:
        FakeRange.Value = property(orig_get, orig_set)
    assert not isinstance(ei.value, ExcelBusy)


def test_busy_excel_on_the_write_means_nothing_was_entered(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.busy = 99
    with pytest.raises(ExcelBusy):
        sink.write_row_fast(None, picks_of(0), name="A")
    assert (15, NAME_COL) not in ws.cells and sink.is_free(sink.slot_for_ticket(251))
    ws.busy = 0
    assert sink.write_row_fast(None, picks_of(0), name="A").ticket == 251


# ------------------------------------------------------------------ F3: connecting
def test_connect_retries_a_busy_excel_and_gives_up_with_excel_busy(tmp_path, no_sleep):
    app, ws = pool_app(tmp_path)
    real_books, fails = app.Workbooks, [3]

    class Books:
        @property
        def Count(self):
            if fails[0] > 0:
                fails[0] -= 1
                raise FakeComError(-2147418111, "Call was rejected by callee.")
            return real_books.Count

        def __call__(self, k):
            return real_books(k)

    app.Workbooks = Books()
    assert ComBackend(app=app).workbook_name == "JORNADA 10 PC.xlsx" and len(no_sleep) == 3      # transient: connected
    fails[0] = 99
    with pytest.raises(ExcelBusy, match="busy"):
        ComBackend(app=app)


def test_connect_with_no_active_workbook_is_a_clear_excel_error(tmp_path):
    app, ws = pool_app(tmp_path)
    app.ActiveWorkbook = None                       # pywin32's Nothing: only a hidden PERSONAL.XLSB is loaded
    with pytest.raises(ExcelError, match="no active workbook") as ei:
        ComBackend(app=app)
    assert not str(ei.value).startswith("Excel error")          # our message, not re-wrapped
    assert ComBackend(workbook_hint="jornada", app=app).workbook_name == "JORNADA 10 PC.xlsx"


def test_connect_errors_from_us_are_not_wrapped_twice_and_foreign_ones_are_wrapped(tmp_path):
    app, ws = pool_app(tmp_path)
    with pytest.raises(ExcelError) as ei:
        ComBackend(workbook_hint="nope", app=app)
    assert str(ei.value) == "no open workbook matches 'nope'"
    app.Workbooks.books.clear()
    with pytest.raises(ExcelError) as ei:
        ComBackend(app=app)
    assert str(ei.value) == "Excel is running but no workbook is open"
    app2, _ = pool_app(tmp_path / "x") if (tmp_path / "x").mkdir() is None else (None, None)

    class Weird:
        @property
        def Name(self):
            raise AttributeError("COM went strange")
    app2.ActiveWorkbook = Weird()
    with pytest.raises(ExcelError, match="Excel error: COM went strange"):
        ComBackend(app=app2)


# ------------------------------------------------------------------ F11: the real GetActiveObject error is kept
def test_excel_is_not_running_keeps_the_friendly_text_and_the_real_error(monkeypatch):
    client = types.ModuleType("win32com.client")

    def get_active(name):
        raise PyWinComError(-2147221021, hresult=-2147221021, msg="Operation unavailable")
    client.GetActiveObject = get_active
    pkg = types.ModuleType("win32com")
    pkg.client = client
    monkeypatch.setitem(sys.modules, "win32com", pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(CoInitialize=lambda: None))
    with pytest.raises(ExcelError) as ei:
        ComBackend()
    msg = str(ei.value)
    assert msg.startswith("Excel is not running - open your jornada workbook first")
    assert "Operation unavailable" in msg and "-2147221021" in msg and "same privileges" in msg


# ------------------------------------------------------------------ F4: the operator's selected block survives
def test_resolve_block_follows_a_block_whose_header_moved(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    old = sink.block_infos()[1].block                          # block 2, column Z, header Z4
    assert old.key == "ENTRY PG1!Z4" and old.first_ticket == 481
    assert sink.resolve_block(old) is not None and sink.resolve_block(old).key == old.key
    ws.user_insert_row(2)
    sink.refresh()                                            # what a stale probe does by itself
    got = sink.resolve_block(old)
    assert got is not None and got.key == "ENTRY PG1!Z5" and got.name_col == old.name_col
    assert got in sink.blocks and sink.resolve_block(None) is not None
    gone = types.SimpleNamespace(sheet="ENTRY PG1", name_col=99, header_row=4, key="ENTRY PG1!CU4", first_ticket=None)
    assert sink.resolve_block(gone) is None and sink.resolve_block(None) in sink.blocks


def test_resolve_block_tells_stacked_blocks_apart_by_their_first_ticket(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    for c, v in zip(range(TICKET_COL, TICKET_COL + 11), ["#", "NOMBRE"] + list(range(1, 10))):
        ws.cells[(70, c)] = v                                 # a second block under the first, same columns
    for i in range(5):
        ws.cells[(71 + i, TICKET_COL)] = 9000 + i
    sink.refresh()
    first, second = sorted((b for b in sink.blocks if b.name_col == NAME_COL), key=lambda b: b.header_row)
    assert (first.first_ticket, second.first_ticket) == (241, 9000)
    ws._shift(2, +80)                                         # 80 rows inserted: 'nearest header' now picks the wrong one
    sink.refresh()
    new_first, new_second = sorted((b for b in sink.blocks if b.name_col == NAME_COL), key=lambda b: b.header_row)
    assert sink.resolve_block(second) is new_second and sink.resolve_block(first) is new_first


def test_writing_to_the_old_block_object_still_lands_in_that_block_after_a_row_insert(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    old = sink.block_infos()[1].block
    sink.write_row_fast(old, picks_of(0), name="A")
    ws.user_insert_row(2)
    rc = sink.write_row_fast(old, picks_of(1), name="B")
    assert rc.ticket == 485 and rc.col == NAME_COL + 12 and ws.cells[(rc.row, TICKET_COL + 12)] == 485


# ------------------------------------------------------------------ F5: error codes / repeated ticket numbers
def test_error_codes_and_booleans_are_not_ticket_numbers():
    assert is_ticket(241) and is_ticket(0) and not is_ticket(True) and not is_ticket(None) and not is_ticket("241")
    for code in (-2146826288, REF, -2146826246, -2146826245):
        assert not is_ticket(code)
    assert is_ticket(-2146826289) and is_ticket(-2146826244)
    vals = [["#", "NOMBRE"] + list(range(1, 10))] + [[241 + i] + [None] * 10 for i in range(3)] + \
           [[REF] + [None] * 10, [245] + [None] * 10]
    skipped = []
    (block,) = discover_blocks("S", Snapshot(1, 1, vals), skipped=skipped)
    assert block.rows == [2, 3, 4] and block.first_ticket == 241 and skipped == []


def test_error_tickets_end_the_block_with_a_warning_and_a_clear_block_full(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    for r in range(15, 30):
        ws.cells[(r, TICKET_COL)] = REF                       # a deleted row broke a =M5+1 chain
    sink.refresh()
    info = sink.block_infos()[0]
    assert info.block.rows[-1] == 14 and "tickets 241-250" in info.label
    assert any("column N" in w and "#REF!" in w and "ticket cell is not a number" in w for w in sink.layout_warnings)
    for i in range(info.free):
        sink.write_row_fast(info.block, picks_of(i), name=f"N{i}")
    with pytest.raises(BlockFull, match=r"block ends at row 14 because the ticket cell is not a number \(row 15: the "
                                        r"Excel error #REF!\)") as ei:
        sink.write_row_fast(info.block, picks_of(0), name="X")
    assert ei.value.block is sink.resolve_block(info.block)
    assert all(ws.cells.get((r, NAME_COL)) is None for r in range(15, 30))             # nothing under the broken tickets


def test_repeated_ticket_numbers_are_reported_and_a_late_name_finds_its_own_row(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.cells[(16, TICKET_COL)] = 251                          # somebody typed 251 twice (rows 15 and 16)
    sink.refresh()
    assert ("ENTRY PG1", 251) in sink.dup_tickets
    assert any("ticket 251 appears 2 times" in w and "rows 15, 16" in w for w in sink.layout_warnings)
    block = sink.block_infos()[0].block
    a = sink.write_row_fast(block, picks_of(0), name="")
    b = sink.write_row_fast(block, picks_of(1), name="")
    assert (a.ticket, a.row, b.ticket, b.row) == (251, 15, 251, 16)
    ws.user_insert_row(2)                                     # A is row 16 now, B row 17 - same ticket number
    sink.update_name(b, "NAME OF B")
    assert b.row == 17 and ws.cells[(17, NAME_COL)] == "NAME OF B" and ws.cells.get((16, NAME_COL)) is None
    sink.update_name(a, "NAME OF A")
    assert a.row == 16 and ws.cells[(16, NAME_COL)] == "NAME OF A" and ws.cells[(17, NAME_COL)] == "NAME OF B"
    ws.user_insert_row(2)
    sink.undo(b)                                              # clears B's row, not A's
    assert ws.row(18) == [None] * 10 and ws.cells[(17, NAME_COL)] == "NAME OF A"


def test_a_clean_sheet_has_no_layout_warnings(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    assert sink.layout_warnings == [] and sink.dup_tickets == set()
    s2 = ExcelSink(OpenpyxlBackend(str(make_pool_workbook(tmp_path / "f.xlsx"))))
    assert s2.layout_warnings == []


# ------------------------------------------------------------------ F6: formulas that show nothing are not free
def formula_sheet(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    for key in [(15, NAME_COL), (15, NAME_COL + 3)]:          # name and pick 3 of ticket 251: =IF(A5="","",A5) -> ''
        ws.cells[key] = ""
        ws.formulas.add(key)
    sink.refresh()
    return sink, be, app, ws


def test_formula_cells_showing_nothing_are_never_free_and_survive(tmp_path):
    sink, be, app, ws = formula_sheet(tmp_path)
    assert not sink.is_free(sink.slot_for_ticket(251)) and sink.block_infos()[0].free == 41
    assert any("2 name/pick cell(s) hold formulas" in w and "row 15" in w for w in sink.layout_warnings)
    rc = sink.write_row_fast(None, picks_of(0), name="A")
    assert rc.ticket == 252 and {(15, NAME_COL), (15, NAME_COL + 3)} <= ws.formulas
    sink.undo(rc)
    assert {(15, NAME_COL), (15, NAME_COL + 3)} <= ws.formulas


def test_an_explicit_ticket_on_a_formula_row_is_taken_and_an_overwrite_says_so(tmp_path):
    sink, be, app, ws = formula_sheet(tmp_path)
    with pytest.raises(SlotTaken):
        sink.write_row_fast(None, picks_of(0), name="A", ticket=251)
    assert {(15, NAME_COL), (15, NAME_COL + 3)} <= ws.formulas
    rc = sink.write_row_fast(None, picks_of(0), name="A", ticket=251, allow_overwrite=True)
    assert ws.cells[(15, NAME_COL)] == "A" and (15, NAME_COL) not in ws.formulas
    assert any("formula" in w and "Undo cannot bring" in w for w in rc.warnings)
    assert sink.snaps["ENTRY PG1"].formula_cells.isdisjoint({(15, NAME_COL + k) for k in range(10)})


def test_formulas_are_read_once_per_refresh_and_cost_no_extra_round_trip_per_slip(tmp_path):
    app, ws = pool_app(tmp_path)
    be = ComBackend(app=app)
    sink = ExcelSink(be)
    assert app.meter.ops["formula_get"] == 2                              # ENTRY PG1 + ENTRY PG2, once
    watch = Watch(be, app)
    sink.write_row_fast(None, picks_of(0), name="A")
    ops, trips, _ = watch.com()
    assert ops == Counter(value_get=1, value_set=1) and trips == 6 and ops["formula_get"] == 0


def test_snapshot_blank_treats_formula_cells_as_data():
    snap = Snapshot(1, 1, [[None] * 5], formula_cells={(1, 3)})
    assert snap.blank(1, 1, 2) and not snap.blank(1, 1, 3) and not snap.blank(1, 3, 1) and snap.blank(1, 4, 2)


def test_file_backend_also_protects_formulas_that_show_nothing(tmp_path):
    path = make_pool_workbook(tmp_path / "w.xlsx")
    wb = openpyxl.load_workbook(path)
    wb["ENTRY PG1"].cell(15, NAME_COL).value = '=IF(A1="","",A1)'
    wb["ENTRY PG1"].cell(15, NAME_COL + 2).value = "=IF(A1=1,\"L\",\"\")"
    wb.save(path)
    sink = ExcelSink(OpenpyxlBackend(str(path)))
    rc = sink.write_row_fast(None, picks_of(0), name="A")
    assert rc.ticket == 252
    assert openpyxl.load_workbook(path)["ENTRY PG1"].cell(15, NAME_COL).value == '=IF(A1="","",A1)'


# ------------------------------------------------------------------ F7: a block cut short by a ticket cell that is not a number
def test_text_ticket_cuts_the_block_with_a_warning_and_the_right_message(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.cells[(24, TICKET_COL)] = "260 "                                    # ticket 260 typed as text
    sink.refresh()
    info = sink.block_infos()[0]
    assert "tickets 241-259" in info.label                                  # the strict rule is kept
    assert any("block ends at row 23 because the ticket cell is not a number (row 24: '260')" in w
               for w in sink.layout_warnings)
    for i in range(info.free):
        sink.write_row_fast(info.block, picks_of(i), name=f"N{i}")
    with pytest.raises(BlockFull, match=r"block ends at row 23 because the ticket cell is not a number") as ei:
        sink.write_row_fast(info.block, picks_of(0), name="X")
    assert "tickets 241-259" in str(ei.value) and ws.cells.get((24, NAME_COL)) is None


def test_blank_spacer_row_in_the_ticket_column_is_explained(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.user_insert_row(20)
    sink.refresh()
    assert any("block ends at row 19 because the ticket cell is not a number (row 20: empty)" in w
               for w in sink.layout_warnings)


def test_a_block_that_simply_ends_is_not_a_warning(tmp_path):
    sink, be, app, ws = fresh(tmp_path, filled=(52, 0, 0))
    ws.cells[(57, TICKET_COL)] = "TOTAL"                                     # a footer label under the block
    sink.refresh()
    assert sink.layout_warnings == []
    with pytest.raises(BlockFull) as ei:
        sink.write_row_fast(sink.block_infos()[0].block, picks_of(0), name="X")
    assert str(ei.value).endswith("every row already has data") and "ENTRY PG1 column N" in str(ei.value)


def test_header_without_numeric_tickets_says_why_instead_of_no_header(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    for r in range(5, 57):
        for off in (0, 12, 24):
            ws.cells[(r, TICKET_COL + off)] = f"{ws.cells[(r, TICKET_COL + off)]}"      # every ticket became text
    with pytest.raises(ExcelError, match=r"header found .*has no ticket number under it \(row 5 is '241'\)"):
        sink.refresh()


# ------------------------------------------------------------------ F8: the snapshot is as wide as the data, never wider than 256
def test_inflated_used_range_is_not_read_sideways(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.used_extra = (0, 16384 - 35)                                          # whole-row formatting: 16384 columns "used"
    ws.read_log.clear()
    snap = be.snapshot("ENTRY PG1")
    assert ws.read_log == [(54, 36)] and snap.last_col == 48 and snap.notes == []         # M..AX, the real data
    ExcelSink(be)                                                             # and the sink still finds its blocks


def test_a_sheet_that_really_is_wider_than_256_columns_is_capped_with_a_warning(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.cells[(5, 400)] = "stray"
    ws.read_log.clear()
    sink.refresh()
    snap = sink.snaps["ENTRY PG1"]
    assert max(c for _, c in ws.read_log) == 256 and snap.last_col - snap.left + 1 == 256
    assert any("only the first 256 columns" in w for w in sink.layout_warnings)
    assert len(sink.blocks) == 3


def test_without_cells_find_the_cap_still_holds(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.used_extra = (0, 16384)
    ws.find_error = PyWinComError(E_FORMAT_DENIED)
    ws.read_log.clear()
    sink.refresh()
    assert max(c for _, c in ws.read_log) == 256 and len(sink.blocks) == 3
    assert any("256" in w for w in sink.layout_warnings)


def test_find_that_hits_a_busy_excel_is_retried(tmp_path, no_sleep):
    sink, be, app, ws = fresh(tmp_path)
    ws.find_error = FakeComError(-2147418111, "busy")
    with pytest.raises(ExcelBusy):
        be.snapshot("ENTRY PG1")
    assert len(no_sleep) == be.RETRY_TRIES


def test_file_backend_caps_the_width_too(tmp_path):
    path = make_pool_workbook(tmp_path / "w.xlsx")
    wb = openpyxl.load_workbook(path)
    wb["ENTRY PG1"].cell(5, 400).value = "stray"
    wb.save(path)
    sink = ExcelSink(OpenpyxlBackend(str(path)))
    assert sink.snaps["ENTRY PG1"].last_col == 256 and any("only the first 256" in w for w in sink.layout_warnings)


# ------------------------------------------------------------------ SOAK-8: BlockFull
def test_block_full_names_the_block_and_carries_it(tmp_path):
    sink, be, app, ws = fresh(tmp_path, filled=(52, 0, 0))
    full = sink.block_infos()[0].block
    with pytest.raises(BlockFull, match=r"no free ticket in block ENTRY PG1 column N \(tickets 241-292\)") as ei:
        sink.write_row_fast(full, picks_of(0), name="X")
    assert ei.value.block is sink.resolve_block(full)
    assert ei.value.block.key == full.key and ei.value.block.name_col == NAME_COL


def test_all_blocks_full_is_block_full_too(tmp_path):
    sink, be, app, ws = fresh(tmp_path, filled=(52, 52, 52))
    with pytest.raises(BlockFull, match="no free ticket in block"):
        sink.write_row_fast(None, picks_of(0), name="X")


def test_a_vanished_block_is_a_plain_excel_error_that_names_it(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    old = sink.block_infos()[1].block
    ws.cells.pop((4, NAME_COL + 12))                  # block 2 loses its NOMBRE header ...
    ws.user_insert_row(2)                             # ... and the sheet moves, so the write notices and re-reads
    with pytest.raises(ExcelError, match="ENTRY PG1!Z4") as ei:
        sink.write_row_fast(old, picks_of(0), name="X")
    assert not isinstance(ei.value, BlockFull)


# ------------------------------------------------------------------ F9: text that Excel would parse
PARSED = ["MAR 5", "MAY 10", "JUN 3", "5-MAR", "DIC 12", "ene 3", "5 AGO 2026", "MARZO 2026", "1E3", "5%", "TRUE",
          "false", "(5)", "$5", "10:30", "5 PM", "10:30AM", "'TONY", "#REF!", "12/10", "1,5", "=1+1", "+5", "-ANA",
          "@HOME", "2024", "1.5.2020", "SEP-5", "5/SEP"]
LITERAL = ["RECA 30 RRR", "TONY 5000 RRR", "CLEMENTE O. RRR", "MARIA 5 RRR", "JOSE E", "30 RECA", "MAR 5 RRR",
           "MARIO", "ANA-LUZ", "E5", "A1", "N/A"]


@pytest.mark.parametrize("name", PARSED)
def test_names_excel_would_parse_are_stored_as_the_same_text(tmp_path, name):
    assert excel_would_convert(name), "test oracle: Excel really would convert this"
    assert needs_text_prefix(name)
    sink, be, app, ws = fresh(tmp_path)
    rc = sink.write_row_fast(None, picks_of(0), name=name)
    key = (rc.row, NAME_COL)
    assert ws.cells[key] == name and key in ws.texts and key not in ws.converted and key not in ws.formulas
    rc2 = sink.write_row_fast(None, picks_of(1), name="")
    sink.update_name(rc2, name)                                                # the late name too
    key2 = (rc2.row, NAME_COL)
    assert ws.cells[key2] == name and key2 in ws.texts and key2 not in ws.converted


@pytest.mark.parametrize("name", LITERAL)
def test_ordinary_names_get_no_prefix_and_no_extra_round_trip(tmp_path, name):
    assert not excel_would_convert(name) and not needs_text_prefix(name)
    sink, be, app, ws = fresh(tmp_path)
    watch = Watch(be, app)
    rc = sink.write_row_fast(None, picks_of(0), name=name)
    assert ws.cells[(rc.row, NAME_COL)] == name and (rc.row, NAME_COL) not in ws.texts
    assert watch.com()[1] == 6 and watch.com()[0]["numfmt"] == 0


def test_the_prefix_rule_covers_everything_the_oracle_says_excel_converts():
    import itertools
    months = ["MAR", "MAY", "jun", "Dic", "SEP", "AGO", "ENE", "ABR", "JUL", "OCT", "NOV", "DEC", "JAN", "APR", "AUG"]
    for m, n, sep in itertools.product(months, ["1", "5", "12", "2026"], [" ", "-", "/", ""]):
        for s in (f"{m}{sep}{n}", f"{n}{sep}{m}"):
            if excel_would_convert(s):
                assert needs_text_prefix(s), s
    for s in ["1", "12", "1.5", "1,000", "0.5", "1e5", "10:30", "1/2", "1 1/2", "(1)", "-1", "50%", "$1.50"]:
        assert excel_would_convert(s) and needs_text_prefix(s), s


def test_a_text_formatted_name_cell_gets_the_plain_string_not_a_visible_apostrophe(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    ws.text_cells.add((15, NAME_COL))                                          # the name cell has NumberFormat "@"
    watch = Watch(be, app)
    rc = sink.write_row_fast(None, picks_of(0), name="30")
    assert ws.cells[(15, NAME_COL)] == "30"                                    # no literal "'30"
    assert watch.com()[0]["numfmt"] == 1 and watch.com()[1] == 8               # one extra format check (2 trips), rarely
    sink.update_name(rc, "1-2-3", expect="30")                                 # the read-back still matches
    assert ws.cells[(15, NAME_COL)] == "1-2-3"
    other = sink.write_row_fast(None, picks_of(1), name="MAR 5")               # a General cell next door: prefix
    assert (other.row, NAME_COL) in ws.texts and ws.cells[(other.row, NAME_COL)] == "MAR 5"


# ------------------------------------------------------------------ F10: ScreenUpdating is never left off
class RejectingApp(FakeApp):
    """Excel goes into cell-edit mode as soon as ScreenUpdating is switched off: further sets are rejected."""
    editing = False
    fail_restores = 0

    @property
    def ScreenUpdating(self):
        return self.screen

    @ScreenUpdating.setter
    def ScreenUpdating(self, v):
        if self.editing or (v and self.fail_restores > 0):
            if v and self.fail_restores > 0:
                self.fail_restores -= 1
            raise FakeComError(-2147418111, "Call was rejected by callee.")
        self.screen = v
        self.screen_log.append(v)


def test_a_rejected_restore_is_retried_at_once(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    app.__class__ = RejectingApp
    app.fail_restores = 1                                                     # the first restore bounces, the second works
    rc = sink.write_row_fast(None, picks_of(0), name="A")
    assert app.screen is True and rc.warnings == [] and app.screen_log[-2:] == [False, True]


def test_screen_updating_stuck_off_is_reported_and_healed_by_the_next_call(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    app.__class__ = RejectingApp
    app.fail_restores = 99
    rc = sink.write_row_fast(None, picks_of(0), name="A")                     # values landed, the restore failed twice
    assert app.screen is False and be._screen_stuck
    assert any("ScreenUpdating" in w for w in rc.warnings)
    app.fail_restores = 0                                                     # Excel is free again (the operator hit Esc)
    be.read_row("ENTRY PG1", 15, NAME_COL, 1)                                 # ANY next call puts the window back
    assert app.screen is True and not be._screen_stuck


def test_a_write_that_fails_while_excel_edits_does_not_leave_the_screen_off(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    app.__class__ = RejectingApp
    ws.busy = 99
    app.editing = False
    orig = ws.check_busy

    def check_busy():                                                         # Excel enters edit mode mid-write
        app.editing = True
        orig()
    ws.check_busy = check_busy
    with pytest.raises(ExcelBusy):
        sink.write_row_fast(None, picks_of(0), name="A")
    assert app.screen is False and be._screen_stuck                           # could not be restored in edit mode ...
    app.editing = False
    ws.busy = 0
    ws.check_busy = orig
    sink.write_row_fast(None, picks_of(0), name="A")                          # ... the next write puts it right
    assert app.screen is True and not be._screen_stuck


def test_screen_updating_is_not_touched_back_on_when_it_could_not_be_switched_off(tmp_path):
    sink, be, app, ws = fresh(tmp_path)
    app.__class__ = RejectingApp
    app.editing = True
    app.screen_log.clear()
    sink.write_row_fast(None, picks_of(0), name="A")
    assert app.screen_log == [] and not be._screen_stuck


# ------------------------------------------------------------------ F12: the file backend
def test_a_crash_while_saving_leaves_the_workbook_untouched(tmp_path, monkeypatch):
    path = make_pool_workbook(tmp_path / "JORNADA 10 PC.xlsx")
    sink = ExcelSink(OpenpyxlBackend(str(path)))
    before = path.read_bytes()
    import openpyxl.writer.excel as W
    orig = W.ExcelWriter.write_data

    def boom(self):
        self._write_worksheets()                                             # part of the archive is written ...
        raise OSError(28, "No space left on device")                         # ... then the disk fills up
    monkeypatch.setattr(W.ExcelWriter, "write_data", boom)
    with pytest.raises(ExcelError, match="could not save"):
        sink.write_row_fast(None, picks_of(0), name="X")
    monkeypatch.setattr(W.ExcelWriter, "write_data", orig)
    assert path.read_bytes() == before and openpyxl.load_workbook(path)["ENTRY PG1"].cell(15, NAME_COL).value is None
    assert sorted(os.listdir(tmp_path)) == ["JORNADA 10 PC.xlsx"]            # no temp file left behind
    assert sink.write_row_fast(None, picks_of(0), name="X").ticket == 251    # and it still works afterwards


def workbook_saved_by_excel(path):
    """ENTRY PG1 whose ticket column is a formula chain (=M5+1) WITH the values Excel caches next to it."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ENTRY PG1"
    ws["M4"], ws["N4"] = "#", "NOMBRE"
    for k in range(9):
        ws.cell(4, 15 + k, k + 1)
    ws["M5"] = 241
    for r in range(6, 40):
        ws.cell(r, 13, f"=M{r - 1}+1")
    raw = path.with_name("raw.xlsx")
    wb.save(raw)
    with zipfile.ZipFile(raw) as zi, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                x = re.sub(r'(<c r="M(\d+)"[^>]*><f>[^<]*</f>)<v\s*/>|(<c r="M(\d+)"[^>]*><f>[^<]*</f>)<v></v>',
                           lambda m: f'{m[1] or m[3]}<v>{236 + int(m[2] or m[4])}</v>', data.decode())
                data = x.encode()
            zo.writestr(item, data)
    raw.unlink()
    return path


def test_formula_tickets_survive_our_own_saves_in_the_file_backend(tmp_path):
    path = workbook_saved_by_excel(tmp_path / "JORNADA 10 PC.xlsx")
    assert openpyxl.load_workbook(path, data_only=True)["ENTRY PG1"]["M7"].value == 243      # cached by "Excel"
    sink = ExcelSink(OpenpyxlBackend(str(path)))
    tickets = [sink.write_row_fast(None, picks_of(i), name=f"N{i}").ticket for i in range(3)]
    assert tickets == [241, 242, 243]                                                       # slip 2 no longer fails
    assert openpyxl.load_workbook(path, data_only=True)["ENTRY PG1"]["M7"].value is None     # openpyxl dropped the cache


@pytest.mark.parametrize("name", ["w.xlsm", "w.xls", "w.xlsb"])
def test_macro_and_legacy_workbooks_are_refused_as_files(tmp_path, name):
    with pytest.raises(ExcelError, match="only plain .xlsx"):
        OpenpyxlBackend(str(tmp_path / name))
