import openpyxl
import pytest

from quiniela_scanner.excel_sink import (ComBackend, ExcelError, ExcelSink, OpenpyxlBackend, SlotTaken,
                                         discover_blocks, needs_text_prefix, Snapshot)
from helpers import make_pool_workbook
from test_excel_fast import FakeBook, FakeSheet, pool_app


@pytest.fixture
def sink(tmp_path):
    path = make_pool_workbook(tmp_path / "JORNADA 10 PC.xlsx")
    return ExcelSink(OpenpyxlBackend(str(path))), path


def test_discovers_all_three_blocks(sink):
    s, _ = sink
    infos = s.block_infos()
    assert [i.block.name_col for i in infos] == [14, 26, 38]        # N, Z, AL
    assert [i.free for i in infos] == [42, 49, 52]
    assert [i.suffix for i in infos] == ["RRR", "CH", ""]
    assert infos[0].first_free.ticket == 251 and infos[1].first_free.ticket == 484
    assert infos[0].block.ticket_col == 13
    assert "tickets 241-292" in infos[0].label


def test_next_slot_and_default_block(sink):
    s, _ = sink
    assert s.default_block().first_free.ticket == 251
    assert s.next_slot().ticket == 251
    assert s.slot_for_ticket(493).row == 5 + 12


def test_write_puts_name_and_picks_and_leaves_T_formula(sink):
    s, path = sink
    slot = s.next_slot()
    picks = list("LVLLLLLLE")
    s.write(slot, "CLEMENTE O. RRR", picks)
    ws = openpyxl.load_workbook(path)["ENTRY PG1"]
    r = slot.row
    assert ws.cell(r, 14).value == "CLEMENTE O. RRR"
    assert [ws.cell(r, 15 + k).value for k in range(9)] == picks
    assert str(ws.cell(r, 24).value).startswith("=COUNTIF")         # T untouched
    assert ws.cell(r, 13).value == 251                                # ticket # untouched
    assert s.next_slot().ticket == 252                                # advances


def test_empty_box_becomes_F1_in_size_9_and_undo_restores_font(sink):
    s, path = sink
    slot = s.next_slot()
    rc = s.write(slot, "X", ["L", "", "E", "", "", "", "", "", "V"])
    ws = openpyxl.load_workbook(path)["ENTRY PG1"]
    assert [ws.cell(slot.row, 15 + k).value for k in range(9)] == ["L", "F1", "E", "F1", "F1", "F1", "F1", "F1", "V"]
    sizes = [ws.cell(slot.row, 15 + k).font.sz for k in range(9)]
    assert sizes == [11, 9, 11, 9, 9, 9, 9, 9, 11]                    # only the F1 boxes are size 9
    assert ws.cell(slot.row, 14).font.sz == 11                          # the name is untouched
    s.undo(rc)
    ws = openpyxl.load_workbook(path)["ENTRY PG1"]
    assert all(ws.cell(slot.row, 14 + k).value is None for k in range(10))
    assert [ws.cell(slot.row, 15 + k).font.sz for k in range(9)] == [11] * 9


def test_missing_trailing_picks_count_as_empty_and_marker_can_be_switched_off(tmp_path):
    p = make_pool_workbook(tmp_path / "w.xlsx")
    s = ExcelSink(OpenpyxlBackend(str(p)))
    slot = s.next_slot()
    s.write(slot, "SHORT", ["L", "E"])                                  # only 2 picks supplied
    ws = openpyxl.load_workbook(p)["ENTRY PG1"]
    assert [ws.cell(slot.row, 15 + k).value for k in range(9)] == ["L", "E"] + ["F1"] * 7
    off = ExcelSink(OpenpyxlBackend(str(p)), blank_marker="")
    slot2 = off.next_slot()
    off.write(slot2, "PLAIN", ["L", ""] + [""] * 7)
    ws = openpyxl.load_workbook(p)["ENTRY PG1"]
    assert [ws.cell(slot2.row, 15 + k).value for k in range(9)] == ["L"] + [None] * 8
    assert ws.cell(slot2.row, 16).font.sz == 11


def test_refuses_to_overwrite_then_allows(sink):
    s, _ = sink
    taken = s.slot_for_ticket(241)
    with pytest.raises(SlotTaken):
        s.write(taken, "NOPE", list("LLLLLLLLL"))
    s.write(taken, "FORCED", list("LLLLLLLLL"), allow_overwrite=True)


def test_undo_restores_empty(sink):
    s, path = sink
    slot = s.next_slot()
    rc = s.write(slot, "TEMP", list("EEEEEEEEE"))
    s.undo(rc)
    ws = openpyxl.load_workbook(path)["ENTRY PG1"]
    assert all(ws.cell(slot.row, 14 + k).value is None for k in range(10))
    assert s.next_slot().ticket == slot.ticket


def test_name_never_becomes_formula_or_date(sink):
    s, path = sink
    slot = s.next_slot()
    s.write(slot, "=1+1", list("LLLLLLLLL"))
    ws = openpyxl.load_workbook(path)["ENTRY PG1"]
    assert ws.cell(slot.row, 14).value == "=1+1" and ws.cell(slot.row, 14).data_type == "s"
    assert needs_text_prefix("1-2-3") and needs_text_prefix("=SUM(A1)") and not needs_text_prefix("1-2-3 RRR")


def test_format_name_tag_after_before_and_off(sink):
    s, _ = sink
    assert s.format_name("clemente  o.") == "CLEMENTE O."                          # no tag
    assert s.format_name("clemente  o.", "rrr") == "CLEMENTE O. RRR"                # default: after
    assert s.format_name("clemente o.", "rrr", "before") == "RRR CLEMENTE O."
    assert s.format_name("ana", "E", "before", sep="-") == "E-ANA"


def test_tag_is_not_doubled_but_single_letters_are_not_confused_with_name_letters(sink):
    s, _ = sink
    assert s.format_name("TONY 5000 RRR", "RRR") == "TONY 5000 RRR"                 # slip already has it
    assert s.format_name("RRR TONY", "RRR", "before") == "RRR TONY"
    assert s.format_name("TONY RRR", "RRR", "before") == "RRR TONY RRR"             # other end: add it
    # "E" is a letter OF these names, not a separate tag word -> must still be added
    assert s.format_name("JOSE", "E") == "JOSE E"
    assert s.format_name("ELENA", "E", "before") == "E ELENA"
    assert s.format_name("ANA E", "E") == "ANA E"                                   # already tagged


def test_full_block_is_reported(tmp_path):
    p = make_pool_workbook(tmp_path / "w.xlsx", filled=(52, 0, 0))
    s = ExcelSink(OpenpyxlBackend(str(p)))
    infos = s.block_infos()
    assert infos[0].free == 0 and infos[0].first_free is None
    assert s.default_block().block.name_col == 26                    # first block with room


def test_missing_header_is_a_clear_error(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "ENTRY PG1"
    wb.active["A1"] = "nothing here"
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ExcelError, match="header"):
        ExcelSink(OpenpyxlBackend(str(tmp_path / "bad.xlsx")))


def test_block_without_ticket_column():
    vals = [["NOMBRE"] + list(range(1, 10))] + [[None] * 10 for _ in range(5)]
    blocks = discover_blocks("S", Snapshot(1, 1, vals), rows_without_tickets=5)
    assert len(blocks) == 1 and blocks[0].ticket_col is None and blocks[0].rows == [2, 3, 4, 5, 6]


# ------------------------------------------------------------------ fake Excel COM object model
# (the instrumented fake lives in test_excel_fast.py and is shared)
def test_com_backend_against_fake_excel_object_model(tmp_path):
    app, sheet = pool_app(tmp_path)
    cells = sheet.cells
    app.Workbooks.books.append(FakeBook("OTHER.xlsx", [FakeSheet("ENTRY PG1", {}, app.meter)], app.meter))
    be = ComBackend(app=app)
    assert be.workbook_name == "JORNADA 10 PC.xlsx" and "live" in be.label
    sink = ExcelSink(be)
    slot = sink.next_slot()
    assert slot.ticket == 251
    sink.write(slot, "CLEMENTE O. RRR", list("LVLLLLLLE"))
    assert cells[(slot.row, 14)] == "CLEMENTE O. RRR"
    assert [cells[(slot.row, 15 + k)] for k in range(9)] == list("LVLLLLLLE")
    assert sink.next_slot().ticket == 252
    # an empty box: "F1" in font size 9 through Range.Value / Font.Size, restored by undo
    slot2 = sink.next_slot()
    rc = sink.write(slot2, "ANA", ["L", "", "E", "V", "V", "V", "V", "V", "V"])
    assert cells[(slot2.row, 16)] == "F1"
    assert sheet.fonts[(slot2.row, 16)] == 9
    assert sheet.fonts.get((slot2.row, 15), 11) == 11                    # normal cells keep their size
    sink.undo(rc)
    assert (slot2.row, 16) not in cells and sheet.fonts[(slot2.row, 16)] == 11
    assert ComBackend(workbook_hint="other", app=app).workbook_name == "OTHER.xlsx"
    with pytest.raises(ExcelError):
        ComBackend(workbook_hint="nope", app=app)
