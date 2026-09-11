from types import SimpleNamespace

from app.core.config import Settings
from app.ocr import docai


def _seg(start, end):
    return SimpleNamespace(start_index=start, end_index=end)


def _cell(start, end):
    return SimpleNamespace(
        layout=SimpleNamespace(text_anchor=SimpleNamespace(text_segments=[_seg(start, end)]))
    )


def _row(spans):
    return SimpleNamespace(cells=[_cell(a, b) for a, b in spans])


def _fake_document():
    # full_text laid out so offsets slice cleanly
    text = "ItemAmountOwn damage8500"
    table = SimpleNamespace(
        header_rows=[_row([(0, 4), (4, 10)])],           # Item | Amount
        body_rows=[_row([(10, 20), (20, 24)])],          # Own damage | 8500
    )
    return SimpleNamespace(text=text, pages=[SimpleNamespace(tables=[table])])


def test_parse_tables_reads_header_and_rows():
    tables = docai.parse_tables(_fake_document())
    assert len(tables) == 1
    t = tables[0]
    assert t.header == ["Item", "Amount"]
    assert t.rows == [["Item", "Amount"], ["Own damage", "8500"]]


def test_parse_tables_empty_document():
    empty = SimpleNamespace(text="", pages=[SimpleNamespace(tables=[])])
    assert docai.parse_tables(empty) == []


def test_is_configured():
    assert not docai.is_configured(Settings())
    assert not docai.is_configured(Settings(docai_project="p"))  # processor_id missing
    assert docai.is_configured(Settings(docai_project="p", docai_processor_id="x"))
