"""원본 XML에서 총자산 칸만 바꾸는 저장 로직"""
from nav_total_assets.workbook import set_cell_number


def test_fill_empty_styled_cell_keeps_style():
    xml = '<sheetData><row r="2"><c r="V2" s="1"/><c r="W2" s="55"/><c r="X2" s="54"/></row></sheetData>'
    out = set_cell_number(xml, "W2", 123)
    assert '<c r="W2" s="55"><v>123</v></c>' in out and '<c r="X2" s="54"/>' in out


def test_replace_shared_string_drops_type():
    xml = '<sheetData><row r="57"><c r="W57" s="56" t="s"><v>383</v></c></row></sheetData>'
    assert set_cell_number(xml, "W57", 9) == '<sheetData><row r="57"><c r="W57" s="56"><v>9</v></c></row></sheetData>'


def test_insert_missing_cell_in_column_order():
    xml = '<sheetData><row r="3"><c r="B3"><v>1</v></c><c r="Z3"><v>2</v></c></row></sheetData>'
    out = set_cell_number(xml, "W3", 5)
    assert out.index('r="B3"') < out.index('r="W3"') < out.index('r="Z3"')


def test_self_closing_and_missing_rows():
    xml = '<sheetData><row r="2"/><row r="9"><c r="A9"/></row></sheetData>'
    out = set_cell_number(xml, "W2", 1)
    assert '<row r="2"><c r="W2"><v>1</v></c></row>' in out
    out = set_cell_number(out, "W5", 2)
    assert out.index('r="2"') < out.index('<row r="5">') < out.index('r="9"')


def test_does_not_touch_similar_refs():
    xml = '<sheetData><row r="1"><c r="W1"/><c r="W10"/></row><row r="10"><c r="W10" s="3"/></row></sheetData>'
    out = set_cell_number(xml, "W10", 7)
    assert '<row r="1"><c r="W1"/><c r="W10"/></row>' in out and '<c r="W10" s="3"><v>7</v></c>' in out
