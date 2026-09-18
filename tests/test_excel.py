from __future__ import annotations

import pytest

from office_docs_mcp.excel.reader import excel_get_metadata, excel_read_sheet
from office_docs_mcp.excel.writer import (
    excel_append_rows,
    excel_create_workbook,
    excel_write_cell,
    excel_write_range,
)


def test_excel_create_workbook(tmp_path):
    f = tmp_path / "new_book.xlsx"
    res = excel_create_workbook(str(f), sheet_names=["Summary", "Data"])
    assert "Created" in res
    assert f.exists()

    meta = excel_get_metadata(str(f))
    assert meta["sheet_names"] == ["Summary", "Data"]
    assert meta["active_sheet"] == "Summary"


def test_excel_write_and_read_cell(tmp_path):
    f = tmp_path / "cell_test.xlsx"
    excel_create_workbook(str(f), sheet_names=["Sheet1"])

    res_write = excel_write_cell(str(f), "Sheet1", "B2", "Hello openpyxl")
    assert "Written" in res_write

    # Read as markdown
    content = excel_read_sheet(str(f), "Sheet1", start_row=1, end_row=3, start_col=1, end_col=3)
    assert "Hello openpyxl" in content


def test_excel_write_range_and_read_raw(tmp_path):
    f = tmp_path / "range_test.xlsx"
    excel_create_workbook(str(f), sheet_names=["Scores"])

    data = [
        ["Name", "Score", "Grade"],
        ["Alice", 95, "A"],
        ["Bob", 80, "B"],
    ]
    excel_write_range(str(f), "Scores", "A1", data)

    # Read as raw
    raw_data = excel_read_sheet(
        str(f), "Scores", start_row=1, end_row=3, start_col=1, end_col=3, format="raw"
    )
    assert raw_data == data


def test_excel_append_rows(tmp_path):
    f = tmp_path / "append_test.xlsx"
    excel_create_workbook(str(f), sheet_names=["Log"])
    excel_write_range(str(f), "Log", "A1", [["ID", "Message"]])

    excel_append_rows(str(f), "Log", [[1, "First"], [2, "Second"]])

    meta = excel_get_metadata(str(f))
    sheet_info = next(s for s in meta["sheets"] if s["name"] == "Log")
    assert sheet_info["max_row"] == 3

    raw_data = excel_read_sheet(str(f), "Log", start_row=1, end_row=3, start_col=1, end_col=2, format="raw")
    assert len(raw_data) == 3
    assert raw_data[2] == [2, "Second"]


def test_excel_read_sheet_invalid_name(tmp_path):
    f = tmp_path / "sample.xlsx"
    excel_create_workbook(str(f))
    with pytest.raises(ValueError, match="Sheet 'NonExistent' not found"):
        excel_read_sheet(str(f), sheet_name="NonExistent")


def test_excel_formula_handling(tmp_path):
    f = tmp_path / "formula.xlsx"
    excel_create_workbook(str(f), sheet_names=["Calc"])
    excel_write_cell(str(f), "Calc", "A1", 10)
    excel_write_cell(str(f), "Calc", "A2", 20)
    excel_write_cell(str(f), "Calc", "A3", "=SUM(A1:A2)")

    # Read with evaluate_formulas=False should show the formula string
    formula_view = excel_read_sheet(
        str(f), "Calc", start_row=3, end_row=3, start_col=1, end_col=1, evaluate_formulas=False, format="raw"
    )
    assert formula_view == [["=SUM(A1:A2)"]]
