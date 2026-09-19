from __future__ import annotations

import pytest

from office_docs_mcp.word.reader import (
    word_get_outline,
    word_read_paragraphs,
    word_read_table,
)
from office_docs_mcp.word.writer import (
    word_append_paragraph,
    word_append_table_row,
    word_create_document,
    word_write_table_cell,
)


def test_word_create_and_outline(tmp_path):
    f = tmp_path / "sample.docx"
    res = word_create_document(str(f), title="Project Overview")
    assert "Created" in res
    assert f.exists()

    word_append_paragraph(str(f), "Section 1: Background", style="Heading 1")
    word_append_paragraph(str(f), "This is the first paragraph describing the project.")

    outline = word_get_outline(str(f))
    assert outline["total_paragraphs"] >= 3
    assert any(h["text"] == "Project Overview" for h in outline["headings"])
    assert any(h["text"] == "Section 1: Background" for h in outline["headings"])


def test_word_read_paragraphs_slicing(tmp_path):
    f = tmp_path / "paragraphs.docx"
    word_create_document(str(f))
    for i in range(10):
        word_append_paragraph(str(f), f"Paragraph number {i}")

    # Read paragraphs 2 to 5 (count=4)
    paras = word_read_paragraphs(str(f), start_idx=2, count=4)
    assert len(paras) == 4
    assert paras[0]["text"] == "Paragraph number 2"
    assert paras[3]["text"] == "Paragraph number 5"


def test_word_table_operations(tmp_path):
    f = tmp_path / "table_doc.docx"
    word_create_document(str(f))

    # Add paragraph
    word_append_paragraph(str(f), "Table Section", style="Heading 1")

    # Add a table with initial row
    from docx import Document

    doc = Document(str(f))
    tbl = doc.add_table(rows=1, cols=2)
    tbl.rows[0].cells[0].text = "Item"
    tbl.rows[0].cells[1].text = "Price"
    doc.save(str(f))

    # Append row
    word_append_table_row(str(f), table_idx=0, row_data=["Apple", "$1.00"])

    # Check outline table count
    outline = word_get_outline(str(f))
    assert outline["total_tables"] == 1

    # Read table as markdown
    tbl_md = word_read_table(str(f), table_idx=0, format="markdown")
    assert "Item" in tbl_md
    assert "Apple" in tbl_md
    assert "$1.00" in tbl_md

    # Update cell
    word_write_table_cell(str(f), table_idx=0, row_idx=1, col_idx=1, text="$1.50")
    tbl_raw = word_read_table(str(f), table_idx=0, format="raw")
    assert tbl_raw[1][1] == "$1.50"


def test_word_invalid_table_index(tmp_path):
    f = tmp_path / "empty.docx"
    word_create_document(str(f))
    with pytest.raises(IndexError, match="Table index 0 out of range"):
        word_read_table(str(f), table_idx=0)


def test_word_invalid_style(tmp_path):
    f = tmp_path / "style_doc.docx"
    word_create_document(str(f))
    with pytest.raises(ValueError, match="Style 'NonExistentStyle' not found"):
        word_append_paragraph(str(f), "Text", style="NonExistentStyle")
