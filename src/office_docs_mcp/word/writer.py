from __future__ import annotations

from docx import Document

from office_docs_mcp.common.file_utils import ensure_parent_dir, validate_file_path

VALID_WORD_EXTS = (".docx",)


def word_create_document(file_path: str, title: str | None = None) -> str:
    """Create a new empty Word document with an optional title.

    Args:
        file_path: Path where the document will be saved (.docx).
        title: Optional document title to add as the first heading.

    Returns:
        Status confirmation message.
    """
    path = ensure_parent_dir(file_path)
    if path.suffix.lower() not in VALID_WORD_EXTS:
        path = path.with_suffix(".docx")

    doc = Document()
    if title:
        doc.add_heading(title, level=0)

    doc.save(str(path))
    return f"Created Word document at {path}"


def word_append_paragraph(
    file_path: str,
    text: str,
    style: str | None = None,
) -> str:
    """Append a paragraph or heading to the end of a Word document.

    Args:
        file_path: Path to the Word document (.docx).
        text: Paragraph text content.
        style: Optional style name (e.g., 'Heading 1', 'Heading 2', 'List Bullet').

    Returns:
        Status confirmation message.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))

    if style and style.lower().startswith("heading"):
        # Extract heading level if possible
        try:
            level = int(style.split()[-1])
            doc.add_heading(text, level=level)
        except (ValueError, IndexError):
            doc.add_paragraph(text, style=style)
    else:
        doc.add_paragraph(text, style=style)

    doc.save(str(path))
    return f"Appended paragraph to {path}"


def word_append_table_row(
    file_path: str,
    table_idx: int,
    row_data: list[str],
) -> str:
    """Append a row of cells to an existing table in a Word document.

    Args:
        file_path: Path to the Word document (.docx).
        table_idx: 0-based index of the table.
        row_data: List of strings for each column cell in the new row.

    Returns:
        Status confirmation message.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))

    if table_idx < 0 or table_idx >= len(doc.tables):
        raise IndexError(
            f"Table index {table_idx} out of range. Document has {len(doc.tables)} tables."
        )

    table = doc.tables[table_idx]
    row = table.add_row()
    for c_idx, val in enumerate(row_data):
        if c_idx < len(row.cells):
            row.cells[c_idx].text = str(val)

    doc.save(str(path))
    return f"Appended row to table {table_idx} in {path}"


def word_write_table_cell(
    file_path: str,
    table_idx: int,
    row_idx: int,
    col_idx: int,
    text: str,
) -> str:
    """Update text in a specific cell of a table in a Word document.

    Args:
        file_path: Path to the Word document (.docx).
        table_idx: 0-based index of the table.
        row_idx: 0-based row index.
        col_idx: 0-based column index.
        text: New text content for the cell.

    Returns:
        Status confirmation message.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))

    if table_idx < 0 or table_idx >= len(doc.tables):
        raise IndexError(
            f"Table index {table_idx} out of range. Document has {len(doc.tables)} tables."
        )

    table = doc.tables[table_idx]
    if row_idx < 0 or row_idx >= len(table.rows):
        raise IndexError(f"Row index {row_idx} out of range in table {table_idx}.")

    row = table.rows[row_idx]
    if col_idx < 0 or col_idx >= len(row.cells):
        raise IndexError(f"Column index {col_idx} out of range in table {table_idx}.")

    row.cells[col_idx].text = text
    doc.save(str(path))
    return f"Updated cell [{row_idx}, {col_idx}] of table {table_idx} in {path}"
