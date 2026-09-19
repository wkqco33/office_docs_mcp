from __future__ import annotations

from typing import Any

from docx import Document

from office_docs_mcp.common.file_utils import format_as_markdown_table, validate_file_path

VALID_WORD_EXTS = (".docx",)


def word_get_outline(file_path: str) -> dict[str, Any]:
    """Get the outline of a Word document including headings and structure.

    Args:
        file_path: Path to the Word document (.docx).

    Returns:
        Dictionary containing paragraph count, table count, and list of headings.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))

    headings = []
    for idx, p in enumerate(doc.paragraphs):
        style_name = p.style.name if p.style else ""
        text = p.text.strip()
        if not text:
            continue
        if "Heading" in style_name or "Title" in style_name:
            headings.append(
                {
                    "paragraph_index": idx,
                    "level": style_name,
                    "text": text,
                }
            )

    return {
        "file_path": str(path),
        "total_paragraphs": len(doc.paragraphs),
        "total_tables": len(doc.tables),
        "headings": headings,
    }


def word_read_paragraphs(
    file_path: str,
    start_idx: int = 0,
    count: int = 30,
) -> list[dict[str, Any]]:
    """Read a slice of paragraphs from a Word document.

    Indices are 0-based.

    Args:
        file_path: Path to the Word document (.docx).
        start_idx: Starting paragraph index (0-based). Default 0.
        count: Number of paragraphs to read. Default 30.

    Returns:
        List of dictionaries with 'index', 'style', and 'text'.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))

    total = len(doc.paragraphs)
    actual_start = max(0, min(start_idx, total))
    actual_end = min(actual_start + count, total)

    results = []
    for idx in range(actual_start, actual_end):
        p = doc.paragraphs[idx]
        results.append(
            {
                "index": idx,
                "style": p.style.name if p.style else "",
                "text": p.text,
            }
        )

    return results


def word_read_table(
    file_path: str,
    table_idx: int = 0,
    format: str = "markdown",
) -> str | list[list[str]]:
    """Read contents of a specific table in a Word document.

    Args:
        file_path: Path to the Word document (.docx).
        table_idx: 0-based index of the table. Default 0.
        format: Output format, either 'markdown' (default) or 'raw' (list of lists).

    Returns:
        Markdown table string or 2D list of strings.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))

    if table_idx < 0 or table_idx >= len(doc.tables):
        raise IndexError(
            f"Table index {table_idx} out of range. Document has {len(doc.tables)} tables."
        )

    table = doc.tables[table_idx]
    rows_data: list[list[str]] = []
    for row in table.rows:
        row_vals = [cell.text.strip() for cell in row.cells]
        rows_data.append(row_vals)

    if format == "raw":
        return rows_data

    if not rows_data:
        return ""

    headers = rows_data[0]
    data_rows = rows_data[1:] if len(rows_data) > 1 else []
    return format_as_markdown_table(headers, data_rows)


def word_search(
    file_path: str,
    query: str,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> list[dict[str, Any]]:
    """Search for a text query across Word paragraphs and tables.

    Args:
        file_path: Path to the Word document (.docx).
        query: Text to search for.
        case_sensitive: Whether search is case sensitive.
        max_results: Maximum results to return (default 50).

    Returns:
        List of match dictionaries detailing location, text, and snippets.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_WORD_EXTS)
    doc = Document(str(path))
    results: list[dict[str, Any]] = []

    norm_query = query if case_sensitive else query.lower()

    # 1. Search paragraphs
    for p_idx, p in enumerate(doc.paragraphs):
        p_text = p.text
        cmp_text = p_text if case_sensitive else p_text.lower()
        if norm_query in cmp_text:
            results.append(
                {
                    "location_type": "paragraph",
                    "paragraph_idx": p_idx,
                    "style": p.style.name if p.style else None,
                    "text": p_text,
                }
            )
            if len(results) >= max_results:
                return results

    # 2. Search tables
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                c_text = cell.text
                cmp_text = c_text if case_sensitive else c_text.lower()
                if norm_query in cmp_text:
                    results.append(
                        {
                            "location_type": "table",
                            "table_idx": t_idx,
                            "row_idx": r_idx,
                            "col_idx": c_idx,
                            "text": c_text,
                        }
                    )
                    if len(results) >= max_results:
                        return results

    return results
