from __future__ import annotations

from typing import Any

from mcp.server.mcpserver import MCPServer

from office_docs_mcp.excel.reader import excel_get_metadata as _excel_get_metadata
from office_docs_mcp.excel.reader import excel_read_sheet as _excel_read_sheet
from office_docs_mcp.excel.writer import (
    excel_append_rows as _excel_append_rows,
)
from office_docs_mcp.excel.writer import (
    excel_create_workbook as _excel_create_workbook,
)
from office_docs_mcp.excel.writer import (
    excel_write_cell as _excel_write_cell,
)
from office_docs_mcp.excel.writer import (
    excel_write_range as _excel_write_range,
)
from office_docs_mcp.powerpoint.reader import ppt_get_outline as _ppt_get_outline
from office_docs_mcp.powerpoint.reader import ppt_read_slide as _ppt_read_slide
from office_docs_mcp.powerpoint.writer import (
    ppt_add_slide as _ppt_add_slide,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_create_presentation as _ppt_create_presentation,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_update_slide_text as _ppt_update_slide_text,
)
from office_docs_mcp.word.reader import word_get_outline as _word_get_outline
from office_docs_mcp.word.reader import (
    word_read_paragraphs as _word_read_paragraphs,
)
from office_docs_mcp.word.reader import word_read_table as _word_read_table
from office_docs_mcp.word.writer import (
    word_append_paragraph as _word_append_paragraph,
)
from office_docs_mcp.word.writer import (
    word_append_table_row as _word_append_table_row,
)
from office_docs_mcp.word.writer import (
    word_create_document as _word_create_document,
)
from office_docs_mcp.word.writer import (
    word_write_table_cell as _word_write_table_cell,
)

mcp = MCPServer("office-docs-mcp")

# ==============================================================================
# Excel Tools
# ==============================================================================


@mcp.tool()
def excel_get_metadata(file_path: str) -> dict[str, Any]:
    """Get metadata from an Excel workbook (.xlsx/.xlsm).

    Returns sheet names, active sheet, dimensions (max rows/cols), and header preview.
    Call this first before reading large sheets to inspect the structure.

    Args:
        file_path: Absolute or relative path to the Excel file.
    """
    return _excel_get_metadata(file_path)


@mcp.tool()
def excel_read_sheet(
    file_path: str,
    sheet_name: str | None = None,
    start_row: int = 1,
    end_row: int = 50,
    start_col: int = 1,
    end_col: int = 20,
    evaluate_formulas: bool = True,
    format: str = "markdown",
) -> str | list[list[Any]]:
    """Read a rectangular cell range from an Excel sheet.

    Indices are 1-based (row 1 is first row, col 1 is 'A').

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name (defaults to active sheet).
        start_row: Starting row index (1-based, inclusive, default 1).
        end_row: Ending row index (1-based, inclusive, default 50).
        start_col: Starting column index (1-based, inclusive, default 1).
        end_col: Ending column index (1-based, inclusive, default 20).
        evaluate_formulas: If True (default), returns calculated values. If False, returns raw formula strings.
        format: 'markdown' (default formatted table) or 'raw' (2D list of values).
    """
    return _excel_read_sheet(
        file_path=file_path,
        sheet_name=sheet_name,
        start_row=start_row,
        end_row=end_row,
        start_col=start_col,
        end_col=end_col,
        evaluate_formulas=evaluate_formulas,
        format=format,
    )


@mcp.tool()
def excel_write_cell(file_path: str, sheet_name: str, coordinate: str, value: Any) -> str:
    """Write a value to a specific cell (e.g. 'A1', 'C10') in an Excel sheet.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name.
        coordinate: Cell coordinate (e.g., 'A1', 'B5').
        value: The value to write (string, number, boolean, or formula like '=SUM(A1:A5)').
    """
    return _excel_write_cell(file_path, sheet_name, coordinate, value)


@mcp.tool()
def excel_write_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    data: list[list[Any]],
) -> str:
    """Write a 2D matrix of data starting from a designated top-left cell in an Excel sheet.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name.
        start_cell: Top-left cell coordinate (e.g. 'A1').
        data: 2D array of rows and column values to write.
    """
    return _excel_write_range(file_path, sheet_name, start_cell, data)


@mcp.tool()
def excel_append_rows(file_path: str, sheet_name: str, rows: list[list[Any]]) -> str:
    """Append one or more rows to the end of an Excel sheet.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name.
        rows: List of rows (each row being a list of cell values).
    """
    return _excel_append_rows(file_path, sheet_name, rows)


@mcp.tool()
def excel_create_workbook(file_path: str, sheet_names: list[str] | None = None) -> str:
    """Create a new empty Excel workbook (.xlsx).

    Args:
        file_path: Path where the new Excel file will be created.
        sheet_names: Optional list of sheet names to initialize. Defaults to ['Sheet'].
    """
    return _excel_create_workbook(file_path, sheet_names)


# ==============================================================================
# Word Tools
# ==============================================================================


@mcp.tool()
def word_get_outline(file_path: str) -> dict[str, Any]:
    """Get the outline of a Word document (.docx).

    Returns total paragraph count, table count, and all headings with their levels.

    Args:
        file_path: Path to the Word document (.docx).
    """
    return _word_get_outline(file_path)


@mcp.tool()
def word_read_paragraphs(
    file_path: str,
    start_idx: int = 0,
    count: int = 30,
) -> list[dict[str, Any]]:
    """Read a chunk of paragraphs from a Word document (.docx).

    Indices are 0-based.

    Args:
        file_path: Path to the Word document.
        start_idx: Starting paragraph index (0-based, default 0).
        count: Number of paragraphs to retrieve (default 30).
    """
    return _word_read_paragraphs(file_path, start_idx=start_idx, count=count)


@mcp.tool()
def word_read_table(
    file_path: str,
    table_idx: int = 0,
    format: str = "markdown",
) -> str | list[list[str]]:
    """Read a table from a Word document (.docx).

    Args:
        file_path: Path to the Word document.
        table_idx: 0-based index of the table (default 0).
        format: 'markdown' (default) or 'raw' (2D list of cell text).
    """
    return _word_read_table(file_path, table_idx=table_idx, format=format)


@mcp.tool()
def word_append_paragraph(
    file_path: str,
    text: str,
    style: str | None = None,
) -> str:
    """Append a paragraph or heading to the end of a Word document (.docx).

    Args:
        file_path: Path to the Word document.
        text: Text to append.
        style: Optional style (e.g., 'Heading 1', 'Heading 2', 'List Bullet', or None for normal body).
    """
    return _word_append_paragraph(file_path, text, style=style)


@mcp.tool()
def word_append_table_row(
    file_path: str,
    table_idx: int,
    row_data: list[str],
) -> str:
    """Append a new row of data to an existing table in a Word document.

    Args:
        file_path: Path to the Word document.
        table_idx: 0-based index of the table.
        row_data: List of strings for each column in the new row.
    """
    return _word_append_table_row(file_path, table_idx, row_data)


@mcp.tool()
def word_write_table_cell(
    file_path: str,
    table_idx: int,
    row_idx: int,
    col_idx: int,
    text: str,
) -> str:
    """Update text in a specific cell of a table in a Word document.

    Args:
        file_path: Path to the Word document.
        table_idx: 0-based index of the table.
        row_idx: 0-based row index.
        col_idx: 0-based column index.
        text: New text content for the cell.
    """
    return _word_write_table_cell(file_path, table_idx, row_idx, col_idx, text)


@mcp.tool()
def word_create_document(file_path: str, title: str | None = None) -> str:
    """Create a new empty Word document (.docx) with an optional title.

    Args:
        file_path: Path where the document will be saved.
        title: Optional title text.
    """
    return _word_create_document(file_path, title)


# ==============================================================================
# PowerPoint Tools
# ==============================================================================


@mcp.tool()
def ppt_get_outline(file_path: str) -> dict[str, Any]:
    """Get the outline of a PowerPoint presentation (.pptx).

    Returns total slide count and slide-by-slide titles and shape counts.

    Args:
        file_path: Path to the PowerPoint file.
    """
    return _ppt_get_outline(file_path)


@mcp.tool()
def ppt_read_slide(file_path: str, slide_idx: int) -> dict[str, Any]:
    """Read all text content and table data from a specific slide in a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file.
        slide_idx: 0-based index of the slide.
    """
    return _ppt_read_slide(file_path, slide_idx)


@mcp.tool()
def ppt_add_slide(
    file_path: str,
    title: str,
    content: str | None = None,
    layout_idx: int = 1,
) -> str:
    """Add a new slide to a PowerPoint presentation (.pptx).

    Args:
        file_path: Path to the PowerPoint file.
        title: Title for the slide.
        content: Optional content text for the body placeholder.
        layout_idx: Slide layout index (0=Title Slide, 1=Title and Content). Default 1.
    """
    return _ppt_add_slide(file_path, title, content=content, layout_idx=layout_idx)


@mcp.tool()
def ppt_update_slide_text(
    file_path: str,
    slide_idx: int,
    shape_idx: int,
    text: str,
) -> str:
    """Update the text of a specific shape on a slide in a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file.
        slide_idx: 0-based slide index.
        shape_idx: 0-based shape index within the slide.
        text: New text content for the shape.
    """
    return _ppt_update_slide_text(file_path, slide_idx, shape_idx, text)


@mcp.tool()
def ppt_create_presentation(file_path: str, title: str | None = None) -> str:
    """Create a new empty PowerPoint presentation (.pptx) with an optional title slide.

    Args:
        file_path: Path where the presentation will be saved.
        title: Optional title for the first slide.
    """
    return _ppt_create_presentation(file_path, title)


def get_server() -> MCPServer:
    """Return the configured MCPServer instance."""
    return mcp
