from __future__ import annotations

from typing import Any

from mcp.server.mcpserver import MCPServer

from office_docs_mcp.excel.reader import (
    excel_get_metadata as _excel_get_metadata,
)
from office_docs_mcp.excel.reader import (
    excel_read_sheet as _excel_read_sheet,
)
from office_docs_mcp.excel.reader import (
    excel_search as _excel_search,
)
from office_docs_mcp.excel.writer import (
    excel_append_rows as _excel_append_rows,
)
from office_docs_mcp.excel.writer import (
    excel_create_workbook as _excel_create_workbook,
)
from office_docs_mcp.excel.writer import (
    excel_manage_sheets as _excel_manage_sheets,
)
from office_docs_mcp.excel.writer import (
    excel_write_cell as _excel_write_cell,
)
from office_docs_mcp.excel.writer import (
    excel_write_range as _excel_write_range,
)
from office_docs_mcp.powerpoint.reader import (
    ppt_get_outline as _ppt_get_outline,
)
from office_docs_mcp.powerpoint.reader import (
    ppt_read_notes as _ppt_read_notes,
)
from office_docs_mcp.powerpoint.reader import (
    ppt_read_slide as _ppt_read_slide,
)
from office_docs_mcp.powerpoint.reader import (
    ppt_search as _ppt_search,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_add_slide as _ppt_add_slide,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_add_table as _ppt_add_table,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_create_presentation as _ppt_create_presentation,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_update_notes as _ppt_update_notes,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_update_slide_text as _ppt_update_slide_text,
)
from office_docs_mcp.word.reader import (
    word_get_outline as _word_get_outline,
)
from office_docs_mcp.word.reader import (
    word_read_paragraphs as _word_read_paragraphs,
)
from office_docs_mcp.word.reader import (
    word_read_table as _word_read_table,
)
from office_docs_mcp.word.reader import (
    word_search as _word_search,
)
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
    word_delete_paragraph as _word_delete_paragraph,
)
from office_docs_mcp.word.writer import (
    word_replace_text as _word_replace_text,
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
def excel_search(
    file_path: str,
    query: str,
    sheet_name: str | None = None,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> list[dict[str, Any]]:
    """Search for a text query across Excel worksheet cells.

    Args:
        file_path: Absolute or relative path to the Excel file.
        query: Text to search for.
        sheet_name: Specific sheet name to search in (searches all sheets if omitted).
        case_sensitive: Whether matching is case-sensitive (default: False).
        max_results: Maximum results to return (default: 50).
    """
    return _excel_search(
        file_path=file_path,
        query=query,
        sheet_name=sheet_name,
        case_sensitive=case_sensitive,
        max_results=max_results,
    )


@mcp.tool()
def excel_write_cell(
    file_path: str,
    sheet_name: str,
    coordinate: str,
    value: Any,
    backup: bool = False,
) -> str:
    """Write a value to a specific cell (e.g. 'A1', 'C10') in an Excel sheet.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name.
        coordinate: Cell coordinate (e.g., 'A1', 'B5').
        value: The value to write (string, number, boolean, or formula like '=SUM(A1:A5)').
        backup: If True, create a backup of the original file before modifying.
    """
    return _excel_write_cell(
        file_path=file_path,
        sheet_name=sheet_name,
        coordinate=coordinate,
        value=value,
        backup=backup,
    )


@mcp.tool()
def excel_write_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    data: list[list[Any]],
    backup: bool = False,
) -> str:
    """Write a 2D matrix of data starting from a designated top-left cell in an Excel sheet.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name.
        start_cell: Top-left cell coordinate (e.g. 'A1').
        data: 2D array of rows and column values to write.
        backup: If True, create a backup of the original file before modifying.
    """
    return _excel_write_range(
        file_path=file_path,
        sheet_name=sheet_name,
        start_cell=start_cell,
        data=data,
        backup=backup,
    )


@mcp.tool()
def excel_append_rows(
    file_path: str,
    sheet_name: str,
    rows: list[list[Any]],
    backup: bool = False,
) -> str:
    """Append one or more rows to the end of an Excel sheet.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Target sheet name.
        rows: List of rows (each row being a list of cell values).
        backup: If True, create a backup of the original file before modifying.
    """
    return _excel_append_rows(file_path=file_path, sheet_name=sheet_name, rows=rows, backup=backup)


@mcp.tool()
def excel_manage_sheets(
    file_path: str,
    action: str,
    sheet_name: str,
    new_name: str | None = None,
    backup: bool = False,
) -> str:
    """Manage sheets in an Excel workbook (add, rename, delete, copy).

    Args:
        file_path: Path to the Excel file.
        action: Action to perform: 'add', 'rename', 'delete', or 'copy'.
        sheet_name: Target sheet name.
        new_name: New name for the sheet (required for 'rename' and 'copy').
        backup: If True, create a backup of the original file before modifying.
    """
    return _excel_manage_sheets(
        file_path=file_path,
        action=action,
        sheet_name=sheet_name,
        new_name=new_name,
        backup=backup,
    )


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
        case_sensitive: Whether matching is case-sensitive (default: False).
        max_results: Maximum results to return (default: 50).
    """
    return _word_search(
        file_path=file_path,
        query=query,
        case_sensitive=case_sensitive,
        max_results=max_results,
    )


@mcp.tool()
def word_append_paragraph(
    file_path: str,
    text: str,
    style: str | None = None,
    backup: bool = False,
) -> str:
    """Append a paragraph or heading to the end of a Word document (.docx).

    Args:
        file_path: Path to the Word document.
        text: Text to append.
        style: Optional style (e.g., 'Heading 1', 'Heading 2', 'List Bullet', or None for normal body).
        backup: If True, create a backup of the original file before modifying.
    """
    return _word_append_paragraph(file_path, text, style=style, backup=backup)


@mcp.tool()
def word_append_table_row(
    file_path: str,
    table_idx: int,
    row_data: list[str],
    backup: bool = False,
) -> str:
    """Append a new row of data to an existing table in a Word document.

    Args:
        file_path: Path to the Word document.
        table_idx: 0-based index of the table.
        row_data: List of strings for each column in the new row.
        backup: If True, create a backup of the original file before modifying.
    """
    return _word_append_table_row(file_path, table_idx, row_data, backup=backup)


@mcp.tool()
def word_write_table_cell(
    file_path: str,
    table_idx: int,
    row_idx: int,
    col_idx: int,
    text: str,
    backup: bool = False,
) -> str:
    """Update text in a specific cell of a table in a Word document.

    Args:
        file_path: Path to the Word document.
        table_idx: 0-based index of the table.
        row_idx: 0-based row index.
        col_idx: 0-based column index.
        text: New text content for the cell.
        backup: If True, create a backup of the original file before modifying.
    """
    return _word_write_table_cell(
        file_path=file_path,
        table_idx=table_idx,
        row_idx=row_idx,
        col_idx=col_idx,
        text=text,
        backup=backup,
    )


@mcp.tool()
def word_replace_text(
    file_path: str,
    find_text: str,
    replace_text: str,
    count: int = -1,
    backup: bool = False,
) -> str:
    """Find and replace text across paragraphs and tables in a Word document.

    Args:
        file_path: Path to the Word document (.docx).
        find_text: Text string to find.
        replace_text: Replacement text string.
        count: Maximum number of replacements (-1 for all occurrences).
        backup: If True, create a backup of the original file before modifying.
    """
    return _word_replace_text(
        file_path=file_path,
        find_text=find_text,
        replace_text=replace_text,
        count=count,
        backup=backup,
    )


@mcp.tool()
def word_delete_paragraph(
    file_path: str,
    paragraph_idx: int,
    backup: bool = False,
) -> str:
    """Delete a specific paragraph from a Word document by its 0-based index.

    Args:
        file_path: Path to the Word document (.docx).
        paragraph_idx: 0-based index of the paragraph to delete.
        backup: If True, create a backup of the original file before modifying.
    """
    return _word_delete_paragraph(
        file_path=file_path,
        paragraph_idx=paragraph_idx,
        backup=backup,
    )


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
def ppt_search(
    file_path: str,
    query: str,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> list[dict[str, Any]]:
    """Search for a text query across PowerPoint slides, shapes, tables, and notes.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        query: Text to search for.
        case_sensitive: Whether matching is case-sensitive (default: False).
        max_results: Maximum results to return (default: 50).
    """
    return _ppt_search(
        file_path=file_path,
        query=query,
        case_sensitive=case_sensitive,
        max_results=max_results,
    )


@mcp.tool()
def ppt_read_notes(file_path: str, slide_idx: int) -> str:
    """Read the speaker notes of a specific slide in a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        slide_idx: 0-based index of the slide.
    """
    return _ppt_read_notes(file_path=file_path, slide_idx=slide_idx)


@mcp.tool()
def ppt_update_notes(
    file_path: str,
    slide_idx: int,
    notes_text: str,
    backup: bool = False,
) -> str:
    """Update speaker notes for a specific slide in a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        slide_idx: 0-based slide index.
        notes_text: Text content to set as the speaker notes.
        backup: If True, create a backup of the original file before modifying.
    """
    return _ppt_update_notes(
        file_path=file_path,
        slide_idx=slide_idx,
        notes_text=notes_text,
        backup=backup,
    )


@mcp.tool()
def ppt_add_slide(
    file_path: str,
    title: str,
    content: str | None = None,
    layout_idx: int = 1,
    backup: bool = False,
) -> str:
    """Add a new slide to a PowerPoint presentation (.pptx).

    Args:
        file_path: Path to the PowerPoint file.
        title: Title for the slide.
        content: Optional content text for the body placeholder.
        layout_idx: Slide layout index (0=Title Slide, 1=Title and Content). Default 1.
        backup: If True, create a backup of the original file before modifying.
    """
    return _ppt_add_slide(
        file_path,
        title,
        content=content,
        layout_idx=layout_idx,
        backup=backup,
    )


@mcp.tool()
def ppt_update_slide_text(
    file_path: str,
    slide_idx: int,
    shape_idx: int,
    text: str,
    backup: bool = False,
) -> str:
    """Update the text of a specific shape on a slide in a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file.
        slide_idx: 0-based slide index.
        shape_idx: 0-based shape index within the slide.
        text: New text content for the shape.
        backup: If True, create a backup of the original file before modifying.
    """
    return _ppt_update_slide_text(
        file_path,
        slide_idx,
        shape_idx,
        text,
        backup=backup,
    )


@mcp.tool()
def ppt_add_table(
    file_path: str,
    slide_idx: int,
    rows: int,
    cols: int,
    data: list[list[str]] | None = None,
    left: float = 1.0,
    top: float = 2.0,
    width: float = 8.0,
    height: float = 4.0,
    backup: bool = False,
) -> str:
    """Add a table to a specific slide in a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        slide_idx: 0-based index of the slide.
        rows: Number of table rows.
        cols: Number of table columns.
        data: Optional 2D list of cell text values.
        left: Left offset position in inches (default: 1.0).
        top: Top offset position in inches (default: 2.0).
        width: Table width in inches (default: 8.0).
        height: Table height in inches (default: 4.0).
        backup: If True, create a backup of the original file before modifying.
    """
    return _ppt_add_table(
        file_path=file_path,
        slide_idx=slide_idx,
        rows=rows,
        cols=cols,
        data=data,
        left=left,
        top=top,
        width=width,
        height=height,
        backup=backup,
    )


@mcp.tool()
def ppt_create_presentation(file_path: str, title: str | None = None) -> str:
    """Create a new empty PowerPoint presentation (.pptx) with an optional title slide.

    Args:
        file_path: Path where the presentation will be saved.
        title: Optional title for the first slide.
    """
    return _ppt_create_presentation(file_path, title)


# ==============================================================================
# MCP Prompts
# ==============================================================================


@mcp.prompt()
def analyze_spreadsheet(file_path: str, focus_area: str | None = None) -> str:
    """Prompt workflow for analyzing an Excel workbook.

    Args:
        file_path: Target Excel file path.
        focus_area: Optional specific domain or metric to focus on.
    """
    focus_str = f" with a focus on '{focus_area}'" if focus_area else ""
    return (
        f"Please perform an in-depth data analysis on the Excel file at '{file_path}'{focus_str}.\n\n"
        "Follow this systematic workflow:\n"
        "1. First, call `excel_get_metadata` to understand available sheets, row/col counts, and headers.\n"
        "2. If needed, use `excel_search` to quickly locate key terms or totals.\n"
        "3. Read necessary data chunks using `excel_read_sheet`.\n"
        "4. Synthesize your findings into a clear executive summary, trend insights, and next actions."
    )


@mcp.prompt()
def create_presentation_outline(topic: str, slide_count: int = 5) -> str:
    """Prompt workflow for structuring and drafting a PowerPoint presentation.

    Args:
        topic: Topic or theme of the presentation.
        slide_count: Target number of slides (default 5).
    """
    return (
        f"Please create a structured {slide_count}-slide presentation on '{topic}'.\n\n"
        "Workflow:\n"
        "1. Plan the narrative structure: slide titles, key takeaway points, and speaker notes.\n"
        "2. Initialize the presentation using `ppt_create_presentation`.\n"
        "3. For each slide, invoke `ppt_add_slide` with concise bullet points.\n"
        "4. Add comprehensive talking points using `ppt_update_notes` for the speaker's script.\n"
        "5. Review with `ppt_get_outline` to ensure clear flow."
    )


def get_server() -> MCPServer:
    """Return the configured MCPServer instance."""
    return mcp
