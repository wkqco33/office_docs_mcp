from __future__ import annotations

from typing import Any

from pptx import Presentation

from office_docs_mcp.common.file_utils import validate_file_path

VALID_PPT_EXTS = (".pptx",)


def ppt_get_outline(file_path: str) -> dict[str, Any]:
    """Get the outline of a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file (.pptx).

    Returns:
        Dictionary with total slide count and list of slides with their titles.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_PPT_EXTS)
    prs = Presentation(str(path))

    slides_summary = []
    for idx, slide in enumerate(prs.slides):
        title_text = ""
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            title_text = slide.shapes.title.text.strip()
        elif len(slide.shapes) > 0 and slide.shapes[0].has_text_frame:
            title_text = slide.shapes[0].text.strip()

        slides_summary.append(
            {
                "slide_index": idx,
                "title": title_text,
                "shapes_count": len(slide.shapes),
            }
        )

    return {
        "file_path": str(path),
        "total_slides": len(prs.slides),
        "slides": slides_summary,
    }


def ppt_read_slide(file_path: str, slide_idx: int) -> dict[str, Any]:
    """Read all text contents and tables from a specific slide.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        slide_idx: 0-based index of the slide.

    Returns:
        Dictionary containing slide details and shapes data.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_PPT_EXTS)
    prs = Presentation(str(path))

    if slide_idx < 0 or slide_idx >= len(prs.slides):
        raise IndexError(
            f"Slide index {slide_idx} out of range. Presentation has {len(prs.slides)} slides."
        )

    slide = prs.slides[slide_idx]
    title_text = ""
    if slide.shapes.title and slide.shapes.title.has_text_frame:
        title_text = slide.shapes.title.text.strip()

    shapes_data = []
    for s_idx, shape in enumerate(slide.shapes):
        shape_info: dict[str, Any] = {
            "shape_index": s_idx,
            "name": shape.name,
            "has_text": shape.has_text_frame,
            "is_table": shape.has_table,
        }

        if shape.has_text_frame:
            shape_info["text"] = shape.text.strip()

        if shape.has_table:
            table_rows = []
            for row in shape.table.rows:
                table_rows.append([c.text.strip() for c in row.cells])
            shape_info["table_data"] = table_rows

        shapes_data.append(shape_info)

    return {
        "slide_index": slide_idx,
        "title": title_text,
        "shapes": shapes_data,
    }
