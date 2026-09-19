from __future__ import annotations

from pptx import Presentation

from office_docs_mcp.common.file_utils import (
    create_backup,
    ensure_parent_dir,
    validate_file_path,
)

VALID_PPT_EXTS = (".pptx",)


def ppt_create_presentation(file_path: str, title: str | None = None) -> str:
    """Create a new PowerPoint presentation with an optional title slide.

    Args:
        file_path: Path where the presentation will be saved (.pptx).
        title: Optional title for the first title slide.

    Returns:
        Status confirmation message.
    """
    path = ensure_parent_dir(file_path)
    if path.suffix.lower() not in VALID_PPT_EXTS:
        path = path.with_suffix(".pptx")

    prs = Presentation()
    if title:
        # Layout 0 is typically the title slide layout
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        if slide.shapes.title:
            slide.shapes.title.text = title

    prs.save(str(path))
    return f"Created PowerPoint presentation at {path}"


def ppt_add_slide(
    file_path: str,
    title: str,
    content: str | None = None,
    layout_idx: int = 1,
    backup: bool = False,
) -> str:
    """Add a new slide to a PowerPoint presentation.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        title: Title of the slide.
        content: Optional body text for the slide content placeholder.
        layout_idx: Slide layout index (0=Title, 1=Title and Content). Default 1.
        backup: If True, create a backup of the original file before modifying.

    Returns:
        Status confirmation message.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_PPT_EXTS)
    bak_msg = ""
    if backup:
        bak = create_backup(path)
        if bak:
            bak_msg = f" (backup created: {bak.name})"

    prs = Presentation(str(path))

    if layout_idx < 0 or layout_idx >= len(prs.slide_layouts):
        layout_idx = 1

    slide_layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(slide_layout)

    if slide.shapes.title:
        slide.shapes.title.text = title

    if content:
        if len(slide.placeholders) > 1:
            body_shape = slide.placeholders[1]
            tf = body_shape.text_frame
            tf.text = content
        else:
            from pptx.util import Inches

            tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
            tb.text_frame.word_wrap = True
            tb.text_frame.text = content

    prs.save(str(path))
    new_idx = len(prs.slides) - 1
    return f"Added slide #{new_idx} with title '{title}' to {path}{bak_msg}"


def ppt_update_slide_text(
    file_path: str,
    slide_idx: int,
    shape_idx: int,
    text: str,
    backup: bool = False,
) -> str:
    """Update the text of a specific shape on a slide.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        slide_idx: 0-based slide index.
        shape_idx: 0-based shape index within the slide.
        text: New text content for the shape.
        backup: If True, create a backup of the original file before modifying.

    Returns:
        Status confirmation message.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_PPT_EXTS)
    bak_msg = ""
    if backup:
        bak = create_backup(path)
        if bak:
            bak_msg = f" (backup created: {bak.name})"

    prs = Presentation(str(path))

    if slide_idx < 0 or slide_idx >= len(prs.slides):
        raise IndexError(
            f"Slide index {slide_idx} out of range. Presentation has {len(prs.slides)} slides."
        )

    slide = prs.slides[slide_idx]
    if shape_idx < 0 or shape_idx >= len(slide.shapes):
        raise IndexError(f"Shape index {shape_idx} out of range on slide {slide_idx}.")

    shape = slide.shapes[shape_idx]
    if not shape.has_text_frame:
        raise ValueError(f"Shape {shape_idx} does not support text editing.")

    shape.text_frame.text = text
    prs.save(str(path))
    return f"Updated text on slide {slide_idx}, shape {shape_idx} in {path}{bak_msg}"


def ppt_update_notes(
    file_path: str,
    slide_idx: int,
    notes_text: str,
    backup: bool = False,
) -> str:
    """Update the speaker notes of a specific slide.

    Args:
        file_path: Path to the PowerPoint file (.pptx).
        slide_idx: 0-based slide index.
        notes_text: Speaker notes text to set for the slide.
        backup: If True, create a backup of the original file before modifying.

    Returns:
        Status confirmation message.
    """
    path = validate_file_path(file_path, expected_extensions=VALID_PPT_EXTS)
    bak_msg = ""
    if backup:
        bak = create_backup(path)
        if bak:
            bak_msg = f" (backup created: {bak.name})"

    prs = Presentation(str(path))

    if slide_idx < 0 or slide_idx >= len(prs.slides):
        raise IndexError(
            f"Slide index {slide_idx} out of range. Presentation has {len(prs.slides)} slides."
        )

    slide = prs.slides[slide_idx]
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text

    prs.save(str(path))
    return f"Updated notes on slide {slide_idx} in {path}{bak_msg}"


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
        data: Optional 2D list of cell text data.
        left: Left position in inches (default 1.0).
        top: Top position in inches (default 2.0).
        width: Table width in inches (default 8.0).
        height: Table height in inches (default 4.0).
        backup: If True, create a backup before modifying.

    Returns:
        Status confirmation message.
    """
    from pptx.util import Inches

    path = validate_file_path(file_path, expected_extensions=VALID_PPT_EXTS)
    bak_msg = ""
    if backup:
        bak = create_backup(path)
        if bak:
            bak_msg = f" (backup created: {bak.name})"

    prs = Presentation(str(path))
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        raise IndexError(
            f"Slide index {slide_idx} out of range. Presentation has {len(prs.slides)} slides."
        )

    slide = prs.slides[slide_idx]
    shape = slide.shapes.add_table(
        rows, cols, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    table = shape.table

    if data:
        for r_idx, row in enumerate(data):
            if r_idx < rows:
                for c_idx, val in enumerate(row):
                    if c_idx < cols:
                        table.cell(r_idx, c_idx).text = str(val)

    prs.save(str(path))
    return f"Added table ({rows}x{cols}) to slide {slide_idx} in {path}{bak_msg}"
