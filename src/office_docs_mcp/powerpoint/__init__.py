"""PowerPoint presentation manipulation tools using python-pptx."""

from __future__ import annotations

from .reader import ppt_get_outline, ppt_read_notes, ppt_read_slide, ppt_search
from .writer import (
    ppt_add_chart,
    ppt_add_flowchart,
    ppt_add_slide,
    ppt_add_table,
    ppt_create_presentation,
    ppt_update_notes,
    ppt_update_slide_text,
)

__all__ = [
    "ppt_add_chart",
    "ppt_add_flowchart",
    "ppt_add_slide",
    "ppt_add_table",
    "ppt_create_presentation",
    "ppt_get_outline",
    "ppt_read_notes",
    "ppt_read_slide",
    "ppt_search",
    "ppt_update_notes",
    "ppt_update_slide_text",
]
