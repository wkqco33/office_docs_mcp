"""PowerPoint presentation manipulation tools using python-pptx."""

from __future__ import annotations

from .reader import ppt_get_outline, ppt_read_slide
from .writer import (
    ppt_add_slide,
    ppt_create_presentation,
    ppt_update_slide_text,
)

__all__ = [
    "ppt_add_slide",
    "ppt_create_presentation",
    "ppt_get_outline",
    "ppt_read_slide",
    "ppt_update_slide_text",
]
