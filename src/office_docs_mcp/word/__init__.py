"""Word document manipulation tools using python-docx."""

from __future__ import annotations

from .reader import (
    word_get_outline,
    word_read_paragraphs,
    word_read_table,
    word_search,
)
from .writer import (
    word_append_paragraph,
    word_append_table_row,
    word_create_document,
    word_delete_paragraph,
    word_replace_text,
    word_write_table_cell,
)

__all__ = [
    "word_append_paragraph",
    "word_append_table_row",
    "word_create_document",
    "word_delete_paragraph",
    "word_get_outline",
    "word_read_paragraphs",
    "word_read_table",
    "word_replace_text",
    "word_search",
    "word_write_table_cell",
]
