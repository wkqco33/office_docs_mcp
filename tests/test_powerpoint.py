from __future__ import annotations

import pytest

from office_docs_mcp.powerpoint.reader import (
    ppt_get_outline,
    ppt_read_slide,
)
from office_docs_mcp.powerpoint.writer import (
    ppt_add_slide,
    ppt_create_presentation,
    ppt_update_slide_text,
)


def test_ppt_create_and_outline(tmp_path):
    f = tmp_path / "deck.pptx"
    res = ppt_create_presentation(str(f), title="Quarterly Review")
    assert "Created" in res
    assert f.exists()

    outline = ppt_get_outline(str(f))
    assert outline["total_slides"] == 1
    assert outline["slides"][0]["title"] == "Quarterly Review"


def test_ppt_add_and_read_slide(tmp_path):
    f = tmp_path / "deck_slides.pptx"
    ppt_create_presentation(str(f), title="Main Title")

    ppt_add_slide(str(f), title="Agenda", content="1. Finance\n2. Product\n3. Operations")

    outline = ppt_get_outline(str(f))
    assert outline["total_slides"] == 2
    assert outline["slides"][1]["title"] == "Agenda"

    slide_data = ppt_read_slide(str(f), slide_idx=1)
    assert slide_data["title"] == "Agenda"
    # Find shape containing the content
    content_found = any("Finance" in s.get("text", "") for s in slide_data["shapes"])
    assert content_found


def test_ppt_update_slide_text(tmp_path):
    f = tmp_path / "deck_edit.pptx"
    ppt_create_presentation(str(f), title="Initial Title")

    # The title shape is typically shape 0
    res = ppt_update_slide_text(str(f), slide_idx=0, shape_idx=0, text="Updated Title")
    assert "Updated" in res

    slide_data = ppt_read_slide(str(f), slide_idx=0)
    assert slide_data["shapes"][0]["text"] == "Updated Title"


def test_ppt_invalid_slide_index(tmp_path):
    f = tmp_path / "deck_empty.pptx"
    ppt_create_presentation(str(f))
    with pytest.raises(IndexError, match="Slide index 99 out of range"):
        ppt_read_slide(str(f), slide_idx=99)


def test_ppt_add_slide_fallback_textbox(tmp_path):
    f = tmp_path / "deck_fallback.pptx"
    ppt_create_presentation(str(f))
    # Layout 5 is Title-only (has only 1 placeholder)
    ppt_add_slide(str(f), title="Title Only", content="Fallback Content", layout_idx=5)

    slide_data = ppt_read_slide(str(f), slide_idx=0)
    assert slide_data["title"] == "Title Only"
    content_found = any("Fallback Content" in s.get("text", "") for s in slide_data["shapes"])
    assert content_found


def test_ppt_speaker_notes(tmp_path):
    from office_docs_mcp.powerpoint.reader import ppt_read_notes
    from office_docs_mcp.powerpoint.writer import ppt_update_notes

    f = tmp_path / "notes_deck.pptx"
    ppt_create_presentation(str(f), title="Slide with Notes")

    # Initial notes should be empty
    initial_notes = ppt_read_notes(str(f), slide_idx=0)
    assert initial_notes == ""

    # Update notes with backup=True
    res = ppt_update_notes(
        str(f), slide_idx=0, notes_text="Remember to mention Q3 results.", backup=True
    )
    assert "Updated notes" in res

    # Check updated notes
    updated_notes = ppt_read_notes(str(f), slide_idx=0)
    assert "Remember to mention Q3 results." in updated_notes

    # Check backup file exists
    backups = list(tmp_path.glob("notes_deck.pptx.*.bak"))
    assert len(backups) == 1


def test_ppt_search(tmp_path):
    from office_docs_mcp.powerpoint.reader import ppt_search
    from office_docs_mcp.powerpoint.writer import ppt_update_notes

    f = tmp_path / "search_deck.pptx"
    ppt_create_presentation(str(f), title="Global Launch Plan")
    ppt_add_slide(str(f), title="Market Analysis", content="Target audience is Enterprise users.")
    ppt_update_notes(str(f), slide_idx=1, notes_text="Highlight Enterprise scalability.")

    # Search case-insensitive across slides, shapes, and notes
    results = ppt_search(str(f), query="enterprise")
    assert len(results) == 2
    types = {r["source"] for r in results}
    assert "shape" in types
    assert "notes" in types

    # Case-sensitive search
    exact_results = ppt_search(str(f), query="Global Launch", case_sensitive=True)
    assert len(exact_results) == 1
    assert exact_results[0]["slide_idx"] == 0
    assert exact_results[0]["source"] == "shape"


def test_ppt_add_table(tmp_path):
    from office_docs_mcp.powerpoint.reader import ppt_read_slide
    from office_docs_mcp.powerpoint.writer import ppt_add_table

    f = tmp_path / "table_deck.pptx"
    ppt_create_presentation(str(f), title="Table Demo")

    data = [
        ["Feature", "Status"],
        ["OAuth", "Done"],
        ["Billing", "In Progress"],
    ]
    res = ppt_add_table(str(f), slide_idx=0, rows=3, cols=2, data=data, backup=True)
    assert "Added table (3x2)" in res

    slide_data = ppt_read_slide(str(f), slide_idx=0)
    table_shape = next(s for s in slide_data["shapes"] if s.get("table_data"))
    assert table_shape["table_data"] == data
