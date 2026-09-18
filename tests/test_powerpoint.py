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
