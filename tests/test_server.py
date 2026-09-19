from __future__ import annotations

import pytest

from office_docs_mcp.server import get_server


@pytest.mark.anyio
async def test_server_tools_registered():
    server = get_server()
    tools = await server.list_tools()
    tool_names = {t.name for t in tools}

    # Verify Excel tools
    assert "excel_get_metadata" in tool_names
    assert "excel_read_sheet" in tool_names
    assert "excel_search" in tool_names
    assert "excel_manage_sheets" in tool_names
    assert "excel_write_cell" in tool_names
    assert "excel_write_range" in tool_names
    assert "excel_append_rows" in tool_names
    assert "excel_create_workbook" in tool_names

    # Verify Word tools
    assert "word_get_outline" in tool_names
    assert "word_read_paragraphs" in tool_names
    assert "word_read_table" in tool_names
    assert "word_search" in tool_names
    assert "word_replace_text" in tool_names
    assert "word_delete_paragraph" in tool_names
    assert "word_append_paragraph" in tool_names
    assert "word_append_table_row" in tool_names
    assert "word_write_table_cell" in tool_names
    assert "word_create_document" in tool_names

    # Verify PowerPoint tools
    assert "ppt_get_outline" in tool_names
    assert "ppt_read_slide" in tool_names
    assert "ppt_read_notes" in tool_names
    assert "ppt_update_notes" in tool_names
    assert "ppt_search" in tool_names
    assert "ppt_add_slide" in tool_names
    assert "ppt_add_table" in tool_names
    assert "ppt_update_slide_text" in tool_names
    assert "ppt_create_presentation" in tool_names


@pytest.mark.anyio
async def test_server_prompts_registered():
    server = get_server()
    prompts = await server.list_prompts()
    prompt_names = {p.name for p in prompts}
    assert "analyze_spreadsheet" in prompt_names
    assert "create_presentation_outline" in prompt_names


@pytest.mark.anyio
async def test_server_call_excel_workflow(tmp_path):
    server = get_server()
    f = tmp_path / "server_test.xlsx"

    # 1. Create workbook
    res = await server.call_tool(
        "excel_create_workbook", {"file_path": str(f), "sheet_names": ["Main"]}
    )
    assert not res.is_error

    # 2. Write cell
    res_write = await server.call_tool(
        "excel_write_cell",
        {"file_path": str(f), "sheet_name": "Main", "coordinate": "A1", "value": "MCP Test"},
    )
    assert not res_write.is_error

    # 3. Read sheet
    res_read = await server.call_tool(
        "excel_read_sheet",
        {
            "file_path": str(f),
            "sheet_name": "Main",
            "start_row": 1,
            "end_row": 1,
            "start_col": 1,
            "end_col": 1,
        },
    )
    assert not res_read.is_error
    assert "MCP Test" in str(res_read.content)
