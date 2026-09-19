from __future__ import annotations

import json

import office_docs_mcp
import office_docs_mcp.common
import office_docs_mcp.excel
import office_docs_mcp.powerpoint
import office_docs_mcp.word
from office_docs_mcp.main import build_root


def test_build_root():
    root = build_root()
    assert root is not None
    assert root.name == "office-docs-mcp"


def test_cli_tools_plain(capsys):
    root = build_root()
    code = root.execute(["tools", "--plain"])
    assert code == 0
    captured = capsys.readouterr()
    assert "excel_get_metadata:" in captured.out
    assert "word_get_outline:" in captured.out
    assert "ppt_get_outline:" in captured.out


def test_cli_tools_json(capsys):
    root = build_root()
    code = root.execute(["tools", "--json"])
    assert code == 0
    captured = capsys.readouterr()
    tools = json.loads(captured.out)
    tool_names = {t["name"] for t in tools}
    assert "excel_read_sheet" in tool_names
    assert "word_read_paragraphs" in tool_names
    assert "ppt_read_slide" in tool_names


def test_cli_config(capsys):
    root = build_root()
    code = root.execute(["config"])
    assert code == 0
    captured = capsys.readouterr()
    cfg = json.loads(captured.out)
    assert cfg["app"]["name"] == "office_docs_mcp"


def test_public_api_exports():
    # Top-level exports
    assert hasattr(office_docs_mcp, "__version__")
    assert hasattr(office_docs_mcp, "get_server")
    assert hasattr(office_docs_mcp, "mcp")
    assert "__version__" in office_docs_mcp.__all__

    # Common exports
    assert "validate_file_path" in office_docs_mcp.common.__all__
    assert "ensure_parent_dir" in office_docs_mcp.common.__all__
    assert "format_as_markdown_table" in office_docs_mcp.common.__all__
    assert "get_platform_config_path" in office_docs_mcp.common.__all__

    # Excel exports
    assert "excel_get_metadata" in office_docs_mcp.excel.__all__
    assert "excel_read_sheet" in office_docs_mcp.excel.__all__
    assert "excel_create_workbook" in office_docs_mcp.excel.__all__

    # Word exports
    assert "word_get_outline" in office_docs_mcp.word.__all__
    assert "word_read_paragraphs" in office_docs_mcp.word.__all__
    assert "word_create_document" in office_docs_mcp.word.__all__

    # PowerPoint exports
    assert "ppt_get_outline" in office_docs_mcp.powerpoint.__all__
    assert "ppt_read_slide" in office_docs_mcp.powerpoint.__all__
    assert "ppt_create_presentation" in office_docs_mcp.powerpoint.__all__
