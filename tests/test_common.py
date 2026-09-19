from __future__ import annotations

import pytest

from office_docs_mcp.common.file_utils import (
    ensure_parent_dir,
    format_as_markdown_table,
    validate_file_path,
)


def test_validate_file_path_existing(tmp_path):
    f = tmp_path / "test.xlsx"
    f.touch()
    validated = validate_file_path(f, expected_extensions=(".xlsx",))
    assert validated == f


def test_validate_file_path_missing(tmp_path):
    f = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        validate_file_path(f, must_exist=True)


def test_validate_file_path_invalid_extension(tmp_path):
    f = tmp_path / "test.txt"
    f.touch()
    with pytest.raises(ValueError, match="Invalid file extension"):
        validate_file_path(f, expected_extensions=(".xlsx", ".xlsm"))


def test_ensure_parent_dir(tmp_path):
    deep_path = tmp_path / "subdir" / "another" / "doc.docx"
    ensured = ensure_parent_dir(deep_path)
    assert ensured.parent.exists()


def test_format_as_markdown_table():
    headers = ["Col A", "Col B"]
    rows = [["1", "2"], ["3", "4"]]
    md = format_as_markdown_table(headers, rows)
    expected = "| Col A | Col B |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |"
    assert md == expected


def test_create_backup(tmp_path):
    from office_docs_mcp.common.file_utils import create_backup

    f = tmp_path / "data.xlsx"
    f.write_text("content", encoding="utf-8")

    # Custom suffix
    bak = create_backup(f, suffix=".bak")
    assert bak is not None
    assert bak.exists()
    assert bak.name == "data.xlsx.bak"
    assert bak.read_text(encoding="utf-8") == "content"

    # Default timestamp suffix
    bak_ts = create_backup(f)
    assert bak_ts is not None
    assert bak_ts.exists()
    assert bak_ts.name.startswith("data.xlsx.") and bak_ts.name.endswith(".bak")

    # Non-existent file
    missing = tmp_path / "not_exist.xlsx"
    assert create_backup(missing) is None
