from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import Any


def validate_file_path(
    file_path: str | Path,
    expected_extensions: tuple[str, ...] | None = None,
    must_exist: bool = True,
) -> Path:
    """Validate that the file path is safe and matches expectations."""
    path = Path(file_path).resolve()

    if must_exist and not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    if expected_extensions:
        ext = path.suffix.lower()
        if ext not in expected_extensions:
            raise ValueError(
                f"Invalid file extension '{ext}'. Expected one of: {', '.join(expected_extensions)}"
            )

    return path


def ensure_parent_dir(file_path: str | Path) -> Path:
    """Ensure the parent directory of the file exists and return resolved Path."""
    path = Path(file_path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def format_as_markdown_table(headers: list[str], rows: Iterable[list[Any]]) -> str:
    """Format 2D data rows with headers into a standard Markdown table."""
    norm_headers = [str(h) if h is not None else "" for h in headers]
    if not norm_headers:
        return ""

    col_count = len(norm_headers)
    table_lines = [
        "| " + " | ".join(norm_headers) + " |",
        "| " + " | ".join(["---"] * col_count) + " |",
    ]

    for row in rows:
        row_padded = list(row) + [""] * (col_count - len(row))
        cells = [
            str(c).replace("\n", " ").replace("|", "\\|") if c is not None else ""
            for c in row_padded[:col_count]
        ]
        table_lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(table_lines)
