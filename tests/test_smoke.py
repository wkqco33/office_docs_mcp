from __future__ import annotations

from office_docs_mcp.main import build_root


def test_build_root():
    root = build_root()
    assert root is not None
