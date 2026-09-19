"""office_docs_mcp - Model Context Protocol (MCP) server for Office documents."""

from __future__ import annotations

from .server import get_server, mcp

__version__ = "0.2.0"

__all__ = [
    "__version__",
    "get_server",
    "mcp",
]
