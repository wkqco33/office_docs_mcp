from __future__ import annotations

from mcp.server.mcpserver import MCPServer

mcp = MCPServer("office-docs-mcp")


def get_server() -> MCPServer:
    """Return the configured MCPServer instance."""
    return mcp
