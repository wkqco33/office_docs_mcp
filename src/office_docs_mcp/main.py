from __future__ import annotations

import asyncio
import json
import sys

from wpycli import Command, ConfigSettings, LoggingSettings

APP_NAME = "office_docs_mcp"


def build_root() -> Command:
    root = Command(
        use="office-docs-mcp",
        short="Model Context Protocol (MCP) server for Office documents (Excel, Word, PowerPoint)",
        long=(
            "office-docs-mcp is a Model Context Protocol server that allows LLMs to inspect\n"
            "and modify Excel, Word, and PowerPoint files with token-efficient read and write operations.\n\n"
            "Usage:\n"
            "  office-docs-mcp [command] [flags]\n\n"
            "Examples:\n"
            "  office-docs-mcp serve\n"
            "  office-docs-mcp tools --plain\n"
            "  office-docs-mcp tools --json"
        ),
        version="0.1.0",
    )
    root.add_persistent_string_flag("config", help="Path to config.toml", shorthand="c")
    root.add_persistent_string_flag("dotenv", help="Path to .env file")
    root.add_persistent_string_flag(
        "log-level", help="Override log level (DEBUG, INFO, WARNING, ERROR)"
    )
    root.add_persistent_string_flag("log-file", help="Override log file path")
    root.enable_no_color_flag()
    root.add_completion_command()

    root.configure_runtime(
        config=ConfigSettings(
            defaults={
                "app": {
                    "name": APP_NAME,
                },
                "logging": {"level": "INFO", "file": None},
            },
            files=("config.toml",),
            dotenv=".env",
            env_prefix="OFFICE_DOCS_MCP",
            file_flag="config",
            dotenv_flag="dotenv",
        ),
        logging=LoggingSettings(
            logger_name="office_docs_mcp",
            level_flag="log-level",
            log_file_flag="log-file",
        ),
    )

    config = Command(
        use="config",
        short="Display effective runtime configuration.",
        long="Display the current runtime configuration parsed from flags, files, and environment.",
        run=run_config,
    )

    serve = Command(
        use="serve",
        short="Start the Office Docs MCP server.",
        long=(
            "Start the Office Docs Model Context Protocol server.\n\n"
            "Examples:\n"
            "  office-docs-mcp serve\n"
            "  office-docs-mcp serve --transport stdio"
        ),
        run=run_serve,
    )
    serve.add_string_flag(
        "transport",
        help="Transport type ('stdio' or 'sse')",
        default="stdio",
        shorthand="t",
    )

    tools = Command(
        use="tools",
        short="List all registered MCP tools.",
        long=(
            "List all available tools provided by the Office Docs MCP server.\n\n"
            "Examples:\n"
            "  office-docs-mcp tools\n"
            "  office-docs-mcp tools --plain\n"
            "  office-docs-mcp tools --json"
        ),
        run=run_tools,
    )
    tools.add_bool_flag("json", help="Output tools as JSON")
    tools.add_bool_flag("plain", help="Output tools as plain text (name: description)")

    root.add_command(config, serve, tools)
    return root


def run_config(ctx) -> int:
    rendered = json.dumps(ctx.config.as_dict(), indent=2, sort_keys=True)
    print(rendered)
    return 0


def run_serve(ctx) -> int:
    from .server import get_server

    transport = ctx.flags.get("transport", "stdio")
    server = get_server()
    if ctx.logger is not None:
        ctx.logger.info("Starting Office Docs MCP server", extra={"transport": transport})
    server.run(transport=transport)
    return 0


def run_tools(ctx) -> int:
    from .server import get_server

    server = get_server()
    tools_list = asyncio.run(server.list_tools())

    if ctx.flags.get("json"):
        payload = [
            {
                "name": tool.name,
                "description": tool.description,
                "parameters": getattr(tool, "inputSchema", None),
            }
            for tool in tools_list
        ]
        print(json.dumps(payload, indent=2))
        return 0

    if ctx.flags.get("plain"):
        for tool in tools_list:
            summary = (tool.description or "").strip().split("\n")[0]
            print(f"{tool.name}: {summary}")
        return 0

    print(f"Registered MCP Tools ({len(tools_list)}):\n")
    for tool in tools_list:
        summary = (tool.description or "").strip().split("\n")[0]
        print(f"  - {tool.name:<26} {summary}")
    return 0


def main() -> int:
    return build_root().execute()


if __name__ == "__main__":
    sys.exit(main())
