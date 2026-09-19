from __future__ import annotations

import asyncio
import json
import sys
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from wpycli import Command, ConfigSettings, LoggingSettings
from wpycli.runtime import _flag_override

from .common.config_manager import (
    DEFAULT_CONFIG,
    dump_toml,
    get_nested_key,
    init_config,
    resolve_config_path,
    update_config_file,
)

APP_NAME = "office_docs_mcp"


class SafeConfigSettings(ConfigSettings):
    """ConfigSettings that tolerates missing files and loads only existing configurations."""

    def build(self, flag_values: Mapping[str, Any]) -> Any:
        from wconfig import load_config

        files = [f for f in self.files if Path(f).exists()]
        file_override = _flag_override(flag_values, self.file_flag)
        if file_override is not None and Path(str(file_override)).exists():
            files.append(str(file_override))

        dotenv = self.dotenv
        dotenv_override = _flag_override(flag_values, self.dotenv_flag)
        if dotenv_override is not None:
            dotenv = str(dotenv_override)

        return load_config(
            defaults=self.defaults,
            files=tuple(files),
            dotenv=dotenv,
            env=self.env,
            env_prefix=self.env_prefix,
            env_prefix_separator=self.env_prefix_separator,
            env_nested_delimiter=self.env_nested_delimiter,
        )


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
            "  office-docs-mcp config show\n"
            "  office-docs-mcp tools --plain"
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
        config=SafeConfigSettings(
            defaults=DEFAULT_CONFIG,
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

    # Config parent command
    config = Command(
        use="config",
        short="Manage and inspect runtime configuration.",
        long=(
            "Manage office-docs-mcp configuration (show, init, path, set, get).\n\n"
            "Subcommands:\n"
            "  init    Initialize a new config.toml file with default values\n"
            "  show    Display effective configuration or a specific key\n"
            "  path    Display the path of the active config file\n"
            "  set     Set a configuration key-value pair\n"
            "  get     Get the value of a configuration key\n\n"
            "Examples:\n"
            "  office-docs-mcp config\n"
            "  office-docs-mcp config show\n"
            "  office-docs-mcp config init\n"
            "  office-docs-mcp config path\n"
            "  office-docs-mcp config set logging.level DEBUG\n"
            "  office-docs-mcp config get logging.level"
        ),
        run=run_config_show,
    )
    config.add_bool_flag("json", help="Output configuration in JSON format")
    config.add_bool_flag("toml", help="Output configuration in TOML format")

    config_show = Command(
        use="show [key]",
        short="Display current configuration or a specific key.",
        long=(
            "Display effective runtime configuration or inspect a specific key.\n\n"
            "Examples:\n"
            "  office-docs-mcp config show\n"
            "  office-docs-mcp config show logging.level\n"
            "  office-docs-mcp config show --toml"
        ),
        run=run_config_show,
    )
    config_show.add_bool_flag("json", help="Output configuration in JSON format")
    config_show.add_bool_flag("toml", help="Output configuration in TOML format")

    config_init = Command(
        use="init",
        short="Initialize a new config.toml file with default values.",
        long=(
            "Create a new configuration file with default settings.\n\n"
            "Examples:\n"
            "  office-docs-mcp config init\n"
            "  office-docs-mcp config init --force\n"
            "  office-docs-mcp config init --path custom.toml"
        ),
        run=run_config_init,
    )
    config_init.add_bool_flag("force", help="Overwrite existing configuration file", shorthand="f")
    config_init.add_string_flag(
        "path", help="Target configuration file path (default: config.toml)", shorthand="p"
    )

    config_path = Command(
        use="path",
        short="Display the path of the active config file.",
        long=(
            "Display the path of the configuration file.\n\n"
            "Examples:\n"
            "  office-docs-mcp config path\n"
            "  office-docs-mcp config path --json"
        ),
        run=run_config_path,
    )
    config_path.add_bool_flag("json", help="Output path details in JSON format")

    config_set = Command(
        use="set <key> <value>",
        short="Set a configuration key-value pair in config.toml.",
        long=(
            "Set a configuration key and value in the active config file.\n"
            "Supports dotted keys and automatic type conversion (bool, int, float, list).\n\n"
            "Examples:\n"
            "  office-docs-mcp config set logging.level DEBUG\n"
            "  office-docs-mcp config set app.name my_server"
        ),
        run=run_config_set,
    )
    config_set.add_string_flag(
        "path", help="Target configuration file path to update", shorthand="p"
    )

    config_get = Command(
        use="get <key>",
        short="Get the value of a configuration key.",
        long=(
            "Retrieve the value of a specific configuration key.\n\n"
            "Examples:\n"
            "  office-docs-mcp config get logging.level\n"
            "  office-docs-mcp config get app.name"
        ),
        run=run_config_get,
    )
    config_get.add_bool_flag("json", help="Output value in JSON format")

    config.add_command(config_show, config_init, config_path, config_set, config_get)

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


def run_config_show(ctx) -> int:
    cfg_dict = ctx.config.as_dict() if ctx.config is not None else DEFAULT_CONFIG

    if ctx.args:
        key = ctx.args[0]
        try:
            val = get_nested_key(cfg_dict, key)
        except KeyError:
            print(f"Error: Configuration key '{key}' not found.", file=ctx.stderr)
            return 1

        if ctx.flags.get("json") or isinstance(val, (dict, list)):
            print(json.dumps(val, indent=2))
        elif ctx.flags.get("toml") and isinstance(val, Mapping):
            print(dump_toml({key: val}))
        elif isinstance(val, bool):
            print("true" if val else "false")
        else:
            print(str(val))
        return 0

    if ctx.flags.get("toml"):
        print(dump_toml(cfg_dict))
    else:
        print(json.dumps(cfg_dict, indent=2, sort_keys=True))
    return 0


def run_config_init(ctx) -> int:
    target_path = resolve_config_path(ctx.flags.get("path"), ctx.flags.get("config"))
    force = bool(ctx.flags.get("force"))
    try:
        created = init_config(target_path, force=force)
        print(f"Initialized configuration at {created}")
        return 0
    except FileExistsError as exc:
        print(f"Error: {exc}", file=ctx.stderr)
        return 1
    except Exception as exc:
        print(f"Error initializing configuration: {exc}", file=ctx.stderr)
        return 1


def run_config_path(ctx) -> int:
    target_path = resolve_config_path(ctx.flags.get("config"))
    if ctx.flags.get("json"):
        payload = {
            "path": str(target_path),
            "exists": target_path.exists(),
        }
        print(json.dumps(payload, indent=2))
    else:
        print(str(target_path))
    return 0


def run_config_set(ctx) -> int:
    if len(ctx.args) < 2:
        print("Error: 'set' requires <key> and <value> arguments.", file=ctx.stderr)
        return 2

    key = ctx.args[0]
    raw_val = ctx.args[1]
    target_path = resolve_config_path(ctx.flags.get("path"), ctx.flags.get("config"))

    try:
        parsed_val, updated_path = update_config_file(target_path, key, raw_val)
        print(f"Set '{key}' = {parsed_val!r} in {updated_path}")
        return 0
    except Exception as exc:
        print(f"Error setting configuration key: {exc}", file=ctx.stderr)
        return 1


def run_config_get(ctx) -> int:
    if not ctx.args:
        print("Error: 'get' requires a <key> argument.", file=ctx.stderr)
        return 2

    key = ctx.args[0]
    cfg_dict = ctx.config.as_dict() if ctx.config is not None else DEFAULT_CONFIG

    try:
        val = get_nested_key(cfg_dict, key)
    except KeyError:
        print(f"Error: Configuration key '{key}' not found.", file=ctx.stderr)
        return 1

    if ctx.flags.get("json") or isinstance(val, (dict, list)):
        print(json.dumps(val, indent=2))
    elif isinstance(val, bool):
        print("true" if val else "false")
    else:
        print(str(val))
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
