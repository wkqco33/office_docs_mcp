# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Note on Versioning**: Version `0.y.z` indicates initial development. Public APIs and tool signatures may undergo refinements based on feedback before reaching `1.0.0`.

## [Unreleased]

### Added
- **Platform-Standard Default Configuration Paths**:
  - Support for XDG Base Directory specification on Linux (`~/.config/office_docs_mcp/config.toml`), Application Support on macOS, and AppData on Windows.
  - Added `--local` (`-l`) flag across `config init`, `config path`, and `config set` to easily operate on local project `./config.toml`.
  - Hierarchical runtime configuration merging: base defaults -> platform config -> local project config -> CLI `--config` override -> environment variables.
- **Config Management Subcommands (`config *`)**:
  - `config init`: Initialize default `config.toml` file with `--force`, `--local`, and `--path` options.
  - `config show`: View merged runtime configuration or inspect a specific key in JSON or TOML format (`--json`, `--toml`).
  - `config path`: Display the active configuration file path with optional `--json` and `--local` flags.
  - `config set <key> <value>`: Modify configuration keys with automatic type parsing (boolean, integer, float, list, string) and atomic writing.
  - `config get <key>`: Retrieve specific configuration values by dotted path.
- **Config Manager Utility**: `src/office_docs_mcp/common/config_manager.py` for safe TOML serialization, platform directory resolution, and nested key management.
- **Automated PyPI Release Workflow**:
  - Added `.github/workflows/release.yml` triggered on version tags (`v*`).
  - Automated testing, package building with `uv build`, PyPI Trusted Publishing (OIDC), and GitHub Releases creation with attached distribution artifacts.

### Added
- **MCP Server**: FastMCP server with Stdio and SSE transport support via `office-docs-mcp serve`.
- **Excel Tools (`excel_*`)**:
  - `excel_get_metadata`: Inspect workbook structure, sheet names, dimensions, and column previews.
  - `excel_read_sheet`: Range reading with 1-based indexing, formula toggle (`evaluate_formulas`), and markdown/raw output.
  - `excel_write_cell`: Write values or formulas to specific cells.
  - `excel_write_range`: Continuous 2D block data insertion.
  - `excel_append_rows`: Append rows to sheet end.
  - `excel_create_workbook`: Initialize blank workbook with custom sheet names.
- **Word Tools (`word_*`)**:
  - `word_get_outline`: Document outline with heading tree, paragraph, and table counts.
  - `word_read_paragraphs`: Chunked 0-based paragraph reading with style info.
  - `word_read_table`: Read tables as Markdown or 2D matrix.
  - `word_append_paragraph`: Append paragraphs or headings with automatic heading level parsing.
  - `word_append_table_row`: Append rows to existing tables.
  - `word_write_table_cell`: Update individual table cells.
  - `word_create_document`: Create blank `.docx` documents with optional title heading.
- **PowerPoint Tools (`ppt_*`)**:
  - `ppt_get_outline`: Slide count, titles, and shape statistics.
  - `ppt_read_slide`: Extract text frames and table data from slides.
  - `ppt_add_slide`: Add slide with title and content placeholder (with automatic fallback to text boxes).
  - `ppt_update_slide_text`: Update text in designated slide shapes.
  - `ppt_create_presentation`: Create blank `.pptx` presentations with optional title slide.
- **CLI Commands**:
  - `office-docs-mcp serve`: Launch the MCP server.
  - `office-docs-mcp tools`: List registered tools with `--plain` and `--json` machine-readable output flags.
  - `office-docs-mcp config`: Display effective configuration.
  - Persistent `--no-color` flag and shell completion generation (`completion <bash|zsh|fish>`).
- **Open-Source & Governance**:
  - MIT License (`LICENSE`).
  - Contribution guidelines (`CONTRIBUTING.md`).
  - Security disclosure policy (`SECURITY.md`).
  - GitHub Actions CI workflow with matrix testing across Python 3.12 and 3.13.
