# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Note on Versioning**: Version `0.y.z` indicates initial development. Public APIs and tool signatures may undergo refinements based on feedback before reaching `1.0.0`.

## [Unreleased]

## [0.2.0] - 2026-09-19

### Added
- **Full-Text Search Tools**:
  - `excel_search`: Keyword search across all worksheets or specific sheets with row preview.
  - `word_search`: Text search across paragraphs and tables with location details.
  - `ppt_search`: Search across slide shapes, tables, and speaker notes.
- **PowerPoint Speaker Notes & Tables**:
  - `ppt_read_notes`: Retrieve speaker notes for a specific slide.
  - `ppt_update_notes`: Add or update talking points/presentation scripts for slides.
  - `ppt_add_table`: Create structured data tables on slides with custom positioning and dimensions.
- **Excel Sheet Lifecycle Management**:
  - `excel_manage_sheets`: Add, rename, copy, or delete sheets with safety checks.
- **Word Document Editing**:
  - `word_replace_text`: Find and replace text across paragraphs and tables with optional occurrence limit.
  - `word_delete_paragraph`: Remove a specific paragraph by index.
- **MCP Prompts**:
  - `@mcp.prompt() analyze_spreadsheet`: Step-by-step workflow guide for spreadsheet analysis.
  - `@mcp.prompt() create_presentation_outline`: Structured workflow guide for presentation creation.
- **Safe File Backup (`backup=True`)**:
  - Added timestamped `.bak` backup option to all write operations (`excel_write_cell`, `excel_write_range`, `excel_append_rows`, `excel_manage_sheets`, `word_append_paragraph`, `word_append_table_row`, `word_write_table_cell`, `word_replace_text`, `word_delete_paragraph`, `ppt_add_slide`, `ppt_add_table`, `ppt_update_slide_text`, `ppt_update_notes`).

## [0.1.1] - 2026-09-19

### Added
- **MCP Client Integration**: Added `mcp.example.json` with configuration examples for Claude Desktop and other MCP clients.
- **Environment & CLI Options**: Added documentation and examples for environment variables (`OFFICE_DOCS_MCP_LOGGING__LEVEL`, `OFFICE_DOCS_MCP_LOGGING__FILE`, `XDG_CONFIG_HOME`) and CLI flags (`--log-file`, `--log-level`, `-c/--config`).

## [0.1.0] - 2026-09-19

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
- **CLI Commands & Configuration**:
  - `office-docs-mcp serve`: Launch the MCP server.
  - `office-docs-mcp tools`: List registered tools with `--plain` and `--json` machine-readable output flags.
  - `office-docs-mcp config show`: Display effective configuration or specific key (`--json`, `--toml`).
  - `office-docs-mcp config init`: Initialize configuration with default values (`--force`, `--local`, `--path`).
  - `office-docs-mcp config path`: Display active configuration file path (`--json`, `--local`).
  - `office-docs-mcp config set <key> <value>`: Modify configuration keys with automatic type parsing and atomic writing.
  - `office-docs-mcp config get <key>`: Retrieve specific configuration values by dotted path.
  - Platform-standard configuration directory support (Linux XDG, macOS Application Support, Windows AppData).
  - Persistent `--no-color` flag and shell completion generation (`completion <bash|zsh|fish>`).
- **Open-Source, Governance & CI/CD**:
  - MIT License (`LICENSE`).
  - Contribution guidelines (`CONTRIBUTING.md`).
  - Security disclosure policy (`SECURITY.md`).
  - GitHub Actions CI workflow with matrix testing across Python 3.12 and 3.13.
  - GitHub Actions automated release pipeline (`release.yml`) for PyPI Trusted Publishing and GitHub Releases upon version tag push.
