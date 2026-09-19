# Contributing to office_docs_mcp

Thank you for your interest in contributing to `office_docs_mcp`! This project follows a strict Test-Driven Development (TDD) workflow and high code-quality standards.

---

## 1. Development Setup

1. Ensure you have Python >= 3.12 and [`uv`](https://docs.astral.sh/uv/) installed.
2. Clone the repository and install dependencies:
   ```bash
   uv sync --group dev
   ```
3. Run test suite and linters to verify your local environment:
   ```bash
   uv run pytest
   uv run ruff check .
   uv run ruff format --check .
   ```

---

## 2. Test-Driven Development (TDD) Workflow

All new features and bugfixes must follow the Red-Green-Refactor cycle:

1. **Red**: Write a failing unit or integration test under `tests/` using pytest's `tmp_path` fixture for safe file I/O.
2. **Green**: Implement the minimal code in `src/office_docs_mcp/` to make the test pass.
3. **Refactor**: Clean up the implementation, optimize performance, and ensure clean docstrings and type annotations.
4. **Lint & Format**:
   ```bash
   uv run ruff check . --fix
   uv run ruff format .
   ```

---

## 3. Pull Request Guidelines

- Keep pull requests focused on a single change or feature.
- Ensure all CI checks pass.
- Update `CHANGELOG.md` under `[Unreleased]` with your changes.
- Update `README.md` and docstrings if public tool signatures or behavior change.
- Never commit binary files, credentials, or generated files.

---

## 4. Release Process

1. Update the version in `pyproject.toml`, `src/office_docs_mcp/__init__.py`, and `src/office_docs_mcp/main.py`.
2. Move relevant unreleased notes in `CHANGELOG.md` to the new version heading with release date.
3. Commit and push changes to `main`.
4. Create and push a semver tag:
   ```bash
   git tag v0.1.0
   git push origin v0.1.0
   ```
5. GitHub Actions (`.github/workflows/release.yml`) will run tests, build distribution wheels, publish to PyPI using Trusted Publishing (OIDC), and create a GitHub Release.
