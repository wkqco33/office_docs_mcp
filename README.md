# office_docs_mcp

wpycli, wpyconf, wpylog 조합으로 시작하는 CLI 프로젝트입니다.

## 시작하기
1. `uv sync`
2. `uv run office-docs-mcp hello Copilot`

(`.env` 파일은 기본값으로 이미 생성되어 있습니다. 필요하면 직접 수정하세요.)

## 구성
- 공통 의존성: `wpyconf`, `wpylog` (import: `wconfig`, `wlogger`)
- `wpycli` 기반 커맨드 구조와 `wpyconf`/`wpylog` 런타임 연동이 포함되어 있습니다.
- 패키지 인덱스: 공식 PyPI
- 패키지 경로: `src/office_docs_mcp`

## SQLite
- 표준 라이브러리 `sqlite3` 기반 초기화 헬퍼가 포함됩니다.
- 기본 DB 경로: `data/office_docs_mcp.db`

## 개발
- 개발 의존성 설치: `uv sync --group dev`
- 린트: `uv run ruff check .`
- 테스트: `uv run pytest`
