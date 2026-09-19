# office_docs_mcp

[![CI](https://github.com/wkqco33/office_docs_mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/wkqco33/office_docs_mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python: 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](pyproject.toml)

LLM(대형 언어 모델)이 오피스 문서(Excel, Word, PowerPoint)를 안정적이고 토큰 효율적으로 읽고 쓸 수 있도록 지원하는 **Model Context Protocol (MCP)** 서버입니다.

> **버전 안내**: 본 프로젝트는 현재 초기 개발 단계(`0.y.z`)이며, 사용자 피드백에 따라 도구 인터페이스가 지속적으로 개선되고 있습니다.

---

## 주요 기능

- **Excel (`.xlsx`, `.xlsm`)**:
  - `excel_get_metadata`: 시트 목록, 크기, 컬럼 요약 조회
  - `excel_read_sheet`: 지정된 행/열 범위(1-based)를 Markdown 테이블 또는 2차원 배열로 읽기 (수식/값 토글 지원)
  - `excel_search`: 워크시트 전체 또는 특정 시트의 셀에서 키워드 검색
  - `excel_write_cell`: 단일 셀(`A1` 등) 값/수식 쓰기 (`backup=True` 지원)
  - `excel_write_range`: 시작 셀부터 2차원 데이터 연속 기입 (`backup=True` 지원)
  - `excel_append_rows`: 시트 마지막 행 뒤에 데이터 추가 (`backup=True` 지원)
  - `excel_manage_sheets`: 시트 추가(`add`), 이름변경(`rename`), 복사(`copy`), 삭제(`delete`) (`backup=True` 지원)
  - `excel_create_workbook`: 새 빈 엑셀 파일 생성
- **Word (`.docx`)**:
  - `word_get_outline`: 문서 헤딩(제목) 목록 및 단락/표 개수 조회
  - `word_read_paragraphs`: 단락 슬라이싱 읽기 (0-based 페이징)
  - `word_read_table`: 표 내용을 Markdown 테이블 또는 2차원 배열로 읽기
  - `word_search`: 문서 내 단락 및 표 셀 대상 키워드 검색
  - `word_append_paragraph`: 문서 끝에 단락/헤딩 추가 (헤딩 레벨 자동 처리, `backup=True` 지원)
  - `word_append_table_row`: 특정 표에 새 행 추가 (`backup=True` 지원)
  - `word_write_table_cell`: 특정 표의 셀 텍스트 수정 (`backup=True` 지원)
  - `word_replace_text`: 문서 전체 텍스트 일괄/부분 치환 (`backup=True` 지원)
  - `word_delete_paragraph`: 특정 단락 삭제 (`backup=True` 지원)
  - `word_create_document`: 새 워드 문서 생성
- **PowerPoint (`.pptx`)**:
  - `ppt_get_outline`: 전체 슬라이드 수 및 각 슬라이드 제목/셰이프 요약 조회
  - `ppt_read_slide`: 특정 슬라이드의 텍스트 상자, 표, 발표자 메모 내용 추출
  - `ppt_read_notes`: 특정 슬라이드의 발표자 메모(Speaker Notes) 읽기
  - `ppt_update_notes`: 발표 대본/슬라이드 메모 작성 및 수정 (`backup=True` 지원)
  - `ppt_search`: 슬라이드 텍스트, 표, 발표자 메모 대상 키워드 검색
  - `ppt_add_slide`: 제목과 본문을 포함하는 새 슬라이드 추가 (`backup=True` 지원)
  - `ppt_add_table`: 슬라이드 내 표 생성 (`backup=True` 지원)
  - `ppt_update_slide_text`: 슬라이드 내 특정 셰이프 텍스트 수정 (`backup=True` 지원)
  - `ppt_create_presentation`: 새 프레젠테이션 파일 생성
- **MCP Prompts (`@mcp.prompt()`)**:
  - `analyze_spreadsheet`: 엑셀 데이터 구조 파악부터 심층 분석까지의 체계적 워크플로우 가이드
  - `create_presentation_outline`: 주어진 주제에 대한 슬라이드 구성 및 발표자 대본 작성 가이드
- **안전한 파일 백업**:
  - 모든 쓰기 도구에서 `backup: bool = True` 지정 시 수정 전 원본을 타임스탬프 백업 파일(`.bak`)로 자동 보관

---

## 시작하기

### 1. 요구 사항
- Python >= 3.12
- [uv](https://docs.astral.sh/uv/) 패키지 관리자

### 2. 설치 및 가상환경 동기화
```bash
uv sync --group dev
```

### 3. CLI 명령어

```bash
# MCP 서버 실행 (기본 Stdio 트랜스포트)
uv run office-docs-mcp serve

# SSE 트랜스포트로 실행
uv run office-docs-mcp serve --transport sse

# 등록된 MCP 도구 목록 확인 (사람 가독형)
uv run office-docs-mcp tools

# 기계 판독용 단일 라인 목록
uv run office-docs-mcp tools --plain

# 기계 판독용 JSON 스키마 덤프
uv run office-docs-mcp tools --json

# 설정 관리 서브명령어 (각 플랫폼 기본 경로 표준 지원)
# - Linux: ~/.config/office_docs_mcp/config.toml ($XDG_CONFIG_HOME)
# - macOS: ~/Library/Application Support/office_docs_mcp/config.toml
# - Windows: %APPDATA%/office_docs_mcp/config.toml
uv run office-docs-mcp config show                  # 현재 병합된 설정 전체 확인 (JSON)
uv run office-docs-mcp config show --toml           # TOML 형식으로 확인
uv run office-docs-mcp config init                  # 플랫폼 기본 경로에 config.toml 생성
uv run office-docs-mcp config init --local          # 현재 작업 디렉토리에 ./config.toml 생성
uv run office-docs-mcp config path                  # 플랫폼 기본 설정 파일 경로 확인 (--json 지원)
uv run office-docs-mcp config path --local          # 로컬 설정 파일 경로 확인
uv run office-docs-mcp config set logging.level DEBUG # 플랫폼 설정 파일 키 변경
uv run office-docs-mcp config set --local logging.level DEBUG # 로컬 설정 파일 키 변경
uv run office-docs-mcp config get logging.level       # 특정 설정 값 조회

# 셸 자동완성 스크립트 생성
uv run office-docs-mcp completion bash > /etc/bash_completion.d/office-docs-mcp
```

### 4. MCP 클라이언트 연동 설정 (Claude Desktop, Antigravity `agy`, Cursor 등)

별도의 저장소 클론이나 사전 설치 없이 `uvx`를 통해 곧바로 연동할 수 있습니다:

#### Antigravity CLI (`agy`) 및 Antigravity IDE / 2.0
- **전역 설정 (모든 프로젝트에서 사용)**: `~/.gemini/config/mcp_config.json`
- **프로젝트별 로컬 설정**: 프로젝트 루트의 `.mcp.json`

```json
{
  "mcpServers": {
    "office-docs": {
      "command": "uvx",
      "args": ["office-docs-mcp", "serve"]
    }
  }
}
```

#### Claude Desktop 설정
`claude_desktop_config.json`에 동일하게 추가합니다:

```json
{
  "mcpServers": {
    "office-docs": {
      "command": "uvx",
      "args": ["office-docs-mcp", "serve"]
    }
  }
}
```

#### 환경 변수(`env`) 및 디버깅 옵션 설정 예시

MCP stdio 통신 시 stdout 오염을 방지하면서 디버그 로그를 파일에 기록하거나 설정을 커스텀할 수 있습니다:

```json
{
  "mcpServers": {
    "office-docs": {
      "command": "uvx",
      "args": [
        "office-docs-mcp",
        "--log-level", "DEBUG",
        "--log-file", "/tmp/office_docs_mcp.log",
        "serve"
      ],
      "env": {
        "OFFICE_DOCS_MCP_LOGGING__LEVEL": "DEBUG",
        "OFFICE_DOCS_MCP_LOGGING__FILE": "/tmp/office_docs_mcp.log"
      }
    }
  }
}
```

#### 로컬 개발 소스 디렉토리에서 연동할 경우:

```json
{
  "mcpServers": {
    "office-docs": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/office_docs_mcp",
        "run",
        "office-docs-mcp",
        "serve"
      ]
    }
  }
}
```

#### 지원되는 환경 변수 및 CLI 옵션

| 분류 | 이름 / 플래그 | 설명 | 기본값 / 예시 |
| :--- | :--- | :--- | :--- |
| **환경 변수** | `OFFICE_DOCS_MCP_LOGGING__LEVEL` | 로깅 레벨 설정 | `INFO` (`DEBUG`, `INFO`, `WARNING`, `ERROR`) |
| **환경 변수** | `OFFICE_DOCS_MCP_LOGGING__FILE` | 서버 로그 출력 파일 경로 | `/path/to/office_docs_mcp.log` |
| **환경 변수** | `OFFICE_DOCS_MCP_APP__NAME` | 앱 식별자 이름 | `office_docs_mcp` |
| **환경 변수** | `XDG_CONFIG_HOME` | 기본 설정 파일 저장 디렉토리 (Linux) | `~/.config` |
| **CLI 옵션** | `--log-level <LEVEL>` | 명령줄에서 로그 레벨 직접 지정 | `DEBUG`, `INFO` 등 |
| **CLI 옵션** | `--log-file <PATH>` | 명령줄에서 로그 파일 경로 직접 지정 | `/path/to/logfile.log` |
| **CLI 옵션** | `-c, --config <PATH>` | 사용할 `config.toml` 경로 명시적 지정 | `~/.config/office_docs_mcp/config.toml` |
| **CLI 옵션** | `--no-color` | 터미널/로그 ANSI 색상 코드 비활성화 | 플래그 |
| **CLI 옵션** | `-t, --transport <stdio\|sse>` | MCP 전송 프로토콜 지정 (`serve` 명령어) | `stdio` |

전체 설정 템플릿은 [mcp.example.json](mcp.example.json) 파일에서 확인하실 수 있습니다.

### Pi 연동 (`uvx`)

Pi에서 MCP를 사용하려면 `pi-mcp-adapter`가 필요합니다. 이 저장소에는 Pi가 자동으로 읽는 프로젝트 설정 파일 `.mcp.json`이 포함되어 있습니다.

```bash
# 최초 1회
pi install npm:pi-mcp-adapter

# 저장소 루트에서 실행하면 .mcp.json을 자동으로 읽음
pi

# 또는 설정 파일을 명시
pi --mcp-config .mcp.json
```

Pi에서 `/mcp`로 `office-docs` 서버 상태를 확인하고, 설정을 변경한 뒤에는 `/reload`를 실행하세요. `uvx`가 PyPI에서 `office-docs-mcp`를 준비해 MCP stdio 서버를 실행하므로 별도 가상환경 활성화가 필요 없습니다. 서버는 첫 도구 호출 때 시작됩니다.

### ncli로 사용 가이드 문서 추가

동일한 안내를 ncli 노트로 저장하려면 로그인 후 다음을 실행하세요.

```bash
ncli login
ncli add \\
  --title "Office Docs MCP를 Pi에서 uvx로 사용하기" \\
  --content-file docs/pi-mcp-ncli-guide.md \\
  --category other \\
  --no-input
```

자세한 절차는 [docs/pi-mcp-ncli-guide.md](docs/pi-mcp-ncli-guide.md)를 참고하세요.

---

## 개발 및 테스트 (TDD)

```bash
# 전체 단위 및 통합 테스트 실행
uv run pytest -v

# 린트 및 코드 스타일 검증
uv run ruff check .
uv run ruff format --check .
```

자세한 기여 방법 및 개발 가이드는 아래 문서를 참고하세요:
- [AGENTS.md](AGENTS.md): AI 에이전트 및 개발자를 위한 설계 원칙 및 TDD 지침
- [CONTRIBUTING.md](CONTRIBUTING.md): 기여 가이드라인
- [CHANGELOG.md](CHANGELOG.md): 변경 이력
- [SECURITY.md](SECURITY.md): 보안 정책
- [LICENSE](LICENSE): MIT License
