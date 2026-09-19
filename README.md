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
  - `excel_write_cell`: 단일 셀(`A1` 등) 값/수식 쓰기
  - `excel_write_range`: 시작 셀부터 2차원 데이터 연속 기입
  - `excel_append_rows`: 시트 마지막 행 뒤에 데이터 추가
  - `excel_create_workbook`: 새 빈 엑셀 파일 생성
- **Word (`.docx`)**:
  - `word_get_outline`: 문서 헤딩(제목) 목록 및 단락/표 개수 조회
  - `word_read_paragraphs`: 단락 슬라이싱 읽기 (0-based 페이징)
  - `word_read_table`: 표 내용을 Markdown 테이블 또는 2차원 배열로 읽기
  - `word_append_paragraph`: 문서 끝에 단락/헤딩 추가 (헤딩 레벨 자동 처리)
  - `word_append_table_row`: 특정 표에 새 행 추가
  - `word_write_table_cell`: 특정 표의 셀 텍스트 수정
  - `word_create_document`: 새 워드 문서 생성
- **PowerPoint (`.pptx`)**:
  - `ppt_get_outline`: 전체 슬라이드 수 및 각 슬라이드 제목/셰이프 요약 조회
  - `ppt_read_slide`: 특정 슬라이드의 텍스트 상자 및 표 내용 추출
  - `ppt_add_slide`: 제목과 본문을 포함하는 새 슬라이드 추가 (플레이스홀더 부재 시 자동 텍스트박스 폴백)
  - `ppt_update_slide_text`: 슬라이드 내 특정 셰이프 텍스트 수정
  - `ppt_create_presentation`: 새 프레젠테이션 파일 생성

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

### 4. Claude Desktop / MCP 클라이언트 연동 설정 예시

별도의 저장소 클론이나 사전 설치 없이 `uvx`를 통해 곧바로 연동할 수 있습니다:

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
