# AGENTS.md - Office Docs MCP Server

이 문서는 AI 에이전트와 개발자가 `office_docs_mcp` 프로젝트의 구조, 설계 원칙, TDD 개발 워크플로우를 이해하고 일관성 있게 협업할 수 있도록 돕는 지침서입니다.

---

## 1. 프로젝트 개요 (Project Overview)

`office_docs_mcp`는 LLM(대형 언어 모델)이 오피스 문서(Excel, Word, PowerPoint)를 안정적이고 효율적으로 읽고 쓸 수 있도록 지원하는 **Model Context Protocol (MCP)** 서버입니다.

### 핵심 설계 철학
1. **LLM 친화적 인터페이스 (Token & Context Conscious)**:
   - 전체 문서를 무작정 텍스트로 덤프하지 않고, **개요 조회(Outline/Metadata) -> 범위/페이징 읽기(Chunked Read) -> 정밀 쓰기(Precise Write)** 흐름을 제공합니다.
   - 반환 데이터는 토큰 효율이 높고 가독성이 좋은 Markdown 테이블 또는 정형화된 JSON 형식을 취합니다.
2. **테스트 주도 개발 (TDD)**:
   - 모든 기능은 요구사항 정의 -> 실패하는 테스트 작성(Red) -> 최소 구현(Green) -> 리팩토링(Refactor) 순서로 개발합니다.
   - 실제 파일 IO 테스트는 `pytest`의 `tmp_path` 픽스처를 활용하여 독립적이고 빠르게 수행합니다.
3. **안전한 파일 조작 (Safe File I/O)**:
   - 파일 유효성 검사(경로 존재 여부, 확장자 검사)를 철저히 수행합니다.
   - 기존 파일 덮어쓰기 시 백업 또는 안전 장치를 제공합니다.

---

## 2. 기술 스택 및 의존성 (Tech Stack)

- **Language Runtime**: Python >= 3.12
- **Package & Environment Manager**: `uv`
- **MCP Framework**: `mcp` (Official Python MCP SDK - FastMCP)
- **Office Document Libraries**:
  - **Excel (`.xlsx`, `.xlsm`)**: `openpyxl` (표준, 셀 단위 조작, 수식 및 서식 지원)
  - **Word (`.docx`)**: `python-docx` (단락, 표, 헤딩 조작 표준)
  - **PowerPoint (`.pptx`)**: `python-pptx` (슬라이드, 셰이프, 텍스트, 표 조작 표준)
- **Linting & Formatting**: `ruff`
- **Testing**: `pytest`

---

## 3. 디렉토리 구조 (Directory Structure)

```text
office_docs_mcp/
├── AGENTS.md                  # 에이전트용 개발 가이드 (본 문서)
├── README.md                  # 사용자용 프로젝트 설명서
├── pyproject.toml             # uv 패키지 및 의존성 정의
├── config.toml                # 런타임 설정 파일
├── src/
│   └── office_docs_mcp/
│       ├── __init__.py
│       ├── main.py            # CLI 진입점 (serve 명령어 제공)
│       ├── server.py          # FastMCP 서버 인스턴스 및 도구 바인딩
│       ├── common/            # 공통 유틸리티 (파일 검증, 페이징 등)
│       │   ├── __init__.py
│       │   └── file_utils.py
│       ├── excel/             # Excel 전용 로직 (.xlsx)
│       │   ├── __init__.py
│       │   ├── reader.py
│       │   └── writer.py
│       ├── word/              # Word 전용 로직 (.docx)
│       │   ├── __init__.py
│       │   ├── reader.py
│       │   └── writer.py
│       └── powerpoint/        # PowerPoint 전용 로직 (.pptx)
│           ├── __init__.py
│           ├── reader.py
│           └── writer.py
└── tests/
    ├── __init__.py
    ├── test_smoke.py
    ├── test_excel.py          # Excel 서비스 및 도구 단위 테스트
    ├── test_word.py           # Word 서비스 및 도구 단위 테스트
    ├── test_powerpoint.py     # PowerPoint 서비스 및 도구 단위 테스트
    └── test_server.py         # MCP 도구 등록 및 실행 통합 테스트
```

---

## 4. MCP 도구 명세 (Tool Specifications)

### 4.1 Excel 도구군 (`excel_*`)
- `excel_get_metadata(file_path: str)`: 시트 목록, 시트별 최대 행/열 수, 컬럼 헤더 요약.
- `excel_read_sheet(file_path: str, sheet_name: str | None = None, start_row: int = 1, end_row: int = 50, start_col: int = 1, end_col: int = 20, format: str = "markdown")`:
  - 특정 시트의 지정된 행/열 범위를 읽어 반환 (기본 1-based 인덱스 사용, LLM 직관성 일치).
- `excel_write_cell(file_path: str, sheet_name: str, coordinate: str, value: Any)`:
  - 예: `A1` 또는 행/열 번호로 단일 셀 값 변경.
- `excel_write_range(file_path: str, sheet_name: str, start_cell: str, data: list[list[Any]])`:
  - 시작 셀(`A1` 등)부터 2차원 리스트 데이터를 연속 기입.
- `excel_append_rows(file_path: str, sheet_name: str, rows: list[list[Any]])`:
  - 시트 맨 마지막 행 뒤에 새로운 행들 추가.
- `excel_create_workbook(file_path: str, sheet_names: list[str] | None = None)`:
  - 빈 엑셀 파일 생성.

### 4.2 Word 도구군 (`word_*`)
- `word_get_outline(file_path: str)`: 문서의 헤딩(제목) 목록, 전체 단락 수, 표(Table) 개수 등 구조 요약.
- `word_read_paragraphs(file_path: str, start_idx: int = 0, count: int = 30)`:
  - 단락 단위로 본문 텍스트 슬라이싱 읽기.
- `word_append_paragraph(file_path: str, text: str, style: str | None = None)`:
  - 문서 끝에 새 단락 또는 헤딩 추가.
- `word_read_table(file_path: str, table_idx: int = 0, format: str = "markdown")`:
  - 특정 순번의 표 내용을 Markdown 표나 2차원 배열로 읽기.
- `word_append_table_row(file_path: str, table_idx: int, row_data: list[str])`:
  - 특정 표에 행 데이터 추가.
- `word_create_document(file_path: str, title: str | None = None)`:
  - 빈 워드 문서 생성.

### 4.3 PowerPoint 도구군 (`ppt_*`)
- `ppt_get_outline(file_path: str)`: 전체 슬라이드 수, 각 슬라이드의 제목 및 포함된 셰이프 종류 요약.
- `ppt_read_slide(file_path: str, slide_idx: int)`:
  - 특정 슬라이드의 모든 텍스트 상자 및 표 내용 추출.
- `ppt_add_slide(file_path: str, title: str, content: str | None = None, layout_idx: int = 1)`:
  - 제목과 본문을 포함하는 새 슬라이드 생성.
- `ppt_update_slide_text(file_path: str, slide_idx: int, shape_idx: int, text: str)`:
  - 특정 슬라이드의 지정된 셰이프 텍스트 수정.
- `ppt_create_presentation(file_path: str, title: str | None = None)`:
  - 빈 프레젠테이션 파일 생성.

---

## 5. TDD 워크플로우 (Test-Driven Development)

에이전트가 새로운 도구나 기능을 추가할 때는 **반드시** 다음 절차를 준수합니다:

1. **테스트 우선 작성 (`Red`)**:
   - `tests/test_<format>.py`에 테스트 케이스를 먼저 작성합니다.
   - 임시 파일은 반드시 pytest의 `tmp_path: pathlib.Path`를 사용하여 디스크 오염을 방지합니다.
   - 예시:
     ```python
     def test_excel_write_and_read(tmp_path):
         file_path = tmp_path / "sample.xlsx"
         excel_create_workbook(str(file_path))
         excel_write_cell(str(file_path), "Sheet", "B2", "Hello MCP")
         result = excel_read_sheet(
             str(file_path), "Sheet", start_row=2, end_row=2, start_col=2, end_col=2
         )
         assert "Hello MCP" in result
     ```
2. **최소 코드 구현 (`Green`)**:
   - `src/office_docs_mcp/<format>/` 하위에 비즈니스 로직을 구현합니다.
   - `uv run pytest tests/test_<format>.py` 실행 후 테스트가 통과하는지 확인합니다.
3. **리팩토링 및 린트 (`Refactor`)**:
   - 예외 처리(파일 누락, 인덱스 초과 등)를 강화합니다.
   - `uv run ruff check .` 및 `uv run ruff format .`을 실행하여 컨벤션을 유지합니다.
4. **MCP 도구 등록 및 검증**:
   - `server.py`에 `@mcp.tool()` 데코레이터를 추가하고 문서화 docstring을 작성합니다.

---

## 6. 개발 필수 명령어 (Development Commands)

모든 작업은 `uv`를 통해 수행합니다:

```bash
# 가상환경 동기화 및 패키지 설치
uv sync --group dev

# 새로운 패키지 추가 (예: 의존성 추가 시)
uv add openpyxl python-docx python-pptx mcp

# 테스트 실행
uv run pytest

# 특정 테스트 파일 단독 실행
uv run pytest tests/test_excel.py -v

# 린트 및 포맷 검사
uv run ruff check .
uv run ruff format --check .

# MCP 서버 로컬 실행 (Stdio 모드)
uv run office-docs-mcp serve
```

---

## 7. 구현 시 주의사항 (Guidelines & Pitfalls)

1. **1-Based vs 0-Based 인덱스**:
   - 엑셀은 사용자/LLM 친화적으로 **1-based** (행 1, 열 1 = A1)를 기본값으로 통일합니다.
   - 워드/파워포인트의 인덱스(단락, 슬라이드, 표)는 API 설명서(docstring)에 기준(0-based)을 명확하게 명시합니다.
2. **Excel 수식(Formula) 처리**:
   - `openpyxl`에서 `data_only=False`일 때는 수식 문자열(`=SUM(A1:A10)`), `data_only=True`일 때는 직전 저장된 계산값을 읽습니다.
   - 읽기 도구 파라미터에 `evaluate_formulas: bool = True` 옵션을 두어 LLM이 필요에 따라 수식 자체 또는 계산값을 선택할 수 있게 합니다.
3. **대용량 파일 방어**:
   - 기본 `read` 호출 시 반환 최대 행/단락 수에 상한선(`max_rows=100`, `max_paragraphs=50` 등)을 두고, 필요 시 페이징을 유도합니다.
4. **타입 힌트와 Docstring**:
   - FastMCP는 함수의 Type Hint와 Docstring을 파싱하여 LLM 도구의 JSON Schema와 Description을 자동 생성합니다.
   - 따라서 모든 도구 함수에는 **상세한 설명, 파라미터 의미, 사용 예시**를 반드시 Docstring에 기재합니다.
