# office_docs_mcp

LLM(대형 언어 모델)이 오피스 문서(Excel, Word, PowerPoint)를 안정적이고 토큰 효율적으로 읽고 쓸 수 있도록 지원하는 **Model Context Protocol (MCP)** 서버입니다.

## 주요 기능

- **Excel (`.xlsx`, `.xlsm`)**:
  - `excel_get_metadata`: 시트 목록, 크기, 컬럼 요약 조회
  - `excel_read_sheet`: 지정된 행/열 범위(1-based)를 Markdown 테이블 또는 2차원 배열로 읽기
  - `excel_write_cell`: 단일 셀(`A1` 등) 값 쓰기
  - `excel_write_range`: 시작 셀부터 2차원 데이터 연속 기입
  - `excel_append_rows`: 시트 마지막 행 뒤에 데이터 추가
  - `excel_create_workbook`: 새 빈 엑셀 파일 생성
- **Word (`.docx`)**:
  - `word_get_outline`: 문서 헤딩(제목) 목록 및 단락/표 개수 조회
  - `word_read_paragraphs`: 단락 슬라이싱 읽기 (0-based 페이징)
  - `word_read_table`: 표 내용을 Markdown 테이블로 읽기
  - `word_append_paragraph`: 문서 끝에 단락/헤딩 추가
  - `word_append_table_row`: 특정 표에 새 행 추가
  - `word_write_table_cell`: 특정 표의 셀 텍스트 수정
  - `word_create_document`: 새 워드 문서 생성
- **PowerPoint (`.pptx`)**:
  - `ppt_get_outline`: 전체 슬라이드 수 및 각 슬라이드 제목/셰이프 요약 조회
  - `ppt_read_slide`: 특정 슬라이드의 텍스트 상자 및 표 내용 추출
  - `ppt_add_slide`: 제목과 본문을 포함하는 새 슬라이드 추가
  - `ppt_update_slide_text`: 슬라이드 내 특정 셰이프 텍스트 수정
  - `ppt_create_presentation`: 새 프레젠테이션 파일 생성

## 시작하기

### 1. 설치 및 동기화 (uv)
```bash
uv sync --group dev
```

### 2. MCP 서버 실행 (Stdio 모드)
```bash
uv run office-docs-mcp serve
```

### 3. Claude Desktop / Antigravity 등 MCP 클라이언트 연동 설정 예시
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

## 개발 및 테스트 (TDD)

```bash
# 전체 단위 및 통합 테스트 실행
uv run pytest

# 린트 및 코드 포맷팅
uv run ruff check .
uv run ruff format .
```

상세한 개발 지침 및 설계 원칙은 [AGENTS.md](AGENTS.md) 문서를 참고하세요.
