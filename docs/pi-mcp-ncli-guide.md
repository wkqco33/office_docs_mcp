# Pi에서 Office Docs MCP 사용하기

## 1. 사전 조건

- `uv`가 설치되어 있고 PATH에 있어야 합니다.
- Pi에 `pi-mcp-adapter`가 설치되어 있어야 합니다.

```bash
pi install npm:pi-mcp-adapter
```

## 2. 프로젝트 설정

이 저장소의 `.mcp.json`에 `uvx` 기반 서버 설정이 포함되어 있습니다.

```json
{
  "mcpServers": {
    "office-docs": {
      "command": "uvx",
      "args": ["office-docs-mcp", "serve"],
      "lifecycle": "lazy"
    }
  }
}
```

저장소 루트에서 Pi를 시작하면 설정을 자동으로 읽습니다. 명시적으로 지정하려면 다음처럼 실행합니다.

```bash
pi --mcp-config .mcp.json
```

Pi 안에서 `/mcp`를 실행하여 서버 상태를 확인하고, 설정을 변경했다면 `/reload`를 실행합니다.

## 3. 연결 확인

서버 패키지와 도구 목록은 다음 명령으로 확인할 수 있습니다.

```bash
uvx office-docs-mcp tools --plain
```

Pi에서는 MCP 프록시 도구로 `office` 또는 `excel`, `word`, `ppt` 관련 도구를 검색하여 사용할 수 있습니다. 서버는 첫 도구 호출 때 시작됩니다(`lifecycle: lazy`).

## 4. ncli에 이 가이드 문서 추가

이 문서를 ncli 노트로 추가하려면 저장소 루트에서 실행합니다.

```bash
ncli add \
  --title "Office Docs MCP를 Pi에서 uvx로 사용하기" \
  --content-file docs/pi-mcp-ncli-guide.md \
  --category other \
  --no-input
```

본문을 직접 전달하려면 다음처럼 사용할 수 있습니다.

```bash
ncli add -t "Office Docs MCP 설정" -c "$(cat docs/pi-mcp-ncli-guide.md)" --category other --no-input
```

파일을 노트에 첨부하고 싶다면 `--file`을 추가합니다.

```bash
ncli add \
  --title "Office Docs MCP 설정 파일" \
  --content "Pi MCP 설정과 사용 가이드" \
  --file .mcp.json \
  --category other \
  --no-input
```

`ncli` 로그인이 필요하면 먼저 `ncli login`을 실행합니다. 생성된 노트는 `ncli list` 또는 `ncli search --title "Office Docs MCP"`로 확인할 수 있습니다.
