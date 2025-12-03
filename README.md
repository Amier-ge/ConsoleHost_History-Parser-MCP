# ConsoleHost Parser MCP Server

ConsoleHost_history.txt Parser를 MCP(Model Context Protocol) Server로 제공


**대상 파일 위치:**
```
%USERPROFILE%\AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt
```

## 요구 사항

### 디렉토리 구조

이 MCP 서버는 `ConsoleHost_Parser` 라이브러리에 의존합니다. **두 폴더가 반드시 같은 디렉토리 내에 위치해야 합니다.**

```
Parent_Directory/
├── ConsoleHost_Parser/         ← 원본 ConsoleHost 파서 라이브러리
│   ├── consoleHost_parser.py
│   └── ...
│
└── ConsoleHost_Parser_MCP/     ← 이 MCP 서버
    ├── mcp_server.py
    └── ...
```

> ⚠️ `ConsoleHost_Parser` 폴더가 없거나 다른 위치에 있으면 이미지 추출 기능이 정상적으로 작동하지 않습니다.

## 설치

```bash
pip install -r requirements.txt
```

## 사용법

### MCP 서버 실행

```bash
python mcp_server.py
```

### 환경 세팅

```json
{
  "mcpServers": {
    "consolehost-parser": {
      "command": "[Python Path]",
      "args": [
        "[mcp_server.py Path]"
      ],
      "env": {
        "PYTHONPATH": "C:/Users/home/Desktop/Made_Tools/ConsoleHost_Parser_MCP"
      }
    }
  }
}
```

## 제공 도구

| 도구 | 설명 |
|------|------|
| `extract_consolehost_history` | ConsoleHost_history.txt 파일을 JSON으로 파싱 |
| `extract_from_image` | E01/DD 이미지에서 파일 추출 및 JSON 파싱 |
| `get_info` | 도구 정보 조회 |

## 출력 예시

### extract_consolehost_history

```json
{
  "success": true,
  "file_path": "C:\\...\\ConsoleHost_history.txt",
  "file_size_bytes": 1234,
  "total_lines": 50,
  "command_count": 45,
  "encoding": "utf-8",
  "commands": [
    {"line_number": 1, "command": "Get-Process"},
    {"line_number": 2, "command": "cd C:\\Users"},
    {"line_number": 3, "command": "dir"}
  ]
}
```

### extract_from_image

```json
{
  "success": true,
  "image_path": "C:\\Evidence\\disk.E01",
  "files_found": 2,
  "extracted_files": [
    {
      "username": "Administrator",
      "source_path": "/Users/Administrator/.../ConsoleHost_history.txt",
      "command_count": 30,
      "commands": [...]
    }
  ]
}
```

## 지원 인코딩

- UTF-8
- UTF-8-BOM
- CP949
- EUC-KR
- Latin-1

## 지원 이미지

- E01
- DD/Raw

## 작성자

Amier-ge
