import os
import sys
from pathlib import Path
from datetime import datetime
from typing import Any

from mcp.server.fastmcp import FastMCP

# pytsk3, pyewf 임포트
try:
    import pytsk3
except ImportError:
    print("[ERROR] pytsk3 is not installed. Install it with: pip install pytsk3", file=sys.stderr)
    sys.exit(1)

try:
    import pyewf
except ImportError:
    print("[ERROR] pyewf is not installed. Install it with: pip install pyewf-python", file=sys.stderr)
    sys.exit(1)


mcp = FastMCP("consolehost-parser")

VERSION = "1.0.0"


class EWFImgInfo(pytsk3.Img_Info):

    def __init__(self, ewf_handle):
        self._ewf_handle = ewf_handle
        super(EWFImgInfo, self).__init__(url="", type=pytsk3.TSK_IMG_TYPE_EXTERNAL)

    def close(self):
        self._ewf_handle.close()

    def read(self, offset, size):
        self._ewf_handle.seek(offset)
        return self._ewf_handle.read(size)

    def get_size(self):
        return self._ewf_handle.get_media_size()


def open_image(image_path: str):
    image_path = os.path.abspath(image_path)

    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image file not found: {image_path}")

    ext = os.path.splitext(image_path)[1].lower()

    if ext in ['.e01', '.ex01', '.s01']:
        filenames = pyewf.glob(image_path)
        ewf_handle = pyewf.handle()
        ewf_handle.open(filenames)
        return EWFImgInfo(ewf_handle)
    else:
        return pytsk3.Img_Info(image_path)


def find_consolehost_history(fs, path: str = "/", results: list = None) -> list:
    if results is None:
        results = []

    target_filename = "consolehost_history.txt"
    target_path_parts = ["appdata", "roaming", "microsoft", "windows", "powershell", "psreadline"]

    try:
        directory = fs.open_dir(path)
    except Exception:
        return results

    for entry in directory:
        try:
            name = entry.info.name.name
            if isinstance(name, bytes):
                name = name.decode('utf-8', errors='replace')

            if name in ['.', '..']:
                continue

            if path == "/":
                full_path = f"/{name}"
            else:
                full_path = f"{path}/{name}"

            # 파일인 경우
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_REG:
                if name.lower() == target_filename:
                    path_lower = full_path.lower()
                    if all(part in path_lower for part in target_path_parts):
                        file_size = entry.info.meta.size
                        results.append({
                            'path': full_path,
                            'size': file_size,
                            'entry': entry
                        })

            # 디렉토리인 경우 - 필요한 경로만 탐색
            elif entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                name_lower = name.lower()
                if name_lower in ['users', 'documents and settings'] or \
                   name_lower in target_path_parts or \
                   'appdata' in full_path.lower():
                    find_consolehost_history(fs, full_path, results)

                elif path.lower() in ['/users', '/documents and settings']:
                    find_consolehost_history(fs, full_path, results)

        except Exception:
            continue

    return results


def extract_file_content(entry) -> bytes | None:
    try:
        file_size = entry.info.meta.size
        offset = 0
        chunk_size = 1024 * 1024  # 1MB chunks
        data = b""

        while offset < file_size:
            available = min(chunk_size, file_size - offset)
            chunk = entry.read_random(offset, available)
            if not chunk:
                break
            data += chunk
            offset += len(chunk)

        return data
    except Exception:
        return None


def get_username_from_path(file_path: str) -> str:
    path_parts = file_path.split('/')
    for i, part in enumerate(path_parts):
        if part.lower() == 'users' and i + 1 < len(path_parts):
            return path_parts[i + 1]
    return "unknown"


def parse_commands(content_bytes: bytes) -> tuple[list[dict], str | None]:
    encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'latin-1']

    for encoding in encodings:
        try:
            content = content_bytes.decode(encoding)
            lines = content.splitlines()
            commands = []

            for line_num, line in enumerate(lines, start=1):
                stripped = line.strip()
                if stripped:
                    commands.append({
                        "line_number": line_num,
                        "command": stripped
                    })

            return commands, encoding
        except (UnicodeDecodeError, UnicodeError):
            continue

    return [], None


@mcp.tool()
def extract_consolehost_history(image_path: str, output_dir: str) -> dict[str, Any]:

    # 이미지 파일 존재 확인
    if not Path(image_path).exists():
        return {
            "success": False,
            "error": f"Image file not found: {image_path}"
        }

    # 이미지 열기
    try:
        img = open_image(image_path)
    except Exception as e:
        return {
            "success": False,
            "error": f"Failed to open image: {e}"
        }

    # 파티션 정보 확인
    try:
        volume = pytsk3.Volume_Info(img)
        partitions = list(volume)
    except Exception:
        partitions = None

    all_results = []

    # 파티션이 있는 경우
    if partitions:
        for part_num, part in enumerate(partitions):
            if part.flags != pytsk3.TSK_VS_PART_FLAG_ALLOC:
                continue

            part_desc = part.desc
            if isinstance(part_desc, bytes):
                part_desc = part_desc.decode('utf-8', errors='replace')

            try:
                fs = pytsk3.FS_Info(img, offset=part.start * 512)
                results = find_consolehost_history(fs)

                for result in results:
                    result['partition'] = part_desc
                    result['partition_num'] = part_num

                all_results.extend(results)
            except Exception:
                continue
    else:
        # 단일 파일시스템
        try:
            fs = pytsk3.FS_Info(img)
            results = find_consolehost_history(fs)
            all_results.extend(results)
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not process filesystem: {e}"
            }

    # 결과가 없는 경우
    if not all_results:
        return {
            "success": True,
            "image_path": image_path,
            "files_found": 0,
            "message": "No ConsoleHost_history.txt files found",
            "extracted_files": []
        }

    # 출력 디렉토리 생성
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    extracted_files = []

    for i, result in enumerate(all_results):
        username = get_username_from_path(result['path'])
        output_filename = f"ConsoleHost_history_{username}_{i+1}.txt"
        output_path = Path(output_dir) / output_filename

        # 파일 내용 추출
        content_bytes = extract_file_content(result['entry'])

        if content_bytes:
            # 파일 저장
            with open(output_path, 'wb') as f:
                f.write(content_bytes)

            # 명령어 파싱
            commands, used_encoding = parse_commands(content_bytes)

            # 메타 파일 저장
            meta_path = str(output_path) + ".meta.txt"
            with open(meta_path, 'w', encoding='utf-8') as f:
                f.write(f"Source Image: {image_path}\n")
                f.write(f"Original Path: {result['path']}\n")
                f.write(f"File Size: {result['size']} bytes\n")
                f.write(f"Extraction Time: {datetime.now().isoformat()}\n")
                if 'partition' in result:
                    f.write(f"Partition: {result['partition']}\n")

            extracted_files.append({
                "username": username,
                "source_path": result['path'],
                "output_path": str(output_path.absolute()),
                "file_size": result['size'],
                "partition": result.get('partition', 'N/A'),
                "encoding": used_encoding,
                "command_count": len(commands),
                "commands": commands
            })

    return {
        "success": True,
        "image_path": image_path,
        "output_dir": output_dir,
        "files_found": len(all_results),
        "files_extracted": len(extracted_files),
        "extracted_files": extracted_files
    }


if __name__ == "__main__":
    mcp.run()
