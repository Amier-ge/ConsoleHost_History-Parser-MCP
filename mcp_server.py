import sys
import os
from pathlib import Path
from datetime import datetime
from typing import Any

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("consolehost-parser")

VERSION = "1.0.0"


def parse_consolehost_history(file_path: str) -> list[dict[str, Any]]:
    commands = []

    encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'latin-1']
    content = None
    used_encoding = None

    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()
                used_encoding = encoding
                break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if content is None:
        raise ValueError(f"Failed to decode file with any of the encodings: {encodings}")

    lines = content.splitlines()

    for line_number, line in enumerate(lines, start=1):
        stripped = line.strip()
        if stripped:  # 빈 줄은 제외
            commands.append({
                "line_number": line_number,
                "command": stripped,
                "raw_line": line
            })

    return commands


@mcp.tool()
def extract_consolehost_history(
    input_path: str,
    include_empty_lines: bool = False,
    include_line_numbers: bool = True
) -> dict[str, Any]:

    path = Path(input_path)

    if not path.exists():
        return {
            "success": False,
            "error": f"File not found: {input_path}"
        }

    if not path.is_file():
        return {
            "success": False,
            "error": f"Path is not a file: {input_path}"
        }

    try:
        # 다양한 인코딩 시도
        encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'latin-1']
        content = None
        used_encoding = None

        for encoding in encodings:
            try:
                with open(input_path, 'r', encoding=encoding) as f:
                    content = f.read()
                    used_encoding = encoding
                    break
            except (UnicodeDecodeError, UnicodeError):
                continue

        if content is None:
            return {
                "success": False,
                "error": f"Failed to decode file with any of the encodings: {encodings}"
            }

        lines = content.splitlines()
        total_lines = len(lines)

        commands = []
        for line_number, line in enumerate(lines, start=1):
            stripped = line.strip()

            if not stripped and not include_empty_lines:
                continue

            command_entry = {
                "command": stripped if stripped else ""
            }

            if include_line_numbers:
                command_entry["line_number"] = line_number

            commands.append(command_entry)

        file_stats = path.stat()

        return {
            "success": True,
            "file_path": str(path.absolute()),
            "file_size_bytes": file_stats.st_size,
            "total_lines": total_lines,
            "command_count": len([c for c in commands if c["command"]]),
            "encoding": used_encoding,
            "commands": commands
        }

    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }


@mcp.tool()
def extract_from_image(
    image_path: str,
    output_dir: str,
    partition: int | None = None
) -> dict[str, Any]:

    try:
        import pytsk3
    except ImportError:
        return {
            "success": False,
            "error": "pytsk3 is not installed. Install it with: pip install pytsk3"
        }

    try:
        import pyewf
    except ImportError:
        return {
            "success": False,
            "error": "pyewf is not installed. Install it with: pip install pyewf-python"
        }

    if not Path(image_path).exists():
        return {
            "success": False,
            "error": f"Image file not found: {image_path}"
        }

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

    def open_image(img_path):
        img_path = os.path.abspath(img_path)
        ext = os.path.splitext(img_path)[1].lower()

        if ext in ['.e01', '.ex01', '.s01']:
            filenames = pyewf.glob(img_path)
            ewf_handle = pyewf.handle()
            ewf_handle.open(filenames)
            return EWFImgInfo(ewf_handle)
        else:
            return pytsk3.Img_Info(img_path)

    def find_consolehost_files(fs, path="/", results=None):
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

                full_path = f"{path}/{name}" if path != "/" else f"/{name}"

                if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_REG:
                    if name.lower() == target_filename:
                        path_lower = full_path.lower()
                        if all(part in path_lower for part in target_path_parts):
                            results.append({
                                'path': full_path,
                                'size': entry.info.meta.size,
                                'entry': entry
                            })

                elif entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                    name_lower = name.lower()
                    if name_lower in ['users', 'documents and settings'] or \
                       name_lower in target_path_parts or \
                       'appdata' in full_path.lower():
                        find_consolehost_files(fs, full_path, results)
                    elif path.lower() in ['/users', '/documents and settings']:
                        find_consolehost_files(fs, full_path, results)
            except Exception:
                continue

        return results

    def extract_file_content(entry):
        try:
            file_size = entry.info.meta.size
            data = entry.read_random(0, file_size)
            return data
        except Exception as e:
            return None

    def get_username_from_path(file_path):
        path_parts = file_path.split('/')
        for i, part in enumerate(path_parts):
            if part.lower() == 'users' and i + 1 < len(path_parts):
                return path_parts[i + 1]
        return "unknown"

    try:
        img = open_image(image_path)
    except Exception as e:
        return {
            "success": False,
            "error": f"Failed to open image: {e}"
        }

    try:
        volume = pytsk3.Volume_Info(img)
        partitions = list(volume)
    except Exception:
        partitions = None

    all_results = []
    extracted_files = []

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    if partitions:
        for part in partitions:
            if part.flags != pytsk3.TSK_VS_PART_FLAG_ALLOC:
                continue

            part_num = partitions.index(part)
            if partition is not None and part_num != partition:
                continue

            part_desc = part.desc
            if isinstance(part_desc, bytes):
                part_desc = part_desc.decode('utf-8', errors='replace')

            try:
                fs = pytsk3.FS_Info(img, offset=part.start * 512)
                results = find_consolehost_files(fs)

                for result in results:
                    result['partition'] = part_desc
                    result['partition_num'] = part_num

                all_results.extend(results)
            except Exception:
                continue
    else:
        try:
            fs = pytsk3.FS_Info(img)
            results = find_consolehost_files(fs)
            all_results.extend(results)
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not process filesystem: {e}"
            }

    for i, result in enumerate(all_results):
        username = get_username_from_path(result['path'])
        output_filename = f"ConsoleHost_history_{username}_{i+1}.txt"
        output_path = Path(output_dir) / output_filename

        content_bytes = extract_file_content(result['entry'])

        if content_bytes:
            with open(output_path, 'wb') as f:
                f.write(content_bytes)

            commands = []
            encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'latin-1']
            used_encoding = None

            for encoding in encodings:
                try:
                    content = content_bytes.decode(encoding)
                    used_encoding = encoding
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue

            if content:
                lines = content.splitlines()
                for line_num, line in enumerate(lines, start=1):
                    stripped = line.strip()
                    if stripped:
                        commands.append({
                            "line_number": line_num,
                            "command": stripped
                        })

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


@mcp.tool()
def get_info() -> dict[str, Any]:

    return {
        "name": "ConsoleHost History Parser",
        "version": VERSION,
        "author": "Amier-ge",
        "description": "PowerShell ConsoleHost_history.txt Extraction and Parsing Tool",
        "capabilities": [
            "extract_consolehost_history - Parse ConsoleHost_history.txt file to JSON",
            "extract_from_image - Extract and parse from E01/DD forensic images",
            "get_info - Get server information"
        ],
        "target_file": "ConsoleHost_history.txt",
        "default_location": "%USERPROFILE%\\AppData\\Roaming\\Microsoft\\Windows\\PowerShell\\PSReadLine\\ConsoleHost_history.txt",
        "supported_encodings": ["utf-8", "utf-8-sig", "cp949", "euc-kr", "latin-1"],
        "supported_images": ["E01", "DD", "Raw"]
    }


if __name__ == "__main__":
    mcp.run()
