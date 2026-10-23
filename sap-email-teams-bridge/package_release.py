# -*- coding: utf-8 -*-
"""Script đóng gói bản phát hành độc lập SAP_EMAIL_TEAMS_BRIDGE thành file ZIP.

Hỗ trợ 2 bản phát hành:
1. Standard Release (Dung lượng siêu nhẹ ~68 KB cho máy đã cài sẵn Python).
2. Portable All-In-One Release (Đã tích hợp sẵn Python 3.12 Runtime Portable Siêu Tinh Gọn ~12.99 MB,
   đã loại bỏ các thư viện rác/test không dùng, không cần cài đặt Python, không cần quyền Admin).
"""

import os
import sys
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
VERSION = "v1.0.7"
STANDARD_ZIP = BASE_DIR / f"SAP_EMAIL_TEAMS_BRIDGE_RELEASE_{VERSION}.zip"
PORTABLE_ZIP = BASE_DIR / f"SAP_EMAIL_TEAMS_BRIDGE_PORTABLE_{VERSION}.zip"

EXTERNAL_RUNTIME = Path(r"G:\AI-Manager\work\releases\ma_control_portable_build\runtime")
LOCAL_RUNTIME = BASE_DIR / "runtime"

EXCLUDE_SOURCE_DIRS = {".git", "__pycache__", ".pytest_cache", ".vscode", "runtime"}
EXCLUDE_EXTS = {".pyc", ".zip", ".pdb"}

# Các thư mục rác / test / tài liệu dev không sử dụng trong bridge để tối ưu dung lượng
EXCLUDE_RUNTIME_DIRS = {
    "__pycache__", ".pytest_cache", "test", "tests", "ensurepip", "idlelib", "tkinter",
    "turtledemo", "PIL", "pygments", "pydantic", "pydantic_core", "_pytest",
    "pytest", "fastapi", "starlette", "uvicorn", "anyio", "httpcore", "httpx",
    "click", "venv", "pydoc_data", "lib2to3", "_testcapi", "_testinternalcapi",
    "unittest"
}

# Các file dll/pyd/txt thừa không cần thiết (GIỮ LẠI sqlite3.dll và _sqlite3.pyd cho hàng đợi bền vững)
EXCLUDE_RUNTIME_FILES = {
    "NEWS.txt", "tcl86t.dll", "tk86t.dll", "_tkinter.pyd",
    "_ctypes_test.pyd", "_testcapi.pyd", "_testinternalcapi.pyd",
    "_testmultiphase.pyd", "xxlimited.pyd", "xxlimited_35.pyd"
}


def clean_old_zips():
    """Xóa các bản zip cũ để tránh nhầm lẫn."""
    for zf in BASE_DIR.glob("*.zip"):
        try:
            zf.unlink()
            print(f"[CLEANUP] Đã xóa bản zip cũ: {zf.name}")
        except Exception:
            pass


def create_standard_zip():
    print(f"\n[PACKAGE-STANDARD] Đang đóng gói Standard Release ({VERSION}) từ: {BASE_DIR}")
    count = 0
    with zipfile.ZipFile(STANDARD_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_SOURCE_DIRS]
            for file in files:
                file_path = Path(root) / file
                if file_path.suffix in EXCLUDE_EXTS:
                    continue
                if file_path.name.endswith(".zip"):
                    continue
                rel_path = file_path.relative_to(BASE_DIR)
                zf.write(file_path, arcname=str(rel_path))
                count += 1

    print(f"[SUCCESS] Đã tạo file Standard ({count} files): {STANDARD_ZIP.name}")
    print(f"[SIZE] Dung lượng: {os.path.getsize(STANDARD_ZIP):,} bytes ({os.path.getsize(STANDARD_ZIP)/1024:.1f} KB)")


def create_portable_zip():
    runtime_source = LOCAL_RUNTIME if LOCAL_RUNTIME.exists() else EXTERNAL_RUNTIME
    if not runtime_source.exists():
        print(f"[SKIP-PORTABLE] Không tìm thấy runtime tại {runtime_source}")
        return

    print(f"\n[PACKAGE-PORTABLE] Đang đóng gói Portable All-In-One ({VERSION} Kèm Python 3.12 Runtime Siêu Tinh Gọn)...")
    count = 0
    with zipfile.ZipFile(PORTABLE_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        # 1. Gói các file mã nguồn & công cụ
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_SOURCE_DIRS]
            for file in files:
                file_path = Path(root) / file
                if file_path.suffix in EXCLUDE_EXTS:
                    continue
                if file_path.name.endswith(".zip"):
                    continue
                rel_path = file_path.relative_to(BASE_DIR)
                zf.write(file_path, arcname=str(rel_path))
                count += 1

        # 2. Gói thư mục runtime/ sau khi lọc bỏ rác
        for root, dirs, files in os.walk(runtime_source):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_RUNTIME_DIRS]
            for file in files:
                if file in EXCLUDE_RUNTIME_FILES:
                    continue
                file_path = Path(root) / file
                if file_path.suffix in EXCLUDE_EXTS:
                    continue
                rel_path = Path("runtime") / file_path.relative_to(runtime_source)
                zf.write(file_path, arcname=str(rel_path))
                count += 1

    print(f"[SUCCESS] Đã tạo file Portable All-In-One ({count} files): {PORTABLE_ZIP.name}")
    print(f"[SIZE] Dung lượng: {os.path.getsize(PORTABLE_ZIP):,} bytes ({os.path.getsize(PORTABLE_ZIP)/1024/1024:.2f} MB)")


if __name__ == "__main__":
    clean_old_zips()
    create_standard_zip()
    create_portable_zip()
