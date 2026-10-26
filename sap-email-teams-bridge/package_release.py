# -*- coding: utf-8 -*-
"""Script đóng gói bản phát hành độc lập SAP_EMAIL_TEAMS_BRIDGE thành file ZIP.

Hỗ trợ 3 bản phát hành:
1. Update Patch Release (Dung lượng siêu nhẹ ~45 KB, chỉ chứa các file cập nhật mã nguồn & cấu hình).
2. Standard Release (Dung lượng ~170 KB cho máy đã cài sẵn Python).
3. Portable All-In-One Release (Đã tích hợp sẵn Python 3.12 Runtime Portable Siêu Tinh Gọn ~13 MB).
"""

import os
import sys
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
VERSION = "v1.0.66"
PATCH_ZIP = BASE_DIR / f"SAP_EMAIL_TEAMS_BRIDGE_UPDATE_PATCH_{VERSION}.zip"
STANDARD_ZIP = BASE_DIR / f"SAP_EMAIL_TEAMS_BRIDGE_RELEASE_{VERSION}.zip"
PORTABLE_ZIP = BASE_DIR / f"SAP_EMAIL_TEAMS_BRIDGE_PORTABLE_{VERSION}.zip"

EXTERNAL_RUNTIME = Path(r"G:\AI-Manager\work\releases\ma_control_portable_build\runtime")
LOCAL_RUNTIME = BASE_DIR / "runtime"

EXCLUDE_SOURCE_DIRS = {".git", "__pycache__", ".pytest_cache", ".vscode", "runtime"}
EXCLUDE_EXTS = {".pyc", ".zip", ".pdb", ".db", ".sqlite", ".log"}

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


def create_patch_zip():
    """Tạo gói Patch cập nhật siêu nhẹ chỉ gồm các file vừa thay đổi."""
    print(f"\n[PACKAGE-PATCH] Đang đóng gói Bản Cập Nhật Siêu Nhẹ ({VERSION}) - Chỉ các file vừa update...")
    count = 0
    patch_includes = [
        "main.py",
        "sap_email_teams_bridge.py",
        "README.md",
        "START_BRIDGE.cmd",
        "START_BACKGROUND.vbs"
    ]
    with zipfile.ZipFile(PATCH_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for fname in patch_includes:
            fpath = BASE_DIR / fname
            if fpath.exists():
                zf.write(fpath, arcname=fname)
                count += 1

        for folder in ["app", "tests", "docs"]:
            folder_path = BASE_DIR / folder
            if folder_path.exists():
                for root, _, files in os.walk(folder_path):
                    if "__pycache__" in root or ".pytest_cache" in root:
                        continue
                    for file in files:
                        if file.endswith(".py"):
                            file_path = Path(root) / file
                            rel_path = file_path.relative_to(BASE_DIR)
                            zf.write(file_path, arcname=str(rel_path))
                            count += 1

    print(f"[SUCCESS] Đã tạo file Update Patch Siêu Nhẹ ({count} files): {PATCH_ZIP.name}")
    print(f"[SIZE] Dung lượng: {os.path.getsize(PATCH_ZIP):,} bytes ({os.path.getsize(PATCH_ZIP)/1024:.1f} KB)")


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
    create_patch_zip()
    create_standard_zip()
    create_portable_zip()

    # Tự động đồng bộ ra thư mục Downloads và Artifacts
    import shutil
    destinations = [
        Path(r"C:\Users\taikt\Downloads"),
        Path(r"C:\Users\taikt\.gemini\antigravity\brain\393da05a-90fd-4275-aec3-dee4a841da63")
    ]
    for dest in destinations:
        if dest.exists():
            for zf in [PATCH_ZIP, STANDARD_ZIP, PORTABLE_ZIP]:
                if zf.exists():
                    try:
                        shutil.copy2(zf, dest / zf.name)
                        print(f"[COPIED] {zf.name} -> {dest}")
                    except Exception as e:
                        print(f"[COPY WARN] Không thể copy sang {dest}: {e}")
