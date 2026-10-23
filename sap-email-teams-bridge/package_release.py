# -*- coding: utf-8 -*-
"""Script đóng gói bản phát hành độc lập SAP_EMAIL_TEAMS_BRIDGE thành file ZIP.

Hỗ trợ 2 bản phát hành:
1. Standard Release (Dung lượng siêu nhẹ ~63 KB cho máy đã cài sẵn Python).
2. Portable All-In-One Release (Đã tích hợp sẵn Python 3.12 Runtime Portable ~47 MB,
   không cần cài đặt Python, không cần quyền Admin, giải nén là chạy ngay).
"""

import os
import sys
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
STANDARD_ZIP = BASE_DIR / "SAP_EMAIL_TEAMS_BRIDGE_RELEASE_v1.0.0.zip"
PORTABLE_ZIP = BASE_DIR / "SAP_EMAIL_TEAMS_BRIDGE_PORTABLE_v1.0.0.zip"

EXTERNAL_RUNTIME = Path(r"G:\AI-Manager\work\releases\ma_control_portable_build\runtime")
LOCAL_RUNTIME = BASE_DIR / "runtime"

EXCLUDE_DIRS = {".git", "__pycache__", ".pytest_cache", ".vscode", "runtime"}
EXCLUDE_EXTS = {".pyc", ".zip"}


def create_standard_zip():
    print(f"\n[PACKAGE-STANDARD] Đang đóng gói Standard Release từ: {BASE_DIR}")
    count = 0
    with zipfile.ZipFile(STANDARD_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            for file in files:
                file_path = Path(root) / file
                if file_path.suffix in EXCLUDE_EXTS:
                    continue
                if file_path in (STANDARD_ZIP, PORTABLE_ZIP):
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

    print(f"\n[PACKAGE-PORTABLE] Đang đóng gói Portable All-In-One (Kèm Python 3.12 Runtime)...")
    count = 0
    with zipfile.ZipFile(PORTABLE_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1. Gói các file mã nguồn & công cụ
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            for file in files:
                file_path = Path(root) / file
                if file_path.suffix in EXCLUDE_EXTS:
                    continue
                if file_path in (STANDARD_ZIP, PORTABLE_ZIP):
                    continue
                rel_path = file_path.relative_to(BASE_DIR)
                zf.write(file_path, arcname=str(rel_path))
                count += 1

        # 2. Gói toàn bộ thư mục runtime/
        for root, dirs, files in os.walk(runtime_source):
            for file in files:
                file_path = Path(root) / file
                rel_path = Path("runtime") / file_path.relative_to(runtime_source)
                zf.write(file_path, arcname=str(rel_path))
                count += 1

    print(f"[SUCCESS] Đã tạo file Portable All-In-One ({count} files): {PORTABLE_ZIP.name}")
    print(f"[SIZE] Dung lượng: {os.path.getsize(PORTABLE_ZIP):,} bytes ({os.path.getsize(PORTABLE_ZIP)/1024/1024:.2f} MB)")


if __name__ == "__main__":
    create_standard_zip()
    create_portable_zip()
