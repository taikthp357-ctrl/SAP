# -*- coding: utf-8 -*-
"""Script đóng gói bản phát hành độc lập SAP_EMAIL_TEAMS_BRIDGE thành file ZIP."""

import os
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DIST_ZIP = BASE_DIR / "SAP_EMAIL_TEAMS_BRIDGE_RELEASE_v1.0.0.zip"

EXCLUDE_DIRS = {".git", "__pycache__", ".pytest_cache", ".vscode"}
EXCLUDE_EXTS = {".pyc", ".zip"}


def create_release_zip():
    print(f"[PACKAGE] Dang dong goi release tu: {BASE_DIR}")
    count = 0
    with zipfile.ZipFile(DIST_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            for file in files:
                file_path = Path(root) / file
                if file_path.suffix in EXCLUDE_EXTS:
                    continue
                if file_path == DIST_ZIP:
                    continue
                rel_path = file_path.relative_to(BASE_DIR)
                zf.write(file_path, arcname=str(rel_path))
                count += 1

    print(f"[SUCCESS] Da tao file dong goi ({count} files): {DIST_ZIP.name}")
    print(f"[SIZE] Dung luong: {os.path.getsize(DIST_ZIP):,} bytes")


if __name__ == "__main__":
    create_release_zip()
