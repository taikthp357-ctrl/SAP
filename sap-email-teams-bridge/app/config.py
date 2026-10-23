# -*- coding: utf-8 -*-
"""Cấu hình ứng dụng SAP Email Teams Bridge kèm bộ tự sửa lỗi định dạng JSON."""

import re
import json
from pathlib import Path

DEFAULT_CONFIG = {
    "teams_webhook_url": "https://prod-xx.southeastasia.logic.azure.com:443/workflows/YOUR_WEBHOOK_URL_HERE",
    "outlook_folder_name": "SAP_BOT",
    "poll_interval_seconds": 10,
    "default_plant": "3093",
    "max_alv_rows": 500,
    "keep_alive_idle_seconds": 900
}


def load_config(config_path: Path = None) -> dict:
    """Tải cấu hình từ config.json, tự động sửa lỗi gõ nhầm dấu ngoặc kép hoặc quên ngoặc URL."""
    if config_path is None:
        config_path = Path(__file__).resolve().parent.parent / "config.json"

    if not config_path.exists():
        return dict(DEFAULT_CONFIG)

    try:
        with open(config_path, "r", encoding="utf-8-sig") as f:
            raw_content = f.read()

        # 1. Thử parse trực tiếp chuẩn JSON
        try:
            data = json.loads(raw_content)
            merged = dict(DEFAULT_CONFIG)
            merged.update(data)
            return merged
        except Exception:
            pass

        # 2. Cơ chế tự động sửa lỗi phổ biến của người dùng khi sửa bằng Notepad
        fixed = raw_content
        # Thay thế dấu nháy kép cong của Windows/Word bằng nháy thẳng
        fixed = fixed.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
        # Tự bọc ngoặc kép nếu người dùng dán URL không có dấu ngoặc kép
        fixed = re.sub(
            r'("teams_webhook_url"\s*:\s*)(https?://[^\s,}\"]+)',
            r'\1"\2"',
            fixed
        )
        # Xóa dấu phẩy thừa ở phần tử cuối cùng trước dấu }
        fixed = re.sub(r',\s*\}', '}', fixed)

        data = json.loads(fixed)
        print("[CONFIG] Đã tự động hiệu chỉnh và nạp thành công cấu hình từ config.json!")
        merged = dict(DEFAULT_CONFIG)
        merged.update(data)
        return merged

    except Exception as e:
        print(f"[CONFIG WARNING] Không đọc được config.json ({e}), dùng mặc định.")
        return dict(DEFAULT_CONFIG)
