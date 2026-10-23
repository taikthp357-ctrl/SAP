# -*- coding: utf-8 -*-
"""Cấu hình ứng dụng SAP Email Teams Bridge."""

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
    """Tải cấu hình từ config.json, nếu không có thì trả về mặc định."""
    if config_path is None:
        config_path = Path(__file__).resolve().parent.parent / "config.json"

    if config_path.exists():
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                merged = dict(DEFAULT_CONFIG)
                merged.update(data)
                return merged
        except Exception as e:
            print(f"[CONFIG WARNING] Không đọc được config.json ({e}), dùng mặc định.")
            return dict(DEFAULT_CONFIG)

    return dict(DEFAULT_CONFIG)
