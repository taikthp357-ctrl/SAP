# -*- coding: utf-8 -*-
"""Cấu hình ứng dụng SAP Email Teams Bridge kèm bộ tự sửa lỗi định dạng JSON."""

import re
import json
from pathlib import Path

DEFAULT_CONFIG = {
    "data_dir": "data",
    "allow_sap_writes": True,
    "enable_multi_tab": True,
    "max_batch_tabs": 2,
    "sap_system": "",
    "sap_client": "",
    "sap_user": "",
    "trusted_intake_senders": [],
    "allow_nvl_reservations": False,
    "teams_webhook_url": "https://prod-xx.southeastasia.logic.azure.com:443/workflows/YOUR_WEBHOOK_URL_HERE",
    "outlook_folder_name": "SAP_BOT",
    "poll_interval_seconds": 10,
    "default_plant": "3093",
    "max_alv_rows": 15000,
    "keep_alive_idle_seconds": 900,
    "sap_number_format": "AUTO",
    "strict_session_safety": True,
    "auto_create_session": True,
    "prefer_dedicated_session": True,
    "user_reserved_sessions": 3,
    "max_sap_sessions": 6,
    "max_bot_sessions": 3,
    "sap_session_index": 0,
    "sap_connection_index": 0,
    "max_webhook_retries": 3,
    "allowed_senders": [],
    "warehouse_slocs": ["1005", "3093"],
    "sloc_master": {
        "1005": "Kho NVL",
        "1013": "Kho Production COI,BLN",
        "200C": "Kho hóa chất,keo",
        "201Q": "GR hàng OS (Tier1) cho xưởng khác",
        "201R": "Sloc trống",
        "3093": "Kho FG",
        "600O": "Hàng lỗi thời",
        "60VR": "Hàng return NCC",
        "80B2": "Hàng hỏng hủy",
        "9004": "Hàng thiếu/thừa của NCC",
        "3001": "Line Hàn",
        "701A": "Hàng lỗi công đoạn Hàn",
        "9999": "Sloc QC",
        "3002": "Line Blank",
        "3003": "Line Dập",
        "TRANSIT": "Hàng đang trung chuyển (In-Transit)",
        "4007": "Kho Xưởng 4007",
        "4004": "Kho Xưởng 4004",
        "604D": "Hàng lỗi công đoạn dập",
        "701B": "Hàng hold line Dập",
        "3006": "Sloc PE",
        "P000": "Try out công đoạn Dập",
        "200E": "Waiting"
    }
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
        raise ValueError("config.json không hợp lệ; sửa cấu hình trước khi chạy Bridge") from e
