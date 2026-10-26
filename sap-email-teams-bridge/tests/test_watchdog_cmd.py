# -*- coding: utf-8 -*-
"""Kiểm thử cấu trúc Watchdog trong START_BRIDGE.cmd."""
from pathlib import Path


def test_start_bridge_cmd_watchdog():
    cmd_path = Path(__file__).resolve().parent.parent / "START_BRIDGE.cmd"
    assert cmd_path.exists(), f"File {cmd_path} does not exist"

    content = cmd_path.read_text(encoding="utf-8", errors="ignore")

    # 1. Phải có nhãn vòng lặp WATCHDOG_LOOP và lệnh quay lại vòng lặp
    assert ":WATCHDOG_LOOP" in content, "Thieu nhan :WATCHDOG_LOOP trong START_BRIDGE.cmd"
    assert "goto :WATCHDOG_LOOP" in content, "Thieu lenh goto :WATCHDOG_LOOP trong START_BRIDGE.cmd"

    # 2. Phải có thời gian chờ 5s tự động khởi động lại
    assert "5" in content, "Thieu thoi gian cho 5s trong START_BRIDGE.cmd"
    assert ("timeout /t 5" in content or "ping 127.0.0.1 -n 6" in content or "ping -n 6 127.0.0.1" in content), "Thieu co che delay 5s giua cac lan restart"

    # 3. Phải hiển thị thông báo watchdog
    assert "[WATCHDOG" in content, "Thieu thong bao [WATCHDOG cho nguoi van hanh"
