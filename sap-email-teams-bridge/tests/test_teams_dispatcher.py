# -*- coding: utf-8 -*-
"""Unit tests for TeamsDispatcher."""

import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.teams_dispatcher import TeamsDispatcher, MAX_TEAMS_PAYLOAD_BYTES


def test_build_html_table_empty():
    html = TeamsDispatcher.build_html_table([], {"col": "Col Label"})
    assert "Không tìm thấy dữ liệu" in html


def test_build_html_table_with_records():
    records = [
        {"product": "1002345", "bin": "A-01-01", "qty": "100"},
        {"product": "1002346", "bin": "B-02-03", "qty": "50"}
    ]
    headers_map = {
        "product": "Mã hàng",
        "bin": "Vị trí (Bin)",
        "qty": "Số lượng"
    }
    html = TeamsDispatcher.build_html_table(records, headers_map)
    assert "<table" in html
    assert "Mã hàng" in html
    assert "1002345" in html
    assert "A-01-01" in html
    assert "B-02-03" in html


def test_send_to_teams_dry_run():
    res = TeamsDispatcher.send_to_teams(
        webhook_url="https://prod-xx.southeastasia.logic.azure.com:443/workflows/YOUR_WEBHOOK_URL_HERE",
        request_id="TEST-001",
        command="VITRI",
        status="SUCCEEDED",
        sender="test@company.com",
        summary="Test summary",
        records=[{"product": "1002345"}],
        headers_map={"product": "Mã hàng"}
    )
    assert res is True


def test_payload_budgeting_under_28kb():
    # Giả lập 100 dòng dữ liệu lớn từ MB51 hoặc Posting Change
    large_records = []
    for i in range(100):
        large_records.append({
            "product": f"BIW320137{i:02d}AA",
            "batch": f"LOT{i:03d}_20261023",
            "bin": f"APFL-BIN-{i:03d}-ROW-01",
            "grDate": f"23.10.2026",
            "grTime": f"14:{i%60:02d}:00",
            "qty": str(100 + i),
            "unit": "PCS",
            "cat": "F2"
        })

    headers_map = {
        "product": "Mã hàng", "batch": "Batch", "bin": "Vị trí",
        "grDate": "Ngày GR", "grTime": "Giờ GR", "qty": "SL",
        "unit": "ĐVT", "cat": "Loại"
    }

    meta = {"messageId": "12345", "command": "VITRI", "senderEmail": "test@congty.com"}
    payload, data = TeamsDispatcher.prepare_safe_payload(
        "REQ-TEST-LARGE", meta, "SUCCEEDED", "Tra cứu 100 dòng", large_records, headers_map
    )

    # Đảm bảo tuyệt đối dung lượng dưới 26KB (tránh vượt trần 28KB của Teams)
    assert len(data) <= MAX_TEAMS_PAYLOAD_BYTES
    # Đảm bảo bảng có chứa thông báo tổng dòng
    assert payload["totalRecords"] == 100
    assert payload["displayedRecords"] <= 12
    assert "Đang hiển thị" in payload["table_html"]


def test_build_kpi_badges_vitri():
    records = [
        {"product": "BIW32013765AA", "batch": "B2610", "bin": "BIN-A-01", "qty": "150", "unit": "PC"},
        {"product": "BIW32013765AA", "batch": "B2611", "bin": "BIN-A-02", "qty": "50", "unit": "PC"}
    ]
    badge_html = TeamsDispatcher.build_kpi_badges(records, "VITRI")
    assert "2 Vị trí Bin" in badge_html
    assert "2 Lô (Batch)" in badge_html
    assert "200 PC" in badge_html


def test_build_help_html():
    help_html = TeamsDispatcher.build_help_html("Anh Ba")
    assert "Anh Ba" in help_html
    assert "@Bot check vị trí" in help_html
    assert "32013765" in help_html
