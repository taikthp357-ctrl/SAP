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
    assert "VL06I" in help_html
    assert "ME2N" in help_html
    assert "32013765" in help_html


def test_build_kpi_badges_ib_treo():
    records = [
        {"ibNo": "180001", "product": "BIW100", "batch": "LOT1", "qty": "50", "unit": "PC", "status": "Treo (Chưa GR)"},
        {"ibNo": "180002", "product": "BIW100", "batch": "LOT2", "qty": "30", "unit": "PC", "status": "Treo (Chưa GR)"}
    ]
    badges = TeamsDispatcher.build_kpi_badges(records, "IB_TREO")
    assert "2 Đơn IB treo" in badges
    assert "80 PC" in badges
    assert "2 Lô" in badges


def test_build_kpi_badges_sto():
    records = [
        {"stoNo": "360001", "product": "BIW100", "orderQty": "100", "deliveredQty": "60", "remainingQty": "40", "unit": "PC"},
        {"stoNo": "360002", "product": "BIW100", "orderQty": "50", "deliveredQty": "50", "remainingQty": "0", "unit": "PC"}
    ]
    badges = TeamsDispatcher.build_kpi_badges(records, "STO")
    assert "2 Đơn STO" in badges
    assert "150 PC" in badges
    assert "40 PC" in badges


def test_kpi_computed_over_full_dataset_not_just_12_display():
    """Đảm bảo KPI tính trên toàn bộ tập dữ liệu (ví dụ 20 dòng) thay vì chỉ 12 dòng hiển thị."""
    records = [
        {"product": f"P{i:02d}", "batch": f"B{i:02d}", "bin": f"BIN-{i:02d}", "qty": "1", "unit": "PC"}
        for i in range(20)
    ]
    headers_map = {"product": "Mã hàng", "batch": "Lô", "bin": "Bin", "qty": "SL"}
    meta = {"messageId": "M01", "command": "VITRI"}

    payload, data = TeamsDispatcher.prepare_safe_payload(
        "REQ-20-ROWS", meta, "SUCCEEDED", "Tra cứu 20 dòng", records, headers_map
    )

    # Hiển thị tối đa 12 dòng nhưng KPI phải cộng đủ 20 PC
    assert payload["totalRecords"] == 20
    assert payload["displayedRecords"] <= 12
    assert "20 PC" in payload["table_html"]
    assert "20 Vị trí Bin" in payload["table_html"]


def test_kpi_badges_sto_thousand_comma_formatting():
    """Kiểm tra STO có số lượng '1,000' không bị parse thành 1.0."""
    records = [
        {"stoNo": "360001", "product": "BIW100", "orderQty": "1,000", "deliveredQty": "200", "remainingQty": "800", "unit": "PC"},
        {"stoNo": "360002", "product": "BIW100", "orderQty": "500", "deliveredQty": "300", "remainingQty": "200", "unit": "PC"}
    ]
    badges = TeamsDispatcher.build_kpi_badges(records, "STO")
    assert "1,500 PC" in badges
    assert "1,000 PC" in badges


def test_extreme_payload_reduction_under_26kb():
    """Kiểm tra trường hợp dữ liệu cực lớn (nhiều text dài) luôn giữ payload <= 26,000 bytes."""
    massive_records = []
    for i in range(50):
        massive_records.append({
            "colA": f"Long descriptive text for item number {i} " * 20,
            "colB": f"Additional notes and warehouse explanations {i} " * 20,
            "qty": "100"
        })
    headers_map = {"colA": "Mô tả A", "colB": "Mô tả B", "qty": "Số lượng"}
    meta = {"messageId": "MASSIVE-01", "command": "MB52"}

    payload, data = TeamsDispatcher.prepare_safe_payload(
        "REQ-MASSIVE", meta, "SUCCEEDED", "Tra cứu cực lớn", massive_records, headers_map
    )

    assert len(data) <= MAX_TEAMS_PAYLOAD_BYTES
    assert payload["totalRecords"] == 50


def test_prepare_safe_payload_with_empty_records_no_crash():
    """Kiểm tra phản hồi không có dòng dữ liệu (records=[]) không bao giờ bị UnboundLocalError cur_html."""
    meta = {"messageId": "EMPTY-01", "command": "VITRI"}
    # Trường hợp 1: records rỗng, không có custom_html
    payload, data = TeamsDispatcher.prepare_safe_payload(
        "REQ-EMPTY-1", meta, "SUCCEEDED", "Không có dữ liệu", [], {"col": "Cột"}
    )
    assert len(data) <= MAX_TEAMS_PAYLOAD_BYTES
    assert payload["totalRecords"] == 0
    assert "Không tìm thấy dữ liệu" in payload["table_html"]

    # Trường hợp 2: records rỗng, có custom_html (ví dụ PING hoặc thẻ lỗi)
    custom = "<div>Thẻ thông báo lỗi</div>"
    payload2, data2 = TeamsDispatcher.prepare_safe_payload(
        "REQ-EMPTY-2", meta, "FAILED", "Tra cứu lỗi", [], {}, custom_html=custom
    )
    assert len(data2) <= MAX_TEAMS_PAYLOAD_BYTES
    assert "Thẻ thông báo lỗi" in payload2["table_html"]


def test_prepare_safe_payload_with_massive_summary_strict_26kb():
    """Kiểm tra summary cực dài (> 35.000 ký tự) vẫn đảm bảo payload <= 26.000 bytes."""
    massive_summary = "Bản báo cáo tóm tắt rất dài " * 2000  # ~56.000 bytes
    meta = {"messageId": "SUMMARY-01", "command": "TON"}
    payload, data = TeamsDispatcher.prepare_safe_payload(
        "REQ-MASSIVE-SUM", meta, "SUCCEEDED", massive_summary, [], {}
    )
    assert len(data) <= MAX_TEAMS_PAYLOAD_BYTES
    assert len(payload["summary"]) <= 300



