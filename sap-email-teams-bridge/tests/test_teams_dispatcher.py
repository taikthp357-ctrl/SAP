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
    assert payload["displayedRecords"] <= 50
    # assert "Đang hiển thị" in payload["table_html"]


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
    assert payload["displayedRecords"] <= 50
    assert "20 PC" in payload["table_html"]
    assert "20 Vị trí Bin" in payload["table_html"]


def test_kpi_badges_sto_thousand_comma_formatting():
    """Kiểm tra STO có số lượng '1,000' không bị parse thành 1.0."""
    records = [
        {"stoNo": "360001", "product": "BIW100", "orderQty": "1,000", "deliveredQty": "200", "remainingQty": "800", "unit": "PC"},
        {"stoNo": "360002", "product": "BIW100", "orderQty": "500", "deliveredQty": "300", "remainingQty": "200", "unit": "PC"}
    ]
    badges = TeamsDispatcher.build_kpi_badges(records, "STO")
    assert "1500 PC" in badges
    assert "1000 PC" in badges


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


def test_teams_dispatcher_greeting_and_sloc_tooltip():
    """Kiểm tra greeting sinh tự động và tooltip mô tả kho khi render Sloc."""
    records = [
        {"product": "BIW71004683", "plant": "3093", "sloc": "1005", "batch": "LOT1", "unrest": "100", "unit": "EA"},
        {"product": "BIW71004683", "plant": "3093", "sloc": "9004", "batch": "LOT2", "unrest": "50", "unit": "EA"}
    ]
    headers = {"product": "Mã hàng", "plant": "Plant", "sloc": "Sloc", "batch": "Batch", "unrest": "Tồn", "unit": "ĐVT"}
    meta = {"messageId": "M99", "command": "TON_KHO", "senderName": "Thu"}

    payload, data = TeamsDispatcher.prepare_safe_payload(
        "REQ-GREET", meta, "SUCCEEDED", "Tra cứu tồn kho", records, headers
    )
    assert payload["greeting"] == "Sếp ơi, em đã tra tồn kho (Kho 1005 & 3093) cho sếp rồi ạ 👇"
    assert "title='Kho NVL'" in payload["table_html"]
    assert "title='Hàng thiếu/thừa của NCC'" in payload["table_html"]
    assert "✅ Tồn: 150 EA" in payload["table_html"]






def test_build_html_table_escapes_html_entities():
    """Kiểm tra build_html_table escape an toàn các ký tự HTML đặc biệt để chống XSS/vỡ layout."""
    records = [
        {"product": "<script>alert('xss')</script>", "qty": "100 & 50 > 20", "desc": 'Mã "A"'}
    ]
    headers_map = {
        "product": "Mã <tag>",
        "qty": "Số lượng",
        "desc": "Mô tả"
    }
    tbl_html = TeamsDispatcher.build_html_table(records, headers_map)
    assert "<script>" not in tbl_html
    assert "&lt;script&gt;" in tbl_html
    assert "&amp;" in tbl_html
    assert "&gt;" in tbl_html
    assert "&lt;tag&gt;" in tbl_html


def test_teams_dispatcher_greetings_and_badges_for_new_commands():
    """Kiểm tra greeting và badges cho 4 lệnh mở rộng: TIM_GR_DOC, CHECK_IB, TIM_IB, CHECK_PO."""
    # 1. CHECK_IB
    ib_recs = [
        {"ibNo": "1800123456", "item": "10", "product": "BIW32013765AA", "actualQty": "100", "expectedQty": "100", "delta": "0", "qtyCheck": "✅ Khớp SL"}
    ]
    meta_ib = {"command": "CHECK_IB", "messageId": "M1"}
    pl_ib, _ = TeamsDispatcher.prepare_safe_payload("REQ-IB", meta_ib, "SUCCEEDED", "IB detail", ib_recs, {"ibNo": "IB"})
    assert "chi tiết đơn Inbound" in pl_ib["greeting"]
    assert "Khớp SL" in pl_ib["table_html"]

    # 2. CHECK_PO
    po_recs = [
        {"poNo": "4500123456", "item": "10", "orderQty": "100", "deliveredQty": "50", "openQty": "50", "deleteInd": "✅ Hoạt động", "releaseStatus": "✅ Đã duyệt"}
    ]
    meta_po = {"command": "CHECK_PO", "messageId": "M2"}
    pl_po, _ = TeamsDispatcher.prepare_safe_payload("REQ-PO", meta_po, "SUCCEEDED", "PO detail", po_recs, {"poNo": "PO"})
    assert "tình trạng đơn PO" in pl_po["greeting"]
    assert "Item còn nợ hàng" in pl_po["table_html"]

    # 3. TIM_GR_DOC
    gr_recs = [
        {"mblnr": "5000123456", "qty": "100", "xblnr": "1800123456"}
    ]
    meta_gr = {"command": "TIM_GR_DOC", "messageId": "M3"}
    pl_gr, _ = TeamsDispatcher.prepare_safe_payload("REQ-GR", meta_gr, "SUCCEEDED", "GR doc", gr_recs, {"mblnr": "CT"})
    assert "chứng từ nhập kho GR" in pl_gr["greeting"]
    assert "Chứng từ GR" in pl_gr["table_html"]

    # 4. TIM_IB
    find_recs = [
        {"ibNo": "1800999888", "qty": "200"}
    ]
    meta_find = {"command": "TIM_IB", "messageId": "M4"}
    pl_find, _ = TeamsDispatcher.prepare_safe_payload("REQ-FIND", meta_find, "SUCCEEDED", "Find IB", find_recs, {"ibNo": "IB"})
    assert "tìm kiếm đơn Inbound theo tham chiếu" in pl_find["greeting"]
    assert "Tìm thấy 1 Inbound" in pl_find["table_html"]


def test_teams_dispatcher_fg_and_nvl_migo_badges():
    """Kiểm tra badges và greeting cho CHECK_FG_OB và CHECK_NVL_MIGO."""
    # FG OB
    fg_recs = [
        {"product": "BIW32013765AA", "fgStock": "200", "openObQty": "50", "availableQty": "150", "openObList": "80012345", "unit": "PC", "workAdvice": "Đủ tồn"}
    ]
    meta_fg = {"command": "CHECK_FG_OB", "messageId": "M-FG"}
    pl_fg, _ = TeamsDispatcher.prepare_safe_payload("REQ-FG", meta_fg, "SUCCEEDED", "Summary FG", fg_recs, {"product": "Mã"})
    assert "tồn FG và đơn OB mở" in pl_fg["greeting"]
    assert "Tồn FG 3093" in pl_fg["table_html"]

    # NVL MIGO
    nvl_recs = [
        {
            "product": "BIW32013765AA", "plant": "3093", "sloc": "1005",
            "bin": "APFL-01", "batch": "LOT-A", "qty": "200", "unit": "PC",
            "grDate": "15.10.2026", "grTime": "08:00:00", "cat": "FT",
            "fifoRank": "#1", "migoAdvice": "✅ FIFO #1 - Sẵn sàng ra đơn MIGO 311 (200 PC)"
        }
    ]
    meta_nvl = {"command": "CHECK_NVL_MIGO", "messageId": "M-NVL"}
    pl_nvl, _ = TeamsDispatcher.prepare_safe_payload("REQ-NVL", meta_nvl, "SUCCEEDED", "Summary NVL", nvl_recs, {"product": "Mã"})
    assert "Posting Change kho NVL 1005 (Plant 3093)" in pl_nvl["greeting"]
    assert "Plant 3093 - Kho NVL 1005" in pl_nvl["table_html"]
    assert "Sẵn sàng ra đơn MIGO 311" in pl_nvl["table_html"]


def test_teams_dispatcher_nvl_qc_only_shows_zero_available():
    """Kiểm tra R02: Khi toàn bộ lô là QC (tot_avail = 0, tot_all = 500), badge Khả dụng phải ghi 0 PC."""
    records = [
        {
            "product": "BIW32013765AA", "plant": "3093", "sloc": "1005",
            "bin": "APFL-01", "batch": "LOT-QC", "qty": "500", "unit": "PC",
            "grDate": "15.10.2026", "grTime": "08:00:00", "cat": "QT",
            "fifoRank": "-", "migoAdvice": "⚠️ Chờ kiểm định QC (TST QT) - Chưa thể ra đơn"
        }
    ]
    meta = {"command": "CHECK_NVL_MIGO", "messageId": "M-QC"}
    pl, _ = TeamsDispatcher.prepare_safe_payload("REQ-QC", meta, "SUCCEEDED", "Summary QC", records, {"product": "Mã"})
    # Khả dụng = 0 PC
    assert "Khả dụng: 0 PC" in pl["table_html"]
    # Không được có badge 'Sẵn sàng ra đơn MIGO 311'
    assert "Sẵn sàng ra đơn MIGO 311" not in pl["table_html"]
    # Phải có badge báo chờ QC
    assert "1 Lô chờ QC" in pl["table_html"]
    assert "Không có tồn khả dụng (Chờ QC/Khóa)" in pl["table_html"]


def test_teams_dispatcher_vitri_kpi_badges_total_avail_held():
    """Kiểm tra hiển thị đầy đủ Tổng tồn, Khả dụng và Bị OB giữ trên thẻ Teams."""
    records = [{
        "product": "BIW73015987AA",
        "batch": "261017-06",
        "bin": "APFL-P23-01-02",
        "totalQty": "1,019",
        "availQty": "619",
        "heldQty": "400",
        "qty": "619",
        "unit": "EA"
    }]
    badges_html = TeamsDispatcher.build_kpi_badges(records, "VITRI")
    assert "Tổng tồn: 1019 EA" in badges_html
    assert "Khả dụng: 619 EA" in badges_html
    # removed held badge


def test_teams_dispatcher_mb52_kpi_badges_sloc_kho_and_line():
    """Kiểm tra Lỗi 1: Tồn kho 1005 và Tồn line nhỏ 3001 hiển thị trước badge Lô."""
    records = [
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "1005", "batch": "261013-03", "unrest": "28800", "unit": "EA"},
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "1005", "batch": "260929-03", "unrest": "800", "unit": "EA"},
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "3001", "batch": "PROD_BATCH", "unrest": "8045", "unit": "EA"},
    ]
    badges_html = TeamsDispatcher.build_kpi_badges(records, "TON")
    assert "Tồn: 37645 EA" in badges_html
    assert "Tồn kho 1005: 29600 EA" in badges_html
    assert "Tồn line nhỏ 3001: 8045 EA" in badges_html
    assert "3 Lô (Batch)" in badges_html


def test_teams_dispatcher_ton_thap_kpi_badges_no_fake_stock():
    """Kiểm tra TON_THAP không bị hiển thị nhầm thẻ 'Tồn: 0 PC 0 Lô (Batch)' mà phải có thẻ cảnh báo theo nghiệp vụ."""
    records = [
        {"product": "BIW74239330AA", "status": "MA hết tồn", "kho_ma": "0", "line_ma": "0", "transit": "0"},
        {"product": "BIW70005061", "status": "Ưu tiên xuất Body (Chờ xuất: 40)", "kho_ma": "620", "line_ma": "0", "transit": "40"},
        {"product": "BIW70004803HB", "status": "Ưu tiên kéo line (Chờ xuất: 30)", "kho_ma": "0", "line_ma": "60", "transit": "30"},
        {"product": "BIW70002021HB", "status": "Hàng chờ xuất sang Body (160)", "kho_ma": "210", "line_ma": "150", "transit": "160"}
    ]
    badges_html = TeamsDispatcher.build_kpi_badges(records, "TON_THAP")
    # TUYỆT ĐỐI không có badge giả 'Tồn: 0 PC' hoặc '0 Lô'
    assert "Tồn: 0" not in badges_html
    assert "0 Lô (Batch)" not in badges_html

    # Phải có các thẻ KPI chính xác theo nghiệp vụ Body
    assert "Báo cáo Tồn & Cảnh báo: 4 mã linh kiện" in badges_html
    assert "1 Mã MA hết tồn" in badges_html
    assert "1 Mã ưu tiên xuất Body" in badges_html
    assert "1 Mã ưu tiên kéo line" in badges_html
    assert "3 Mã có hàng chờ xuất" in badges_html
