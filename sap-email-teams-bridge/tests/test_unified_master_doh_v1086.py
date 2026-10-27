# -*- coding: utf-8 -*-
"""Comprehensive verification suite for Unified Master Data, FEFO/FIFO, 3-Layer Intent Shield, and Shift DOH Report."""

import pytest
from unittest.mock import MagicMock
from decimal import Decimal
from app.nvl_master import NvlMasterService
from app.nvl_service import NvlService
from app.email_parser import SmartCommandParser
from app.sap_engine import SapEngine
from app.core.context import RequestContext, CommandResult
from app.handlers.nvl_handler import NvlHandler
from app.handlers.stock_handler import StockHandler


def test_nvl_master_vendor_and_bin_zone_rules():
    """Kiểm tra Master Data đọc được NCC thật, FEFO keo 200C và quy tắc Bin vật lý."""
    master = NvlMasterService.load_master_data(force_reload=True)
    assert master is not None
    assert "box_catalog" in master
    assert "fefo_materials" in master

    # 1. Kiểm tra NCC thật (từ THEO DOI ODER NVL files)
    henkel_ncc = NvlMasterService.get_vendor("BIW71005512")
    assert "HENKEL" in henkel_ncc.upper() or "SINGAPORE" in henkel_ncc.upper()

    # 2. Kiểm tra Keo Henkel SLoc 200C
    assert NvlMasterService.is_henkel_glue("BIW71005512") is True
    assert NvlMasterService.is_henkel_glue("BIW71005612") is True
    assert NvlMasterService.is_henkel_glue("BIW80001331") is False

    # 3. Kiểm tra Bin vật lý mặt đất vs tầng cao vs rack xếp chồng
    assert NvlMasterService.is_ground_bin("APFL-HR1-01-01-1") is True
    assert NvlMasterService.is_ground_bin("APFL-HR1-01-01-2") is False  # Tầng 2
    assert NvlMasterService.is_ground_bin("APFL-HR1-01-01-3") is False  # Tầng 3
    assert NvlMasterService.is_ground_bin("APFL-P01-01-01") is True     # P ground box
    assert NvlMasterService.is_ground_bin("APFL-P22-01-01") is False    # P22 RV5 rack chồng
    assert NvlMasterService.is_ground_bin("APFL-P24-01-01") is False    # P24 Rack cam chồng
    assert NvlMasterService.is_ground_bin("APFL-ZF10-01-01") is True    # ZF thùng kiện xếp sàn
    assert NvlMasterService.is_ground_bin("APFL-TH05-01-01") is True    # TH thùng kiện xếp sàn


def test_fefo_fifo_and_target_bin_priority(tmp_path):
    """Kiểm tra FEFO & FIFO là Số 1 và Số 2 tuyệt đối, sau đó đến target bin và mặt đất."""
    tmp_db = tmp_path / "test_fefo_fifo.db"
    orders = [
        {"material": "MAT_FEFO_TEST", "qty": 10, "station": "LINE_01", "unit": "PC", "target_bin": "APFL-ZF02-01-01"}
    ]

    # Giả lập tồn:
    # Lot A: Nhập cũ hơn (FIFO) nhưng HSD xa hơn
    # Lot B: Nhập mới hơn nhưng HSD cận kề hơn (FEFO ưu tiên trước!)
    # Lot C: Trùng target_bin nhưng HSD xa
    stock_records = [
        {"product": "MAT_FEFO_TEST", "batch": "LOT_OLD_GR", "bin": "APFL-HR1-01-01-2", "availQty": 10,
         "unit": "PC", "plant": "3093", "sloc": "1005", "cat": "UNRESTRICTED",
         "grDate": "01.10.2026", "grTime": "08:00:00", "hsd": "30.12.2026"},
        {"product": "MAT_FEFO_TEST", "batch": "LOT_FEFO_EXPIRING", "bin": "APFL-HR1-01-01-2", "availQty": 10,
         "unit": "PC", "plant": "3093", "sloc": "1005", "cat": "UNRESTRICTED",
         "grDate": "20.10.2026", "grTime": "08:00:00", "hsd": "05.11.2026"},
        {"product": "MAT_FEFO_TEST", "batch": "LOT_TARGET_BIN", "bin": "APFL-ZF02-01-01", "availQty": 10,
         "unit": "PC", "plant": "3093", "sloc": "1005", "cat": "UNRESTRICTED",
         "grDate": "25.10.2026", "grTime": "08:00:00", "hsd": "30.12.2026"},
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)
    assert plan["is_full_allocation"] is True
    item = plan["allocated_items"][0]
    # FEFO là số 1 tuyệt đối: Lô LOT_FEFO_EXPIRING phải được bốc đầu tiên!
    assert item["batch"] == "LOT_FEFO_EXPIRING"


def test_intent_shield_prevents_att_diverted_to_ra_don():
    """Hàng rào 3 lớp: Dán bảng 10-digit DOC tuyệt đối không bị nhận nhầm thành RA_DON_NVL."""
    text_with_doc_table = """
    SAP att mail
    5001234567    2026    KV08
    5001234568    2026    KV08
    5001234569    2026    KV08
    """
    res = SmartCommandParser.parse(text_with_doc_table, default_plant="3093")
    assert res is not None
    # Ý định phải khóa chặt là ATT_MIGO, TUYỆT ĐỐI không bị ghi đè sang RA_DON_NVL
    assert res["command"] == "ATT_MIGO"
    assert not res.get("nvlOrders")

    # Trường hợp dán bảng số DOC và số phiếu (không có năm)
    text_docs_tickets = """
    SAP đính kèm số phiếu
    5009876543    99123
    5009876544    99124
    """
    res2 = SmartCommandParser.parse(text_with_doc_table, default_plant="3093")
    assert res2["command"] == "ATT_MIGO"


def test_nvl_handler_blocks_doc_numbers_if_misrouted():
    """Layer 3: NvlHandler chủ động từ chối nếu nhận lệnh chứa từ khóa ATT và không có lệnh ra đơn rõ ràng."""
    mock_ctx = MagicMock()
    mock_ctx.command = "RA_DON_NVL"
    mock_ctx.meta = {
        "rawText": "SAP att mail 5001234567 2026",
        "orders": [{"material": "5001234567", "qty": 2026, "station": "LINE", "unit": "PC"}]
    }

    handler = NvlHandler()
    result = handler.handle(mock_ctx)
    assert result.status == "REJECTED_INTENT_CONFLICT"
    assert "đính kèm" in result.summary.lower()


def test_shift_doh_warning_report():
    """WP5: Báo cáo tồn thấp & DOH theo ca xuất đúng 6 cột, không suy đoán Inbound ảo."""
    engine = SapEngine.__new__(SapEngine)
    engine.session = None

    # Mock check_stock_mb52 trả về tồn SLoc 3001 và 1005
    def mock_mb52(materials=None, plant="3093", slocs=None, max_rows=15000):
        records = []
        for m in (materials or []):
            if m == "BIW_LOW_DOH":
                records.append({"product": m, "plant": "3093", "sloc": "3001", "unrest": "40", "unit": "PC"})
                records.append({"product": m, "plant": "3093", "sloc": "1005", "unrest": "500", "unit": "PC"})
            elif m == "BIW_HIGH_DOH":
                records.append({"product": m, "plant": "3093", "sloc": "3001", "unrest": "600", "unit": "PC"})
                records.append({"product": m, "plant": "3093", "sloc": "1005", "unrest": "1000", "unit": "PC"})
        return records

    engine.check_stock_mb52 = mock_mb52

    report = engine.check_shift_doh_warning(
        materials=["BIW_LOW_DOH", "BIW_HIGH_DOH"],
        model_name="VF6",
        shift_name="S6",
        threshold=200
    )

    assert len(report) == 1
    row = report[0]
    # Kiểm tra 6 cột chuẩn xác
    assert row["product"] == "BIW_LOW_DOH"
    assert "ncc" in row
    assert row["stock_3001"] == 40
    assert row["stock_1005"] == 500
    assert row["doh"] == 40  # 40 / US=1 = 40 (< 200)
    assert row["model"] == "VF6"
    assert "inbound" not in row
    assert "eta" not in row
