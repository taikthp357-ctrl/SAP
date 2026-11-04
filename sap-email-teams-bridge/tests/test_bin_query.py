# -*- coding: utf-8 -*-
"""Unit & Integration tests cho tính năng tra cứu vị trí theo Storage Bin EWM (/SCWM/POST)
Hỗ trợ Multi-SKU (nhiều mã hàng trong 1 Bin), chuẩn hóa Bin, và tìm kiếm wildcard.
"""
from unittest.mock import MagicMock, create_autospec
import pytest
from app.email_parser import SmartCommandParser, normalize_bin_query
from app.core.context import RequestContext
from app.handlers.stock_handler import StockHandler
from app.sap_engine import SapEngine
from app.router import CommandRouter


def test_normalize_bin_query():
    """Kiểm tra toàn diện thuật toán chuẩn hóa chuỗi Bin kho."""
    cases = [
        # Zero-padding 3 phân đoạn
        ("zf01-01-01", "ZF01-01-01"),
        ("zf1-1-1", "ZF01-01-01"),
        ("p1-1-1", "P01-01-01"),
        ("p01-01-01", "P01-01-01"),
        # Giữ nguyên wildcard rõ ràng
        ("p01-01-*", "P01-01-*"),
        ("p1-1-*", "P01-01-*"),
        ("zf1-1-*", "ZF01-01-*"),
        # Implicit wildcard prefix cho phân đoạn đơn
        ("zf01", "ZF01*"),
        ("zf1", "ZF01*"),
        ("p01", "P01*"),
        ("p1", "P01*"),
        # Tọa độ 4 phân đoạn đặc thù VinFast VF02
        ("APFL-HR08-08-01", "APFL-HR08-08-01"),
        ("apfl-hr8-8-1", "APFL-HR08-08-01"),
        ("HR08-08-01", "HR08-08-01"),
        ("hr8-8-1", "HR08-08-01"),
        # Bóc tách tiền tố từ khóa
        ("bin: zf1-1-1", "ZF01-01-01"),
        ("vị trí: p1-1-1", "P01-01-01"),
        ("kệ: zf1", "ZF01*"),
        ("ô bin: APFL-HR08-08-01", "APFL-HR08-08-01"),
    ]
    for raw, expected in cases:
        assert normalize_bin_query(raw) == expected, f"Failed: {raw} -> {normalize_bin_query(raw)} != {expected}"


def test_email_parser_bin_intents():
    """Kiểm tra SmartCommandParser bóc tách đúng lệnh VITRI kèm Bin và không bị lẫn mã hàng."""
    queries = [
        ("check bin ZF01-01-01", "VITRI", "ZF01-01-01", []),
        ("check bin zf1-1-1", "VITRI", "ZF01-01-01", []),
        ("bin p1-1-1", "VITRI", "P01-01-01", []),
        ("check bin zf01", "VITRI", "ZF01*", []),
        ("check bin p01-01-*", "VITRI", "P01-01-*", []),
        ("vi tri APFL-HR08-08-01", "VITRI", "APFL-HR08-08-01", []),
        ("kiem tra vi tri bin HR08-08-01", "VITRI", "HR08-08-01", []),
        ("check vi tri 32013765", "VITRI", "", ["32013765"]),
        ("check vi tri 32013765 o bin ZF01-01-01", "VITRI", "ZF01-01-01", ["32013765"]),
        ("check bin zf1", "VITRI", "ZF01*", []),
        ("check bin p1", "VITRI", "P01*", []),
    ]
    for q, exp_cmd, exp_bin, exp_mats in queries:
        res = SmartCommandParser.parse(q)
        assert res.get("command") == exp_cmd, f"Query '{q}': cmd={res.get('command')}"
        assert res.get("bin") == exp_bin, f"Query '{q}': bin={res.get('bin')}"
        assert res.get("materials") == exp_mats, f"Query '{q}': mats={res.get('materials')}"
        assert res.get("valid") is True, f"Query '{q}': valid={res.get('valid')}"


def test_stock_handler_multi_sku_bin_query():
    """Kiểm tra StockHandler xử lý tra cứu Storage Bin chứa nhiều mã hàng (Multi-SKU)
    Mô phỏng dữ liệu thực tế từ ảnh chụp màn hình /SCWM/POST của người dùng tại Kho VF02:
    Bin APFL-HR08-08-01 chứa 3 mã: BIW70011279, BIW73015190HB, BIW73015477HB
    """
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_stock_by_bin.return_value = [
        {
            "product": "BIW70011279", "availQty": "11", "batch": "260923-10",
            "bin": "APFL-HR08-08-01", "cat": "FT", "totalQty": "11", "heldQty": "0",
            "unit": "PC", "grDate": "23.09.2026", "grTime": "16:55"
        },
        {
            "product": "BIW73015190HB", "availQty": "258", "batch": "261024-20",
            "bin": "APFL-HR08-08-01", "cat": "FT", "totalQty": "258", "heldQty": "0",
            "unit": "PC", "grDate": "24.10.2026", "grTime": "19:33"
        },
        {
            "product": "BIW73015477HB", "availQty": "0", "batch": "261027-17",
            "bin": "APFL-HR08-08-01", "cat": "FT", "totalQty": "400", "heldQty": "400",
            "unit": "PC", "grDate": "27.10.2026", "grTime": "16:45"
        }
    ]

    ctx = RequestContext(
        request_id="REQ-BIN-001",
        command="VITRI",
        bin="APFL-HR08-08-01",
        meta={"bin": "APFL-HR08-08-01"},
        sap=mock_sap,
        resolved_materials=[]
    )

    handler = StockHandler()
    result = handler.handle(ctx)

    assert result.status == "SUCCEEDED"
    assert len(result.records) == 3
    # Xác minh gọi hàm check_stock_by_bin thay vì check_position_posting_change
    mock_sap.check_stock_by_bin.assert_called_once_with("APFL-HR08-08-01", plant="3093")
    assert "3 dòng (3 mã hàng)" in result.summary
    assert "APFL-HR08-08-01" in result.summary
    # Kiểm tra mã thứ 3 có heldQty = 400
    assert result.records[2]["heldQty"] == "400"
    assert result.records[2]["availQty"] == "0"


def test_stock_handler_bin_query_empty():
    """Kiểm tra khi Bin không có hàng trên EWM."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_stock_by_bin.return_value = []

    ctx = RequestContext(
        request_id="REQ-BIN-EMPTY",
        command="BIN",
        bin="ZF01-01-01",
        meta={"bin": "ZF01-01-01"},
        sap=mock_sap,
        resolved_materials=[]
    )

    handler = StockHandler()
    result = handler.handle(ctx)

    assert result.status == "SUCCEEDED"
    assert len(result.records) == 0
    assert "Không tìm thấy vật tư nào tại Bin 'ZF01-01-01'" in result.summary


def test_stock_handler_bin_with_material_filter():
    """Kiểm tra khi người dùng hỏi đích danh mã hàng ở vị trí Bin cụ thể."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_stock_by_bin.return_value = [
        {"product": "BIW70011279", "availQty": "11", "batch": "260923-10", "bin": "APFL-HR08-08-01", "totalQty": "11", "heldQty": "0"},
        {"product": "BIW73015190HB", "availQty": "258", "batch": "261024-20", "bin": "APFL-HR08-08-01", "totalQty": "258", "heldQty": "0"}
    ]

    ctx = RequestContext(
        request_id="REQ-BIN-MAT",
        command="VITRI",
        bin="APFL-HR08-08-01",
        meta={"bin": "APFL-HR08-08-01"},
        sap=mock_sap,
        resolved_materials=["BIW70011279"]
    )

    handler = StockHandler()
    result = handler.handle(ctx)

    assert result.status == "SUCCEEDED"
    assert len(result.records) == 1
    assert result.records[0]["product"] == "BIW70011279"
    assert "Vị trí lưu kho (/SCWM/POST) cho mã BIW70011279 tại Bin 'APFL-HR08-08-01'" in result.summary


def test_router_dispatch_bin_query():
    """Kiểm tra CommandRouter điều phối lệnh BIN tới đúng StockHandler."""
    router = CommandRouter()
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_stock_by_bin.return_value = [
        {"product": "BIW12345", "availQty": "50", "batch": "LOT1", "bin": "ZF01-01-01", "totalQty": "50", "heldQty": "0"}
    ]

    ctx = RequestContext(
        request_id="REQ-ROUTER-BIN",
        command="BIN",
        bin="ZF01-01-01",
        meta={"bin": "ZF01-01-01"},
        sap=mock_sap,
        resolved_materials=[]
    )

    result = router.dispatch(ctx)
    assert result.status == "SUCCEEDED"
    assert len(result.records) == 1
    assert result.records[0]["product"] == "BIW12345"
