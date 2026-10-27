# -*- coding: utf-8 -*-
"""Unit tests cho CommandRouter và các Handler nghiệp vụ độc lập."""
import pytest
from unittest.mock import MagicMock, patch, create_autospec
from app.sap_engine import SapEngine
from app.core.context import RequestContext, CommandResult
from app.router import CommandRouter, router
from app.handlers import (
    BaseHandler,
    AttHandler,
    NvlHandler,
    StockHandler,
    LogisticsHandler,
    ProductionHandler,
    BbbgHandler,
    SystemHandler,
)


def test_router_registration():
    """Kiểm tra CommandRouter đăng ký đủ các lệnh nghiệp vụ chính."""
    r = CommandRouter()
    
    # System
    assert isinstance(r.get_handler("PING"), SystemHandler)
    assert isinstance(r.get_handler("WHO_IS_TAI"), SystemHandler)
    assert isinstance(r.get_handler("HELP"), SystemHandler)
    
    # Stock
    assert isinstance(r.get_handler("VITRI"), StockHandler)
    assert isinstance(r.get_handler("TON"), StockHandler)
    assert isinstance(r.get_handler("TON_KHO"), StockHandler)
    assert isinstance(r.get_handler("LICHSU"), StockHandler)
    assert isinstance(r.get_handler("TON_THAP"), StockHandler)
    assert isinstance(r.get_handler("LICHSU_BIN"), StockHandler)
    assert isinstance(r.get_handler("TIM_GR_DOC"), StockHandler)
    
    # Logistics
    assert isinstance(r.get_handler("CHECK_IB"), LogisticsHandler)
    assert isinstance(r.get_handler("TIM_IB"), LogisticsHandler)
    assert isinstance(r.get_handler("IB_TREO"), LogisticsHandler)
    assert isinstance(r.get_handler("CHECK_PO"), LogisticsHandler)
    assert isinstance(r.get_handler("STO"), LogisticsHandler)
    assert isinstance(r.get_handler("TRA_CUU_ASN"), LogisticsHandler)
    
    # Production
    assert isinstance(r.get_handler("CHECK_BOM"), ProductionHandler)
    assert isinstance(r.get_handler("CHECK_OB"), ProductionHandler)
    assert isinstance(r.get_handler("CHECK_FG_OB"), ProductionHandler)
    assert isinstance(r.get_handler("DIEU_PHOI_OB"), ProductionHandler)
    assert isinstance(r.get_handler("CHECK_QUEUE"), ProductionHandler)
    assert isinstance(r.get_handler("CHECK_RES"), ProductionHandler)
    
    # BBBG
    assert isinstance(r.get_handler("TAO_BBBG"), BbbgHandler)
    assert isinstance(r.get_handler("XEM_BBBG"), BbbgHandler)
    
    # Att & ZFIR266
    assert isinstance(r.get_handler("CHECK_ATT"), AttHandler)
    assert isinstance(r.get_handler("ZFIR266"), AttHandler)
    assert isinstance(r.get_handler("ATT_MIGO"), AttHandler)
    
    # NVL
    assert isinstance(r.get_handler("RA_DON_NVL"), NvlHandler)
    assert isinstance(r.get_handler("MIGO_311"), NvlHandler)
    assert isinstance(r.get_handler("CHECK_NVL_MIGO"), NvlHandler)


def test_router_dispatch_ping():
    """Kiểm tra router dispatch lệnh PING thành công."""
    ctx = RequestContext(
        request_id="REQ-TEST-001",
        command="PING",
        meta={"senderName": "Tester", "requestId": "REQ-TEST-001"}
    )
    res = router.dispatch(ctx)
    assert res.status == "SUCCEEDED"
    assert "PONG" in res.summary or "PONG" in (res.custom_html or "")
    assert res.custom_html is not None


def test_router_dispatch_who_is_tai():
    """Kiểm tra router dispatch lệnh WHO_IS_TAI thành công."""
    ctx = RequestContext(
        request_id="REQ-TEST-002",
        command="WHO_IS_TAI",
        meta={"senderName": "Tester"}
    )
    res = router.dispatch(ctx)
    assert res.status == "SUCCEEDED"
    assert "Phạm Tiến Tài" in (res.custom_html or "")
    assert "Kiến trúc sư trưởng" in res.summary


def test_router_dispatch_unknown_command():
    """Kiểm tra lệnh không xác định trả về status FAILED rõ ràng."""
    ctx = RequestContext(
        request_id="REQ-TEST-003",
        command="UNKNOWN_CMD_XYZ"
    )
    res = router.dispatch(ctx)
    assert res.status == "FAILED"
    assert "Lệnh không được hỗ trợ" in res.summary


def test_stock_handler_vitri_dispatch():
    """Kiểm tra StockHandler xử lý lệnh VITRI với mock sap."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_position_posting_change.return_value = [
        {"product": "BIW32013765AA", "batch": "B01", "bin": "BIN-A01", "totalQty": "100", "availQty": "100", "unit": "PC", "grDate": "20261026", "grTime": "100000", "cat": "Tồn kho"}
    ]
    ctx = RequestContext(
        request_id="REQ-STOCK-001",
        command="VITRI",
        sap=mock_sap,
        resolved_materials=["BIW32013765AA"]
    )
    res = router.dispatch(ctx)
    assert res.status == "SUCCEEDED"
    assert len(res.records) == 1
    assert res.records[0]["bin"] == "BIN-A01"
    assert "Vị trí lưu kho" in res.summary


def test_stock_handler_smart_combo_check():
    """Kiểm tra StockHandler tự động combo check MB52 khi EWM trả về 0 dòng."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_position_posting_change.return_value = []
    mock_sap.check_stock_mb52.return_value = [
        {"product": "BIW32013765AA", "batch": "B01", "sloc": "1005", "unrest": "50", "unit": "PC"}
    ]
    ctx = RequestContext(
        request_id="REQ-STOCK-COMBO",
        command="VITRI",
        sap=mock_sap,
        resolved_materials=["BIW32013765AA"]
    )
    res = router.dispatch(ctx)
    assert res.status == "SUCCEEDED"
    assert len(res.records) == 1
    assert "SLoc 1005 (Ngoài EWM)" in res.records[0]["bin"]
    assert "MB52 đang có tồn" in res.summary


def test_logistics_handler_dispatch():
    """Kiểm tra LogisticsHandler xử lý lệnh CHECK_IB."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_inbound_detail_vl06i.return_value = [
        {"ibNo": "1800012345", "item": "000010", "product": "BIW32013765AA", "actualQty": 50.0, "expectedQty": 50.0, "delta": 0.0, "qtyCheck": "OK", "unit": "PC", "sloc": "1005", "status": "Pending"}
    ]
    ctx = RequestContext(
        request_id="REQ-LOG-001",
        command="CHECK_IB",
        sap=mock_sap,
        meta={"inboundNo": "1800012345"},
        resolved_materials=["BIW32013765AA"]
    )
    res = router.dispatch(ctx)
    assert res.status == "SUCCEEDED"
    assert len(res.records) == 1
    assert res.records[0]["ibNo"] == "1800012345"


def test_production_handler_dispatch():
    """Kiểm tra ProductionHandler xử lý lệnh CHECK_BOM."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_bom_cs03.return_value = [
        {"parent": "PAR01", "component": "CMP01", "description": "Part 1", "qty": "2", "unit": "PC", "status": "CÓ BOM (Active)"}
    ]
    ctx = RequestContext(
        request_id="REQ-PROD-001",
        command="CHECK_BOM",
        sap=mock_sap,
        resolved_materials=["PAR01"]
    )
    res = router.dispatch(ctx)
    assert res.status == "SUCCEEDED"
    assert len(res.records) == 1
    assert "CÓ BOM" in res.records[0]["status"]


def test_bbbg_handler_dispatch():
    """Kiểm tra BbbgHandler xử lý lệnh XEM_BBBG."""
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_history_mb51.return_value = []
    mock_sap.check_inbound_treo_vl06i.return_value = []
    with patch("app.handlers.bbbg_handler.BbbgService") as mock_bbbg:
        mock_bbbg.reconcile.return_value = {
            "batch_id": "BATCH-01",
            "matched_count": 0,
            "total_items": 0,
            "items": []
        }
        mock_bbbg.generate_mail_preview.return_value = {"html_preview": "<div>Preview</div>"}
        ctx = RequestContext(
            request_id="REQ-BBBG-001",
            command="XEM_BBBG",
            sap=mock_sap
        )
        res = router.dispatch(ctx)
        assert res.status == "SUCCEEDED"
        assert "Xem trước BBBG FG" in res.summary


def test_att_handler_check_att_no_docs():
    """Kiểm tra AttHandler báo lỗi khi không có số chứng từ."""
    mock_sap = create_autospec(SapEngine, instance=True)
    ctx = RequestContext(
        request_id="REQ-ATT-001",
        command="CHECK_ATT",
        sap=mock_sap,
        meta={}
    )
    with pytest.raises(ValueError, match="Vui lòng cung cấp ít nhất một số chứng từ DOC"):
        router.dispatch(ctx)
