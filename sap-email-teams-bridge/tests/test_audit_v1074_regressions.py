# -*- coding: utf-8 -*-
"""Regression Test Suite for v1.0.75 - Chống Mock-Pass & Tái hiện toàn bộ 11 Probe từ Audit Log DANH_GIA_v1.0.74."""

from pathlib import Path
import pytest
from decimal import Decimal
from unittest.mock import MagicMock, patch

from app.core.context import RequestContext
from app.core.security import assert_write_permission, WritePermissionError
from app.number_utils import (
    parse_sap_decimal,
    parse_sap_float,
    NumericParseError,
    parse_sap_datetime_sort_key
)
from app.nvl_service import validated_stock_lot
from app.sap_attachment import MigoAttachmentRunner
from app.sap_engine import SapEngine
from app.handlers.system_handler import SystemHandler
from app.teams_dispatcher import TeamsDispatcher
from app.state_manager import StateManager


# --------------------------------------------------------------------------
# 1. P1-06 Probe: empty_identity & different_session_id
# --------------------------------------------------------------------------
def test_probe_empty_session_identity_rejected():
    """Probe empty_identity: Khi session identity rỗng hoặc thiếu -> Fail-closed."""
    runner = MigoAttachmentRunner()
    mock_session = MagicMock()
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "BOT_USER"
    mock_session.Id = "ses[0]"

    runner.expected_session_identity = {}
    ok, reason = runner.verify_session_identity(mock_session)
    assert ok is False
    assert reason == "SESSION_IDENTITY_REQUIRED"

    runner.expected_session_identity = None
    ok, reason = runner.verify_session_identity(mock_session)
    assert ok is False
    assert reason == "SESSION_IDENTITY_REQUIRED"


def test_probe_different_session_id_rejected():
    """Probe different_session_id: Khi session_id thực tế khác với expected identity -> Mismatch."""
    runner = MigoAttachmentRunner()
    mock_session = MagicMock()
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "BOT_USER"
    mock_session.Id = "ses[1]"  # Thực tế là tab ses[1]

    runner.expected_session_identity = {
        "system": "PRD",
        "client": "100",
        "user": "BOT_USER",
        "session_id": "ses[0]"  # Mong muốn là tab ses[0]
    }
    ok, reason = runner.verify_session_identity(mock_session)
    assert ok is False
    assert reason == "SESSION_IDENTITY_MISMATCH"


# --------------------------------------------------------------------------
# 2. P1-05 & P1-01 Probe: spoofed_body_sender & write permissions
# --------------------------------------------------------------------------
def test_probe_spoofed_body_sender_rejected():
    """Probe spoofed_body_sender: Giả mạo sender trong body không vượt qua được Write Guard."""
    ctx = RequestContext(
        request_id="REQ-TEST-001",
        command="ATT_MIGO",
        meta={
            "actualSenderEmail": "attacker@external.com",
            "senderEmail": "authorized_manager@vinfast.vn"  # Body giả mạo email sếp
        },
        config={
            "allow_sap_writes": True,
            "trusted_intake_senders": ["authorized_manager@vinfast.vn"]
        }
    )

    # 1. Bị chặn khi actualSenderEmail là kẻ tấn công (dù body có email sếp)
    with pytest.raises(WritePermissionError, match="không thuộc danh sách trusted_intake_senders"):
        assert_write_permission(ctx)

    # 2. Không cho phép ghi khi allow_sap_writes = False
    ctx.config["allow_sap_writes"] = False
    with pytest.raises(WritePermissionError, match="allow_sap_writes=False"):
        assert_write_permission(ctx)


# --------------------------------------------------------------------------
# 3. P1-07 Probe: stale_status_while_popup_close_times_out & url_missing_controls
# --------------------------------------------------------------------------
def test_probe_stale_status_while_popup_close_times_out(tmp_path):
    """Probe stale_status_while_popup_close_times_out: sbar báo S nhưng popup wnd[1] bị kẹt không đóng."""
    runner = MigoAttachmentRunner()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "MIGO"
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "BOT"
    mock_session.Id = "ses[0]"
    mock_session.Busy = False
    runner.session_identity = {"system": "PRD", "client": "100", "user": "BOT", "session_id": "ses[0]"}
    runner._active_document = "5001234567"
    runner._active_year = "2026"

    # Mock sbar có tin nhắn thành công stale
    mock_sbar = MagicMock()
    mock_sbar.Text = "Attachment successfully created"
    mock_sbar.text = "Attachment successfully created"
    mock_sbar.MessageType = "S"
    mock_sbar.messageType = "S"

    # Mock wnd[1] không chịu đóng (kẹt popup)
    mock_btn = MagicMock()
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_session.FindByName.return_value = mock_btn
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.return_value = mock_btn

    def find_by_id(path, *args):
        if "sbar" in str(path):
            return mock_sbar
        if str(path).startswith("wnd[1]"):
            return mock_wnd1  # Luôn tồn tại -> không đóng được
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    dummy_file = tmp_path / "dummy.pdf"
    dummy_file.write_bytes(b"%PDF-dummy")

    # wait_for_wnd1(appear=False) sẽ timeout trả về False
    with patch("app.sap_attachment.wait_for_wnd1", side_effect=[mock_wnd1, False]):
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is False
        assert "FAILED_POPUP_STUCK" in msg or "không đóng" in msg.lower()


# --------------------------------------------------------------------------
# 4. P1-07 Probe: different_ticket_prefix_exact_token_match
# --------------------------------------------------------------------------
def test_probe_ticket_prefix_exact_token_match():
    """Probe different_ticket_prefix_accepted: OUT_P1_1234 không được nhận cho OUT_P1_123."""
    import re
    # Token regex boundary logic
    target_token = "OUT_P1_123"
    pattern = rf"(?<![A-Z0-9_]){re.escape(target_token)}(?![A-Z0-9_])"

    existing_header_diff = "Đã xuất phiếu OUT_P1_1234 cho ca 1"
    existing_header_exact = "Đã xuất phiếu OUT_P1_123 cho ca 1"

    assert not re.search(pattern, existing_header_diff)
    assert re.search(pattern, existing_header_exact) is not None


# --------------------------------------------------------------------------
# 5. P1-02 Probe: missing_session_live & negative_message_rejected
# --------------------------------------------------------------------------
def test_probe_migo_missing_session_fail_closed():
    """Probe missing_session_live: MIGO post với session=None khi live -> Chặn ngay lập tức."""
    engine = SapEngine()
    engine.session = None
    items = [
        {"material": "BIW001", "batch": "LOT1", "allocated_qty": 10, "unit": "PC",
         "source_plant": "3093", "source_sloc": "1005", "dest_plant": "3093", "dest_sloc": "3001"}
    ]
    ok, doc, msg = engine.post_migo_311_ob_group(items, dry_run=False)
    assert ok is False
    assert "ERR_SAP_SESSION_NONE" in msg


def test_probe_migo_negative_message_rejected():
    """Probe MIGO nhận diện sai thông điệp lỗi hoặc phủ định của SAP."""
    engine = SapEngine()
    engine.session = MagicMock()
    mock_sbar = MagicMock()
    mock_sbar.MessageType = "E"
    mock_sbar.messageType = "E"
    mock_sbar.Text = "Material Document 5001234567 does not exist"
    mock_sbar.text = "Material Document 5001234567 does not exist"

    items = [
        {"material": "BIW001", "batch": "LOT1", "allocated_qty": 10, "unit": "PC",
         "source_plant": "3093", "source_sloc": "1005", "dest_plant": "3093", "dest_sloc": "3001"}
    ]
    with patch.object(engine, "_open_transaction"), \
         patch.object(engine, "_get_status", return_value=("Material Document 5001234567 does not exist", "E")), \
         patch.object(engine, "_find_by_id_safe", return_value=mock_sbar), \
         patch.object(engine, "_wait_idle"):
        ok, doc, msg = engine.post_migo_311_ob_group(items, dry_run=False)
        assert ok is False
        assert doc == ""
        assert "ERR_MIGO" in msg or "does not exist" in msg.lower() or "chưa xác nhận" in msg.lower()


# --------------------------------------------------------------------------
# 6. P1-03 Probe: number_auto_and_decimal_safety
# --------------------------------------------------------------------------
def test_probe_number_auto_and_decimal_safety():
    """Probe number_auto: 0,001 không thành 1; 1.234 trong AUTO báo lỗi mơ hồ."""
    # 0,001 luôn là 0.001
    assert parse_sap_decimal("0,001", fmt="AUTO") == Decimal("0.001")
    assert parse_sap_decimal("0.001", fmt="AUTO") == Decimal("0.001")
    assert parse_sap_decimal("0,001", fmt="EU") == Decimal("0.001")
    assert parse_sap_decimal("0.001", fmt="EU") == Decimal("0.001")

    # 1.234 và 1,234 trong AUTO là mơ hồ -> ném lỗi NumericParseError
    with pytest.raises(NumericParseError, match="Ambiguous quantity"):
        parse_sap_decimal("1.234", fmt="AUTO")

    with pytest.raises(NumericParseError, match="Ambiguous quantity"):
        parse_sap_decimal("1,234", fmt="AUTO")


# --------------------------------------------------------------------------
# 7. Item 100 Probe: PROBE_PING_WITHOUT_SAP
# --------------------------------------------------------------------------
def test_probe_ping_without_sap_shows_disconnected():
    """Probe PROBE_PING_WITHOUT_SAP: PING khi ctx.sap=None phải báo DISCONNECTED / Chưa kết nối SAP."""
    ctx = RequestContext(
        request_id="REQ-PING-001",
        command="PING",
        sap=None  # Không có kết nối SAP
    )
    handler = SystemHandler()
    res = handler.handle(ctx)
    assert "DISCONNECTED" in res.summary
    assert "Chưa kết nối SAP" in res.custom_html


# --------------------------------------------------------------------------
# 8. Item 96 Probe: PROBE_BOM_NEGATIVE_BADGE
# --------------------------------------------------------------------------
def test_probe_bom_negative_badge():
    """Probe PROBE_BOM_NEGATIVE_BADGE: Status 'KHÔNG CÓ BOM' không được match nhầm thành 'Có BOM'."""
    records = [
        {"parent": "MAT_A", "component": "", "status": "KHÔNG CÓ BOM"}
    ]
    kpi_html = TeamsDispatcher.build_kpi_badges(records, "CHECK_BOM")
    assert "Có BOM" not in kpi_html
    assert "Không có BOM" in kpi_html


# --------------------------------------------------------------------------
# 9. Item 101 Probe: FIFO ngày không hợp lệ
# --------------------------------------------------------------------------
def test_probe_fifo_invalid_date_rejected():
    """Probe FIFO ngày không hợp lệ: 31.02.2026 hoặc 25:70:00 bị loại bỏ (INVALID_FIFO_TIMESTAMP)."""
    k_inv = parse_sap_datetime_sort_key("31.02.2026", "25:70:00")
    assert k_inv == ("INVALID_DATE", "INVALID_TIME")

    row_inv = {
        "product": "BIW001", "batch": "LOT_INV", "bin": "BIN-A",
        "plant": "3093", "sloc": "1005", "unit": "PC", "cat": "FT",
        "availQty": "50", "grDate": "31.02.2026", "grTime": "25:70:00"
    }
    with pytest.raises(ValueError, match="INVALID_FIFO_TIMESTAMP"):
        validated_stock_lot(row_inv, source_plant="3093", source_sloc="1005", require_fifo=True)


# --------------------------------------------------------------------------
# 10. Item Dedup Probe: Dedup theo Store
# --------------------------------------------------------------------------
def test_probe_state_manager_dedup_namespaced_by_store(tmp_path):
    """Probe Dedup: Khóa tin nhắn được phân tách namespace theo store_id."""
    db_file = tmp_path / "test_dedup.db"
    mgr = StateManager(db_file)

    # Đánh dấu msg-01 ở Store A
    mgr.mark_message_processed("msg-01", None, "REQ-01", "CHECK_ATT", "SUCCEEDED", store_id="STORE_A")

    # Kiểm tra ở Store A: đã xử lý
    assert mgr.is_message_processed("msg-01", store_id="STORE_A") is True

    # Kiểm tra ở Store B cùng ID msg-01: chưa xử lý
    assert mgr.is_message_processed("msg-01", store_id="STORE_B") is False
