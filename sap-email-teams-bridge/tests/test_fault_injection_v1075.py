# -*- coding: utf-8 -*-
"""Fault Injection Test Suite v1.0.75 - Kiểm thử tiêm lỗi độc lập chống Mock-Pass.

Bao gồm 4 kịch bản tiêm lỗi trọng yếu:
1. Tiêm lỗi rớt kết nối COM giữa chừng (COM drop mid-flight).
2. Tiêm lỗi popup lỗi/treo không đóng được (Stalled popup / modal dialog).
3. Tiêm lỗi rollback DB SQLite sau khi post SAP thành công (Ambiguous reservation).
4. Tiêm lỗi người gửi ngoài allowlist cố tình kích hoạt ghi qua body email (Spoofed body sender).
"""

from decimal import Decimal
from pathlib import Path
import pytest
from unittest.mock import MagicMock, patch, create_autospec

from app.core.context import RequestContext, CommandResult
from app.core.security import assert_write_permission, WritePermissionError
from app.sap_engine import SapEngine
from app.sap_attachment import MigoAttachmentRunner
from app.handlers.nvl_handler import NvlHandler
from app.handlers.att_handler import AttHandler
from app.handlers.system_handler import SystemHandler
from app.nvl_service import NvlService
from app.material_resolver import MaterialResolver
from app.state_manager import StateManager


# ==============================================================================
# PHẦN 1: MÔ PHỎNG RỚT KẾT NỐI COM GIỮA CHỪNG (COM DROP MID-FLIGHT)
# ==============================================================================

def test_fault_injection_com_drop_during_migo_311():
    """Mô phỏng rớt kết nối COM RPC giữa chừng khi đang chạy MIGO 311.
    Hệ thống phải fail-closed lập tức, không được báo thành công hoặc tạo số chứng từ giả."""
    engine = SapEngine()
    mock_session = MagicMock()
    engine.session = mock_session

    # Giả lập COM RPC drop bất thình lình khi findById gọi vào SAP GUI
    def com_drop_side_effect(id_str, *args, **kwargs):
        if "wnd[0]" in str(id_str):
            raise RuntimeError("0x800706BA: The RPC server is unavailable (COM connection dropped)")
        return MagicMock()

    mock_session.findById.side_effect = com_drop_side_effect

    items = [
        {
            "material": "BIW32013765AA",
            "batch": "LOT-2026A",
            "allocated_qty": Decimal("10"),
            "unit": "PC",
            "source_plant": "3093",
            "source_sloc": "1005",
            "dest_plant": "3093",
            "dest_sloc": "3001",
        }
    ]

    with patch.object(engine, "_open_transaction", autospec=True):
        ok, doc, msg = engine.post_migo_311_ob_group(items, dry_run=False)

    # Khẳng định Fail-Closed:
    assert ok is False
    assert doc == ""
    assert "RPC server is unavailable" in msg or "Lỗi thực thi SAP MIGO 311" in msg


def test_fault_injection_com_drop_during_attachment_upload(tmp_path):
    """Mô phỏng rớt kết nối COM khi đang mở hộp thoại upload GOS attachment."""
    runner = MigoAttachmentRunner()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "MIGO"
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "BOT_USER"
    mock_session.Id = "ses[0]"
    mock_session.Busy = False

    runner.session_identity = {
        "system": "PRD",
        "client": "100",
        "user": "BOT_USER",
        "session_id": "ses[0]",
    }
    runner._active_document = "5001234567"
    runner._active_year = "2026"

    # Giả lập COM văng Exception khi gọi hàm COM của SAP GUI
    mock_session.FindByName.side_effect = Exception("0x80010108: The object invoked has disconnected from its clients")
    mock_session.findById.side_effect = Exception("0x80010108: The object invoked has disconnected from its clients")

    test_pdf = tmp_path / "test_doc.pdf"
    test_pdf.write_bytes(b"%PDF-1.4 test payload")

    ok, msg = runner.attach_file_to_current_migo(mock_session, test_pdf)

    # Khẳng định an toàn: fail-closed và trả về mã lỗi rõ ràng
    assert ok is False
    assert "GOS_TRIGGER_FAILED" in msg or "0x80010108" in msg or "disconnected" in msg.lower()


def test_fault_injection_com_drop_during_ping_healthcheck():
    """Mô phỏng rớt kết nối COM khi bot chạy lệnh PING để kiểm tra trạng thái sống."""
    mock_engine = create_autospec(SapEngine, instance=True)
    mock_session = MagicMock()
    type(mock_session.Info).SystemName = property(
        fget=MagicMock(side_effect=Exception("COM connection timeout: server dead"))
    )
    mock_engine.session = mock_session

    ctx = RequestContext(
        request_id="REQ-PING-FAULT",
        command="PING",
        sap=mock_engine,
    )

    handler = SystemHandler()
    res = handler.handle(ctx)

    assert "DISCONNECTED" in res.summary
    assert "Mất kết nối SAP" in (res.custom_html or "")


# ==============================================================================
# PHẦN 2: MÔ PHỎNG POPUP LỖI / TREO KHÔNG ĐÓNG ĐƯỢC (STALLED POPUP)
# ==============================================================================

def test_fault_injection_popup_stuck_during_gos_attachment(tmp_path):
    """Mô phỏng hộp thoại GOS wnd[1] bị kẹt không đóng sau khi gửi lệnh upload.
    Dù thanh trạng thái có thể có tin nhắn cũ, hệ thống vẫn phải báo FAILED_POPUP_STUCK."""
    runner = MigoAttachmentRunner()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "MIGO"
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "BOT_USER"
    mock_session.Id = "ses[0]"
    mock_session.Busy = False

    runner.session_identity = {
        "system": "PRD",
        "client": "100",
        "user": "BOT_USER",
        "session_id": "ses[0]",
    }
    runner._active_document = "5001234567"
    runner._active_year = "2026"

    mock_btn = MagicMock()
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"

    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.return_value = mock_btn

    # Statusbar giả lập tin nhắn thành công lừa dối (stale success message)
    mock_sbar = MagicMock()
    mock_sbar.Text = "Attachment successfully created"
    mock_sbar.text = "Attachment successfully created"
    mock_sbar.MessageType = "S"
    mock_sbar.messageType = "S"

    def find_by_id_mock(path, *args):
        if "sbar" in str(path):
            return mock_sbar
        if str(path).startswith("wnd[1]"):
            return mock_wnd1  # Luôn tồn tại popup không tắt
        return MagicMock()

    mock_session.findById.side_effect = find_by_id_mock
    mock_session.FindByName.return_value = mock_btn

    test_file = tmp_path / "vat_invoice.pdf"
    test_file.write_bytes(b"%PDF-invoice-test")

    # wait_for_wnd1(appear=False) trả về False (hộp thoại không đóng được)
    with patch("app.sap_attachment.wait_for_wnd1", side_effect=[mock_wnd1, False]):
        ok, msg = runner.attach_file_to_current_migo(mock_session, test_file)

    assert ok is False
    assert "FAILED_POPUP_STUCK" in msg
    assert "không đóng" in msg.lower()


def test_fault_injection_migo_popup_blocking_error():
    """Mô phỏng MIGO gặp popup lỗi nghiêm trọng của SAP khi điền dòng phân bổ."""
    engine = SapEngine()
    mock_session = MagicMock()
    engine.session = mock_session

    items = [
        {
            "material": "BIW001",
            "batch": "LOT1",
            "allocated_qty": Decimal("10"),
            "unit": "PC",
            "source_plant": "3093",
            "source_sloc": "1005",
            "dest_plant": "3093",
            "dest_sloc": "3001",
        }
    ]

    mock_sbar = MagicMock()
    mock_sbar.MessageType = "E"
    mock_sbar.messageType = "E"
    mock_sbar.Text = "Deficit of SL Storage location unrestricted-use 10 PC"
    mock_sbar.text = "Deficit of SL Storage location unrestricted-use 10 PC"

    with patch.object(engine, "_open_transaction", autospec=True), \
         patch.object(engine, "_find_by_id_safe", return_value=mock_sbar), \
         patch.object(engine, "_get_status", return_value=(mock_sbar.Text, "E")), \
         patch.object(engine, "_wait_idle", autospec=True):
        ok, doc, msg = engine.post_migo_311_ob_group(items, dry_run=False)

    assert ok is False
    assert doc == ""
    assert "Deficit" in msg or "E" in msg or "Lỗi SAP MIGO" in msg


# ==============================================================================
# PHẦN 3: MÔ PHỎNG ROLLBACK DB SQLITE SAU KHI POST SAP THÀNH CÔNG
# ==============================================================================

def test_fault_injection_sqlite_sync_failure_post_migo(tmp_path):
    """Mô phỏng tình huống SAP MIGO post thành công nhưng SQLite gặp sự cố khóa/rollback.
    Hệ thống bắt buộc phải gán AMBIGUOUS_RESERVATION và lưu số chứng từ SAP để kiểm tra đối soát,
    tuyệt đối không được gán FAILED khiến người dùng nhấn post trùng chứng từ."""
    db_file = tmp_path / "bridge_test.db"
    state_mgr = StateManager(db_file)
    resolver = MaterialResolver()

    mock_sap = create_autospec(SapEngine, instance=True)
    # 1. Trả về tồn kho đủ để phân bổ (chuẩn hóa cat="FT", ngày DD.MM.YYYY, giờ HH:MM:SS)
    mock_sap.check_position_posting_change.return_value = [
        {
            "product": "BIW32013765AA",
            "batch": "LOT-A",
            "bin": "BIN-01",
            "plant": "3093",
            "sloc": "1005",
            "availQty": "100",
            "totalQty": "100",
            "unit": "PC",
            "grDate": "26.10.2026",
            "grTime": "08:00:00",
            "cat": "FT",
        }
    ]
    # 2. SAP Post MIGO 311 thành công rực rỡ, trả về số chứng từ 5001234567
    mock_sap.post_migo_311_ob_group.return_value = (
        True,
        "5001234567",
        "Material document 5001234567 posted",
    )

    ctx = RequestContext(
        request_id="REQ-NVL-FAULT-DB",
        command="RA_DON_NVL",
        sap=mock_sap,
        resolver=resolver,
        state_mgr=state_mgr,
        resolved_materials=["BIW32013765AA"],
        meta={
            "actualSenderEmail": "authorized_manager@vinfast.vn",
            "qty": 10,
            "station": "KV08",
            "unit": "PC",
            "sloc": "1005",
        },
        config={
            "allow_sap_writes": True,
            "trusted_intake_senders": ["authorized_manager@vinfast.vn"],
            "allow_nvl_reservations": True,
        },
    )

    handler = NvlHandler()

    # Mô phỏng SQLite ném lỗi Disk I/O hoặc Database Locked khi ghi nhận confirm_plan_reservation
    with patch.object(
        NvlService,
        "confirm_plan_reservation",
        side_effect=RuntimeError("sqlite3.OperationalError: database is locked / rollback initiated"),
    ):
        result = handler.handle(ctx)

    # Khẳng định bảo vệ an toàn:
    assert result.status == "AMBIGUOUS_RESERVATION"
    assert "AMBIGUOUS_RESERVATION" in result.summary
    assert "5001234567" in result.summary
    assert result.audit_target == "5001234567"
    assert "đối soát kiểm tra" in result.summary


def test_fault_injection_sqlite_reservation_released_on_sap_post_failure(tmp_path):
    """Mô phỏng khi SAP Post thất bại: giữ chỗ (reservation) trong SQLite phải được giải phóng."""
    db_file = tmp_path / "bridge_test_fail.db"
    state_mgr = StateManager(db_file)
    resolver = MaterialResolver()

    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_position_posting_change.return_value = [
        {
            "product": "BIW32013765AA",
            "batch": "LOT-A",
            "bin": "BIN-01",
            "plant": "3093",
            "sloc": "1005",
            "availQty": "100",
            "totalQty": "100",
            "unit": "PC",
            "grDate": "26.10.2026",
            "grTime": "08:00:00",
            "cat": "FT",
        }
    ]
    # SAP Post thất bại
    mock_sap.post_migo_311_ob_group.return_value = (
        False,
        "",
        "ERR_MIGO: Hết tồn kho trên SAP",
    )

    ctx = RequestContext(
        request_id="REQ-NVL-SAP-FAIL",
        command="RA_DON_NVL",
        sap=mock_sap,
        resolver=resolver,
        state_mgr=state_mgr,
        resolved_materials=["BIW32013765AA"],
        meta={
            "actualSenderEmail": "authorized_manager@vinfast.vn",
            "qty": 10,
            "station": "KV08",
            "unit": "PC",
            "sloc": "1005",
        },
        config={
            "allow_sap_writes": True,
            "trusted_intake_senders": ["authorized_manager@vinfast.vn"],
            "allow_nvl_reservations": True,
        },
    )

    handler = NvlHandler()

    with patch.object(NvlService, "release_plan_reservation", autospec=True) as mock_release:
        result = handler.handle(ctx)

        assert result.status == "FAILED"
        mock_release.assert_called_once()


# ==============================================================================
# PHẦN 4: MÔ PHỎNG NGƯỜI GỬI NGOÀI ALLOWLIST CỐ TÌNH KÍCH HOẠT GHI QUA BODY
# ==============================================================================

def test_fault_injection_spoofed_body_sender_ra_don_nvl(tmp_path):
    """Mô phỏng kẻ tấn công gửi mail kích hoạt RA_DON_NVL với body mạo danh người có quyền."""
    db_file = tmp_path / "bridge_test_sec.db"
    state_mgr = StateManager(db_file)
    mock_sap = create_autospec(SapEngine, instance=True)

    ctx = RequestContext(
        request_id="REQ-ATTACK-001",
        command="RA_DON_NVL",
        sap=mock_sap,
        state_mgr=state_mgr,
        meta={
            # Header thật từ transport envelope
            "actualSenderEmail": "attacker@external-domain.com",
            # Nội dung thư cố tình mạo danh Giám đốc / Trưởng ca
            "senderEmail": "authorized_manager@vinfast.vn",
            "senderName": "Tran Van Giam Doc",
            "qty": 50,
            "station": "LINE_01",
            "unit": "PC",
            "sloc": "1005",
        },
        config={
            "allow_sap_writes": True,
            "trusted_intake_senders": ["authorized_manager@vinfast.vn"],
        },
    )

    handler = NvlHandler()

    # Bắt buộc ném WritePermissionError trước khi gọi bất kỳ hàm SAP write nào
    with pytest.raises(WritePermissionError, match="không thuộc danh sách trusted_intake_senders"):
        handler.handle(ctx)

    # Đảm bảo tuyệt đối không có lệnh post MIGO nào được gọi
    mock_sap.post_migo_311_ob_group.assert_not_called()


def test_fault_injection_spoofed_body_sender_att_migo():
    """Mô phỏng kẻ tấn công cố tình đính kèm file độc hại vào MIGO qua lệnh ATT_MIGO."""
    mock_sap = create_autospec(SapEngine, instance=True)

    ctx = RequestContext(
        request_id="REQ-ATTACK-002",
        command="ATT_MIGO",
        sap=mock_sap,
        meta={
            "actualSenderEmail": "hacker@evil.org",
            "senderEmail": "bod_finance@vinfast.vn",  # Body mạo danh kế toán
            "explicitWriteIntent": True,
            "migoDocs": ["5001234567"],
        },
        config={
            "allow_sap_writes": True,
            "trusted_intake_senders": ["bod_finance@vinfast.vn"],
        },
    )

    handler = AttHandler()

    with pytest.raises(WritePermissionError, match="không thuộc danh sách trusted_intake_senders"):
        handler.handle(ctx)


def test_fault_injection_allow_sap_writes_false_blocks_writes(tmp_path):
    """Khi allow_sap_writes=False, dù người gửi đúng là Trưởng ca trong allowlist, lệnh ghi vẫn bị chặn."""
    db_file = tmp_path / "bridge_test_ro.db"
    state_mgr = StateManager(db_file)
    mock_sap = create_autospec(SapEngine, instance=True)

    ctx = RequestContext(
        request_id="REQ-NVL-READONLY-MODE",
        command="RA_DON_NVL",
        sap=mock_sap,
        state_mgr=state_mgr,
        meta={
            "actualSenderEmail": "authorized_manager@vinfast.vn",
            "qty": 10,
            "station": "LINE",
            "unit": "PC",
            "sloc": "1005",
        },
        config={
            "allow_sap_writes": False,  # Hệ thống đang ở chế độ bảo vệ cấm ghi
            "trusted_intake_senders": ["authorized_manager@vinfast.vn"],
        },
    )

    handler = NvlHandler()

    with pytest.raises(WritePermissionError, match="allow_sap_writes=False"):
        handler.handle(ctx)


def test_fault_injection_ke_hoach_nvl_always_preview_safe(tmp_path):
    """Lệnh KE_HOACH_NVL luôn là PREVIEW, không bao giờ thực hiện ghi SAP,
    kể cả khi allow_sap_writes=False hoặc người gửi chưa có trong allowlist."""
    db_file = tmp_path / "bridge_test_preview.db"
    state_mgr = StateManager(db_file)
    resolver = MaterialResolver()

    mock_sap = create_autospec(SapEngine, instance=True)
    mock_sap.check_position_posting_change.return_value = [
        {
            "product": "BIW32013765AA",
            "batch": "LOT-P",
            "bin": "BIN-PREVIEW",
            "plant": "3093",
            "sloc": "1005",
            "availQty": "100",
            "totalQty": "100",
            "unit": "PC",
            "grDate": "26.10.2026",
            "grTime": "08:00:00",
            "cat": "FT",
        }
    ]

    ctx = RequestContext(
        request_id="REQ-NVL-PREVIEW-001",
        command="KE_HOACH_NVL",
        sap=mock_sap,
        resolver=resolver,
        state_mgr=state_mgr,
        resolved_materials=["BIW32013765AA"],
        meta={
            "actualSenderEmail": "unknown_viewer@vinfast.vn",  # Người xem bất kỳ
            "qty": 20,
            "station": "KV08",
            "unit": "PC",
            "sloc": "1005",
        },
        config={
            "allow_sap_writes": False,  # Cấm ghi
            "trusted_intake_senders": [],
        },
    )

    handler = NvlHandler()
    res = handler.handle(ctx)

    assert res.status == "PREVIEW"
    assert "Xem trước kế hoạch" in res.summary
    # Xác nhận mock SAP post không hề được đụng tới
    mock_sap.post_migo_311_ob_group.assert_not_called()
