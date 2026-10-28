# -*- coding: utf-8 -*-
"""Unit tests for v1.0.91:
1. Multipart message chunking (2, 3, 4 parts for long batches).
2. Special characters filename matching in GOS attachments (VAT 6546 + 6562, VAT 6562(2)).
3. Sbar fallback confirmation (DONE_SAP_SBAR) when readback titles is empty.
4. Modal popup auto-dismissal (dismiss_any_modal_popup).
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock
import tempfile
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.teams_dispatcher import TeamsDispatcher, MAX_TEAMS_PAYLOAD_BYTES
from app.sap_attachment import MigoAttachmentRunner, dismiss_any_modal_popup, normalize_text


def test_multipart_payloads_under_threshold_returns_single_part():
    """Khi số lượng dòng <= 15, chỉ trả về đúng 1 payload duy nhất."""
    records = [{"doc": f"50070000{i:02d}", "ticket_no": f"TICK_{i}", "status": "DONE"} for i in range(10)]
    headers = {"doc": "Số DOC", "ticket_no": "Số phiếu", "status": "Trạng thái"}
    meta = {"command": "ATT_MIGO", "senderName": "Sếp"}
    
    parts = TeamsDispatcher.prepare_multipart_payloads("REQ-01", meta, "SUCCEEDED", "Batch 10 DOC", records, headers, chunk_size=15)
    assert len(parts) == 1
    p, d = parts[0]
    assert p["requestId"] == "REQ-01"
    assert "Phần" not in p["summary"]
    assert len(d) <= MAX_TEAMS_PAYLOAD_BYTES


def test_multipart_payloads_splits_into_multiple_parts_for_long_batch():
    """Khi số lượng dòng lớn (VD: 35 chứng từ), tự động tách thành 3 tin nhắn liên tiếp."""
    records = [{"doc": f"50070000{i:02d}", "ticket_no": f"TICK_{i}", "status": "DONE"} for i in range(35)]
    headers = {"doc": "Số DOC", "ticket_no": "Số phiếu", "status": "Trạng thái"}
    meta = {"command": "ATT_MIGO", "senderName": "Sếp"}
    
    parts = TeamsDispatcher.prepare_multipart_payloads("REQ-LARGE", meta, "SUCCEEDED", "Batch 35 DOC", records, headers, chunk_size=15)
    assert len(parts) == 3  # 15 + 15 + 5
    
    # Part 1
    p1, d1 = parts[0]
    assert p1["requestId"] == "REQ-LARGE_p1"
    assert "(Phần 1/3)" in p1["summary"]
    assert "Dữ liệu còn tiếp ở Phần 2/3" in p1["table_html"]
    assert len(d1) <= MAX_TEAMS_PAYLOAD_BYTES
    
    # Part 2
    p2, d2 = parts[1]
    assert p2["requestId"] == "REQ-LARGE_p2"
    assert "(Phần 2/3)" in p2["summary"]
    assert "Dữ liệu còn tiếp ở Phần 3/3" in p2["table_html"]
    assert len(d2) <= MAX_TEAMS_PAYLOAD_BYTES
    
    # Part 3
    p3, d3 = parts[2]
    assert p3["requestId"] == "REQ-LARGE_p3"
    assert "(Phần 3/3)" in p3["summary"]
    assert "Đã hiển thị trọn vẹn 100%" in p3["table_html"]
    assert len(d3) <= MAX_TEAMS_PAYLOAD_BYTES


def _make_mock_att_session():
    mock_session = MagicMock()
    mock_session.Info.SystemName = "TEST"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "TESTUSER"
    mock_session.Info.Transaction = "MIGO"
    mock_session.Id = "/app/con[0]/ses[1]"
    mock_session.Busy = False

    mock_sbar = MagicMock()
    mock_sbar.Text = "Document created"
    mock_sbar.MessageType = "S"

    mock_shell = MagicMock()
    wnd1_active = [False]
    mock_shell.SelectContextMenuItem.side_effect = lambda *a, **k: wnd1_active.__setitem__(0, True)
    mock_btn = MagicMock()
    mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_btn.Press.side_effect = lambda: wnd1_active.__setitem__(0, False)
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.side_effect = lambda name, *args: mock_btn if "btn" in name else MagicMock()

    def find_by_id(path):
        p = str(path)
        if "sbar" in p:
            return mock_sbar
        if "shell" in p:
            return mock_shell
        if "btn" in p:
            return mock_btn
        if p == "wnd[1]" or p.startswith("wnd[1]"):
            if wnd1_active[0]:
                return mock_wnd1
            raise Exception("Control wnd[1] not found")
        return MagicMock()

    mock_session.findById.side_effect = find_by_id
    return mock_session, mock_sbar


def test_special_characters_filename_flexible_matching():
    """Kiểm tra đối soát tên file có ký tự đặc biệt '+' hoặc '(2)' khớp linh hoạt."""
    mock_session, mock_sbar = _make_mock_att_session()
    # sbar bình thường để kiểm tra nhánh readback
    mock_sbar.Text = "Ready"
    mock_sbar.MessageType = "I"
    
    runner = MigoAttachmentRunner(expected_session_identity={"system": "TEST", "client": "100", "user": "TESTUSER", "session_id": mock_session.Id})
    runner._active_document = ("5007158704", "2026")
    
    # File 1: VAT 6546 + 6562.pdf -> SAP title: 'VAT 6546 + 6562'
    with tempfile.NamedTemporaryFile(suffix="_VAT 6546 + 6562.pdf", delete=False) as tf:
        dummy_file1 = Path(tf.name)
    try:
        runner.read_sap_attachment_titles = lambda s: ["VAT 6546 + 6562"]
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file1)
        assert ok is True
        assert "DONE_SAP_READBACK" in msg
    finally:
        dummy_file1.unlink(missing_ok=True)
        
    # File 2: VAT 6562(2).pdf -> SAP title: 'VAT 6562_2'
    mock_session2, mock_sbar2 = _make_mock_att_session()
    mock_sbar2.Text = "Ready"
    mock_sbar2.MessageType = "I"
    runner2 = MigoAttachmentRunner(expected_session_identity={"system": "TEST", "client": "100", "user": "TESTUSER", "session_id": mock_session2.Id})
    runner2._active_document = ("5007158704", "2026")

    with tempfile.NamedTemporaryFile(suffix="_VAT 6562(2).pdf", delete=False) as tf2:
        dummy_file2 = Path(tf2.name)
    try:
        runner2.read_sap_attachment_titles = lambda s: ["VAT 6562_2"]
        ok2, msg2 = runner2.attach_file_to_current_migo(mock_session2, dummy_file2)
        assert ok2 is True
        assert "DONE_SAP_READBACK" in msg2
    finally:
        dummy_file2.unlink(missing_ok=True)


def test_sbar_fallback_confirmation_when_readback_titles_empty():
    """Khi readback trả về rỗng (do MIGO chưa kịp vẽ lại nút OK_REF_ATTA), nhưng sbar có 'Document created' -> DONE_SAP_SBAR."""
    mock_session, mock_sbar = _make_mock_att_session()
    mock_sbar.Text = "Document created"
    mock_sbar.MessageType = "S"

    runner = MigoAttachmentRunner(expected_session_identity={"system": "TEST", "client": "100", "user": "TESTUSER", "session_id": mock_session.Id})
    runner._active_document = ("5007158704", "2026")

    with tempfile.NamedTemporaryFile(suffix="_VAT.pdf", delete=False) as tf:
        dummy_file = Path(tf.name)
    try:
        # Readback popup không mở được hoặc trả về rỗng []
        runner.read_sap_attachment_titles = lambda s: []
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is True
        assert "DONE_SAP_SBAR" in msg
    finally:
        dummy_file.unlink(missing_ok=True)


def test_dismiss_any_modal_popup():
    """Kiểm tra tự động dập tắt popup modal wnd[1] bằng nút OK/Yes."""
    mock_session = MagicMock()
    mock_popup = MagicMock()
    mock_btn_yes = MagicMock()

    mock_popup.findById = lambda path, raise_err=True: mock_btn_yes if "btn" in path else None

    def find_by_id(path, raise_err=True):
        p = str(path)
        if p == "wnd[1]":
            return mock_popup
        if "btn" in p:
            return mock_btn_yes
        return None
    mock_session.findById.side_effect = find_by_id

    dismissed = dismiss_any_modal_popup(mock_session)
    assert dismissed is True
    mock_btn_yes.Press.assert_called_once()
