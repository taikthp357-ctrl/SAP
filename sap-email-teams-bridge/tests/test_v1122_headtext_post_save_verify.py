# -*- coding: utf-8 -*-
"""Unit tests cho tính năng v1.1.22:
1. Post-Save Double Verification: Thẩm định trạng thái lưu thực tế trên SAP, triệt tiêu báo lỗi ảo khi Save gặp COM timeout.
2. Cross-Document Status Bar Sanitization: Lọc bỏ số DOC của chứng từ cũ còn sót lại trong status bar.
3. Strict Session Isolation: Ngăn chặn 2 luồng worker dùng chung 1 SAP session COM.
4. Intra-Batch Self-Healing: Tự động cứu hộ và xác thực lại các DOC bị báo lỗi EXCEPTION/COM timeout.
"""
import pytest
from unittest.mock import MagicMock, patch
from collections import deque
from app.writers.migo_headtext_writer import MigoHeadtextWriter
from app.core.sap_primitives import get_safe_sap_session


def test_v1122_post_save_verify_rescues_com_exception():
    """Kiểm tra khi SendVKey(11) gặp COM exception (Save lag/timeout) nhưng SAP đã lưu xong trong DB -> Trả về SUCCESS (FILLED)."""
    session = MagicMock()
    session.Info.Transaction = "MIGO"
    session.Busy = False

    w0 = MagicMock()
    # Giả lập SendVKey(11) quăng COM error do Callee Rejected
    def send_vkey_mock(key):
        if key == 11:
            raise RuntimeError("(-2147418111, 'Call was rejected by callee')")
        return None

    w0.SendVKey.side_effect = send_vkey_mock

    titl = MagicMock()
    titl.Text = "Change Material Document 4906007995"
    sbar = MagicMock()
    sbar.Text = "Document 4906007995 2026 has been changed"
    sbar.MessageType = "S"

    def find_side_effect(id_str, *args):
        if id_str == "wnd[0]":
            return w0
        if id_str == "wnd[0]/titl":
            return titl
        if id_str == "wnd[0]/sbar":
            return sbar
        return None

    session.findById.side_effect = find_side_effect

    mock_doc_fld = MagicMock()
    mock_doc_fld.Changeable = True
    mock_year_fld = MagicMock()
    mock_head_fld = MagicMock()
    mock_head_fld.Text = ""  # Ban đầu trống

    def smart_find_mock(ses, name, *args, **kwargs):
        if name == "GODYNPRO-MAT_DOC":
            return mock_doc_fld
        if name == "GODYNPRO-DOC_YEAR":
            return mock_year_fld
        if name == "GOHEAD-BKTXT":
            return mock_head_fld
        return None

    with patch("app.writers.migo_headtext_writer.smart_find", side_effect=smart_find_mock), \
         patch("app.writers.migo_headtext_writer.wait_idle", return_value=True), \
         patch("app.writers.migo_headtext_writer.dismiss_popups", return_value=""):

        success, code, msg = MigoHeadtextWriter.fill(
            session=session,
            doc="4906007995",
            year="2026",
            ticket_no="OUT_P1_26801",
            allow_overwrite=False
        )

        assert success is True
        assert code == "FILLED"
        assert "4906007995" in msg


def test_v1122_cross_doc_sbar_leak_sanitization():
    """Kiểm tra nếu sbar còn sót lại thông báo của chứng từ cũ (4906007995) khi đang xử lý DOC mới (4906007456), message rác bị loại bỏ."""
    session = MagicMock()
    session.Info.Transaction = "MIGO"
    session.Busy = False

    w0 = MagicMock()
    titl = MagicMock()
    titl.Text = "Change Material Document"
    sbar = MagicMock()
    # Sbar chứa thông báo của chứng từ KHÁC
    sbar.Text = "Document 4906007995 2026 has been changed"
    sbar.MessageType = "S"

    def find_side_effect(id_str, *args):
        if id_str == "wnd[0]":
            return w0
        if id_str == "wnd[0]/titl":
            return titl
        if id_str == "wnd[0]/sbar":
            return sbar
        return None

    session.findById.side_effect = find_side_effect

    mock_doc_fld = MagicMock()
    mock_doc_fld.Changeable = True
    mock_year_fld = MagicMock()
    mock_head_fld = MagicMock()
    mock_head_fld.Text = ""  # Ban đầu trống

    def smart_find_mock(ses, name, *args, **kwargs):
        if name == "GODYNPRO-MAT_DOC":
            return mock_doc_fld
        if name == "GODYNPRO-DOC_YEAR":
            return mock_year_fld
        if name == "GOHEAD-BKTXT":
            return mock_head_fld
        return None

    with patch("app.writers.migo_headtext_writer.smart_find", side_effect=smart_find_mock), \
         patch("app.writers.migo_headtext_writer.wait_idle", return_value=True), \
         patch("app.writers.migo_headtext_writer.dismiss_popups", return_value=""):

        success, code, msg = MigoHeadtextWriter.fill(
            session=session,
            doc="4906007456",
            year="2026",
            ticket_no="OUT_P1_26808",
            allow_overwrite=False
        )

        assert success is True
        assert code == "FILLED"
        # msg KHÔNG ĐƯỢC CHỨA 4906007995
        assert "4906007995" not in msg
        assert "4906007456" in msg


def test_v1122_session_isolation_rejects_duplicate_sessions():
    """Kiểm tra khi get_safe_sap_session trả về cùng 1 session ID cho 2 tab -> Luồng trùng lặp tự động nhả việc và dừng lại."""
    shared_session = MagicMock()
    shared_session.Id = "/app/con[0]/ses[0]"

    def mock_get_session(conn_idx, tab_idx):
        # Cả 2 tab đều trỏ về shared_session
        return shared_session

    items = [
        {"doc": "4906007995", "ticket_no": "OUT_P1_26801"},
        {"doc": "4906007456", "ticket_no": "OUT_P1_26808"}
    ]

    mock_conn = MagicMock()
    mock_conn.Children.Count = 2

    with patch("app.writers.migo_headtext_writer.ensure_sap_tabs", return_value=2), \
         patch("app.writers.migo_headtext_writer.get_safe_sap_connection", return_value=mock_conn), \
         patch("app.writers.migo_headtext_writer.get_safe_sap_session", side_effect=mock_get_session), \
         patch.object(MigoHeadtextWriter, "fill", return_value=(True, "FILLED", "OK")):

        records, succ = MigoHeadtextWriter.fill_batch_multi_tab(
            items=items,
            year="2026",
            prompt_raw="2 tab",
            connection_index=0
        )

        # Cả 2 phần tử đều phải được xử lý thành công không bị kẹt hay rơi rụng
        assert succ == 2
        assert len(records) == 2


def test_v1122_intra_batch_self_healing_rescues_failed_items():
    """Kiểm tra cơ chế Intra-Batch Self-Healing: Khi một DOC bị EXCEPTION trên luồng song song, bước 6 tự động thẩm định lại và cứu thành công."""
    items = [
        {"doc": "4906007995", "ticket_no": "OUT_P1_26801"},
        {"doc": "4906007456", "ticket_no": "OUT_P1_26808"}
    ]

    mock_s0 = MagicMock(); mock_s0.Id = "/app/con[0]/ses[0]"
    mock_s1 = MagicMock(); mock_s1.Id = "/app/con[0]/ses[1]"
    sessions = [mock_s0, mock_s1]

    mock_conn = MagicMock()
    mock_conn.Children.Count = 2

    call_count = 0
    def mock_fill(session, doc, year, ticket_no, allow_overwrite):
        nonlocal call_count
        call_count += 1
        if doc == "4906007995":
            if call_count <= 2:
                # Lần đầu trong worker thread ném EXCEPTION
                raise RuntimeError("Call rejected by callee")
            else:
                # Lần sau trong bước Self-Healing phát hiện đã có sẵn trên SAP
                return True, "ALREADY_FILLED", "Header Text đã có sẵn số phiếu khớp (OUT_P1_26801)"
        return True, "FILLED", f"Lưu thành công {doc}"

    with patch("app.writers.migo_headtext_writer.ensure_sap_tabs", return_value=2), \
         patch("app.writers.migo_headtext_writer.get_safe_sap_connection", return_value=mock_conn), \
         patch("app.writers.migo_headtext_writer.get_safe_sap_session", side_effect=lambda c, t: sessions[t] if t < len(sessions) else None), \
         patch.object(MigoHeadtextWriter, "fill", side_effect=mock_fill):

        records, succ = MigoHeadtextWriter.fill_batch_multi_tab(
            items=items,
            year="2026",
            prompt_raw="2 tab",
            connection_index=0
        )

        assert succ == 2
        r_map = {r["doc"]: r for r in records}
        assert r_map["4906007995"]["overall_status"] == "DONE"
        assert r_map["4906007995"]["headtext_status"] == "ALREADY_FILLED"
        assert r_map["4906007456"]["overall_status"] == "DONE"


def test_v1122_get_safe_sap_session_no_fallback_to_zero_for_tab_gt_zero():
    """Kiểm tra get_safe_sap_session KHÔNG fallback về tab 0 khi tab_idx > 0 và không tồn tại."""
    mock_conn = MagicMock()
    mock_conn.Children.Count = 1
    mock_s0 = MagicMock()
    mock_s0.Id = "/app/con[0]/ses[0]"
    mock_conn.Children.side_effect = lambda idx: mock_s0 if idx == 0 else None

    mock_app = MagicMock()
    mock_app.Children.Count = 1
    mock_app.Children.return_value = mock_conn

    mock_gui = MagicMock()
    mock_gui.GetScriptingEngine = mock_app

    with patch("app.core.sap_primitives.pythoncom"), \
         patch("app.core.sap_primitives.win32com.client.GetObject", return_value=mock_gui):

        # Tab 0 hợp lệ -> trả về mock_s0
        s0 = get_safe_sap_session(0, 0)
        assert s0 == mock_s0

        # Tab 1 không tồn tại -> PHẢI trả về None (KHÔNG được trả về mock_s0!)
        s1 = get_safe_sap_session(0, 1)
        assert s1 is None
