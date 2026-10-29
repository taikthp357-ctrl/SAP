# -*- coding: utf-8 -*-
"""Unit tests for v1.1.01 features:
1. Multi-Tab Turbo Max (Adaptive 5-6 SAP GUI Tabs, User Tab 1 preservation).
2. Elimination of pythoncom.CoUninitialize in ensure_sap_tabs to keep Outlook and SAP COM apartments intact.
3. SAP Engine case-insensitive session identity check and connection_id reconnect resilience.
"""

from unittest.mock import MagicMock, patch
import pytest

from app.sap_attachment import ensure_sap_tabs, MigoAttachmentRunner
from app.sap_engine import SapEngine, SapExecutionError


class TestSapTabsAndApartmentV1101:
    """Kiểm thử quản lý Tab SAP và căn hộ COM trong v1.1.01."""

    def test_ensure_sap_tabs_caps_at_7_and_does_not_uninitialize_com(self):
        """Đảm bảo ensure_sap_tabs chặn trần ở 7 tabs và KHÔNG gọi CoUninitialize trên thread chính."""
        mock_conn = MagicMock()
        mock_s0 = MagicMock()
        mock_s0.Busy = False
        mock_conn.Children = MagicMock(return_value=mock_s0)
        mock_conn.Children.Count = 1

        def fake_create_session():
            mock_conn.Children.Count += 1

        mock_s0.CreateSession.side_effect = fake_create_session

        mock_app = MagicMock()
        mock_app.Children = MagicMock(return_value=mock_conn)
        mock_app.Children.Count = 1

        mock_sap_gui = MagicMock()
        mock_sap_gui.GetScriptingEngine = mock_app

        with patch("app.sap_attachment.pythoncom") as mock_pythoncom, \
             patch("app.sap_attachment.win32com.client.GetObject", return_value=mock_sap_gui), \
             patch("time.sleep"):
            # Yêu cầu 10 tabs, nhưng ensure_sap_tabs chỉ cho tối đa 7
            final_tabs = ensure_sap_tabs(target_tabs=10, connection_index=0)
            assert final_tabs == 7
            # Đảm bảo pythoncom.CoUninitialize KHÔNG bao giờ được gọi
            mock_pythoncom.CoUninitialize.assert_not_called()

    def test_ensure_sap_tabs_graceful_when_server_caps_at_6(self):
        """Khi server SAP giới hạn 6 session và từ chối mở tab thứ 7, ensure_sap_tabs trả về an toàn 6 tabs."""
        mock_conn = MagicMock()
        mock_s0 = MagicMock()
        mock_s0.Busy = False
        mock_conn.Children = MagicMock(return_value=mock_s0)
        mock_conn.Children.Count = 1

        def fake_create_session():
            if mock_conn.Children.Count >= 6:
                raise RuntimeError("Maximum number of sessions reached (6)")
            mock_conn.Children.Count += 1

        mock_s0.CreateSession.side_effect = fake_create_session

        mock_app = MagicMock()
        mock_app.Children = MagicMock(return_value=mock_conn)
        mock_app.Children.Count = 1

        mock_sap_gui = MagicMock()
        mock_sap_gui.GetScriptingEngine = mock_app

        with patch("app.sap_attachment.pythoncom") as mock_pythoncom, \
             patch("app.sap_attachment.win32com.client.GetObject", return_value=mock_sap_gui), \
             patch("time.sleep"):
            final_tabs = ensure_sap_tabs(target_tabs=7, connection_index=0)
            assert final_tabs == 6
            mock_pythoncom.CoUninitialize.assert_not_called()

    def test_ensure_sap_tabs_picks_idle_session_to_create(self):
        """Đảm bảo tìm session không Busy để gọi CreateSession."""
        mock_conn = MagicMock()
        mock_s0 = MagicMock()
        mock_s0.Busy = True  # Đang bận
        mock_s1 = MagicMock()
        mock_s1.Busy = False # Đang rảnh

        mock_conn.Children = MagicMock(side_effect=lambda idx: [mock_s0, mock_s1][idx])
        mock_conn.Children.Count = 2

        def fake_create():
            mock_conn.Children.Count = 3

        mock_s1.CreateSession.side_effect = fake_create

        mock_app = MagicMock()
        mock_app.Children = MagicMock(return_value=mock_conn)
        mock_app.Children.Count = 1

        mock_sap_gui = MagicMock()
        mock_sap_gui.GetScriptingEngine = mock_app

        with patch("app.sap_attachment.pythoncom"), \
             patch("app.sap_attachment.win32com.client.GetObject", return_value=mock_sap_gui), \
             patch("time.sleep"):
            tabs = ensure_sap_tabs(target_tabs=3, connection_index=0)
            mock_s1.CreateSession.assert_called_once()
            mock_s0.CreateSession.assert_not_called()
            assert tabs == 3


class TestTurboMaxAllocationV1101:
    """Kiểm thử phân bổ Tab Turbo Max trong process_attachment_batch."""

    def test_turbo_max_allocates_up_to_6_bot_tabs_when_7_tabs_available(self):
        """Lệnh 'att max' với 10 items khi có 7 tabs sẽ dùng 6 bot tabs (Tab 2..7) và bảo tồn Tab 1 cho User."""
        mock_sap = MagicMock()
        mock_sap.session_identity = {"system": "PRD", "client": "100", "user": "TAIPH3", "connection_index": 0}
        mock_sap.check_attachments_zfir266.return_value = []

        runner = MigoAttachmentRunner(sap_engine=mock_sap, expected_session_identity=mock_sap.session_identity)
        runner._prompt_context = "Sap att max"

        items = [{"doc": f"50071{i:05d}", "ticket_no": f"TICK{i}"} for i in range(10)]

        with patch("app.sap_attachment.ensure_sap_tabs", return_value=7) as mock_ensure, \
             patch("app.sap_attachment.get_user_idle_seconds", return_value=10.0), \
             patch.object(runner, "prepare_migo_session", return_value=True), \
             patch.object(runner, "process_single_migo_item") as mock_item:

            mock_item.return_value = ({"doc": "123", "overall_status": "DONE", "is_succ": True}, True)

            summary = runner.process_attachment_batch(folder_path="", items=items, fill_headtext=True, dry_run=False, year="2026")

            mock_ensure.assert_called_once_with(7, connection_index=0)
            assert summary["multi_tab_count"] == 6
            assert summary["is_turbo"] is True

    def test_turbo_max_adapts_to_5_bot_tabs_when_6_tabs_available(self):
        """Lệnh 'att max' khi hệ thống chỉ có 6 tabs sẽ dùng 5 bot tabs (Tab 2..6) và bảo tồn Tab 1 cho User."""
        mock_sap = MagicMock()
        mock_sap.session_identity = {"system": "PRD", "client": "100", "user": "TAIPH3", "connection_index": 0}
        mock_sap.check_attachments_zfir266.return_value = []

        runner = MigoAttachmentRunner(sap_engine=mock_sap, expected_session_identity=mock_sap.session_identity)
        runner._prompt_context = "Sap att max"

        items = [{"doc": f"50071{i:05d}", "ticket_no": f"TICK{i}"} for i in range(10)]

        with patch("app.sap_attachment.ensure_sap_tabs", return_value=6) as mock_ensure, \
             patch("app.sap_attachment.get_user_idle_seconds", return_value=10.0), \
             patch.object(runner, "prepare_migo_session", return_value=True), \
             patch.object(runner, "process_single_migo_item") as mock_item:

            mock_item.return_value = ({"doc": "123", "overall_status": "DONE", "is_succ": True}, True)

            summary = runner.process_attachment_batch(folder_path="", items=items, fill_headtext=True, dry_run=False, year="2026")

            mock_ensure.assert_called_once_with(7, connection_index=0)
            assert summary["multi_tab_count"] == 5
            assert summary["is_turbo"] is True


class TestSapEngineIdentityStabilityV1101:
    """Kiểm thử tính ổn định của SAP Engine Identity."""

    def test_check_identity_case_insensitive_and_connection_id_resilience(self):
        """Khớp danh tính không phân biệt hoa thường và không văng lỗi khi connection_id thay đổi."""
        engine = SapEngine(expected_system="PRD", expected_client="100", expected_user="TAIPH3")
        engine.session_identity = {
            "system": "PRD",
            "client": "100",
            "user": "TAIPH3",
            "session_id": "/app/con[0]/ses[0]",
            "connection_id": "0"
        }

        mock_session = MagicMock()
        mock_session.Id = "/app/con[1]/ses[0]"
        mock_info = MagicMock()
        mock_info.SystemName = "prd"
        mock_info.Client = "100"
        mock_info.User = "taiph3"
        mock_session.Info = mock_info

        current = engine._check_identity(mock_session)
        assert current["system"] == "prd"
        assert current["connection_id"] == "1"

    def test_check_identity_raises_on_actual_user_or_system_mismatch(self):
        """Phải cảnh báo dừng lại nếu User hoặc System thực sự thay đổi khác người dùng."""
        engine = SapEngine(expected_system="", expected_client="", expected_user="")
        engine.session_identity = {
            "system": "PRD",
            "client": "100",
            "user": "TAIPH3",
            "session_id": "/app/con[0]/ses[0]",
            "connection_id": "0"
        }

        # Session khác user
        mock_session = MagicMock()
        mock_session.Id = "/app/con[0]/ses[0]"
        mock_info = MagicMock()
        mock_info.SystemName = "PRD"
        mock_info.Client = "100"
        mock_info.User = "HACKER99"
        mock_session.Info = mock_info

        with pytest.raises(SapExecutionError, match="Danh tính phiên SAP thay đổi"):
            engine._check_identity(mock_session)

    def test_ensure_session_handles_case_insensitivity(self):
        """ensure_session kiểm tra danh tính không bị văng lỗi do lệch hoa thường."""
        engine = SapEngine(expected_system="PRD", expected_client="100", expected_user="TAIPH3")
        engine.session_identity = {
            "system": "PRD",
            "client": "100",
            "user": "TAIPH3",
            "session_id": "/app/con[0]/ses[0]",
            "connection_id": "0"
        }

        mock_session = MagicMock()
        mock_session.Id = "/app/con[0]/ses[0]"
        mock_info = MagicMock()
        mock_info.SystemName = "prd"
        mock_info.Client = "100"
        mock_info.User = "taiph3"
        mock_session.Info = mock_info
        mock_session.Busy = False
        engine.session = mock_session

        with patch.object(engine, "_is_session_alive", return_value=True):
            sess = engine.ensure_session()
            assert sess == mock_session
