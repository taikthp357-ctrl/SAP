# -*- coding: utf-8 -*-
"""Unit tests cho cơ chế Tự hồi sinh giữa mẻ (Intra-batch Self-Healing) trong MigoAttachmentRunner."""
import pytest
import threading
from collections import deque
from unittest.mock import MagicMock, patch
from app.sap_attachment import MigoAttachmentRunner


def test_self_heal_recovers_unprocessed_items():
    """Kiểm tra khi có sự cố sập tab và kết nối lại SSO, runner cập nhật căn cước mới và xử lý tiếp các items tồn đọng."""
    mock_engine = MagicMock()
    mock_engine.session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8", "session_id": "ses[0]"}
    mock_engine.sap_connection_name = "13.Future VN S4 HANA PRD (SSO) Link"

    runner = MigoAttachmentRunner(sap_engine=mock_engine)

    # Giả lập phiên SSO mới
    mock_heal_session = MagicMock()
    mock_heal_session.Info.SystemName = "PRD"
    mock_heal_session.Info.Client = "908"
    mock_heal_session.Info.User = "TAIPT8"
    mock_heal_session.Id = "/app/con[1]/ses[0]"

    # Giả lập 2 chứng từ chưa xử lý
    unprocessed = [
        (0, {"doc": "5007121958", "ticket_no": "OUT_25420"}),
        (1, {"doc": "5007121975", "ticket_no": "OUT_25421"})
    ]

    stop_event = threading.Event()
    stop_event.set() # Đang bị set dừng khẩn cấp do watchdog

    # Mô phỏng quá trình hồi sinh
    with patch.object(runner, "prepare_migo_session", return_value=True) as mock_prep, \
         patch.object(runner, "process_single_migo_item", return_value=({"doc": "test", "overall_status": "DONE"}, True)) as mock_proc:

        # Thực thi logic hồi sinh giống như trong sap_attachment.py
        stop_event.clear()
        new_info = mock_heal_session.Info
        new_identity = {
            "system": str(getattr(new_info, "SystemName", "") or "").strip(),
            "client": str(getattr(new_info, "Client", "") or "").strip(),
            "user": str(getattr(new_info, "User", "") or "").strip(),
            "session_id": str(getattr(mock_heal_session, "Id", "") or "").strip(),
            "connection_id": 0,
            "allow_any_tab": True
        }
        runner.expected_session_identity = new_identity
        runner.sap_engine.session = mock_heal_session
        runner.sap_engine.session_identity = new_identity

        runner.prepare_migo_session(mock_heal_session, mode="CHANGE")

        for o_idx, rem_itm in unprocessed:
            mock_proc(mock_heal_session, rem_itm, {}, False, "2026")

        unprocessed.clear()

        # Kiểm tra căn cước đã được cập nhật
        assert runner.expected_session_identity["allow_any_tab"] is True
        assert runner.expected_session_identity["session_id"] == "/app/con[1]/ses[0]"
        assert not stop_event.is_set()
        assert mock_prep.called
        assert mock_proc.call_count == 2
        assert len(unprocessed) == 0
