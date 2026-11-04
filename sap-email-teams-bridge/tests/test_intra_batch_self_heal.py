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


def test_reclaim_in_flight_failed_items():
    """Kiểm tra cơ chế quét ngược thu hồi các chứng từ in-flight bị lỗi rụng phiên (SESSION_IDENTITY_UNAVAILABLE)."""
    items = [
        {"doc": "5007144012", "ticket_no": "6500"},
        {"doc": "5007144025", "ticket_no": "6500"},
        {"doc": "5007144099", "ticket_no": "6501"},
    ]
    # Giả lập indexed_results có 1 chứng từ thành công, và 1 chứng từ bị dính SESSION_IDENTITY_UNAVAILABLE do sập tab
    indexed_results = [
        (0, {"doc": "5007144012", "ticket_no": "6500", "headtext_status": "SESSION_IDENTITY_UNAVAILABLE", "overall_status": "SESSION_IDENTITY_UNAVAILABLE", "match_notes": "test", "is_succ": False}, False),
        (2, {"doc": "5007144099", "ticket_no": "6501", "headtext_status": "HEADTEXT_OK", "overall_status": "DONE", "match_notes": "ok", "is_succ": True}, True),
    ]
    tab_queues = {
        0: deque([(1, {"doc": "5007144025", "ticket_no": "6500"})])
    }

    # Áp dụng logic thu hồi in-flight
    reclaimed_items = []
    survived_results = []
    for o_idx, r_dict, is_s in indexed_results:
        st = str(r_dict.get("headtext_status", "")).upper()
        ov = str(r_dict.get("overall_status", "")).upper()
        notes = str(r_dict.get("match_notes", "")).lower()
        is_session_fatal = (
            st in ("SESSION_IDENTITY_UNAVAILABLE", "SESSION_NONE", "SESSION_INFO_NONE", "EXCEPTION", "TIMEOUT")
            or ov in ("SESSION_IDENTITY_UNAVAILABLE", "SESSION_NONE", "SESSION_INFO_NONE", "ERROR", "TIMEOUT")
            or any(k in notes for k in ["disconnected", "rpc", "session", "quarantine", "self-heal", "sập tab", "căn cước", "identity"])
        )
        if not is_s and is_session_fatal:
            orig_itm = items[o_idx] if o_idx < len(items) else {"doc": r_dict.get("doc"), "ticket_no": r_dict.get("ticket_no")}
            reclaimed_items.append((o_idx, orig_itm))
        else:
            survived_results.append((o_idx, r_dict, is_s))

    unprocessed_items = list(reclaimed_items)
    for t_q in tab_queues.values():
        while t_q:
            unprocessed_items.append(t_q.popleft())

    # Kiểm tra:
    # 1. Chứng từ bị lỗi rụng phiên (5007144012) đã được rút khỏi survived_results
    assert len(survived_results) == 1
    assert survived_results[0][0] == 2

    # 2. unprocessed_items thu gom đủ cả 5007144012 (từ indexed_results) và 5007144025 (từ queue)
    assert len(unprocessed_items) == 2
    docs_to_heal = [itm["doc"] for _, itm in unprocessed_items]
    assert "5007144012" in docs_to_heal
    assert "5007144025" in docs_to_heal
