# -*- coding: utf-8 -*-
"""Unit tests cho StateManager (SQLite hàng đợi và chống trùng lặp)."""

import sys
import tempfile
from pathlib import Path
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.state_manager import StateManager


@pytest.fixture
def temp_state_mgr(tmp_path):
    db_file = tmp_path / "test_bridge_state.db"
    return StateManager(db_file)


def test_processed_messages_flow(temp_state_mgr):
    mgr = temp_state_mgr

    assert not mgr.is_message_processed("MSG-001", "ENTRY-001")

    # Đánh dấu đã xử lý
    mgr.mark_message_processed("MSG-001", "ENTRY-001", "SAP-001", "VITRI", "SUCCEEDED")

    # Kiểm tra lại cả qua message_id lẫn entry_id
    assert mgr.is_message_processed("MSG-001")
    assert mgr.is_message_processed("", "ENTRY-001")
    assert not mgr.is_message_processed("MSG-002", "ENTRY-002")


def test_outbox_queue_flow(temp_state_mgr):
    mgr = temp_state_mgr

    payload = {"requestId": "REQ-100", "summary": "Kết quả test", "status": "SUCCEEDED"}
    mgr.enqueue_outbox("REQ-100", "https://webhook.test/api", payload)

    pending = mgr.get_pending_outbox(max_attempts=3)
    assert len(pending) == 1
    assert pending[0]["req_id"] == "REQ-100"
    assert pending[0]["attempts"] == 0
    assert pending[0]["payload"]["summary"] == "Kết quả test"

    # Ghi nhận thất bại lần 1
    mgr.record_outbox_failure("REQ-100", "Network timeout", max_attempts=3)
    pending_after_fail = mgr.get_pending_outbox(max_attempts=3)
    assert len(pending_after_fail) == 1
    assert pending_after_fail[0]["attempts"] == 1

    # Ghi nhận thành công -> phải bị xóa khỏi outbox
    mgr.record_outbox_success("REQ-100")
    pending_done = mgr.get_pending_outbox(max_attempts=3)
    assert len(pending_done) == 0


def test_outbox_max_attempts_exceeded(temp_state_mgr):
    mgr = temp_state_mgr
    payload = {"requestId": "REQ-EXCEED"}
    mgr.enqueue_outbox("REQ-EXCEED", "https://webhook.test/api", payload)

    # Thử 3 lần
    mgr.record_outbox_failure("REQ-EXCEED", "Fail 1", max_attempts=3)
    mgr.record_outbox_failure("REQ-EXCEED", "Fail 2", max_attempts=3)
    mgr.record_outbox_failure("REQ-EXCEED", "Fail 3", max_attempts=3)

    # Đã vượt quá 3 lần, không được trả về trong pending nữa
    pending = mgr.get_pending_outbox(max_attempts=3)
    assert len(pending) == 0


def test_claim_request_lifecycle_and_helpers(temp_state_mgr):
    mgr = temp_state_mgr
    req_id = "REF-2aef3802b46b52f69f1e112008ac2ae8"

    assert mgr.get_claim_status(req_id) is None
    assert not mgr.has_outbox_result(req_id)

    # Claim lần đầu -> thành công (RUNNING)
    assert mgr.claim_request(req_id) is True
    assert mgr.get_claim_status(req_id) == "RUNNING"

    # Claim lần 2 -> thất bại vì đã claim
    assert mgr.claim_request(req_id) is False

    # Giả lập worker bị ngắt đột ngột và gọi recover_interrupted
    mgr.recover_interrupted()
    assert mgr.get_claim_status(req_id) == "RECOVERY_REQUIRED"

    # Hoàn tất request và lưu outbox
    mgr.complete_request("MSG-REC", "ENTRY-REC", req_id, "VITRI", "SUCCEEDED", "https://flow.test", {"status": "OK"})
    assert mgr.get_claim_status(req_id) == "COMPLETED"
    assert mgr.has_outbox_result(req_id) is True
    assert mgr.is_message_processed("MSG-REC", "ENTRY-REC") is True
