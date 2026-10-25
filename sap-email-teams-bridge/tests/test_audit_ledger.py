# -*- coding: utf-8 -*-
"""Unit tests for Audit Ledger (Đợt 7: Sổ cái kiểm toán bất biến)."""

import pytest
from app.state_manager import StateManager


def test_audit_ledger_log_and_query(tmp_path):
    db_file = tmp_path / "test_audit.db"
    mgr = StateManager(db_path=db_file)

    row_id1 = mgr.log_audit(
        req_id="REQ-001",
        action_type="MIGO_311",
        status="SUCCESS",
        target_doc="4900112233",
        user_name="Nguyễn Văn Kho",
        payload={"material": "BIW32013765AA", "qty": 50},
        message="Xuất chuyển kho thành công"
    )
    assert row_id1 > 0

    row_id2 = mgr.log_audit(
        req_id="REQ-002",
        action_type="VL02N_PICK",
        status="BLOCKED",
        target_doc="8000999999",
        user_name="Lưu Thị Ngọc Mai",
        payload={"deficit_count": 2},
        message="Thiếu tồn kho toàn lô"
    )
    assert row_id2 > row_id1

    # Query all
    logs = mgr.query_audit_logs()
    assert len(logs) == 2
    assert logs[0]["req_id"] == "REQ-002"  # Order by id DESC
    assert logs[0]["action_type"] == "VL02N_PICK"
    assert logs[0]["payload"]["deficit_count"] == 2

    # Query with filter
    filter_migo = mgr.query_audit_logs(action_type="MIGO_311")
    assert len(filter_migo) == 1
    assert filter_migo[0]["target_doc"] == "4900112233"
