# -*- coding: utf-8 -*-
"""Unit tests cho tính năng Đa Tab (Multi-Tab) trong MigoHeadtextWriter."""
import pytest
from unittest.mock import MagicMock, patch
from app.writers.migo_headtext_writer import MigoHeadtextWriter


def test_fill_batch_multi_tab_single_item():
    """Kiểm tra nếu chỉ có 1 DOC thì chạy nhanh trên session đơn."""
    mock_session = MagicMock()
    with patch.object(MigoHeadtextWriter, "fill", return_value=(True, "FILLED", "Lưu thành công")) as mock_fill:
        records, succ = MigoHeadtextWriter.fill_batch_multi_tab(
            items=[{"doc": "4905199093", "ticket_no": "OUT_P1_25422"}],
            year="2026",
            fallback_session=mock_session
        )
        assert succ == 1
        assert len(records) == 1
        assert records[0]["doc"] == "4905199093"
        assert records[0]["ticket_no"] == "OUT_P1_25422"
        assert records[0]["headtext_status"] == "FILLED"
        assert mock_fill.called


def test_fill_batch_multi_tab_parallel_distribution():
    """Kiểm tra phân bổ song song trên 3 Tab khi có nhiều DOC."""
    items = [
        {"doc": f"49051990{i:02d}", "ticket_no": f"OUT_P1_{25420+i}"}
        for i in range(6)
    ]
    mock_sessions = [MagicMock(), MagicMock(), MagicMock()]

    def mock_get_session(conn_idx, tab_idx):
        if 0 <= tab_idx < len(mock_sessions):
            return mock_sessions[tab_idx]
        return None

    def mock_fill(session, doc, year, ticket_no, allow_overwrite):
        return True, "FILLED", f"Lưu thành công {doc}"

    with patch("app.writers.migo_headtext_writer.ensure_sap_tabs", return_value=3), \
         patch("app.writers.migo_headtext_writer.get_safe_sap_session", side_effect=mock_get_session), \
         patch.object(MigoHeadtextWriter, "fill", side_effect=mock_fill):

        records, succ = MigoHeadtextWriter.fill_batch_multi_tab(
            items=items,
            year="2026",
            prompt_raw="SAP điền số phiếu 3 tab",
            connection_index=0
        )

        assert succ == 6
        assert len(records) == 6
        # Đảm bảo đúng thứ tự gốc
        for i, r in enumerate(records):
            assert r["doc"] == items[i]["doc"]
            assert r["ticket_no"] == items[i]["ticket_no"]
            assert r["overall_status"] == "DONE"
            assert r["tab"] in (1, 2, 3)


def test_fill_batch_multi_tab_max_prompt():
    """Kiểm tra nhận diện lệnh 'max' kích hoạt tối đa tab khả dụng."""
    items = [{"doc": f"49051990{i:02d}", "ticket_no": f"OUT_P1_{25420+i}"} for i in range(10)]
    mock_sessions = [MagicMock() for _ in range(4)]

    def mock_get_session(conn_idx, tab_idx):
        if 0 <= tab_idx < len(mock_sessions):
            return mock_sessions[tab_idx]
        return None

    with patch("app.writers.migo_headtext_writer.ensure_sap_tabs", return_value=4) as mock_ensure, \
         patch("app.writers.migo_headtext_writer.get_safe_sap_session", side_effect=mock_get_session), \
         patch.object(MigoHeadtextWriter, "fill", return_value=(True, "FILLED", "OK")):

        records, succ = MigoHeadtextWriter.fill_batch_multi_tab(
            items=items,
            year="2026",
            prompt_raw="SAP điền số phiếu max",
            connection_index=0
        )

        # target_tabs phải là 6 khi gặp 'max'
        assert mock_ensure.call_args[0][0] == 6
        assert succ == 10
        assert len(records) == 10
