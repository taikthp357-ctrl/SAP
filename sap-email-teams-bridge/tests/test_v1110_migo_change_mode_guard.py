# -*- coding: utf-8 -*-
"""Unit tests for MIGO Change vs Display mode enforcement (v1.1.10)."""
from unittest.mock import MagicMock, patch
import pytest

from app.sap_attachment import MigoAttachmentRunner, ensure_migo_mode


class MockComboBox:
    def __init__(self, key="A04", text="Display"):
        self.Key = key
        self.Text = text


def owned_runner(session, engine=None):
    session.Info.SystemName = "TEST"
    session.Info.Client = "100"
    session.Info.User = "TESTUSER"
    session.Info.Transaction = "MIGO"
    session.Id = "/app/con[0]/ses[1]"
    session.Busy = False
    runner = MigoAttachmentRunner(engine, {"system": "TEST", "client": "100", "user": "TESTUSER", "session_id": session.Id})
    runner._active_document = ("4905195632", "2026")
    return runner


def test_ensure_migo_mode_switches_display_to_change():
    """Kiểm tra ensure_migo_mode tự động đổi từ A04 (Display) sang A12 (Change)."""
    mock_session = MagicMock()
    act_cb = MockComboBox("A04", "Display")
    ref_cb = MockComboBox("R01", "Purchase Order")

    def find_by_name(name, ctype):
        if name == "GODYNPRO-ACTION":
            return act_cb
        elif name == "GODYNPRO-REFDOC":
            return ref_cb
        return MagicMock()

    mock_w0 = MagicMock()
    mock_w0.findByName.side_effect = find_by_name
    mock_session.findById.return_value = mock_w0

    ok = ensure_migo_mode(mock_session, target_action="A12", target_refdoc="R02")

    assert ok is True
    assert act_cb.Key == "A12"
    assert ref_cb.Key == "R02"


def test_ensure_migo_mode_text_fallback():
    """Kiểm tra ensure_migo_mode nhận diện được text 'DISPLAY' khi Key rỗng."""
    mock_session = MagicMock()
    act_cb = MockComboBox("", "Display")
    ref_cb = MockComboBox("", "Material Document")

    def find_by_name(name, ctype):
        if name == "GODYNPRO-ACTION":
            return act_cb
        elif name == "GODYNPRO-REFDOC":
            return ref_cb
        return MagicMock()

    mock_w0 = MagicMock()
    mock_w0.findByName.side_effect = find_by_name
    mock_session.findById.return_value = mock_w0

    ok = ensure_migo_mode(mock_session, target_action="A12", target_refdoc="R02")

    assert ok is True
    assert act_cb.Key == "A12"


def test_handle_migo_headtext_blocked_in_display_mode():
    """Kiểm tra handle_migo_headtext chặn ngay nếu cửa sổ MIGO đang là Display."""
    mock_session = MagicMock()
    runner = owned_runner(mock_session)

    title_bar = MagicMock(Text="Display Material Document - Phạm Tiến Tài")
    mock_session.findById.side_effect = lambda cid, *a: title_bar if "titl" in cid else MagicMock()

    code, msg = runner.handle_migo_headtext(mock_session, "OUT_P1_123")

    assert code == "HEADTEXT_BLOCKED"
    assert "Display (A04)" in msg
