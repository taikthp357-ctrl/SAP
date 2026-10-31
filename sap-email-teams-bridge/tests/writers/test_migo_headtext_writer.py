# -*- coding: utf-8 -*-
"""Unit tests cho MigoHeadtextWriter độc lập."""
import pytest
from unittest.mock import MagicMock, patch
from app.writers.migo_headtext_writer import MigoHeadtextWriter


def test_migo_headtext_fill_success():
    """Kiểm tra quy trình điền số phiếu và lưu thành công."""
    session = MagicMock()
    session.Info.Transaction = "MIGO"
    session.Busy = False

    w0 = MagicMock()
    titl = MagicMock()
    titl.Text = "Change Material Document 5007145690 - Phạm Tiến Tài"
    sbar = MagicMock()
    sbar.Text = "Material document 5007145690 changed"
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
    mock_head_fld.Text = ""

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
            doc="5007145690",
            year="2026",
            ticket_no="6381",
            allow_overwrite=False
        )

        assert success is True
        assert code == "FILLED"
        assert mock_head_fld.Text == "6381"
        assert w0.SendVKey.called


def test_migo_headtext_display_mode_blocked():
    """Kiểm tra nếu màn hình nạp nhầm vào Display (Ảnh 2) thì bị chặn và thoát ra."""
    session = MagicMock()
    session.Info.Transaction = "MIGO"
    session.Busy = False

    w0 = MagicMock()
    titl = MagicMock()
    titl.Text = "Display Material Document 5007145690 - Phạm Tiến Tài"
    sbar = MagicMock()
    sbar.Text = ""
    sbar.MessageType = ""

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

    def smart_find_mock(ses, name, *args, **kwargs):
        if name == "GODYNPRO-MAT_DOC":
            return mock_doc_fld
        return None

    with patch("app.writers.migo_headtext_writer.smart_find", side_effect=smart_find_mock), \
         patch("app.writers.migo_headtext_writer.wait_idle", return_value=True), \
         patch("app.writers.migo_headtext_writer.dismiss_popups", return_value=""):

        success, code, msg = MigoHeadtextWriter.fill(
            session=session,
            doc="5007145690",
            year="2026",
            ticket_no="6381"
        )

        assert success is False
        assert code == "DISPLAY_MODE_BLOCKED"
        assert "Display" in msg
