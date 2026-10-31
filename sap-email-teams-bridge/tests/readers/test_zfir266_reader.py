# -*- coding: utf-8 -*-
"""Unit tests cho Zfir266Reader độc lập."""
import pytest
from unittest.mock import MagicMock, patch
from app.readers.zfir266_reader import Zfir266Reader


def test_zfir266_reader_batch_success():
    """Kiểm tra đọc ALV Grid từ ZFIR266 thành công mà không gọi MIGO."""
    session = MagicMock()
    session.Info.Transaction = "ZFIR266"
    session.Busy = False

    w0 = MagicMock()
    sbar = MagicMock()
    sbar.Text = ""

    grid = MagicMock()
    grid.RowCount = 2
    grid.ColumnOrder = MagicMock()
    grid.ColumnOrder.Count = 2
    grid.ColumnOrder.ElementAt.side_effect = lambda idx: ["MBELN", "FILE_NAME"][idx]

    def get_cell_value(r, col):
        if col == "MBELN":
            return ["5007145690", "5007145691"][r]
        if col == "FILE_NAME":
            return ["5007145690_BBBG.pdf", "5007145691_VAT.pdf"][r]
        return ""

    grid.GetCellValue.side_effect = get_cell_value

    def find_side_effect(id_str, *args):
        if id_str == "wnd[0]":
            return w0
        if id_str == "wnd[0]/sbar":
            return sbar
        if "cntlGRID1" in id_str or "usr/shell" in id_str:
            return grid
        if "btn%_S_MBELN_%_APP_%-VALU_PUSH" in id_str:
            return MagicMock()
        return MagicMock()

    session.findById.side_effect = find_side_effect

    with patch("app.readers.zfir266_reader.open_transaction"), \
         patch("app.readers.zfir266_reader.wait_idle", return_value=True), \
         patch("app.readers.zfir266_reader.dismiss_popups", return_value=""), \
         patch("app.readers.zfir266_reader.fill_multiple_selection"), \
         patch("app.readers.zfir266_reader.reset_to_session_manager"):

        records = Zfir266Reader.read(
            session=session,
            docs=["5007145690", "5007145691"],
            year="2026",
            company_code="2301"
        )

        assert len(records) == 2
        r1 = next(r for r in records if r["doc"] == "5007145690")
        assert r1["bbbg_status"] == "CÓ"
        assert r1["vat_status"] == "CHƯA"

        r2 = next(r for r in records if r["doc"] == "5007145691")
        assert r2["bbbg_status"] == "CHƯA"
        assert r2["vat_status"] == "CÓ"


def test_zfir266_reader_empty_no_migo_call():
    """Kiểm tra khi không có file đính kèm, Zfir266Reader báo CHƯA và KHÔNG mở MIGO Display."""
    session = MagicMock()
    session.Info.Transaction = "ZFIR266"
    session.Busy = False

    w0 = MagicMock()
    sbar = MagicMock()
    sbar.Text = ""

    grid = MagicMock()
    grid.RowCount = 0  # Không có kết quả nào

    def find_side_effect(id_str, *args):
        if id_str == "wnd[0]":
            return w0
        if id_str == "wnd[0]/sbar":
            return sbar
        if "cntlGRID1" in id_str or "usr/shell" in id_str:
            return grid
        return None

    session.findById.side_effect = find_side_effect

    with patch("app.readers.zfir266_reader.open_transaction"), \
         patch("app.readers.zfir266_reader.wait_idle", return_value=True), \
         patch("app.readers.zfir266_reader.dismiss_popups", return_value=""), \
         patch("app.readers.zfir266_reader.fill_multiple_selection"), \
         patch("app.readers.zfir266_reader.reset_to_session_manager") as mock_reset:

        records = Zfir266Reader.read(
            session=session,
            docs=["5007145690"],
            year="2026",
            company_code="2301"
        )

        assert len(records) == 1
        r = records[0]
        assert r["bbbg_status"] == "CHƯA"
        assert r["vat_status"] == "CHƯA"
        assert r["files_found"] == "Không có file"
        # Đảm bảo reset session được gọi để giải phóng ALV
        assert mock_reset.called
