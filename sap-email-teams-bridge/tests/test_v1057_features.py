# -*- coding: utf-8 -*-
"""Unit tests for v1.0.57: Wildcard matching and ZFIR266 / _return_to_main_menu fix."""

import pytest
from unittest.mock import MagicMock
from app.email_parser import EmailCommandParser
from app.material_resolver import MaterialResolver
from app.sap_engine import SapEngine


def test_email_parser_wildcard_patterns():
    # 1. Natural language with wildcard BIW*6100*
    r1 = EmailCommandParser.parse("check tồn kho BIW*6100*")
    assert r1 is not None
    assert r1.get("materials") == ["BIW*6100*"]

    # 2. Natural language with *6100* and typo plan 4007
    r2 = EmailCommandParser.parse("SAP tra tồn *6100* plan 4007")
    assert r2 is not None
    assert r2.get("materials") == ["*6100*"]
    assert r2.get("plant") == "4007"

    # 3. JSON payload with wildcard
    body_json = """SAP_REQUEST_BEGIN
{"command": "MB52", "materials": ["BIW*6100*", "*6100*"], "plant": "3093"}
SAP_REQUEST_END"""
    r3 = EmailCommandParser.parse(body_json)
    assert r3 is not None
    assert "BIW*6100*" in r3.get("materials", [])
    assert "*6100*" in r3.get("materials", [])


def test_material_resolver_wildcard_expansion():
    resolver = MaterialResolver()
    resolver.load_master_data()

    # 1. BIW*6100* should resolve to the 5 known codes in DATA_MASTER
    codes, note = resolver.resolve("BIW*6100*")
    assert len(codes) == 5
    assert "BIW32016100AA" in codes
    assert "BIW32026100AB" in codes
    assert "BIW75126100AB" in codes
    assert "BIW75116100AA" in codes
    assert "BIW32216100AA" in codes
    assert "Khớp 5 mã theo mẫu wildcard" in note

    # 2. *6100* should also match all 5 codes
    codes2, note2 = resolver.resolve("*6100*")
    assert len(codes2) == 5
    assert "Khớp 5 mã theo mẫu wildcard" in note2

    # 3. Wildcard for unknown pattern
    codes3, note3 = resolver.resolve("UNKNOWN*999*")
    assert codes3 == ["UNKNOWN*999*"]
    assert "Mẫu wildcard trực tiếp" in note3


def test_sap_engine_return_to_main_menu_and_matches_material():
    engine = SapEngine()
    # Ensure both methods exist and are callable
    assert hasattr(engine, "_return_to_main_menu")
    assert hasattr(engine, "_reset_to_session_manager")

    # Mock _reset_to_session_manager and test _return_to_main_menu calls it
    engine._reset_to_session_manager = MagicMock()
    engine._return_to_main_menu()
    engine._reset_to_session_manager.assert_called_once()

    # Test wildcard material matching in _matches_material
    assert SapEngine._matches_material("BIW32016100AA", ["BIW*6100*"]) is True
    assert SapEngine._matches_material("BIW32016100AA", ["*6100*"]) is True
    assert SapEngine._matches_material("BIW32022638AB", ["BIW*6100*"]) is False


def test_sap_engine_zfir266_alv_robust_reading():
    engine = SapEngine()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "ZFIR266"
    engine.ensure_session = MagicMock(return_value=mock_session)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._reset_to_session_manager = MagicMock()

    mock_bukrs = MagicMock()
    mock_gjahr = MagicMock()
    mock_doc = MagicMock()
    mock_max = MagicMock()
    mock_wnd0 = MagicMock()

    # Simulate ALV Grid with technical column named 'ZFILE_NAME' containing VAT_6489.PDF
    mock_grid = MagicMock()
    mock_grid.RowCount = 1
    mock_col_order = MagicMock()
    mock_col_order.Count = 3
    mock_col_order.ElementAt.side_effect = lambda idx: ["COCODE", "DOCNO", "ZFILE_NAME"][idx]
    mock_grid.ColumnOrder = mock_col_order

    def get_cell(row, col):
        if col == "ZFILE_NAME":
            return "VAT_6489.PDF"
        return ""

    mock_grid.GetCellValue.side_effect = get_cell

    def find_mock(ctl_id):
        if "ctxtP_BUKRS" in ctl_id: return mock_bukrs
        if "txtP_GJAHR" in ctl_id: return mock_gjahr
        if "ctxtS_MBELN-LOW" in ctl_id: return mock_doc
        if "txtP_MAX" in ctl_id: return mock_max
        if ctl_id == "wnd[0]": return mock_wnd0
        if "cntlGRID1" in ctl_id: return mock_grid
        return None

    engine._find_by_id_safe = MagicMock(side_effect=find_mock)

    records = engine.check_attachment_status(["5007145691"], year="2026")
    assert len(records) == 1
    rec = records[0]
    assert rec["doc"] == "5007145691"
    assert rec["bbbg_status"] == "CHƯA"
    assert rec["vat_status"] == "CÓ"
    assert rec["vat_file"] == "VAT_6489.PDF"
    assert rec["status_code"] == "MISSING_BBBG"
    assert "Đã có VAT, thiếu BBBG" in rec["overall_status"]
    engine._reset_to_session_manager.assert_called()


def test_email_parser_check_att_no_folder_path():
    text = """SAP check att
5007145691 | Thành Đạt IN-6489
5007145669 | Thành Đạt IN-6489"""
    r = EmailCommandParser.parse(text)
    assert r is not None
    assert r["command"] == "CHECK_ATT"
    # folderPath must NOT be populated for check att
    assert r.get("folderPath") == ""
    assert len(r.get("attachmentDocs", [])) == 2


def test_sap_engine_zfir266_batch_multi_selection():
    engine = SapEngine()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "ZFIR266"
    engine.ensure_session = MagicMock(return_value=mock_session)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._fill_multiple_selection = MagicMock()
    engine._reset_to_session_manager = MagicMock()

    mock_multi_btn = MagicMock()
    mock_bukrs = MagicMock()
    mock_gjahr = MagicMock()
    mock_max = MagicMock()
    mock_wnd0 = MagicMock()

    # ALV Grid with 2 rows: one for 5007145691 (VAT), one for 5007145669 (BBBG)
    mock_grid = MagicMock()
    mock_grid.RowCount = 2
    mock_col_order = MagicMock()
    mock_col_order.Count = 3
    mock_col_order.ElementAt.side_effect = lambda idx: ["DOCUMENTNO", "FILE_NAME", "AMOUNT"][idx]
    mock_grid.ColumnOrder = mock_col_order

    def get_cell(row, col):
        if col == "DOCUMENTNO":
            return "5007145691" if row == 0 else "5007145669"
        if col == "FILE_NAME":
            return "VAT_6489.PDF" if row == 0 else "5007145669_BBBG.pdf"
        return ""

    mock_grid.GetCellValue.side_effect = get_cell

    def find_mock(ctl_id):
        if "btn%_S_MBELN_%_APP_%-VALU_PUSH" in ctl_id: return mock_multi_btn
        if "ctxtP_BUKRS" in ctl_id: return mock_bukrs
        if "txtP_GJAHR" in ctl_id: return mock_gjahr
        if "txtP_MAX" in ctl_id: return mock_max
        if ctl_id == "wnd[0]": return mock_wnd0
        if "cntlGRID1" in ctl_id: return mock_grid
        return None

    engine._find_by_id_safe = MagicMock(side_effect=find_mock)

    docs = ["5007145691", "5007145669", "5007145670"]
    records = engine.check_attachment_status(docs, year="2026")
    assert len(records) == 3

    # Multi selection should be called with all docs
    engine._fill_multiple_selection.assert_called_once_with(
        "wnd[0]/usr/btn%_S_MBELN_%_APP_%-VALU_PUSH", docs
    )

    rec1 = next(r for r in records if r["doc"] == "5007145691")
    assert rec1["vat_status"] == "CÓ"
    assert rec1["bbbg_status"] == "CHƯA"

    rec2 = next(r for r in records if r["doc"] == "5007145669")
    assert rec2["bbbg_status"] == "CÓ"
    assert rec2["vat_status"] == "CHƯA"

    rec3 = next(r for r in records if r["doc"] == "5007145670")
    assert rec3["bbbg_status"] == "CHƯA"
    assert rec3["vat_status"] == "CHƯA"
    assert rec3["status_code"] == "NO_ATT"


def test_attach_file_verification_without_extension():
    """Kiểm tra đính kèm file VAT 5319.pdf khi SAP trả về tiêu đề không có đuôi 'VAT 5319'."""
    import tempfile
    from pathlib import Path
    from app.sap_attachment import MigoAttachmentRunner

    runner = MigoAttachmentRunner()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "MIGO"
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "TESTUSER"
    mock_session.Id = "ses[0]"
    mock_session.Busy = False
    runner.session_identity = {"system": "PRD", "client": "100", "user": "TESTUSER", "session_id": "ses[0]"}
    runner._active_document = "5006892548"
    runner._active_year = "2026"

    mock_sbar = MagicMock()
    mock_sbar.Text = "Attachment successfully created"
    mock_sbar.MessageType = "S"

    mock_btn = MagicMock()
    mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_session.FindByName.return_value = mock_btn
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.return_value = mock_btn
    wnd1_active = [True]
    mock_btn.Press.side_effect = lambda: wnd1_active.__setitem__(0, False)
    def find_by_id(path, *args):
        p = str(path)
        if "sbar" in p:
            return mock_sbar
        if "btn" in p:
            return mock_btn
        if p == "wnd[1]" or p.startswith("wnd[1]"):
            if wnd1_active[0]:
                return mock_wnd1
            raise Exception("not found")
        return MagicMock()
    mock_session.findById.side_effect = find_by_id

    # Mock SAP GOS Attachment list returning Title without .pdf extension (e.g. 'VAT 5319')
    runner.read_sap_attachment_titles = MagicMock(return_value=["VAT 5319", "BBBG 5319"])

    with tempfile.NamedTemporaryFile(prefix="VAT 5319", suffix=".pdf", delete=False) as tf:
        tf.write(b"dummy vat content")
        vat_file = Path(tf.name)

    try:
        ok, msg = runner.attach_file_to_current_migo(mock_session, vat_file)
        assert ok is True
        assert "DONE" in msg
    finally:
        vat_file.unlink(missing_ok=True)
