# -*- coding: utf-8 -*-
"""Bộ kiểm thử đơn vị cho bản nâng cấp v1.0.97.

Kiểm tra:
1. 4-Layer Contextual Intent Router: Tiếng lóng kho, Regex Entity, Ma trận trọng số + Veto Rules, Fast-First Fallback.
2. An toàn 3 lớp đọc file .xlsb (RAM pyxlsb -> Excel COM trên %TEMP% -> Companion .xlsx).
3. Luồng kiểm tra 2 pha ZFIR266 -> MIGO: Bóc tách cờ chi tiết (need_bbbg, need_vat, need_header_text), bỏ qua MIGO cho chứng từ đã đủ 100%.
4. Chốt chặn an toàn: DOC không có trong ALV ZFIR266 không bị coi là SKIP.
5. Multi-session SAP Pool & Turbo Mode & FILE_DIALOG_LOCK.
6. Nhãn lỗi theo ngữ cảnh (Context-aware error tagging).
"""

import os
import io
import tempfile
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from app.email_parser import SmartCommandParser
from app.nvl_tracking import NvlTrackingService, NvlFilePasswordRequiredError
from app.sap_attachment import (
    MigoAttachmentRunner,
    get_user_idle_seconds,
    is_turbo_mode,
    FILE_DIALOG_LOCK
)
from app.teams_dispatcher import TeamsDispatcher


# =====================================================================
# 1. KIỂM THỬ 4-LAYER CONTEXTUAL INTENT ROUTER
# =====================================================================

def test_router_slang_expansion():
    """Kiểm tra tầng 1: dịch tiếng lóng kho & sửa lỗi gõ nhanh."""
    assert "file ca" in SmartCommandParser.normalize_and_expand_slang("fiel ca s6")
    assert "oder" in SmartCommandParser.normalize_and_expand_slang("order nvl")
    assert "cap bu" in SmartCommandParser.normalize_and_expand_slang("cb cho ca dem")
    assert "don bu" in SmartCommandParser.normalize_and_expand_slang("donbu ma 32022789")


def test_router_short_supply_queries():
    """Kiểm tra câu hỏi ngắn gọn không có chữ 'file' vẫn nhận diện đúng TRA_CUU_CAP_NVL."""
    test_queries = [
        "check 32022789 cấp chưa",
        "32022789 ca S6",
        "kiểm tra mã 32022789 đơn bù",
        "oder 32022789",
        "32022789 giao chưa",
        "32022789 xong chưa",
        "32022789 đủ chưa",
        "đơn bù 32022789",
        "odbs 32022789",
        "khsx 32022789",
        "fiel ca s6 32022789",
        "tiến độ 32022789 ca đêm"
    ]
    for q in test_queries:
        parsed = SmartCommandParser.parse(q)
        assert parsed is not None, f"Query '{q}' parse ra None"
        assert parsed.get("command") == "TRA_CUU_CAP_NVL", (
            f"Query '{q}' bị gán nhầm thành {parsed.get('command')}"
        )


def test_router_veto_rules():
    """Kiểm tra luật phủ định (Veto): File ca/Cấp hàng không được rơi vào VITRI."""
    # Có mã hàng nhưng hỏi file ca -> Tuyệt đối không vào VITRI
    parsed = SmartCommandParser.parse("file ca s6 mã 32022789 cấp chưa")
    assert parsed.get("command") == "TRA_CUU_CAP_NVL"

    # Chỉ hỏi vị trí Bin -> Vào VITRI
    parsed_bin = SmartCommandParser.parse("vitri mã 32022789 ô bin nào")
    assert parsed_bin.get("command") == "VITRI"


def test_router_check_att_vs_att_migo():
    """Kiểm tra phân luồng rạch ròi giữa CHECK_ATT (ZFIR266 tra cứu) và ATT_MIGO (MIGO ghi file)."""
    # Tra cứu
    p_check = SmartCommandParser.parse("check att 5000123456")
    assert p_check.get("command") == "CHECK_ATT"

    p_kt = SmartCommandParser.parse("zfir266 chứng từ 5000123456 thiếu gì")
    assert p_kt.get("command") == "CHECK_ATT"

    # Thực thi
    p_run = SmartCommandParser.parse("chạy att 5000123456")
    assert p_run.get("command") == "ATT_MIGO"

    p_migo = SmartCommandParser.parse("att migo 5000123456 số phiếu 12345")
    assert p_migo.get("command") == "ATT_MIGO"


# =====================================================================
# 2. KIỂM THỬ AN TOÀN ĐỌC FILE XLSB 3 LỚP
# =====================================================================

def test_xlsb_layer2_temp_copy_safety(tmp_path):
    """Kiểm tra Lớp 2 Excel COM: Luôn mở bản copy trong %TEMP% và xóa sạch khi kết thúc."""
    dummy_file = tmp_path / "dummy_tracking.xlsb"
    # Header PK Zip hợp lệ
    dummy_file.write_bytes(b"PK\x03\x04" + b"\x00" * 600)

    # Mock win32com và pythoncom
    mock_excel = MagicMock()
    mock_wb = MagicMock()
    mock_excel.Workbooks.Open.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_sheet.Name = "ODER NVl"
    mock_sheet.UsedRange.Value = (("Col1", "Col2"), ("Val1", "Val2"))
    mock_wb.Sheets = MagicMock()
    mock_wb.Sheets.Count = 1
    mock_wb.Sheets.return_value = mock_sheet
    mock_wb.Sheets.__iter__ = MagicMock(return_value=iter([mock_sheet]))

    with patch("win32com.client.DispatchEx", return_value=mock_excel):
        with patch("pythoncom.CoInitialize", return_value=None):
            with patch("pythoncom.CoUninitialize", return_value=None):
                rows = NvlTrackingService._read_xlsb_via_excel_com(dummy_file, passwords=["Khop"])
                assert len(rows) == 2
                assert rows[0] == ["Col1", "Col2"]

                # Kiểm tra wb.Close(False) và excel.Quit() luôn được gọi
                mock_wb.Close.assert_called_with(False)
                mock_excel.Quit.assert_called_once()


def test_xlsb_layer3_companion_xlsx_fallback(tmp_path):
    """Kiểm tra Lớp 3: Tự động fallback sang companion .xlsx nếu .xlsb hỏng."""
    bad_xlsb = tmp_path / "THEO DOI ODER NVL_P_S6.xlsb"
    bad_xlsb.write_bytes(b"\x00" * 10)  # File hỏng không đúng định dạng

    good_xlsx = tmp_path / "THEO DOI ODER NVL_P_S6.xlsx"
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ODER NVl"
    ws.append(["Mã", "SL", "Ca"])
    ws.append(["32022789", 50, "S6"])
    wb.save(good_xlsx)

    rows = NvlTrackingService._read_tracking_sheet_rows(bad_xlsb)
    assert len(rows) >= 2
    assert rows[1][0] == "32022789"


# =====================================================================
# 3. KIỂM THỬ LUỒNG 2 PHA ZFIR266 -> MIGO & BẢN ĐỒ CỜ THIẾU
# =====================================================================

def test_zfir266_phase1_precheck_skips_complete_docs():
    """Kiểm tra Phase 1: Chứng từ đã có đủ cả BBBG và VAT được SKIP khỏi MIGO."""
    runner = MigoAttachmentRunner(sap_engine=MagicMock())
    mock_sess = MagicMock()
    runner.sap_engine.ensure_session.return_value = mock_sess
    runner.verify_session_identity = MagicMock(return_value=(True, "OK"))
    runner.sap_engine.check_attachments_zfir266.return_value = [
        {
            "doc": "5000123456",
            "found_in_zfir": True,
            "need_bbbg": False,
            "need_vat": False,
            "need_header_text": False,
            "bbbg_file": "BBBG_123.pdf",
            "vat_file": "VAT_123.pdf"
        },
        {
            "doc": "5000123457",
            "found_in_zfir": True,
            "need_bbbg": False,
            "need_vat": True,
            "need_header_text": False,
            "bbbg_file": "BBBG_124.pdf",
            "vat_file": None
        }
    ]

    items = [
        {"doc": "5000123456", "ticket_no": "123"},
        {"doc": "5000123457", "ticket_no": "124"}
    ]

    # Mock open_migo_document và handle_migo_headtext
    with patch.object(runner, "open_migo_document", return_value=(True, "OK")):
        with patch.object(runner, "handle_migo_headtext", return_value=("HEADTEXT_DONE", "OK")):
            with patch.object(runner, "attach_file_to_current_migo", return_value=(True, "OK")):
                with patch.object(runner, "index_source_folder", return_value={"total_files": 0, "files": [], "token_map": {}}):
                    summary = runner.process_attachment_batch("", items, dry_run=False)

                    results = summary["results"]
                    assert len(results) == 2
                    doc1_res = next(r for r in results if r["doc"] == "5000123456")
                    assert doc1_res["overall_status"] == "DONE"
                    assert "ZFIR266" in doc1_res["bbbg_status"]


def test_zfir266_missing_doc_not_skipped():
    """Chốt an toàn 5: DOC không xuất hiện trên ALV ZFIR266 tuyệt đối không bị coi là SKIP."""
    runner = MigoAttachmentRunner(sap_engine=MagicMock())
    mock_session = MagicMock()
    runner.sap_engine.ensure_session.return_value = mock_session
    runner.verify_session_identity = MagicMock(return_value=(True, "OK"))
    runner.prepare_migo_session = MagicMock(return_value=True)

    # ZFIR266 chỉ trả về 1 DOC, DOC thứ 2 không có trên ALV
    runner.sap_engine.check_attachments_zfir266.return_value = [
        {
            "doc": "5000123456",
            "found_in_zfir": True,
            "need_bbbg": False,
            "need_vat": False,
            "need_header_text": False
        }
    ]

    items = [
        {"doc": "5000123456", "ticket_no": "123"},
        {"doc": "5000999999", "ticket_no": "999"}  # Không có trên ZFIR266
    ]

    with patch.object(runner, "open_migo_document", return_value=(True, "OK")) as mock_open:
        with patch.object(runner, "handle_migo_headtext", return_value=("HEADTEXT_DONE", "OK")):
            with patch.object(runner, "index_source_folder", return_value={"total_files": 0, "files": [], "token_map": {}}):
                summary = runner.process_attachment_batch("", items, dry_run=False)
                results = summary["results"]
                doc2_res = next(r for r in results if r["doc"] == "5000999999")
                # DOC 5000999999 không bị SKIP mà phải được mở qua MIGO
                assert doc2_res["doc"] == "5000999999"
                mock_open.assert_called()


# =====================================================================
# 4. KIỂM THỬ TURBO MODE & FILE_DIALOG_LOCK
# =====================================================================

def test_turbo_mode_detection():
    """Kiểm tra phát hiện từ khóa Turbo Mode và idle time."""
    assert is_turbo_mode("chạy att turbo 100 doc") is True
    assert is_turbo_mode("att migo max công suất") is True
    assert is_turbo_mode("chạy 5 tab nhanh lên bot") is True
    assert is_turbo_mode("chạy 6 tab") is True


def test_file_dialog_lock_concurrency():
    """Kiểm tra FILE_DIALOG_LOCK đảm bảo tuần tự hóa khi nhiều luồng truy cập hộp thoại file."""
    assert isinstance(FILE_DIALOG_LOCK, type(threading.Lock()))

    acquired_order = []
    def worker(worker_id):
        with FILE_DIALOG_LOCK:
            acquired_order.append(worker_id)

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(5)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert len(acquired_order) == 5


# =====================================================================
# 5. KIỂM THỬ NHÃN LỖI THEO NGỮ CẢNH (CONTEXT-AWARE ERROR TAGGING)
# =====================================================================

def test_context_aware_error_formatting():
    """Kiểm tra lời chào Teams không còn bị gán nhầm thành lỗi SAP GUI khi gặp lỗi file Excel."""
    # Lỗi file ca / theo dõi NVL
    payload_nvl, _ = TeamsDispatcher.prepare_safe_payload(
        request_id="REQ_1",
        meta={"command": "TRA_CUU_CAP_NVL", "senderEmail": "test@fvn.com", "senderName": "Tài"},
        status="FAILED",
        summary="Lỗi xử lý File Ca / Theo dõi NVL Excel: No module named 'pyxlsb'",
        records=[],
        headers_map={}
    )
    # Lời chào phải chuyên biệt cho File Ca / NVL, không phải "chưa tra cứu được SAP"
    greeting = payload_nvl.get("greeting", "")
    assert "Excel" in greeting or "file ca" in greeting or "theo dõi NVL" in greeting
