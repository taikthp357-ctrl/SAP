# -*- coding: utf-8 -*-
"""Unit tests cho tính năng đọc file .xlsb và xử lý mật khẩu file Order NVL (v1.0.93)."""

import os
import json
import pytest
import openpyxl
from pathlib import Path
from unittest.mock import MagicMock, patch

from app.email_parser import SmartCommandParser
from app.nvl_tracking import NvlTrackingService, NvlFilePasswordRequiredError
from app.handlers.nvl_handler import NvlHandler
from app.core.context import RequestContext, CommandResult


@pytest.fixture
def sample_xlsx_tracking(tmp_path):
    """Tạo file Excel .xlsx giả lập để test đọc dữ liệu."""
    file_path = tmp_path / "261028_THEO DOI ODER NVL_P_S6.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ODER NVl"

    headers = [
        "NCC", "KV", "Packing", "Mã NVL", "SL oder", "Trạm", "Plant", "Sloc Kho", "Sloc Line", "Model",
        "Ca Nhập", "Ca chạy", "OB", "OB1", "SP_1", "SL_1", "OB2", "SP_2", "SL_2", "OB_3",
        "SP_3", "SL_3", "OB_4", "SP_4", "SL_4", "Total cấp", "Gap", "Tình Trạng cấp đơn", "TT nhặt Hàng", "KHSX",
        "NOTE", "Ngày order", "Giờ order", "Người order", "Ca", "SL San Ca", "Ghi Chú San Ca"
    ]
    ws.append(headers)
    ws.append([
        "VINFAST", "KV16", 10, "BIW32022789AA", 380, "ZONE 2", 3093, 1005, 3001, "LIMO",
        "S6", "S1/S2", "1012804739", "", "20-562", 334, "", "", 0, "", "", 0, "", "", 0,
        334, -46, "Partial finished", "Cấp thiếu", "2026-10-28", "NOTE TEST", "17/10/2026", "08:00", "Việt", "S6", 0, ""
    ])
    wb.save(str(file_path))
    return file_path


def test_password_registration():
    """Kiểm tra cơ chế lưu mật khẩu runtime và danh sách mật khẩu đã biết."""
    initial_pwds = NvlTrackingService.get_known_passwords()
    assert "" in initial_pwds
    assert "1" in initial_pwds

    NvlTrackingService.register_password("TestSecretPass2026")
    updated_pwds = NvlTrackingService.get_known_passwords()
    assert "TestSecretPass2026" in updated_pwds
    assert updated_pwds[0] == "TestSecretPass2026"


def test_read_tracking_sheet_rows_xlsx(sample_xlsx_tracking):
    """Kiểm tra đọc dữ liệu trực tiếp từ file .xlsx bằng helper dùng chung."""
    rows = NvlTrackingService._read_tracking_sheet_rows(sample_xlsx_tracking)
    assert len(rows) >= 2
    # Row 0 là header
    assert rows[0][3] == "Mã NVL"
    # Row 1 là dữ liệu
    assert rows[1][3] == "BIW32022789AA"
    assert rows[1][9] == "LIMO"
    assert rows[1][4] == 380


def test_trace_material_allocation_success(sample_xlsx_tracking):
    """Kiểm tra tra cứu thành công mã vật tư và phân tách gap."""
    res = NvlTrackingService.trace_material_allocation("32022789", tracking_file=sample_xlsx_tracking)
    assert res["status"] == "OK"
    assert len(res["records"]) == 1
    rec = res["records"][0]
    assert rec["material"] == "BIW32022789AA"
    assert rec["model"] == "LIMO"
    assert rec["order_qty"] == 380.0
    assert rec["total_cap"] == 334.0
    assert rec["gap"] == -46.0
    assert "20-562" in rec["sp_display"]


def test_encrypted_file_password_required_handling(tmp_path):
    """Kiểm tra khi file bị mã hóa và không có mật khẩu đúng thì ném đúng lỗi NvlFilePasswordRequiredError."""
    fake_enc_file = tmp_path / "261028_THEO DOI ODER NVL_P_S7.xlsb"
    fake_enc_file.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1\x00\x00\x00\x00\x00\x00\x00\x00" + b"\x00" * 100)

    # Đọc với danh sách mật khẩu sai
    with pytest.raises(NvlFilePasswordRequiredError) as exc_info:
        NvlTrackingService._read_tracking_sheet_rows(fake_enc_file, passwords=["wrong_pass_xyz"])

    assert fake_enc_file.name in str(exc_info.value)

    # Khi gọi trace_material_allocation thì không crash mà trả về status PASSWORD_REQUIRED
    res = NvlTrackingService.trace_material_allocation("BIW32022789", tracking_file=fake_enc_file)
    assert res["status"] == "PASSWORD_REQUIRED"
    assert "mật khẩu" in res["summary"].lower()


def test_email_parser_set_nvl_password():
    """Kiểm tra nhận diện câu lệnh cài mật khẩu NVL trên Teams."""
    cmd1 = "mật khẩu nvl: Vinfast@2026"
    parsed1 = SmartCommandParser.parse(cmd1)
    assert parsed1["command"] == "SET_NVL_PASSWORD"
    assert parsed1.get("nvlPassword") == "Vinfast@2026"

    cmd2 = "pass nvl: 123456"
    parsed2 = SmartCommandParser.parse(cmd2)
    assert parsed2["command"] == "SET_NVL_PASSWORD"
    assert parsed2.get("nvlPassword") == "123456"

    cmd3 = "cài mật khẩu order: KhoP@123"
    parsed3 = SmartCommandParser.parse(cmd3)
    assert parsed3["command"] == "SET_NVL_PASSWORD"
    assert parsed3.get("nvlPassword") == "KhoP@123"


def test_nvl_handler_set_password():
    """Kiểm tra NvlHandler xử lý lệnh SET_NVL_PASSWORD và lưu vào runtime."""
    ctx = RequestContext(
        request_id="req_test_pwd",
        command="SET_NVL_PASSWORD",
        sender_email="supervisor@vinfast.vn",
        meta={"nvlPassword": "Vinfast@999", "rawText": "mật khẩu nvl: Vinfast@999"}
    )
    res = NvlHandler().handle(ctx)
    assert res.status == "SUCCEEDED"
    assert "Vinfast@999" in res.summary
    assert "Vinfast@999" in NvlTrackingService.get_known_passwords()
