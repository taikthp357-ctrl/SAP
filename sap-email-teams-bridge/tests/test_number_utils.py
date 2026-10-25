# -*- coding: utf-8 -*-
"""Unit tests cho number_utils: parse_sap_float và format_qty_display."""

import sys
from pathlib import Path
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.number_utils import parse_sap_float, format_qty_display, parse_sap_datetime_sort_key


def test_parse_sap_float_us_format():
    """Kiểm tra định dạng US/SAP chuẩn: dấu phẩy phân cách hàng nghìn, dấu chấm thập phân."""
    assert parse_sap_float("1,000") == 1000.0
    assert parse_sap_float("1,000.50") == 1000.5
    assert parse_sap_float("2,500,000.75") == 2500000.75
    assert parse_sap_float("500") == 500.0
    assert parse_sap_float("0.5") == 0.5


def test_parse_sap_float_vn_de_format():
    """Kiểm tra định dạng Đức / Việt Nam: dấu chấm phân cách hàng nghìn, dấu phẩy thập phân."""
    assert parse_sap_float("1.000,50") == 1000.5
    assert parse_sap_float("12,5") == 12.5
    assert parse_sap_float("12,50") == 12.5
    assert parse_sap_float("2.500.000,00") == 2500000.0


def test_parse_sap_float_trailing_negative():
    """Kiểm tra định dạng số âm có dấu '-' ở đuôi từ SAP."""
    assert parse_sap_float("100.00-") == -100.0
    assert parse_sap_float("1,000-") == -1000.0
    assert parse_sap_float("-50.5") == -50.5


def test_parse_sap_float_spaces_and_blanks():
    """Kiểm tra khoảng trắng, None, rỗng."""
    assert parse_sap_float(" 1 000 ") == 1000.0
    assert parse_sap_float("") == 0.0
    assert parse_sap_float(None) == 0.0
    assert parse_sap_float("abc", default=99.0) == 99.0
    assert parse_sap_float(123) == 123.0
    assert parse_sap_float(45.67) == 45.67


def test_parse_sap_float_corrupted_and_mixed_strings():
    """Kiểm tra chuỗi chứa ký tự chữ và số không gây lỗi ValueError."""
    # 'abc.123' -> trích xuất phần số an toàn
    assert parse_sap_float("abc.123") == 0.123
    assert parse_sap_float("QTY: 1,500 PCS") == 1500.0
    assert parse_sap_float("INVALID_VAL", default=0.0) == 0.0
    assert parse_sap_float({"not": "a number"}, default=0.0) == 0.0


def test_parse_sap_float_explicit_formats():
    """Kiểm tra chỉ định rõ ràng cấu hình số US vs EU/VN."""
    # Chuẩn US: '1.000' là 1.0 (với 3 số thập phân), '1,000' là 1000.0
    assert parse_sap_float("1.000", fmt="US") == 1.0
    assert parse_sap_float("1,000", fmt="US") == 1000.0

    # Chuẩn EU/VN: '1.000' là 1000.0, '1,234' là 1.234
    assert parse_sap_float("1.000", fmt="EU") == 1000.0
    assert parse_sap_float("1,234", fmt="VN") == 1.234
    assert parse_sap_float("1.000,50", fmt="VN") == 1000.5


def test_format_qty_display():
    """Kiểm tra định dạng chuỗi hiển thị số lượng."""
    assert format_qty_display(1000) == "1000"
    assert format_qty_display("1,000") == "1000"
    assert format_qty_display(1000.5) == "1000.5"
    assert format_qty_display("1000.500") == "1000.5"
    assert format_qty_display(0) == "0"
    assert format_qty_display("0.000") == "0"


def test_parse_sap_datetime_sort_key_comprehensive():
    """Kiểm tra R04: Phân giải ngày giờ GR đa định dạng và đẩy ngày thiếu xuống cuối cùng."""
    # 1. Các định dạng ngày phổ biến trong SAP
    k_eu = parse_sap_datetime_sort_key("15.10.2026", "08:30:00")
    assert k_eu == ("20261015", "083000")

    k_slash = parse_sap_datetime_sort_key("15/10/2026", "14:15")
    assert k_slash == ("20261015", "141500")

    k_iso = parse_sap_datetime_sort_key("2026-10-20", "09:00:00")
    assert k_iso == ("20261020", "090000")

    k_raw8 = parse_sap_datetime_sort_key("20261022", "120000")
    assert k_raw8 == ("20261022", "120000")

    # 2. Ngày bị thiếu / N/A / rỗng phải có key lớn nhất để xếp cuối bảng trong sort ascending
    k_empty = parse_sap_datetime_sort_key("", "")
    assert k_empty == ("99999999", "999999")

    k_na = parse_sap_datetime_sort_key("N/A", "-")
    assert k_na == ("99999999", "999999")

    # 3. Kiểm tra sắp xếp danh sách 3 phần tử: 20.10, thiếu ngày, 15.10
    sample = [
        {"batch": "B2", "d": "20.10.2026", "t": "10:00"},
        {"batch": "B_NO_DATE", "d": "", "t": ""},
        {"batch": "B1", "d": "15.10.2026", "t": "08:00"},
    ]
    sample.sort(key=lambda x: parse_sap_datetime_sort_key(x["d"], x["t"]))
    assert sample[0]["batch"] == "B1"          # Cũ nhất -> FIFO #1
    assert sample[1]["batch"] == "B2"          # Nhập sau -> FIFO #2
    assert sample[2]["batch"] == "B_NO_DATE"   # Không có ngày -> Xếp cuối cùng


def test_parse_sap_float_user_production_cases():
    """Kiểm tra các trường hợp thực tế tại xưởng: 14.000, 2.800, 2.050, 450, 16.500, 17.250."""
    # 1. Với cấu hình EU (dấu chấm hàng nghìn)
    assert parse_sap_float("14.000", fmt="EU") == 14000.0
    assert parse_sap_float("2.050", fmt="EU") == 2050.0
    assert parse_sap_float("2.800", fmt="EU") == 2800.0
    assert parse_sap_float("450", fmt="EU") == 450.0
    assert parse_sap_float("16.500", fmt="EU") == 16500.0
    assert parse_sap_float("17.250", fmt="EU") == 17250.0
    assert parse_sap_float("1.019", fmt="EU") == 1019.0

    # 2. Với cấu hình AUTO thông minh
    assert parse_sap_float("2.050") == 2050.0
    assert parse_sap_float("2.800") == 2800.0
    assert parse_sap_float("1.019") == 1019.0
    assert parse_sap_float("17.250") == 17250.0
    assert parse_sap_float("16.500") == 16500.0
    assert parse_sap_float("450") == 450.0
