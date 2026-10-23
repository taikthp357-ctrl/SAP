# -*- coding: utf-8 -*-
"""Unit tests cho number_utils: parse_sap_float và format_qty_display."""

import sys
from pathlib import Path
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.number_utils import parse_sap_float, format_qty_display


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


def test_format_qty_display():
    """Kiểm tra định dạng chuỗi hiển thị số lượng."""
    assert format_qty_display(1000) == "1,000"
    assert format_qty_display("1,000") == "1,000"
    assert format_qty_display(1000.5) == "1,000.5"
    assert format_qty_display("1000.500") == "1,000.5"
    assert format_qty_display(0) == "0"
    assert format_qty_display("0.000") == "0"
