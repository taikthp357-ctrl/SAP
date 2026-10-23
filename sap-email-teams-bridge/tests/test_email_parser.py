# -*- coding: utf-8 -*-
"""Unit tests for EmailCommandParser."""

import sys
from pathlib import Path
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.email_parser import EmailCommandParser


def test_parse_simple_vitri():
    body = """
    Chào bạn,
    LENH: VITRI | MA: 32013765 | BATCH: B2610
    Cảm ơn nhé!
    --
    Nguyễn Văn A - Phòng Quản lý kho
    """
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["command"] == "VITRI"
    assert res["materials"] == ["32013765"]
    assert res["batch"] == "B2610"
    assert res["plant"] == "3093"


def test_parse_ultra_short_single_code():
    # Người dùng chỉ gửi độc nhất mã 32013765 trong nội dung email
    body = "32013765"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["command"] == "VITRI"
    assert res["materials"] == ["32013765"]


def test_parse_ultra_short_code_and_batch():
    # Người dùng chỉ gõ: 32013765 B2610
    body = "32013765 B2610"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["command"] == "VITRI"
    assert res["materials"] == ["32013765"]
    assert res["batch"] == "B2610"


def test_parse_multi_material_with_spaces_and_commas():
    body = "LENH: TON | MA: 32013765, 32024300; 1002347 | PLANT: 3001"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["command"] == "TON"
    assert res["materials"] == ["32013765", "32024300", "1002347"]
    assert res["plant"] == "3001"


def test_parse_lich_su_with_days():
    body = "LENH: LICHSU | MA: 32013765 | NGAY: 14 | PLANT: 3093"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["command"] == "LICHSU"
    assert res["materials"] == ["32013765"]
    assert res["days"] == 14
    assert res["plant"] == "3093"


def test_parse_json_marker_format():
    body = """
    SAP_REQUEST_BEGIN
    {
      "command": "MB52",
      "materials": ["32013765", "32024300"],
      "plant": "3093"
    }
    SAP_REQUEST_END
    """
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["command"] == "MB52"
    assert res["materials"] == ["32013765", "32024300"]
    assert res["plant"] == "3093"


def test_parse_invalid_body():
    body = "Xin chào, hôm nay họp lúc 2h nhé."
    res = EmailCommandParser.parse(body)
    assert res is None
