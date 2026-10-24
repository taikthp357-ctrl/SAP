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


def test_parse_real_power_automate_ping_json():
    mail_body = """
    THẬN TRỌNG: Email này được gửi từ BÊN NGOÀI. Hãy kiểm tra kỹ địa chỉ email người gửi và thận trọng trước khi mở liên kết hoặc tệp đính kèm.
    CAUTION: This email was sent from an EXTERNAL source. Do not click links or open attachments unless you recognize the sender and know the content is safe.

    {
    "schemaVersion": 1,
    "requestId": "08584088749507581526025903895CU04",
    "command": "PING",
    "plant": "3093",
    "dryRun": true,
    "source": "PowerAutomate",
    "note": "Kiem tra nhan lenh qua email. Chua chay SAP."
    }

    If you want to unsubscribe from these emails, please use this form.
    """
    res = EmailCommandParser.parse(mail_body)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "PING"
    assert res["requestId"] == "08584088749507581526025903895CU04"
    assert res["dryRun"] is True
    assert res["plant"] == "3093"


def test_parse_real_mail_v3_with_sender_id_chat_id():
    mail_body = """
    THẬN TRỌNG: Email này là từ bên ngoài của VinFast.
    CAUTION: This email originated from outside of VinFast.

    MSG_ID: 179275160350 SENDER_EMAIL: taipt8@futurevn.com SENDER_NAME: Phạm Tấn Tài SENDER_ID: 29:1abc CHAT_ID: 19:2def RAW_TEXT: @Bot check vị trí giúp mã 32013765

    If you want to unsubscribe from these emails, please use this form.
    """
    res = EmailCommandParser.parse(mail_body, subject="[SAP-REQUEST] 179275160350")
    assert res is not None
    assert res["valid"] is True
    assert res["messageId"] == "179275160350"
    assert res["senderEmail"] == "taipt8@futurevn.com"
    assert res["senderName"] == "Phạm Tấn Tài"
    assert res["senderId"] == "29:1abc"
    assert res["chatId"] == "19:2def"
    assert res["command"] == "VITRI"
    assert res["materials"] == ["32013765"]


def test_parse_check_ib_treo():
    body = "@Bot check ib treo mã 32013765 plant 3093"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "IB_TREO"
    assert res["materials"] == ["32013765"]
    assert res["plant"] == "3093"


def test_parse_natural_language_ib_treo():
    body = "mã này BIW32013765AA có ib treo không"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "IB_TREO"
    assert res["materials"] == ["BIW32013765AA"]


def test_parse_check_sto():
    body = "check sto mã 32013765"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "STO"
    assert res["materials"] == ["32013765"]


def test_parse_natural_language_sto():
    body = "kiểm tra đơn điều chuyển STO cho vật tư BIW32013765AA plant 3093"
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "STO"
    assert res["materials"] == ["BIW32013765AA"]
    assert res["plant"] == "3093"


def test_parse_json_invalid_command_coaching():
    """Kiểm tra JSON có command không hỗ trợ -> trả về Coaching / Trợ giúp."""
    body = """
    SAP_REQUEST_BEGIN
    {
      "command": "UNKNOWN_TCODE_XYZ",
      "materials": ["32013765"],
      "senderEmail": "test@vinfast.vn"
    }
    SAP_REQUEST_END
    """
    res = EmailCommandParser.parse(body)
    assert res is not None
    assert res["valid"] is False
    assert res["needHelpReply"] is True
    assert "không được hỗ trợ" in res["reason"]


def test_parse_json_days_validation():
    """Kiểm tra parse an toàn trường days: không crash với string, clamp hợp lệ."""
    # 1. Chuỗi 'abc' -> mặc định về 7
    body_str = """
    SAP_REQUEST_BEGIN
    {
      "command": "MB51",
      "materials": ["32013765"],
      "days": "abc"
    }
    SAP_REQUEST_END
    """
    res1 = EmailCommandParser.parse(body_str)
    assert res1 is not None
    assert res1["days"] == 7

    # 2. Số âm -> clamp về tối thiểu 1 ngày
    body_neg = """
    SAP_REQUEST_BEGIN
    {
      "command": "MB51",
      "materials": ["32013765"],
      "days": -10
    }
    SAP_REQUEST_END
    """
    res2 = EmailCommandParser.parse(body_neg)
    assert res2 is not None
    assert res2["days"] == 1

    # 3. Số quá lớn -> clamp về tối đa 365 ngày
    body_large = """
    SAP_REQUEST_BEGIN
    {
      "command": "MB51",
      "materials": ["32013765"],
      "days": 1000
    }
    SAP_REQUEST_END
    """
    res3 = EmailCommandParser.parse(body_large)
    assert res3 is not None
    assert res3["days"] == 365


def test_parse_json_materials_type_validation():
    """Kiểm tra bóc tách và lọc danh sách materials: chấp nhận số/chuỗi, loại bỏ rác/None."""
    # 1. Danh sách có chứa số và chuỗi hợp lệ
    body_valid = """
    SAP_REQUEST_BEGIN
    {
      "command": "MB52",
      "materials": [32013765, "32024300", "BIW32013765AA"]
    }
    SAP_REQUEST_END
    """
    res1 = EmailCommandParser.parse(body_valid)
    assert res1 is not None
    assert res1["valid"] is True
    assert res1["materials"] == ["32013765", "32024300", "BIW32013765AA"]

    # 2. Danh sách toàn đối tượng không hợp lệ (None, {}, []) -> báo lỗi thiếu materials
    body_invalid = """
    SAP_REQUEST_BEGIN
    {
      "command": "MB52",
      "materials": [null, {}, [], "", false]
    }
    SAP_REQUEST_END
    """
    res2 = EmailCommandParser.parse(body_invalid)
    assert res2 is not None
    assert res2["valid"] is False
    assert res2["needHelpReply"] is True
    assert "Không tìm thấy danh sách mã vật tư" in res2["reason"]


def test_parse_ton_kho_distinction():
    """Kiểm tra phân biệt 'Tồn kho' (chỉ SLoc 1005 & 3093) vs 'Tồn' (toàn plant 3093)."""
    # 1. Có từ 'kho' -> TON_KHO
    res1 = EmailCommandParser.parse("Tồn kho BIW71004683 @Phạm Tiến Tài")
    assert res1 is not None
    assert res1["valid"] is True
    assert res1["command"] == "TON_KHO"
    assert "BIW71004683" in res1["materials"]

    # 2. Không có từ 'kho' -> TON (Toàn bộ sloc)
    res2 = EmailCommandParser.parse("Check tồn BIW71004683")
    assert res2 is not None
    assert res2["valid"] is True
    assert res2["command"] == "TON"
    assert "BIW71004683" in res2["materials"]

    # 3. JSON TON_KHO
    res3 = EmailCommandParser.parse('{"command": "TON_KHO", "materials": ["BIW71004683"]}')
    assert res3 is not None
    assert res3["command"] == "TON_KHO"


def test_parse_ton_thap_limo():
    """Kiểm tra nhận diện tự nhiên: '@tài check cho tôi 20 mã limo tồn thấp'."""
    cmd = "@tài check cho tôi 20 mã limo tồn thấp"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TON_THAP"
    assert res["model"] == "LIMO"
    assert res["shop"] == "BODY"
    assert res["plants"] == ["3093", "3001"]
    assert res["limit"] == 20


def test_parse_ton_thap_xmd():
    """Kiểm tra nhận diện xe máy điện: 'check tồn thấp xmd 10 mã' -> Plant 3093, 2001, 3001."""
    cmd = "check tồn thấp xmd 10 mã"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TON_THAP"
    assert res["model"] == "XMD"
    assert res["shop"] == "XMD"
    assert res["plants"] == ["3093", "2001", "3001"]
    assert res["limit"] == 10


def test_parse_check_ob():
    """Kiểm tra nhận diện câu hỏi kiểm tra OB: 'check ob có không mã BIW32013800AB'."""
    cmd = "check ob có không mã BIW32013800AB"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_OB"
    assert "BIW32013800AB" in res["materials"]


def test_parse_lichsu_bin():
    """Kiểm tra nhận diện lịch sử bin-to-bin: 'lịch sử bin BIW32013800AB'."""
    cmd = "lịch sử bin BIW32013800AB"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "LICHSU_BIN"
    assert "BIW32013800AB" in res["materials"]


def test_parse_vitri_by_bin():
    """Kiểm tra tra cứu trực tiếp theo tên bin: 'check bin 10-A-01'."""
    cmd = "check bin 10-A-01"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "VITRI"
    assert res["bin"] == "10-A-01"





