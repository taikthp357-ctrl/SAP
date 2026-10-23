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

