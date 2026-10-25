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







def test_parse_tim_gr_doc_json():
    """Kiểm tra parse lệnh TIM_GR_DOC từ JSON."""
    json_cmd = {
        "command": "TIM_GR_DOC",
        "inboundNo": "1800123456",
        "refNo": "BL123456",
        "materialDocument": "5000123456",
        "materials": ["BIW32013765AA"],
        "days": 14,
        "plant": "3093"
    }
    res = EmailCommandParser.parse(json_cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TIM_GR_DOC"
    assert res["inboundNo"] == "1800123456"
    assert res["refNo"] == "BL123456"
    assert res["materialDocument"] == "5000123456"
    assert res["materials"] == ["BIW32013765AA"]
    assert res["days"] == 14
    assert res["plant"] == "3093"


def test_parse_tim_gr_doc_natural():
    """Kiểm tra parse lệnh TIM_GR_DOC từ ngôn ngữ tự nhiên."""
    cmd = "tìm chứng từ gr theo inbound 1800123456"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TIM_GR_DOC"
    assert res["inboundNo"] == "1800123456"


def test_parse_check_ib_json():
    """Kiểm tra parse lệnh CHECK_IB từ JSON với expectedQty."""
    json_cmd = {
        "command": "CHECK_IB",
        "inboundNo": "1800555666",
        "expectedQty": 350.5,
        "plant": "3093"
    }
    res = EmailCommandParser.parse(json_cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_IB"
    assert res["inboundNo"] == "1800555666"
    assert res["expectedQty"] == 350.5


def test_parse_check_ib_natural():
    """Kiểm tra parse lệnh CHECK_IB từ ngôn ngữ tự nhiên."""
    cmd = "chi tiết inbound 1800555666 sl 500"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_IB"
    assert res["inboundNo"] == "1800555666"
    assert res["expectedQty"] == 500.0


def test_parse_tim_ib_json():
    """Kiểm tra parse lệnh TIM_IB từ JSON."""
    json_cmd = {
        "command": "TIM_IB",
        "poNo": "4500123456",
        "refNo": "BILL789"
    }
    res = EmailCommandParser.parse(json_cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TIM_IB"
    assert res["poNo"] == "4500123456"
    assert res["refNo"] == "BILL789"


def test_parse_tim_ib_natural():
    """Kiểm tra parse lệnh TIM_IB từ ngôn ngữ tự nhiên."""
    cmd = "tìm inbound theo po 4500123456"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TIM_IB"
    assert res["poNo"] == "4500123456"


def test_parse_check_po_json():
    """Kiểm tra parse lệnh CHECK_PO từ JSON."""
    json_cmd = {
        "command": "CHECK_PO",
        "poNo": "4500888999",
        "poItem": "20"
    }
    res = EmailCommandParser.parse(json_cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_PO"
    assert res["poNo"] == "4500888999"
    assert res["poItem"] == "20"


def test_parse_check_po_natural():
    """Kiểm tra parse lệnh CHECK_PO từ ngôn ngữ tự nhiên."""
    cmd = "kiểm tra PO 4500888999 item 10"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_PO"
    assert res["poNo"] == "4500888999"
    assert res["poItem"] == "10"


def test_parse_json_parity_ob_and_bin():
    """Kiểm tra tính tương đương tham số JSON cho obNo và bin."""
    json_cmd = {
        "command": "CHECK_OB",
        "obNo": "80012345",
        "materials": ["BIW32013765AA"]
    }
    res = EmailCommandParser.parse(json_cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_OB"
    assert res["obNo"] == "80012345"

    json_bin = {
        "command": "LICHSU_BIN",
        "bin": "10-A-01",
        "materials": ["BIW32013765AA"]
    }
    res_bin = EmailCommandParser.parse(json_bin)
    assert res_bin is not None
    assert res_bin["valid"] is True
    assert res_bin["command"] == "LICHSU_BIN"
    assert res_bin["bin"] == "10-A-01"


def test_parse_check_fg_ob_natural_and_json():
    """Kiểm tra parse lệnh CHECK_FG_OB (Điền OB thành phẩm)."""
    cmd = "điền ob BIW32013765AA"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_FG_OB"
    assert "BIW32013765AA" in res["materials"]

    json_cmd = {"command": "CHECK_FG_OB", "materials": ["BIW32013765AA"]}
    res_j = EmailCommandParser.parse(json_cmd)
    assert res_j["valid"] is True
    assert res_j["command"] == "CHECK_FG_OB"


def test_parse_check_nvl_migo_natural_and_json():
    """Kiểm tra parse lệnh CHECK_NVL_MIGO (Ra đơn MIGO 311 NVL)."""
    cmd = "ra đơn migo 32013765"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_NVL_MIGO"
    assert "32013765" in res["materials"]

    cmd2 = "check nvl 32013765"
    res2 = EmailCommandParser.parse(cmd2)
    assert res2["valid"] is True
    assert res2["command"] == "CHECK_NVL_MIGO"

    json_cmd = {"command": "MIGO_311", "materials": ["32013765"]}
    res_j = EmailCommandParser.parse(json_cmd)
    assert res_j["valid"] is True
    assert res_j["command"] == "CHECK_NVL_MIGO"

    cmd3 = "check postingchange 32013765"
    res3 = EmailCommandParser.parse(cmd3)
    assert res3["valid"] is True
    assert res3["command"] == "CHECK_NVL_MIGO"
    assert "32013765" in res3["materials"]

    json_cmd2 = {"command": "CHECK_POSTING_CHANGE", "materials": ["32013765"]}
    res_j2 = EmailCommandParser.parse(json_cmd2)
    assert res_j2["valid"] is True
    assert res_j2["command"] == "CHECK_NVL_MIGO"


def test_loop_guard_detects_and_drops_bot_echo():
    """Kiểm tra Loop Guard: Phát hiện và drop ngay lập tức các tin nhắn phản hồi của Bot."""
    bot_reply_body = (
        "Sếp ơi, em tra được rồi ạ. Kết quả của sếp đây 👇\n"
        "Vị trí lưu kho & Giờ GR (/SCWM/POST) cho mã: BIW73015987AA\n"
        "Mã tham chiếu: SAP-072101-01"
    )
    res = EmailCommandParser.parse(bot_reply_body)
    assert res is not None
    assert res.get("valid") is False
    assert res.get("isBotEcho") is True
    assert res.get("command") == "LOOP_GUARD"


def test_parse_att_mail_with_unc_and_ticket_numbers():
    """Kiểm tra lệnh 'att mail' đính kèm thư mục UNC và phiếu: Không bị phân loại nhầm sang posting change và không nhận nhầm IP làm vật tư."""
    text = (
        "SAP att mail nhé\n"
        "\\\\10.128.10.72\\fvn$\\test\\folder\n"
        "5007072231  THÀNH CÔNG in - 6156\n"
        "5007072232  THÀNH ĐẠT in 6157"
    )
    res = EmailCommandParser.parse(text)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "ATT_MIGO"
    assert res["materials"] == []
    assert res["folderPath"] == "\\\\10.128.10.72\\fvn$\\test\\folder"
    assert len(res["attachmentItems"]) == 2
    assert res["attachmentItems"][0]["doc"] == "5007072231"
    assert "6156" in res["attachmentItems"][0]["ticket_no"]


def test_parse_mb51_explicit_vs_default_plant():
    """Kiểm tra phân biệt rõ explicitPlant=True khi người dùng có gõ plant, và False khi không gõ."""
    cmd1 = "SAP lịch sử xuất 3 ngày BIW32026700HA"
    res1 = EmailCommandParser.parse(cmd1)
    assert res1 is not None
    assert res1["command"] == "LICHSU"
    assert res1["days"] == 3
    assert "BIW32026700HA" in res1["materials"]
    assert res1["explicitPlant"] is False

    cmd2 = "SAP lịch sử xuất 5 ngày plant 3093 BIW32026700HA"
    res2 = EmailCommandParser.parse(cmd2)
    assert res2 is not None
    assert res2["command"] == "LICHSU"
    assert res2["days"] == 5
    assert res2["plant"] == "3093"
    assert res2["explicitPlant"] is True


def test_parse_ton_model_body_full_list():
    """Kiểm tra nhận diện câu lệnh tồn model Body với tùy chọn nguyên list / full list."""
    cmd = "SAP tồn limo body nguyên list"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["command"] == "TON_THAP"
    assert res["model"] == "LIMO"
    assert res["shop"] == "BODY"
    assert res["fullModelList"] is True
    assert res["limit"] == 100


def test_parse_dien_so_phieu_and_raw_excel_row():
    """Kiểm tra câu lệnh 'điền số phiếu' và cả định dạng copy thẳng từ Excel có mã doc 49xxxxxxx."""
    # 1. Câu lệnh có chữ 'SAP điền số phiếu'
    cmd1 = "SAP điền số phiếu 4995880054\tOUT_P1_23581"
    res1 = EmailCommandParser.parse(cmd1)
    assert res1 is not None
    assert res1["valid"] is True
    assert res1["command"] == "ATT_MIGO"
    assert res1["materials"] == []
    assert res1["bin"] == ""
    assert len(res1["attachmentItems"]) == 1
    assert res1["attachmentItems"][0]["doc"] == "4995880054"
    assert res1["attachmentItems"][0]["ticket_no"] == "OUT_P1_23581"

    # Raw pasted rows are data, not authorization to write SAP.
    res2 = EmailCommandParser.parse("4995880054\tOUT_P1_23581")
    assert res2["valid"] is False
    assert res2["explicitWriteIntent"] is False
    assert res2["detectedIntent"] == "AMBIGUOUS_DOCUMENT_LIST"



def test_parse_check_con_ob():
    """Kiểm tra câu lệnh 'SAP check còn OB mã BIW70007551HA không' nhận diện chuẩn CHECK_OB."""
    cmd = "SAP check còn OB mã BIW70007551HA không"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_OB"
    assert "BIW70007551HA" in res["materials"]


def test_parse_check_bom():
    """Kiểm tra câu lệnh 'SAP check bom BIW70007551HA' nhận diện chuẩn CHECK_BOM."""
    cmd = "SAP check bom BIW70007551HA"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_BOM"
    assert "BIW70007551HA" in res["materials"]


def test_parse_check_queue():
    """Kiểm tra câu lệnh 'SAP check queue 18001234' nhận diện chuẩn CHECK_QUEUE."""
    cmd = "SAP check queue 18001234"
    res = EmailCommandParser.parse(cmd)
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "CHECK_QUEUE"


def test_parse_tao_bbbg_natural():
    res = EmailCommandParser.parse("tạo bbbg fg xưởng hàn plant 3093")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TAO_BBBG"
    assert res["plant"] == "3093"


def test_parse_ra_don_nvl_natural():
    res = EmailCommandParser.parse("ra đơn nvl v31 zone mã BIW32013765AA số lượng 100")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "RA_DON_NVL"
    assert "BIW32013765AA" in res["materials"]


def test_parse_dieu_phoi_ob_natural():
    res = EmailCommandParser.parse("điều phối ob rc65 đơn 8000123456")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "DIEU_PHOI_OB"


def test_parse_tra_cuu_asn_natural():
    res = EmailCommandParser.parse("tra cứu asn xuất xứ container TCLU1234567")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "TRA_CUU_ASN"
