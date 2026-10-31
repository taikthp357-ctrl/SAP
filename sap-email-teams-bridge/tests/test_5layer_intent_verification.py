# -*- coding: utf-8 -*-
"""Kiểm thử Phễu Nhận Diện Đa Tầng (5-Layer Intent Verification):
1. Lọc nhiễu Chữ ký & Quoted Reply/Forward
2. Giảm trọng số Tiêu đề cũ (Subject Hijacking)
3. Chấm điểm Cụm động từ hành động
4. Đối chiếu Hình thái dữ liệu (Payload Shape) & Luật Phủ quyết cứng (Hard Veto)
5. Ngưỡng tự tin (Confidence Gate)
"""
import pytest
from app.email_parser import SmartCommandParser, EmailCommandParser


def test_signature_stripping_and_job_title_noise():
    """Chữ ký chứa 'Chuyên viên Kiểm soát Hóa đơn VAT & MIGO' không được làm sai lệch lệnh điền số phiếu."""
    mail_body = """Điền giúp anh số phiếu 6381 cho các doc này:
5007135730\tIN_6381

Trân trọng,
Nguyễn Văn A - Chuyên viên Kiểm soát Hóa đơn VAT & MIGO Kho 1005
Tel: 0987654321 - Ext: 1234
VinFast Hải Phòng
"""
    res = SmartCommandParser.parse(mail_body, subject="Điền số phiếu")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "MIGO_HEADTEXT", f"Phải là MIGO_HEADTEXT thay vì ATT_MIGO, nhận được: {res['command']}"
    assert "VAT" not in res.get("topZone", "")
    assert "MIGO Kho" not in res.get("topZone", "")


def test_quoted_reply_forward_trap_neutralization():
    """Người dùng forward mail cũ có nội dung 'Check att mail BBBG và VAT', nhưng gõ thêm lệnh mới ở đầu: 'Điền số phiếu'."""
    mail_body = """Điền giúp anh số phiếu cho các doc này nhé:
4905195632\tIN_23581

-----Original Message-----
From: sep@vinfast.vn
Sent: Friday, October 30, 2026 8:00 AM
To: logistics@vinfast.vn
Subject: [SAP] Check att mail BBBG và VAT
Nhờ các bạn kiểm tra đính kèm và att mail gấp toàn bộ BBBG, VAT cho các chứng từ tháng 10.
Các chứng từ thiếu file: 4905195632, 4905195633, 4905195634.
"""
    res = SmartCommandParser.parse(mail_body, subject="Re: [SAP] Check att mail BBBG và VAT")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "MIGO_HEADTEXT", f"Phải là MIGO_HEADTEXT, không bị kéo về ATT_MIGO, nhận được: {res['command']}"


def test_subject_hijacking_mitigation():
    """Subject cũ là 'Re: Đính kèm BBBG và VAT' nhưng nội dung mới gõ lệnh điền số phiếu."""
    mail_body = "Điền giúp em số phiếu 1012993097 vào doc 4905195632 nhé anh."
    res = SmartCommandParser.parse(mail_body, subject="Re: Đính kèm BBBG và VAT tháng 10")
    assert res is not None
    assert res["valid"] is True
    assert res["command"] == "MIGO_HEADTEXT"


def test_payload_shape_cross_check_hard_veto_att_migo():
    """Gõ nhầm 'Đính kèm số phiếu vào MIGO' nhưng 0 file đính kèm và có bảng 2 cột -> Phải chuyển về MIGO_HEADTEXT."""
    mail_body = """Đính kèm số phiếu vào MIGO giúp anh nhé:
4905195632\tPXK-0012
"""
    # Không truyền attachments (hoặc attachments rỗng)
    res = SmartCommandParser.parse(mail_body, attachments=[])
    assert res is not None
    assert res["valid"] is True
    # Nhờ Payload Shape, ATT_MIGO bị VETO (0 file) và MIGO_HEADTEXT được BOOST (+40)
    assert res["command"] == "MIGO_HEADTEXT"
