# -*- coding: utf-8 -*-
"""SAP - Outlook (Mail V3) - Teams Bridge Worker.

Tính năng:
1. Nhận diện ngôn ngữ tự nhiên (Intent & Entity Recognition) từ Mail V3 của Power Automate.
2. Tự động quy đổi mã thông minh (Smart Material Resolver) qua sheet DATA trong DATA_MASTER.xlsx.
3. Tự động khởi động SAP Logon nếu chưa chạy, kết nối lại nếu bị đứt session.
4. Cơ chế Keep-Alive định kỳ chống ngắt kết nối tự động của SAP NetWeaver.
5. Gửi thông báo kết quả định dạng bảng HTML đẹp mắt về kênh Teams qua Webhook.
6. Hỗ trợ cờ --test-parser để kiểm thử bộ phân tích trực tiếp trên console.
"""

import sys
import os
import time
import json
from pathlib import Path
from datetime import datetime
import pythoncom
import win32com.client

# Cho phép import các module trong app/
CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.config import load_config
from app.email_parser import NaturalCommandParser
from app.material_resolver import MaterialResolver
from app.teams_dispatcher import TeamsDispatcher
from app.sap_engine import SapEngine


def run_self_test():
    """Chạy kiểm thử bộ phân tích câu lệnh từ Email V3 mà chưa cần mở Outlook/SAP."""
    resolver = MaterialResolver()

    sample_v3_mails = [
        # Mẫu 1: Email V3 bị dính dòng + có Footer quảng cáo của Microsoft Power Apps + chỉ gõ mã 8 số rút gọn
        "MSG_ID: 172966001122 SENDER_EMAIL: kho01@congty.com SENDER_NAME: Nguyễn Văn Kho "
        "RAW_TEXT: <at>Bot</at> check vị trí giúp em mã 32013765, 32024300 lô B2610 với ạ! "
        "If you want to unsubscribe from these emails, please use this form.",

        # Mẫu 2: Hỏi lịch sử xuất nhập không dấu, gõ tự nhiên
        "MSG_ID: 172966009988<br>SENDER_EMAIL: plan02@congty.com<br>SENDER_NAME: Trần Kế Hoạch<br>"
        "RAW_TEXT: @Bot xem lich su xuat nhap ma BIW73012711CA trong 14 ngay plant 3093",

        # Mẫu 3: Hỏi tồn kho ngắn gọn
        "MSG_ID: 172966005544\nSENDER_EMAIL: sx@congty.com\nSENDER_NAME: Lê Sản Xuất\n"
        "RAW_TEXT: check tồn kho 32013765 plant 3093",

        # Mẫu 4: Người dùng chỉ nhắn duy nhất mã số
        "MSG_ID: 172966007788 SENDER_EMAIL: user@congty.com SENDER_NAME: Anh Ba RAW_TEXT: 32013765 B2610"
    ]

    print("==================================================================")
    print(" CHẠY KIỂM THỬ BỘ PHÂN TÍCH Ý ĐỊNH & QUY ĐỔI MÃ VẬT TƯ (MAIL V3)")
    print("==================================================================")

    for i, mail_text in enumerate(sample_v3_mails, 1):
        parsed = NaturalCommandParser.parse_from_mail(mail_text)
        print(f"\n--- Test Case #{i} ---")
        if not parsed or not parsed.get("valid"):
            print(f"❌ Không hợp lệ: {parsed}")
            continue

        raw_mats = parsed["materials"]
        resolved_mats, notes = resolver.resolve_list(raw_mats)

        print(f"• Người gửi:     {parsed['senderName']} ({parsed['senderEmail']})")
        print(f"• Mã tin (Msg):  {parsed['messageId']}")
        print(f"• Lệnh nhận:     {parsed['command']} (Lý do: {parsed['intentReason']})")
        print(f"• Batch:         {parsed['batch'] or 'Tất cả'}")
        print(f"• Plant:         {parsed['plant']}")
        print(f"• Mã gốc nhận:   {raw_mats}")
        print(f"• Mã SAP quy đổi:{resolved_mats}")
        for note in notes:
            print(f"  └─ {note}")


def get_or_create_processed_folder(bot_folder):
    """Lấy hoặc tự động tạo thư mục con Processed bên trong thư mục BOT."""
    try:
        return bot_folder.Folders["Processed"]
    except Exception:
        try:
            return bot_folder.Folders.Add("Processed")
        except Exception:
            return None


def move_mail_to_processed(mail, bot_folder):
    """Chuyển mail đã đọc vào thư mục Processed để giữ Inbox luôn sạch sẽ."""
    try:
        proc_folder = get_or_create_processed_folder(bot_folder)
        if proc_folder:
            mail.Move(proc_folder)
    except Exception:
        pass


def run_bridge():
    """Hàm chạy chính của tiến trình Worker."""
    config = load_config()

    webhook_url = config.get("teams_webhook_url", "")
    folder_name = config.get("outlook_folder_name", "SAP_BOT")
    poll_sec = config.get("poll_interval_seconds", 8)
    def_plant = config.get("default_plant", "3093")
    max_rows = config.get("max_alv_rows", 300)
    keep_alive_idle = config.get("keep_alive_idle_seconds", 900)

    print("==================================================================")
    print(" SAP - OUTLOOK (MAIL V3) - TEAMS BRIDGE WORKER")
    print(" (0 OPEN PORT / 0 POWERSHELL / AUTO KEEP-ALIVE / INTENT PARSER)")
    print("==================================================================")
    print(f"• Thư mục Outlook: Inbox/{folder_name}")
    print(f"• Chu kỳ quét:     {poll_sec}s")
    print(f"• Plant mặc định:  {def_plant}")
    print(f"• Giới hạn ALV:    {max_rows} dòng")
    if "YOUR_WEBHOOK_URL_HERE" in webhook_url or not webhook_url:
        print("• Trạng thái Webhook: DRY-RUN (In payload ra terminal)")
    else:
        print("• Trạng thái Webhook: LIVE (Gửi trực tiếp lên Power Automate)")
    print("==================================================================\n")

    pythoncom.CoInitialize()

    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        inbox = outlook_ns.GetDefaultFolder(6)  # 6 = olFolderInbox
        bot_folder = inbox.Folders[folder_name]
    except Exception as e:
        raise RuntimeError(
            f"Không tìm thấy thư mục con 'Inbox/{folder_name}' trong Outlook. "
            f"Vui lòng tạo thư mục này trong Outlook Classic trước khi chạy!\nChi tiết: {e}"
        )

    resolver = MaterialResolver()
    sap = SapEngine(max_alv_rows=max_rows, keep_alive_idle_seconds=keep_alive_idle)
    print(f"[READY] Đang lắng nghe thư mới trong 'Inbox/{folder_name}'...")

    seq_counter = 1
    processed_msg_ids = set()

    while True:
        try:
            # 1. Giữ phiên SAP không bị timeout khi nhàn rỗi
            sap.keep_alive()

            # 2. Quét các thư chưa đọc trong thư mục BOT
            unread_items = bot_folder.Items.Restrict("[UnRead] = True")
            for mail in list(unread_items):
                req_id = f"SAP-{datetime.now().strftime('%H%M%S')}-{seq_counter:02d}"
                seq_counter += 1

                raw_body = getattr(mail, "Body", "")
                html_body = getattr(mail, "HTMLBody", "")
                subject = getattr(mail, "Subject", "")

                meta = NaturalCommandParser.parse(raw_body, html_body=html_body, subject=subject, default_plant=def_plant)

                if not meta or not meta.get("valid"):
                    if meta and meta.get("needHelpReply") and meta.get("messageId"):
                        msg_id = meta.get("messageId")
                        if msg_id in processed_msg_ids:
                            mail.UnRead = False
                            mail.Save()
                            move_mail_to_processed(mail, bot_folder)
                            continue
                        processed_msg_ids.add(msg_id)

                        print(f"[COACHING] {req_id} | Người dùng cần hướng dẫn -> Gửi trợ giúp lên Teams...")
                        help_html = TeamsDispatcher.build_help_html(meta.get("senderName", "Đồng nghiệp"))
                        TeamsDispatcher.send_reply(
                            webhook_url,
                            req_id,
                            meta,
                            "INFO",
                            "Trợ lý SAP: Hướng dẫn tra cứu cú pháp nhanh trên Teams",
                            [],
                            {},
                            custom_html=help_html
                        )
                    else:
                        reason = meta.get("reason", "Không rõ") if meta else "Thư rỗng hoặc không liên quan"
                        print(f"[SKIP] {req_id} | Bỏ qua thư: {reason}")

                    mail.UnRead = False
                    mail.Save()
                    move_mail_to_processed(mail, bot_folder)
                    continue

                # Chống xử lý trùng lặp (Anti-duplicate guard)
                msg_id = meta.get("messageId")
                if msg_id:
                    if msg_id in processed_msg_ids:
                        print(f"[DUPLICATE] {req_id} | Message ID '{msg_id}' đã xử lý trước đó. Bỏ qua.")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_processed(mail, bot_folder)
                        continue
                    processed_msg_ids.add(msg_id)
                    if len(processed_msg_ids) > 1000:
                        processed_msg_ids.clear()

                cmd = meta["command"]
                raw_mats = meta["materials"]
                plant = meta["plant"]
                batch = meta["batch"]
                days = meta["days"]

                # 3. Phân giải mã thông minh từ Sheet DATA (vd: 32013765 -> BIW32013765AA)
                resolved_mats, notes = resolver.resolve_list(raw_mats)
                print(f"\n[DETECTED] {req_id} | Người hỏi: {meta['senderName']} ({meta['senderEmail']})")
                print(f"  -> Ý định:   {cmd} ({meta['intentReason']})")
                print(f"  -> Mã gốc:   {raw_mats} ──> Mã SAP: {resolved_mats}")
                print(f"  -> Tham số:  Batch='{batch or 'ALL'}' | Plant='{plant}' | Số ngày={days}")

                try:
                    if cmd in ["VITRI", "VỊ TRÍ", "BIN", "POST"]:
                        records = sap.check_position_posting_change(resolved_mats, batch)
                        headers = {
                            "product": "Mã hàng", "batch": "Batch", "bin": "Vị trí (Bin)",
                            "grDate": "Ngày GR", "grTime": "Giờ GR", "qty": "Số lượng",
                            "unit": "ĐVT", "cat": "Trạng thái"
                        }
                        mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                        summary = f"Vị trí lưu kho & Giờ GR (/SCWM/POST) cho mã: {', '.join(resolved_mats)}{mapping_info}"
                        TeamsDispatcher.send_reply(webhook_url, req_id, meta, "SUCCEEDED", summary, records, headers)

                    elif cmd in ["TON", "TỒN", "MB52"]:
                        records = sap.check_stock_mb52(resolved_mats, plant)
                        headers = {
                            "product": "Mã hàng", "plant": "Plant", "sloc": "Kho (SLoc)",
                            "batch": "Batch", "unrest": "Tồn tự do", "qual": "Chờ kiểm", "unit": "ĐVT"
                        }
                        mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                        summary = f"Tồn kho tổng hợp (MB52 - Plant {plant}) cho mã: {', '.join(resolved_mats)}{mapping_info}"
                        TeamsDispatcher.send_reply(webhook_url, req_id, meta, "SUCCEEDED", summary, records, headers)

                    elif cmd in ["LICHSU", "LỊCH SỬ", "MB51"]:
                        records = sap.check_history_mb51(resolved_mats, days, plant)
                        headers = {
                            "product": "Mã hàng", "docNo": "Số CT", "mvtType": "Loại GD",
                            "postDate": "Ngày HT", "qty": "Số lượng", "unit": "ĐVT",
                            "sloc": "Kho", "batch": "Batch"
                        }
                        mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                        summary = f"Lịch sử giao dịch {days} ngày (MB51) cho mã: {', '.join(resolved_mats)}{mapping_info}"
                        TeamsDispatcher.send_reply(webhook_url, req_id, meta, "SUCCEEDED", summary, records, headers)

                except Exception as sap_err:
                    err_html = f"<p style='color:#d9534f;font-weight:bold;'>❌ Lỗi khi truy xuất SAP GUI: {sap_err}</p>"
                    TeamsDispatcher.send_reply(
                        webhook_url, req_id, meta, "FAILED", "Không thể hoàn tất tra cứu SAP", [], {}, custom_html=err_html
                    )
                    print(f"[SAP ERROR] {sap_err}")

                # 4. Đánh dấu đã đọc và chuyển vào thư mục Processed
                mail.UnRead = False
                mail.Save()
                move_mail_to_processed(mail, bot_folder)

        except Exception as loop_err:
            print(f"[LOOP WARN] {loop_err}")

        time.sleep(poll_sec)


if __name__ == "__main__":
    if "--test-parser" in sys.argv:
        run_self_test()
    else:
        run_bridge()
