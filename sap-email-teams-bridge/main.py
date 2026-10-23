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
import gc
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
from app.sap_engine import SapEngine, SapExecutionError
from app.state_manager import StateManager
from app.number_utils import set_global_sap_number_format


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


def get_or_create_subfolder(parent_folder, folder_name: str):
    """Lấy hoặc tự động tạo thư mục con (Processed, Failed, v.v.)."""
    try:
        return parent_folder.Folders[folder_name]
    except Exception:
        try:
            return parent_folder.Folders.Add(folder_name)
        except Exception:
            return None


def move_mail_to_folder(mail, target_folder):
    """Chuyển mail vào thư mục chỉ định an toàn."""
    if not target_folder:
        return
    try:
        mail.Move(target_folder)
    except Exception as e:
        print(f"[OUTLOOK MOVE WARN] Không thể chuyển thư: {e}")


def move_mail_to_processed(mail, bot_folder):
    """Chuyển mail đã hoàn tất vào thư mục con Processed."""
    proc = get_or_create_subfolder(bot_folder, "Processed")
    move_mail_to_folder(mail, proc)


def move_mail_to_failed(mail, bot_folder):
    """Chuyển mail lỗi nghiệp vụ/không xử lý được vào thư mục con Failed."""
    failed = get_or_create_subfolder(bot_folder, "Failed")
    move_mail_to_folder(mail, failed)


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
    except Exception as e:
        raise RuntimeError(
            f"Không thể kết nối tới ứng dụng Microsoft Outlook qua COM.\n"
            f"Vui lòng đảm bảo Outlook Classic (Desktop) đang được mở trên máy!\nChi tiết: {e}"
        )

    # Tìm hoặc tự động tạo mới thư mục BOT trong Outlook
    bot_folder = None

    # 1. Quét trong Default Inbox
    try:
        for f in inbox.Folders:
            if f.Name.strip().lower() == folder_name.strip().lower():
                bot_folder = f
                print(f"[OUTLOOK] Tìm thấy thư mục '{f.Name}' trong '{inbox.Name}'.")
                break
    except Exception:
        pass

    # 2. Quét tại cấp Root của tài khoản (VD: v.taipt8@futurevn.com -> SAP_BOT)
    if bot_folder is None:
        try:
            root_store = inbox.Parent
            for f in root_store.Folders:
                if f.Name.strip().lower() == folder_name.strip().lower():
                    bot_folder = f
                    print(f"[OUTLOOK] Tìm thấy thư mục '{f.Name}' tại cấp gốc tài khoản '{root_store.Name}'.")
                    break
        except Exception:
            pass

    # 3. Quét tất cả các Store / Mailbox trong Outlook
    if bot_folder is None:
        try:
            for store in outlook_ns.Folders:
                for f in store.Folders:
                    if f.Name.strip().lower() == folder_name.strip().lower():
                        bot_folder = f
                        print(f"[OUTLOOK] Tìm thấy thư mục '{f.Name}' trong hòm thư '{store.Name}'.")
                        break
                if bot_folder:
                    break
        except Exception:
            pass

    # 4. Nếu chưa có ở bất kỳ đâu -> TỰ ĐỘNG TẠO MỚI tại gốc tài khoản hoặc trong Inbox
    if bot_folder is None:
        try:
            try:
                bot_folder = inbox.Parent.Folders.Add(folder_name)
                print(f"[OUTLOOK] Đã tự động tạo mới thư mục '{folder_name}' tại gốc tài khoản!")
            except Exception:
                bot_folder = inbox.Folders.Add(folder_name)
                print(f"[OUTLOOK] Đã tự động tạo mới thư mục '{folder_name}' trong Inbox!")
        except Exception as add_err:
            raise RuntimeError(
                f"Không tìm thấy và không thể tự động tạo thư mục '{folder_name}' trong Outlook.\n"
                f"Vui lòng tạo thư mục mang tên '{folder_name}' trong Outlook trước khi chạy!\n"
                f"Chi tiết: {add_err}"
            )

    set_global_sap_number_format(config.get("sap_number_format", "AUTO"))
    state_mgr = StateManager()
    max_retries = config.get("max_webhook_retries", 3)

    resolver = MaterialResolver()
    sap = SapEngine(
        max_alv_rows=max_rows,
        keep_alive_idle_seconds=keep_alive_idle,
        connection_index=config.get("sap_connection_index", 0),
        session_index=config.get("sap_session_index", 0),
        strict_session_safety=config.get("strict_session_safety", True)
    )
    print(f"[READY] Đang lắng nghe thư mới trong '{bot_folder.Name}'...")

    seq_counter = 1
    allowed_senders = [s.strip().lower() for s in config.get("allowed_senders", []) if s.strip()]

    while True:
        try:
            # 1. Giữ phiên SAP không bị timeout khi nhàn rỗi
            sap.keep_alive()

            # 2. Xử lý hàng đợi Outbox (Retry gửi Webhook trực tiếp mà KHÔNG CẦN gọi lại SAP GUI)
            pending_outbox = state_mgr.get_pending_outbox(max_attempts=max_retries)
            for item in pending_outbox:
                req_out_id = item["req_id"]
                url_out = item["webhook_url"]
                pl_out = item["payload"]
                print(f"[OUTBOX RETRY] Thử gửi lại Webhook {req_out_id} (Lần thử {item['attempts'] + 1}/{max_retries})...")
                if TeamsDispatcher.send_raw_payload(url_out, pl_out):
                    state_mgr.record_outbox_success(req_out_id)
                    print(f"[OUTBOX SUCCESS] Gửi lại thành công Webhook {req_out_id}!")
                else:
                    state_mgr.record_outbox_failure(req_out_id, "Gửi lại thất bại", max_attempts=max_retries)

            # 3. Quét tất cả các thư trong thư mục SAP_BOT (kể cả thư đã mở xem lẫn thư chưa đọc)
            # Mọi thư đã xử lý xong đều được chuyển tự động sang thư mục con 'Processed'.
            # Do đó, bất kỳ thư nào còn nằm trong 'SAP_BOT' đều là yêu cầu đang chờ xử lý!
            folder_items = bot_folder.Items
            try:
                folder_items.Sort("[ReceivedTime]", False)  # Sắp xếp từ cũ đến mới (FIFO)
            except Exception:
                pass

            mail_candidates = []
            for item in list(folder_items):
                try:
                    if getattr(item, "Class", 0) == 43 or hasattr(item, "Subject"):
                        mail_candidates.append(item)
                except Exception:
                    pass

            for mail in mail_candidates:
                entry_id = getattr(mail, "EntryID", "")
                if entry_id and state_mgr.is_message_processed("", entry_id):
                    mail.UnRead = False
                    mail.Save()
                    move_mail_to_processed(mail, bot_folder)
                    continue

                req_id = f"SAP-{datetime.now().strftime('%H%M%S')}-{seq_counter:02d}"
                seq_counter += 1

                # Cô lập hoàn toàn lỗi từng thư để không chặn luồng
                try:
                    # 1. Kiểm tra danh sách người gửi được phép (nếu có cấu hình)
                    if allowed_senders:
                        actual_sender = str(getattr(mail, "SenderEmailAddress", "") or "").strip().lower()
                        if actual_sender and not any(actual_sender == s or actual_sender.endswith(s) for s in allowed_senders):
                            print(f"[SECURITY] {req_id} | Người gửi '{actual_sender}' không thuộc allowed_senders. Chuyển thư sang Failed.")
                            mail.UnRead = False
                            mail.Save()
                            move_mail_to_failed(mail, bot_folder)
                            continue

                    raw_body = getattr(mail, "Body", "")
                    html_body = getattr(mail, "HTMLBody", "")
                    subject = getattr(mail, "Subject", "")

                    meta = NaturalCommandParser.parse(raw_body, html_body=html_body, subject=subject, default_plant=def_plant)

                    # Kiểm tra trùng lặp qua SQLite
                    msg_id = meta.get("messageId", "") if meta else ""
                    if msg_id and state_mgr.is_message_processed(msg_id, entry_id):
                        print(f"[DUPLICATE] {req_id} | Message ID '{msg_id}' đã xử lý thành công trước đó. Bỏ qua.")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_processed(mail, bot_folder)
                        continue

                    if not meta or not meta.get("valid"):
                        if meta and meta.get("needHelpReply") and meta.get("messageId"):
                            print(f"[COACHING] {req_id} | Người dùng cần hướng dẫn -> Gửi trợ giúp lên Teams...")
                            help_html = TeamsDispatcher.build_help_html(meta.get("senderName", "Đồng nghiệp"))
                            payload, data = TeamsDispatcher.prepare_safe_payload(
                                req_id,
                                meta,
                                "INFO",
                                "Trợ lý SAP: Hướng dẫn tra cứu cú pháp nhanh trên Teams",
                                [],
                                {},
                                custom_html=help_html
                            )
                            state_mgr.enqueue_outbox(req_id, webhook_url, payload)
                            if TeamsDispatcher.send_raw_payload(webhook_url, payload):
                                state_mgr.record_outbox_success(req_id)
                            else:
                                state_mgr.record_outbox_failure(req_id, "Gửi trợ giúp lần đầu thất bại", max_attempts=max_retries)

                            state_mgr.mark_message_processed(msg_id, entry_id, req_id, "HELP", "COACHING")
                            mail.UnRead = False
                            mail.Save()
                            move_mail_to_processed(mail, bot_folder)
                        else:
                            reason = meta.get("reason", "Không rõ") if meta else "Thư rỗng hoặc không liên quan"
                            print(f"[SKIP] {req_id} | Bỏ qua thư không hợp lệ: {reason}")
                            mail.UnRead = False
                            mail.Save()
                            move_mail_to_failed(mail, bot_folder)
                        continue

                    cmd = meta["command"]
                    raw_mats = meta["materials"]
                    plant = meta["plant"]
                    batch = meta["batch"]
                    days = meta["days"]

                    # 3. Phân giải mã thông minh từ Sheet DATA (vd: 32013765 -> BIW32013765AA)
                    if cmd != "PING":
                        resolved_mats, notes = resolver.resolve_list(raw_mats)
                        print(f"\n[DETECTED] {req_id} | Người hỏi: {meta['senderName']} ({meta['senderEmail']})")
                        print(f"  -> Ý định:   {cmd} ({meta['intentReason']})")
                        print(f"  -> Mã gốc:   {raw_mats} ──> Mã SAP: {resolved_mats}")
                        print(f"  -> Tham số:  Batch='{batch or 'ALL'}' | Plant='{plant}' | Số ngày={days}")
                    else:
                        resolved_mats = []
                        print(f"\n[DETECTED] {req_id} | Lệnh PING kiểm tra kết nối từ {meta['senderName']}")

                    records = []
                    headers = {}
                    summary = ""
                    custom_html = None

                    try:
                        if cmd == "PING":
                            custom_html = (
                                f"<div style='font-family:Segoe UI,sans-serif;background:#dafbe1;border:1px solid #4ac26b;"
                                f"border-left:4px solid #1a7f37;padding:12px 14px;border-radius:4px;color:#24292f;'>"
                                f"<div style='font-weight:bold;color:#1a7f37;font-size:14px;margin-bottom:6px;'>"
                                f"🏓 PONG! Kết nối hoàn tất thành công!</div>"
                                f"<div style='font-size:12px;line-height:1.6;'>"
                                f"• <strong>Request ID:</strong> {meta.get('requestId') or req_id}<br>"
                                f"• <strong>Thư mục lắng nghe:</strong> Outlook/{bot_folder.Name}<br>"
                                f"• <strong>Plant:</strong> {plant}<br>"
                                f"• <strong>Trạng thái SAP:</strong> Sẵn sàng (0 Open Port / 0 PowerShell)<br>"
                                f"• <strong>Ghi chú:</strong> {meta.get('note', 'Kiểm tra đường truyền hoàn tất.')}"
                                f"</div></div>"
                            )
                            summary = f"Kiểm tra kết nối PING-PONG thành công (Request ID: {meta.get('requestId') or req_id})"

                        elif cmd in ["VITRI", "VỊ TRÍ", "BIN", "POST"]:
                            records = sap.check_position_posting_change(resolved_mats, batch, plant=plant)
                            if not records:
                                print(f"[FALLBACK] /SCWM/POST không có dữ liệu, tự động tra cứu tồn kho qua MB52...")
                                fallback_records = sap.check_stock_mb52(resolved_mats, plant)
                                if fallback_records:
                                    records = fallback_records
                                    headers = {
                                        "product": "Mã hàng", "plant": "Plant", "sloc": "Kho (SLoc)",
                                        "batch": "Batch", "unrest": "Tồn tự do", "qual": "Chờ kiểm",
                                        "bin": "Vị trí (Bin)", "unit": "ĐVT"
                                    }
                                    mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                                    summary = f"Tồn kho & Vị trí (MB52 - Tự động tra cứu) cho mã: {', '.join(resolved_mats)}{mapping_info}"
                                else:
                                    headers = {
                                        "product": "Mã hàng", "batch": "Batch", "bin": "Vị trí (Bin)",
                                        "qty": "Số lượng", "unit": "ĐVT",
                                        "grDate": "Ngày GR", "grTime": "Giờ GR",
                                        "cat": "Trạng thái"
                                    }
                                    mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                                    summary = f"Vị trí lưu kho & Giờ GR (/SCWM/POST) cho mã: {', '.join(resolved_mats)}{mapping_info}"
                            else:
                                headers = {
                                    "product": "Mã hàng", "batch": "Batch", "bin": "Vị trí (Bin)",
                                    "qty": "Số lượng", "unit": "ĐVT",
                                    "grDate": "Ngày GR", "grTime": "Giờ GR",
                                    "cat": "Trạng thái"
                                }
                                mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                                summary = f"Vị trí lưu kho & Giờ GR (/SCWM/POST) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        elif cmd in ["TON", "TỒN", "MB52"]:
                            records = sap.check_stock_mb52(resolved_mats, plant)
                            headers = {
                                "product": "Mã hàng", "plant": "Plant", "sloc": "Kho (SLoc)",
                                "batch": "Batch", "unrest": "Tồn tự do", "qual": "Chờ kiểm",
                                "bin": "Vị trí (Bin)", "unit": "ĐVT"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Tồn kho tổng hợp (MB52 - Plant {plant}) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        elif cmd in ["LICHSU", "LỊCH SỬ", "MB51"]:
                            records = sap.check_history_mb51(resolved_mats, days, plant)
                            headers = {
                                "product": "Mã hàng", "docNo": "Số CT", "mvtType": "Loại GD",
                                "postDate": "Ngày HT", "qty": "Số lượng", "unit": "ĐVT",
                                "sloc": "Kho", "batch": "Batch"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Lịch sử giao dịch {days} ngày (MB51) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        elif cmd in ["IB_TREO", "IB", "INBOUND", "VL06I"]:
                            records = sap.check_inbound_treo_vl06i(resolved_mats, plant=plant)
                            headers = {
                                "ibNo": "Số IB", "refNo": "Số OB / Vận đơn", "product": "Mã hàng",
                                "batch": "Số Lô", "qty": "SL giao", "unit": "ĐVT",
                                "plant": "Plant", "sloc": "Kho nhận", "status": "Trạng thái"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Danh sách đơn Inbound (IB) treo chưa GR (VL06I) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        elif cmd in ["STO", "ME2N", "MB5T", "DIEUCHUYEN"]:
                            records = sap.check_sto_me2n(resolved_mats, plant=plant)
                            headers = {
                                "stoNo": "Số STO", "item": "Item", "product": "Mã hàng",
                                "orderQty": "SL Đặt", "deliveredQty": "Đã xuất", "remainingQty": "Còn nợ",
                                "unit": "ĐVT", "supplyingPlant": "Plant cấp", "receivingPlant": "Plant nhận"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Danh sách đơn điều chuyển STO (ME2N) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        # Chuẩn bị payload chuẩn và đưa vào hàng đợi Outbox
                        payload, data = TeamsDispatcher.prepare_safe_payload(
                            req_id, meta, "SUCCEEDED", summary, records, headers, custom_html=custom_html
                        )
                        state_mgr.enqueue_outbox(req_id, webhook_url, payload)

                        # Thử gửi ngay lập tức
                        if TeamsDispatcher.send_raw_payload(webhook_url, payload):
                            state_mgr.record_outbox_success(req_id)
                            print(f"[COMPLETED] {req_id} | Đã hoàn tất tra cứu SAP & gửi thành công tới Teams!")
                        else:
                            state_mgr.record_outbox_failure(req_id, "Lần gửi đầu tiên qua Webhook thất bại", max_attempts=max_retries)
                            print(f"[OUTBOX QUEUED] {req_id} | Tra cứu SAP xong. Webhook chưa gửi được, đã lưu vào Outbox SQLite để tự động thử lại.")

                        state_mgr.mark_message_processed(msg_id, entry_id, req_id, cmd, "SUCCEEDED")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_processed(mail, bot_folder)

                    except Exception as sap_err:
                        print(f"[SAP ERROR] {req_id} | {sap_err}")
                        err_html = f"<p style='color:#d9534f;font-weight:bold;'>❌ Lỗi khi truy xuất SAP GUI: {sap_err}</p>"
                        fail_summary = f"Lỗi tra cứu SAP: {sap_err}"
                        fail_payload, fail_data = TeamsDispatcher.prepare_safe_payload(
                            req_id, meta, "FAILED", fail_summary, [], {}, custom_html=err_html
                        )
                        state_mgr.enqueue_outbox(req_id, webhook_url, fail_payload)
                        if TeamsDispatcher.send_raw_payload(webhook_url, fail_payload):
                            state_mgr.record_outbox_success(req_id)
                        else:
                            state_mgr.record_outbox_failure(req_id, f"Lỗi Webhook sau lỗi SAP: {sap_err}", max_attempts=max_retries)

                        state_mgr.mark_message_processed(msg_id, entry_id, req_id, cmd, "FAILED")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_failed(mail, bot_folder)

                except Exception as mail_err:
                    print(f"[MAIL ERROR] {req_id} | Lỗi khi xử lý thư: {mail_err}")
                    try:
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_failed(mail, bot_folder)
                    except Exception:
                        pass

        except Exception as loop_err:
            print(f"[LOOP WARN] {loop_err}")

        # Tối ưu giải phóng RAM và COM object sau mỗi vòng lặp chạy nền
        try:
            import gc
            gc.collect()
            pythoncom.CoFreeUnusedLibraries()
        except Exception:
            pass

        time.sleep(poll_sec)


if __name__ == "__main__":
    if "--test-parser" in sys.argv:
        run_self_test()
    else:
        run_bridge()
