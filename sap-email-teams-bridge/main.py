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
    max_rows = config.get("max_alv_rows", 15000)
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
        outlook = outlook_ns
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
    last_sync_time = 0.0

    while True:
        try:
            # Ép Outlook Classic gọi lệnh đồng bộ ngầm (Sync Send/Receive) với Server (mỗi 10s một lần, tránh nghẽn MAPI)
            now_ts = time.time()
            if now_ts - last_sync_time >= 10.0:
                try:
                    outlook.SendAndReceive(False)
                    last_sync_time = now_ts
                except Exception:
                    pass

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
            # 1. Quét thư mục bot_folder (SAP_BOT) - tối đa 30 thư đang chờ
            try:
                b_count = folder_items.Count
                for i in range(1, min(b_count + 1, 31)):
                    try:
                        item = folder_items.Item(i)
                        if getattr(item, "Class", 0) == 43 or hasattr(item, "Subject"):
                            mail_candidates.append(item)
                    except Exception:
                        pass
            except Exception:
                pass

            # 2. TỐC ĐỘ CAO (0ms): Dùng MAPI Restrict để nhặt ngay lập tức các thư [SAP-REQUEST] chưa đọc trong Inbox mà không nạp toàn bộ hòm thư
            if inbox and inbox != bot_folder:
                try:
                    unread_inbox = inbox.Items.Restrict("[UnRead] = True")
                    u_count = unread_inbox.Count
                    for i in range(1, min(u_count + 1, 31)):
                        try:
                            item = unread_inbox.Item(i)
                            subj = str(getattr(item, "Subject", "") or "")
                            if "[SAP-REQUEST]" in subj or "SAP_REQUEST" in subj or "SAP" in subj.upper():
                                mail_candidates.append(item)
                        except Exception:
                            pass
                except Exception:
                    pass

            for mail in mail_candidates:
                entry_id = getattr(mail, "EntryID", "")
                if entry_id and state_mgr.is_message_processed("", entry_id):
                    mail.UnRead = False
                    mail.Save()
                    move_mail_to_processed(mail, bot_folder)
                    continue

                req_id = f"REF-{datetime.now().strftime('%H%M%S')}-{seq_counter:02d}"
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

                    # LOOP GUARD: Nếu phát hiện tin nhắn Bot tự trả lời -> Bỏ qua ngay lập tức
                    if meta and meta.get("isBotEcho"):
                        print(f"[LOOP GUARD] {req_id} | Phát hiện tin nhắn phản hồi của Bot. Bỏ qua ngay lập tức để chống lặp!")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_processed(mail, bot_folder)
                        continue

                    # Kiểm tra trùng lặp qua SQLite
                    msg_id = meta.get("messageId", "") if meta else ""
                    if msg_id and state_mgr.is_message_processed(msg_id, entry_id):
                        print(f"[DUPLICATE] {req_id} | Message ID '{msg_id}' đã xử lý thành công trước đó. Bỏ qua.")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_processed(mail, bot_folder)
                        continue

                    if not meta or not meta.get("valid"):
                        reason = meta.get("reason", "Không rõ") if meta else "Thư rỗng hoặc không liên quan"
                        print(f"[SKIP] {req_id} | Bỏ qua yêu cầu không có mã hợp lệ: {reason}. Không gửi trợ giúp lên Teams.")
                        state_mgr.mark_message_processed(msg_id, entry_id, req_id, "SKIP", "INVALID_INPUT")
                        mail.UnRead = False
                        mail.Save()
                        move_mail_to_processed(mail, bot_folder)
                        continue

                    cmd = meta["command"]
                    raw_mats = meta["materials"]
                    plant = meta["plant"]
                    batch = meta["batch"]
                    days = meta["days"]

                    # 3. Phân giải mã thông minh từ Sheet DATA (vd: 32013765 -> BIW32013765AA)
                    if cmd == "TON_THAP":
                        model_name = meta.get("model", "LIMO")
                        if not raw_mats:
                            resolved_mats = resolver.get_materials_by_model(model_name)
                        else:
                            resolved_mats, notes = resolver.resolve_list(raw_mats)
                        print(f"\n[DETECTED] {req_id} | Người hỏi: {meta['senderName']} ({meta['senderEmail']})")
                        print(f"  -> Ý định:   TON_THAP ({meta['intentReason']})")
                        print(f"  -> Model:    {model_name} (Xưởng: {meta.get('shop', 'BODY')})")
                        print(f"  -> Plants:   {meta.get('plants', ['3093', '3001'])} | Top={meta.get('limit', 20)}")
                        print(f"  -> Mã quét:  {len(resolved_mats)} mã tìm thấy cho model {model_name}")
                    elif cmd in ["CHECK_IB", "TIM_IB", "CHECK_PO", "TIM_GR_DOC"] and not raw_mats:
                        resolved_mats = []
                        print(f"\n[DETECTED] {req_id} | Người hỏi: {meta['senderName']} ({meta['senderEmail']})")
                        print(f"  -> Ý định:   {cmd} ({meta['intentReason']})")
                        print(f"  -> Tham số:  Inbound='{meta.get('inboundNo', '')}' | PO='{meta.get('poNo', '')}' | Ref='{meta.get('refNo', '')}' | Plant='{plant}'")
                    elif cmd != "PING":
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
                            headers = {
                                "product": "Mã hàng", "batch": "Batch", "bin": "Vị trí (Bin)",
                                "totalQty": "Tổng tồn", "availQty": "Khả dụng",
                                "unit": "ĐVT", "grDate": "Ngày GR", "grTime": "Giờ GR",
                                "cat": "Trạng thái"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Vị trí lưu kho & Giờ GR (/SCWM/POST) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        elif cmd in ["TON_KHO", "TONKHO"]:
                            wh_slocs = config.get("warehouse_slocs", ["1005", "3093"])
                            records = sap.check_stock_mb52(resolved_mats, plant, slocs=wh_slocs)
                            headers = {
                                "product": "Mã hàng", "plant": "Plant", "sloc": "Sloc",
                                "batch": "Batch", "unrest": "Tồn", "unit": "ĐVT"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Tồn kho (Kho 1005 & 3093 - Plant {plant}) cho mã: {', '.join(resolved_mats)}{mapping_info}"

                        elif cmd in ["TON", "TỒN", "MB52"]:
                            records = sap.check_stock_mb52(resolved_mats, plant)
                            headers = {
                                "product": "Mã hàng", "plant": "Plant", "sloc": "Sloc",
                                "batch": "Batch", "unrest": "Tồn", "unit": "ĐVT"
                            }
                            mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""
                            summary = f"Tồn tổng hợp (Toàn bộ SLoc - Plant {plant}) cho mã: {', '.join(resolved_mats)}{mapping_info}"

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

                        elif cmd in ["TON_THAP", "LOW_STOCK"]:
                            model_name = meta.get("model", "LIMO")
                            shop_name = meta.get("shop", "BODY")
                            target_plants = meta.get("plants", ["3093", "3001"] if shop_name == "BODY" else ["3093", "2001", "3001"])
                            limit_cnt = meta.get("limit", 20)
                            records = sap.check_low_stock(resolved_mats, target_plants, limit=limit_cnt, model_name=model_name, shop_name=shop_name)
                            if shop_name == "BODY":
                                headers = {
                                    "product": "Material",
                                    "us": "US",
                                    "line_body": "Line Body",
                                    "bd_kho": "BD Kho",
                                    "bd_carset": "BD Carset",
                                    "transit": "Transit",
                                    "kho_ma": "Tồn MA",
                                    "line_ma": "Tồn Line",
                                    "model": "Model",
                                    "station": "Trạm",
                                    "status": "Cảnh báo"
                                }
                            elif shop_name == "XMD":
                                headers = {
                                    "product": "Mã hàng", "station": "Trạm", "model": "Model",
                                    "stock_3093": "Kho 3093", "stock_2001": "XMĐ 2001", "stock_3001": "Sơn 3001",
                                    "total_stock": "Tổng tồn", "unit": "ĐVT", "status": "Cảnh báo"
                                }
                            else:
                                headers = {
                                    "product": "Mã hàng", "station": "Trạm", "model": "Model",
                                    "stock_3093": "Kho FG 3093", "stock_3001": "Line Body 3001",
                                    "total_stock": "Tổng tồn", "unit": "ĐVT", "status": "Cảnh báo"
                                }
                            summary = f"Cảnh báo tồn kho Body - Model {model_name} ({len(records)} mã Ưu tiên xuất theo chuẩn Stock_FG_MA_BODY)"

                        elif cmd in ["CHECK_OB", "VL06O"]:
                            ob_no = meta.get("obNo", "")
                            records = sap.check_open_ob_vl06o(resolved_mats, ob_no=ob_no, plant=plant)
                            headers = {
                                "obNo": "Số OB", "product": "Mã hàng", "qty": "SL xuất", "unit": "ĐVT",
                                "sloc": "Sloc", "batch": "Batch", "shipTo": "Khách nhận", "createDate": "Ngày tạo"
                            }
                            target_label = ob_no or ', '.join(resolved_mats) or 'Tất cả'
                            summary = f"Danh sách đơn Outbound (OB) đang mở (VL06O) cho: {target_label}"

                        elif cmd in ["LICHSU_BIN", "MON", "WHT"]:
                            bin_name = meta.get("bin", "")
                            records = sap.check_bin_history_mon(resolved_mats, bin_name=bin_name, days=days)
                            headers = {
                                "wtNo": "Số WT", "product": "Mã hàng", "batch": "Số lô",
                                "srcBin": "Bin xuất", "destBin": "Bin đích",
                                "qty": "SL chuyển", "unit": "ĐVT", "status": "Trạng thái",
                                "createDate": "Ngày tạo", "user": "Người xử lý"
                            }
                            target_label = f"Bin {bin_name}" if bin_name else f"Mã {', '.join(resolved_mats)}"
                            summary = f"Lịch sử chuyển Bin-to-Bin & Cancel WHT (/SCWM/MON) cho: {target_label}"

                        elif cmd in ["TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR"]:
                            inbound_no = meta.get("inboundNo", "")
                            ref_no = meta.get("refNo", "")
                            mat_doc = meta.get("materialDocument", "")
                            records = sap.check_gr_doc_mb51(inbound_no=inbound_no, ref_no=ref_no, mat_doc=mat_doc, materials=resolved_mats, plant=plant, days=days)
                            headers = {
                                "mblnr": "Số CT GR", "mjahr": "Năm", "item": "Item", "bwart": "Loại GD",
                                "matnr": "Mã hàng", "qty": "SL", "unit": "ĐVT",
                                "xblnr": "Inbound / Ref", "postDate": "Ngày HT", "sloc": "Kho", "batch": "Số lô", "user": "Người tạo"
                            }
                            target_lbl = inbound_no or ref_no or mat_doc or (', '.join(resolved_mats) if resolved_mats else 'Tất cả')
                            summary = f"Chứng từ nhập kho GR (MB51) cho: {target_lbl}"

                        elif cmd in ["CHECK_IB", "IB_DETAIL", "KIEMTRA_IB"]:
                            inbound_no = meta.get("inboundNo", "") or (raw_mats[0] if raw_mats and raw_mats[0].startswith("18") else "")
                            expected_qty = meta.get("expectedQty", 0.0)
                            records = sap.check_inbound_detail_vl06i(inbound_no=inbound_no, materials=resolved_mats, plant=plant, expected_qty=expected_qty)
                            headers = {
                                "ibNo": "Số IB", "item": "Item", "product": "Mã hàng",
                                "actualQty": "SL Inbound", "expectedQty": "SL Yêu cầu", "delta": "Chênh lệch", "qtyCheck": "Đối chiếu SL",
                                "unit": "ĐVT", "sloc": "Kho", "status": "Trạng thái GR",
                                "refPo": "PO tham chiếu", "poItem": "PO Item"
                            }
                            summary = f"Chi tiết đơn Inbound {inbound_no or ', '.join(resolved_mats)} (VL06I Item View)"

                        elif cmd in ["TIM_IB", "FIND_IB"]:
                            po_no = meta.get("poNo", "")
                            ref_no = meta.get("refNo", "")
                            records = sap.find_inbound_vl06i(po_no=po_no, ref_no=ref_no, materials=resolved_mats, plant=plant)
                            headers = {
                                "ibNo": "Số IB", "item": "Item", "product": "Mã hàng",
                                "qty": "SL giao", "unit": "ĐVT", "sloc": "Kho",
                                "status": "Trạng thái", "refPo": "Số PO", "refNo": "Vận đơn / Bill"
                            }
                            target_lbl = po_no or ref_no or (', '.join(resolved_mats) if resolved_mats else 'Tất cả')
                            summary = f"Tìm kiếm Inbound theo tham chiếu: {target_lbl}"

                        elif cmd in ["CHECK_PO", "PO_STATUS", "KIEMTRA_PO"]:
                            po_no = meta.get("poNo", "") or (raw_mats[0] if raw_mats and raw_mats[0].startswith(("45", "55")) else "")
                            po_item = meta.get("poItem", "")
                            records = sap.check_po_me2n(po_no=po_no, materials=resolved_mats, plant=plant, po_item=po_item)
                            headers = {
                                "poNo": "Số PO", "item": "Item", "product": "Mã hàng",
                                "orderQty": "SL Đặt", "deliveredQty": "Đã nhận", "openQty": "Còn nợ",
                                "unit": "ĐVT", "plant": "Plant", "deleteInd": "Xóa/Khóa", "releaseStatus": "Duyệt"
                            }
                            summary = f"Tình trạng đơn đặt hàng PO {po_no or ', '.join(resolved_mats)} (ME2N)"

                        elif cmd in ["CHECK_FG_OB", "DIEN_OB"]:
                            records = sap.check_fg_ob_status(resolved_mats, plant=plant)
                            headers = {
                                "product": "Mã hàng", "fgStock": "Tồn FG 3093",
                                "openObList": "Đơn OB đang mở", "openObQty": "SL Đang giữ",
                                "availableQty": "Khả dụng xuất", "unit": "ĐVT",
                                "workAdvice": "Đánh giá công việc"
                            }
                            summary = f"Điều độ Thành Phẩm (FG) & Đơn OB cho: {', '.join(resolved_mats)}"

                        elif cmd in ["CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL", "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING"]:
                            records = sap.check_nvl_migo_balance(resolved_mats, plant=plant or "3093", batch=batch)
                            headers = {
                                "product": "Mã hàng", "plant": "Plant", "sloc": "Kho",
                                "bin": "Vị trí Bin", "batch": "Số lô",
                                "qty": "SL Khả dụng", "unit": "ĐVT",
                                "grDate": "Ngày GR", "grTime": "Giờ GR",
                                "cat": "Trạng thái", "fifoRank": "Thứ tự FIFO",
                                "migoAdvice": "Khuyến nghị ra đơn MIGO 311"
                            }
                            summary = f"Posting Change NVL (Plant {plant or '3093'} - Kho 1005) cho: {', '.join(resolved_mats)}"

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
