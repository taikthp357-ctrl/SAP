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
import re
import time
import json
import gc
import html
from pathlib import Path
from datetime import datetime, timedelta
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
from app.runtime_control import WorkerLease, data_directory
from app.number_utils import set_global_sap_number_format
from app.sap_attachment import MigoAttachmentRunner
from app.bbbg_service import BbbgService
from app.nvl_service import NvlService
from app.ob_service import ObService
from app.extended_query_service import ExtendedQueryService
from app.core.context import RequestContext
from app.router import router


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


def complete_mail_item(mail, bot_folder, status: str = "processed", outlook_ns=None):
    """
    Đánh dấu mail đã xử lý và di chuyển an toàn vào thư mục con (Processed hoặc Failed).
    Bọc thép chống văng lỗi COM 'Property Item.UnRead can not be set' hoặc object disconnected.
    """
    if mail is None:
        return

    # 1. Thử đánh dấu đã đọc an toàn
    try:
        mail.UnRead = False
        mail.Save()
    except Exception as read_err:
        entry_id = getattr(mail, "EntryID", None)
        recovered = False
        if entry_id and outlook_ns:
            try:
                fresh = outlook_ns.GetItemFromID(entry_id)
                fresh.UnRead = False
                fresh.Save()
                mail = fresh
                recovered = True
            except Exception:
                pass
        if not recovered:
            print(f"[OUTLOOK WARN] Bỏ qua lỗi cập nhật UnRead ({status}): {read_err}")

    # 2. Di chuyển vào thư mục con tương ứng
    try:
        if str(status).lower() == "processed":
            move_mail_to_processed(mail, bot_folder)
        else:
            move_mail_to_failed(mail, bot_folder)
    except Exception as move_err:
        print(f"[OUTLOOK WARN] Bỏ qua lỗi di chuyển thư sang {status}: {move_err}")


def reconnect_outlook_session(folder_name: str, max_retries: int = 3, retry_delay: float = 2.0):
    """
    Tái tạo kết nối COM Outlook sạch hoàn toàn (Clean Apartment Teardown).
    Chống hiện tượng rò rỉ hoặc truy cập con trỏ COM RPC chết dẫn đến lỗi ngắt Access Violation (0xC0000005).
    """
    print("[OUTLOOK RECONNECT] Phát hiện mất kết nối COM Outlook. Đang kết nối lại...")
    app = None
    ns = None
    inbox = None
    bot_f = None

    for attempt in range(1, max_retries + 1):
        try:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            pythoncom.CoInitialize()

            app = win32com.client.Dispatch("Outlook.Application")
            ns = app.GetNamespace("MAPI")
            inbox = ns.GetDefaultFolder(6)  # 6 = olFolderInbox

            # 1. Quét tìm trong Inbox
            try:
                for f in inbox.Folders:
                    if f.Name.strip().lower() == folder_name.strip().lower():
                        bot_f = f
                        break
            except Exception:
                pass

            # 2. Quét tìm trong Store root (Parent của Inbox)
            if bot_f is None:
                try:
                    for f in inbox.Parent.Folders:
                        if f.Name.strip().lower() == folder_name.strip().lower():
                            bot_f = f
                            break
                except Exception:
                    pass

            if bot_f is not None:
                print(f"[OUTLOOK RECONNECT] Đã kết nối lại Outlook thành công (thư mục: '{getattr(bot_f, 'Name', '')}').")
                return app, ns, inbox, bot_f
            else:
                print(f"[OUTLOOK RECONNECT WARN] Đã kết nối MAPI (lần {attempt}/{max_retries}) nhưng chưa tìm thấy thư mục '{folder_name}'.")
        except Exception as recon_err:
            print(f"[OUTLOOK RECONNECT WARN] Thử kết nối lại lần {attempt}/{max_retries} thất bại: {recon_err}")

        time.sleep(retry_delay)

    return app, ns, inbox, bot_f


def deliver_outbox(state_mgr, max_attempts=3):
    for item in state_mgr.get_pending_outbox(max_attempts=max_attempts):
        token = state_mgr.begin_delivery(item["req_id"])
        if token:
            outcome = TeamsDispatcher.deliver_payload(item["webhook_url"], item["payload"])
            state_mgr.finish_delivery(item["req_id"], token, outcome["status"], outcome["detail"], max_attempts=max_attempts)
            detail_str = f" - {outcome['detail']}" if outcome.get("detail") else ""
            print(f"[OUTBOX] {item['req_id']} | {outcome['status']}{detail_str}")


def actual_sender_address(mail):
    address = str(getattr(mail, "SenderEmailAddress", "") or "").strip().lower()
    if str(getattr(mail, "SenderEmailType", "")).upper() == "EX":
        try:
            address = str(mail.Sender.GetExchangeUser().PrimarySmtpAddress).strip().lower()
        except Exception:
            return ""
    return address


def intake_subject(subject):
    return bool(re.match(r"^\s*(?:\[SAP-REQUEST\]|SAP_REQUEST)(?:\s|$|:)", str(subject), re.I))


def run_bridge():
    """Hàm chạy chính của tiến trình Worker."""
    config = load_config()

    data_dir = data_directory(config, CURRENT_DIR)
    lease = WorkerLease(data_dir).acquire()
    state_mgr = StateManager(data_dir / "bridge_state.db")
    state_mgr.recover_interrupted()
    max_retries = config.get("max_webhook_retries", 3)
    # Delivery is independent of SAP and Outlook startup.
    deliver_outbox(state_mgr, max_retries)

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

    resolver = MaterialResolver()
    sap = SapEngine(
        max_alv_rows=max_rows,
        keep_alive_idle_seconds=keep_alive_idle,
        enable_keep_alive=config.get("enable_sap_keep_alive", True),
        connection_index=config.get("sap_connection_index", 0),
        session_index=config.get("sap_session_index", 0),
        strict_session_safety=config.get("strict_session_safety", True),
        auto_create_session=config.get("auto_create_session", True),
        prefer_dedicated_session=config.get("prefer_dedicated_session", True),
        user_reserved_sessions=config.get("user_reserved_sessions", 3),
        max_sap_sessions=config.get("max_sap_sessions", 6),
        max_batch_tabs=config.get("max_batch_tabs", 3),
        expected_system=config.get("sap_system", ""),
        expected_client=config.get("sap_client", ""),
        expected_user=config.get("sap_user", "")
    )
    print(f"[READY] Đang lắng nghe thư mới trong '{bot_folder.Name}'...")

    allowed_senders = [s.strip().lower() for s in config.get("allowed_senders", []) if s.strip()]
    last_sync_time = 0.0

    while not lease.should_stop():
        try:
            # Ép Outlook Classic gọi lệnh đồng bộ ngầm (Sync Send/Receive) với Server (mỗi 10s một lần, tránh nghẽn MAPI)
            now_ts = time.time()
            if now_ts - last_sync_time >= 10.0:
                try:
                    outlook.SendAndReceive(False)
                    last_sync_time = now_ts
                except Exception:
                    pass

            # Deliver durable results before touching SAP or requesting new work.
            deliver_outbox(state_mgr, max_retries)
            sap.keep_alive()

            # 3. Quét tất cả các thư trong thư mục SAP_BOT (kể cả thư đã mở xem lẫn thư chưa đọc)
            # Mọi thư đã xử lý xong đều được chuyển tự động sang thư mục con 'Processed'.
            # Do đó, bất kỳ thư nào còn nằm trong 'SAP_BOT' đều là yêu cầu đang chờ xử lý!
            if bot_folder is None:
                raise RuntimeError("Outlook bot_folder is not connected (-2147220995)")
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
                    unread_inbox = inbox.Items.Restrict('@SQL="urn:schemas:httpmail:read" = 0 AND ("urn:schemas:httpmail:subject" LIKE \'[SAP-REQUEST]%\' OR "urn:schemas:httpmail:subject" LIKE \'SAP_REQUEST%\')')
                    u_count = unread_inbox.Count
                    for i in range(1, min(u_count + 1, 31)):
                        try:
                            item = unread_inbox.Item(i)
                            subj = str(getattr(item, "Subject", "") or "")
                            if intake_subject(subj):
                                mail_candidates.append(item)
                        except Exception:
                            pass
                except Exception:
                    pass

            for mail in mail_candidates:
                if lease.should_stop():
                    break
                entry_id = getattr(mail, "EntryID", "")
                if entry_id and state_mgr.is_message_processed("", entry_id):
                    complete_mail_item(mail, bot_folder, "processed", outlook_ns=outlook)
                    continue

                req_id = StateManager.request_id(str(data_dir), "", entry_id) if entry_id else "UNIDENTIFIED"

                # Cô lập hoàn toàn lỗi từng thư để không chặn luồng
                try:
                    # 1. Kiểm tra danh sách người gửi được phép (nếu có cấu hình)
                    if allowed_senders:
                        actual_sender = actual_sender_address(mail)
                        if actual_sender not in allowed_senders:
                            print(f"[SECURITY] {req_id} | Người gửi '{actual_sender}' không thuộc allowed_senders. Chuyển thư sang Failed.")
                            complete_mail_item(mail, bot_folder, "failed", outlook_ns=outlook)
                            continue

                    raw_body = getattr(mail, "Body", "")
                    html_body = getattr(mail, "HTMLBody", "")
                    subject = getattr(mail, "Subject", "")

                    meta = NaturalCommandParser.parse(raw_body, html_body=html_body, subject=subject, default_plant=def_plant)

                    # LOOP GUARD: Nếu phát hiện tin nhắn Bot tự trả lời -> Bỏ qua ngay lập tức
                    if meta and meta.get("isBotEcho"):
                        print(f"[LOOP GUARD] {req_id} | Phát hiện tin nhắn phản hồi của Bot. Bỏ qua ngay lập tức để chống lặp!")
                        complete_mail_item(mail, bot_folder, "processed", outlook_ns=outlook)
                        continue

                    # Kiểm tra trùng lặp qua SQLite
                    msg_id = meta.get("messageId", "") if meta else ""
                    if (msg_id or entry_id) and state_mgr.is_message_processed(msg_id, entry_id):
                        print(f"[DUPLICATE] {req_id} | Message ID '{msg_id or (entry_id[:12] if entry_id else '')}' đã xử lý thành công trước đó. Bỏ qua.")
                        complete_mail_item(mail, bot_folder, "processed", outlook_ns=outlook)
                        continue

                    if not meta or not meta.get("valid"):
                        reason = meta.get("reason", "Không rõ") if meta else "Thư rỗng hoặc không liên quan"
                        print(f"[SKIP] {req_id} | Bỏ qua yêu cầu không có mã hợp lệ: {reason}. Không gửi trợ giúp lên Teams.")
                        state_mgr.mark_message_processed(msg_id, entry_id, req_id, "SKIP", "INVALID_INPUT")
                        complete_mail_item(mail, bot_folder, "processed", outlook_ns=outlook)
                        continue

                    store_id = str(getattr(getattr(mail, "Parent", None), "StoreID", "") or data_dir)
                    req_id = StateManager.request_id(store_id, msg_id, entry_id)
                    if not state_mgr.claim_request(req_id):
                        claim_status = state_mgr.get_claim_status(req_id)
                        has_outbox = state_mgr.has_outbox_result(req_id)
                        print(f"[RECOVERY] {req_id}: kết quả đã lưu hoặc cần kiểm tra lần chạy gián đoạn (claim: {claim_status}, outbox: {has_outbox}); không chạy SAP lại.")
                        rec_status = "processed" if (claim_status == "COMPLETED" or has_outbox) else "failed"
                        complete_mail_item(mail, bot_folder, rec_status, outlook_ns=outlook)
                        print(f"[RECOVERY] {req_id}: Đã chuyển thư sang thư mục '{rec_status.capitalize()}'.")
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
                    elif cmd not in ("PING", "WHO_IS_TAI"):
                        resolved_mats, notes = resolver.resolve_list(raw_mats)
                        print(f"\n[DETECTED] {req_id} | Người hỏi: {meta['senderName']} ({meta['senderEmail']})")
                        print(f"  -> Ý định:   {cmd} ({meta['intentReason']})")
                        print(f"  -> Mã gốc:   {raw_mats} ──> Mã SAP: {resolved_mats}")
                        print(f"  -> Tham số:  Batch='{batch or 'ALL'}' | Plant='{plant}' | Số ngày={days}")
                    else:
                        resolved_mats = []
                        print(f"\n[DETECTED] {req_id} | Lệnh PING kiểm tra kết nối từ {meta['senderName']}")

                    try:
                        # Chuẩn bị Context dữ liệu chuẩn hóa và điều phối qua Command Router
                        meta["actualSenderEmail"] = actual_sender_address(mail)
                        req_ctx = RequestContext(
                            request_id=req_id,
                            message_id=msg_id,
                            command=cmd,
                            meta=meta,
                            sap=sap,
                            resolver=resolver,
                            state_mgr=state_mgr,
                            config=config,
                            data_dir=data_dir,
                            raw_materials=raw_mats,
                            resolved_materials=resolved_mats,
                            plant=plant,
                            batch=batch,
                            days=days,
                        )
                        cmd_res = router.dispatch(req_ctx)
                        records = cmd_res.records
                        headers = cmd_res.headers
                        summary = cmd_res.summary
                        custom_html = cmd_res.custom_html
                        result_status = cmd_res.status

                        if cmd_res.audit_message:
                            state_mgr.log_audit(
                                req_id,
                                cmd,
                                result_status,
                                target_doc=cmd_res.audit_target,
                                user_name=meta.get("senderName"),
                                message=cmd_res.audit_message,
                            )

                        # Chuẩn bị payload chuẩn (hỗ trợ tự động phân tách 2, 3, 4 tin nhắn nếu danh sách dài) và đưa vào Outbox
                        multipart_payloads = TeamsDispatcher.prepare_multipart_payloads(
                            req_id, meta, result_status, summary, records, headers, custom_html=custom_html, chunk_size=15
                        )
                        for p_idx, (p_load, p_data) in enumerate(multipart_payloads):
                            p_req_id = f"{req_id}#p{p_idx+1}" if len(multipart_payloads) > 1 else req_id
                            if p_idx == 0:
                                state_mgr.complete_request(msg_id, entry_id, p_req_id, cmd, result_status, webhook_url, p_load)
                            else:
                                state_mgr.enqueue_outbox(p_req_id, webhook_url, p_load)
                        deliver_outbox(state_mgr, max_retries)
                        complete_mail_item(mail, bot_folder, "processed", outlook_ns=outlook)

                    except Exception as sap_err:
                        print(f"[SAP ERROR] {req_id} | {sap_err}")
                        is_att_cmd = (cmd or "").upper() in ["ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO"]
                        cmd_u = (cmd or "").upper()
                        is_att_cmd = cmd_u in ["ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO"]
                        is_check_att = cmd_u in ["CHECK_ATT", "CHECK_ATT_DOC", "CHECK_DINH_KEM", "ZFIR266"]
                        is_nvl_tracking = cmd_u in ["TRA_CUU_CAP_NVL", "TINH_TRANG_NVL", "RA_DON_NVL", "SET_FOLDER", "SET_NVL_PASSWORD"]
                        err_str = str(sap_err)
                        if is_nvl_tracking or any(k in err_str.lower() for k in ["pyxlsb", "openpyxl", "msoffcrypto", "excel", "sheet", "theo doi oder", "file ca"]):
                            err_title = "Lỗi xử lý File Ca / Theo dõi NVL Excel"
                        elif is_att_cmd:
                            err_title = "Lỗi khi đính kèm / điền HeadText MIGO"
                        elif is_check_att or "zfir266" in err_str.lower():
                            err_title = "Lỗi khi tra cứu đính kèm ZFIR266"
                        else:
                            err_title = "Lỗi khi truy xuất SAP GUI"
                        err_html = f"<p style='color:#d9534f;font-weight:bold;'>❌ {err_title}: {html.escape(str(sap_err))}</p>"
                        fail_summary = f"{err_title}: {sap_err}"
                        fail_payload, fail_data = TeamsDispatcher.prepare_safe_payload(
                            req_id, meta, "FAILED", fail_summary, [], {}, custom_html=err_html
                        )
                        if not state_mgr.is_message_processed(msg_id, entry_id):
                            state_mgr.complete_request(msg_id, entry_id, req_id, cmd, "FAILED", webhook_url, fail_payload)
                            deliver_outbox(state_mgr, max_retries)

                        complete_mail_item(mail, bot_folder, "failed", outlook_ns=outlook)

                except Exception as mail_err:
                    print(f"[MAIL ERROR] {req_id} | Lỗi khi xử lý thư: {mail_err}")
                    complete_mail_item(mail, bot_folder, "failed", outlook_ns=outlook)

        except Exception as loop_err:
            print(f"[LOOP WARN] {loop_err}")
            err_str = str(loop_err)
            if (
                bot_folder is None
                or "-2147220995" in err_str
                or "not connected to server" in err_str.lower()
                or "disconnected" in err_str.lower()
                or "rpc" in err_str.lower()
                or "nonetype" in err_str.lower()
                or "-2147467260" in err_str
                or "operation aborted" in err_str.lower()
            ):
                outlook_app = None
                outlook_ns = None
                outlook = None
                inbox = None
                bot_folder = None
                new_app, new_ns, new_inbox, new_bot_f = reconnect_outlook_session(folder_name)
                if new_app and new_ns and new_inbox:
                    outlook_app = new_app
                    outlook_ns = new_ns
                    outlook = new_ns
                    inbox = new_inbox
                    bot_folder = new_bot_f

        # Tối ưu giải phóng RAM sau mỗi chu kỳ; loại bỏ CoFreeUnusedLibraries để chống crash combase.dll
        try:
            import gc
            gc.collect()
        except Exception:
            pass

        for _ in range(max(1, int(poll_sec))):
            if lease.should_stop():
                break
            time.sleep(1)
    lease.close()
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    if "--stop" in sys.argv:
        requested = WorkerLease.request_stop(data_directory(load_config(), CURRENT_DIR))
        print("Đã yêu cầu dừng sau tác vụ hiện tại." if requested else "Chưa tìm thấy worker trong thư mục dữ liệu.")
    elif "--test-parser" in sys.argv:
        run_self_test()
    else:
        while True:
            try:
                run_bridge()
                break
            except KeyboardInterrupt:
                print("\n[STOP] Đã nhận tín hiệu dừng từ người dùng.")
                break
            except Exception as fatal_e:
                print(f"\n[FATAL SHIELD] Tiến trình worker gặp lỗi ngoại cảnh: {fatal_e}")
                print("[FATAL SHIELD] Đang tự động phục hồi trong 5 giây...")
                time.sleep(5)
