# -*- coding: utf-8 -*-
"""Bộ phân tích ngôn ngữ tự nhiên thông minh (Smart NLU Parser V2.0).

Tính năng:
1. Chấm điểm ý định theo trọng số (Fuzzy & Context Scoring) cho VITRI, LICHSU, TON, HELP.
2. Làm sạch email Mail V3, cắt bỏ Footer quảng cáo của microsoft@powerapps.com.
3. Bóc tách Metadata (MSG_ID, SENDER_EMAIL, SENDER_NAME, RAW_TEXT).
4. Nhận diện tham số: Plant, Batch, Số ngày, Danh sách Mã vật tư.
5. Hỗ trợ Smart Coaching: Tự động trả về needHelpReply nếu người dùng thiếu mã hàng hoặc hỏi cú pháp.
6. Mặc định thông minh: Nếu chỉ gửi mã hàng mà không gõ lệnh -> Tự động chọn VITRI (/SCWM/POST).
"""

import re
import unicodedata
from typing import Optional, Dict, Any, List


class SmartCommandParser:
    """Bộ bóc tách ý định câu lệnh và thực thể từ Email V3 của Power Automate."""

    INTENT_SCORES = {
        "TON_THAP": [
            (r"\b(?:ton\s*thap|thap|canh\s*bao\s*ton|low\s*stock|doh\s*thap|sap\s*het|thieu\s*hang)\b", 16),
            (r"\b(?:top\s*\d+\s*(?:ma)?\s*ton\s*thap|check\s*(?:cho\s*toi\s*)?\d+\s*ma\s*.*ton\s*thap)\b", 22),
            (r"\b(?:ton\s*thap\s*(?:body|xmd|limo|vf\d+))\b", 20),
        ],
        "CHECK_BOM": [
            (r"\b(?:check\s*bom|kiem\s*tra\s*bom|co\s*bom|bom\s*co\s*khong|co\s*bom\s*khong|dinh\s*muc|cs03)\b", 22),
            (r"\b(?:linh\s*kien\s*con|ma\s*con|thanh\s*phan|bom|subcon)\b", 18),
        ],
        "CHECK_QUEUE": [
            (r"\b(?:check\s*queue|kiem\s*tra\s*queue|ket\s*queue|loi\s*queue|queue\s*ewm|mqueue|scwm/mon\s*queue)\b", 22),
            (r"\b(?:queue|mqueue0001)\b", 18),
        ],
        "CHECK_OB": [
            (r"\b(?:check\s*(?:con\s*)?ob|kiem\s*tra\s*(?:con\s*)?ob|kt\s*ob|con\s*ob|ob\s*co\s*khong|co\s*ob|ob\s*treo|vl06o|don\s*xuat|outbound)\b", 18),
            (r"\b(?:ob\s*nao\s*dang\s*cho|cho\s*xuat|ob\s*mo|ob\s*chua\s*xuat|ob\s*ma|con\s*ob\s*ma)\b", 16),
            (r"\b(?:check\s*don\s*ob|xem\s*ob)\b", 12),
        ],
        "LICHSU_BIN": [
            (r"\b(?:lich\s*su\s*bin|lichsu\s*bin|bin\s*to\s*bin|bintobin|chuyen\s*bin|cancel\s*wht|wht|scwm/mon|/scwm/mon)\b", 18),
            (r"\b(?:ls\s*bin|vet\s*bin|di\s*chuyen\s*bin|lich\s*su\s*dich\s*chuyen)\b", 14),
        ],
        "TIM_GR_DOC": [
            (r"\b(?:tim\s*gr|chung\s*tu\s*gr|gr\s*doc|mb51\s*gr|lich\s*su\s*gr|so\s*gr|tra\s*gr|chung\s*tu\s*nhap)\b", 22),
            (r"\b(?:tra\s*cuu\s*gr|tim\s*chung\s*tu|xem\s*gr|chung\s*tu\s*mb51)\b", 18),
        ],
        "CHECK_IB": [
            (r"\b(?:chi\s*tiet\s*ib|chi\s*tiet\s*inbound|ib\s*detail|check\s*inbound\s*detail|doi\s*chieu\s*ib|so\s*khop\s*ib|kiem\s*tra\s*inbound)\b", 22),
            (r"\b(?:check\s*ib\s*18\d{6,8}|inbound\s*18\d{6,8}|ib\s*18\d{6,8})\b", 22),
        ],
        "TIM_IB": [
            (r"\b(?:tim\s*ib|tim\s*inbound|find\s*ib|find\s*inbound|inbound\s*theo\s*po|inbound\s*theo\s*bill)\b", 22),
        ],
        "CHECK_FG_OB": [
            (r"\b(?:dien\s*ob|di\s*en\s*ob|check\s*ob\s*fg|kha\s*dung\s*ob|ob\s*fg|ton\s*ob)\b", 24),
        ],
        "CHECK_NVL_MIGO": [
            (r"\b(?:posting\s*change|postingchange|check\s*posting|tai\s*posting|ra\s*don\s*migo|migo\s*311|migo|cap\s*migo|chuyen\s*kho\s*311|can\s*doi\s*nvl|check\s*nvl|kiem\s*tra\s*nvl|nvl\s*migo|nvl\s*1005|kho\s*1005)\b", 24),
        ],
        "ATT_MIGO": [
            (r"\b(?:dien\s*so\s*phieu|dien\s*phieu|so\s*phieu|headtext|dien\s*bktxt|bktxt)\b", 35),
            (r"\b(?:att\s*mail|attach\s*mail|att\s*email|attach\s*email|dinh\s*kem\s*mail|dinh\s*kem\s*email)\b", 35),
            (r"\b(?:att\s*vat|att\s*bbbg|dinh\s*kem\s*vat|dinh\s*kem\s*bbbg|migo\s*att|att\s*migo|dinh\s*kem\s*migo)\b", 28),
            (r"\b(?:att\s*file|dinh\s*kem\s*file|gos\s*att|att\s*chung\s*tu)\b", 22),
            (r"\b(?:att\s*vat\s*bbbg|att\s*bbbg\s*vat)\b", 30),
        ],
        "CHECK_PO": [
            (r"\b(?:check\s*po|kiem\s*tra\s*po|tinh\s*trang\s*po|po\s*status|don\s*dat\s*hang|po\s*con\s*no|po\s*da\s*giao)\b", 20),
            (r"\b(?:po\s*45\d{8}|po\s*55\d{8}|me2n\s*po)\b", 22),
        ],
        "VITRI": [
            (r"\b(?:vi\s*tri|vitri|vtri|bin|posting|scwm/post|scwm)\b", 10),
            (r"\b(?:ngay\s*gr|gio\s*gr|check\s*gr|kiem\s*tra\s*gr)\b", 10),
            (r"\b(?:tim\s*hang|o\s*dau|ke\s*nao|nam\s*o|khu\s*vuc|zone)\b", 8),
            (r"\b(?:check\s*bin|kiem\s*tra\s*bin|bin\s*nao|o\s*ke|tra\s*bin)\b", 12),
            (r"\b(?:post|gr)\b", 5),
        ],
        "LICHSU": [
            (r"\b(?:lich\s*su|lichsu|lsu|ls|mb51)\b", 10),
            (r"\b(?:xuat\s*nhap|nhap\s*xuat|giao\s*dich|chung\s*tu|movement|mvt)\b", 9),
            (r"\b(?:da\s*xuat|da\s*nhap|bi\s*tru|tru\s*ton)\b", 7),
        ],
        "TON": [
            (r"\b(?:ton\s*kho|tonkho|check\s*ton|kt\s*ton|kiem\s*tra\s*ton|mb52)\b", 10),
            (r"\b(?:so\s*luong\s*ton|con\s*ton|con\s*hang|stock)\b", 8),
            (r"\b(?:ton)\b", 6),
        ],
        "IB_TREO": [
            (r"\b(?:ib\s*treo|ibtreo|don\s*treo|inbound\s*treo|vl06i|check\s*ib|inbound)\b", 12),
            (r"\b(?:chua\s*gr|chua\s*nhap|hang\s*chua\s*ve|cho\s*nhap|treo)\b", 9),
            (r"\b(?:ib)\b", 6),
        ],
        "STO": [
            (r"\b(?:sto|stock\s*transport|me2n|check\s*sto|dieu\s*chuyen)\b", 12),
            (r"\b(?:hang\s*di\s*duong|di\s*duong|mb5t)\b", 9),
            (r"\b(?:po\s*dieu\s*chuyen|don\s*dieu\s*chuyen)\b", 8),
        ],
        "HELP": [
            (r"\b(?:huong\s*dan|tro\s*giup|help|cu\s*phap|cach\s*dung|menu)\b", 15),
        ]
    }

    STOP_WORDS = {
        "CHECK", "CHEK", "KIEM", "TRA", "KIEMTRA", "KT", "VI", "TRI", "VITRI", "VTRI",
        "TON", "KHO", "TONKHO", "LICH", "SU", "LICHSU", "LSU", "BATCH", "PLANT", "WERKS",
        "NGAY", "DAYS", "DAY", "HOM", "MA", "HANG", "MAHANG", "SAN", "PHAM", "BIN",
        "MB51", "MB52", "POST", "SCWM", "LENH", "GIUP", "VOI", "ANH", "CHI", "EM", "OI",
        "BOT", "SAP", "XEM", "TIM", "CHO", "HOI", "XIN", "BAO", "NHIEU", "XUAT", "NHAP",
        "GIAO", "DICH", "LO", "CHARG", "GR", "DATE", "TIME", "STOCK", "TRONG", "VONG",
        "GAN", "NHAT", "SENDER", "EMAIL", "NAME", "MSG", "RAW", "TEXT", "CON", "KHONG",
        "NAM", "DAU", "NAO", "NHE", "NHA", "THANKS", "CAM", "ON", "DUM", "HO", "STATUS",
        "IB", "TREO", "IBTREO", "STO", "INBOUND", "VL06I", "ME2N", "MB5T", "DIEUCHUYEN", "DIEU", "CHUYEN", "DON",
        "TAI", "TÀI", "@TAI", "@TÀI", "TOI", "TÔI", "MINH", "MÌNH", "SEP", "SẾP", "NHO", "NHỜ",
        "BAN", "BẠN", "GIÚP", "VỚI", "CÓ", "ĐANG", "CHỜ", "TOP", "TOP20", "TOP10", "THAP", "THẤP",
        "NHẤT", "CỦA", "HOẶC", "HOAC", "XMD", "LIMO", "BODY", "XMĐ", "XE", "MAY", "DIEN", "OB", "WHT", "POSTING", "POSTINGCHANGE", "MIGO", "NVL", "CHANGE"
    }

    @staticmethod
    def remove_accents(text: str) -> str:
        """Chuyển tiếng Việt có dấu thành không dấu (chữ thường) để so khớp chuẩn xác."""
        if not text:
            return ""
        nfd = unicodedata.normalize("NFD", text)
        no_acc = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
        return no_acc.replace("đ", "d").replace("Đ", "D").lower()

    @classmethod
    def clean_v3_email(cls, raw_body: str, html_body: str = "") -> str:
        """Làm sạch email từ Mail V3, giữ lại cấu trúc dòng và cắt bỏ Footer quảng cáo."""
        if html_body and ("<table" in html_body.lower() or "<tr" in html_body.lower()):
            source = html_body
        else:
            source = html_body if (not raw_body or len(raw_body.strip()) < 5) else raw_body

        import html
        text = html.unescape(source)
        text = text.replace("\xa0", " ").replace("\u200b", "").replace("\u3000", " ")
        text = re.sub(r"<br\s*/?>|</p>|</div>|</tr>|</li>", "\n", text, flags=re.IGNORECASE)
        text = re.sub(r"</td>", "\t", text, flags=re.IGNORECASE)
        text = re.sub(r"<at[^>]*>.*?</at>", " ", text, flags=re.IGNORECASE | re.DOTALL)
        text = re.sub(r"<[^>]+>", " ", text)

        for marker in [
            "If you want to unsubscribe",
            "Nếu bạn muốn hủy đăng ký",
            "please use this form",
            "vui lòng sử dụng biểu mẫu này",
            "Microsoft Corporation"
        ]:
            idx = text.lower().find(marker.lower())
            if idx != -1:
                text = text[:idx]

        return text.strip()

    @classmethod
    def parse(cls, raw_body: Any, html_body: str = "", subject: str = "", default_plant: str = "3093") -> Optional[Dict[str, Any]]:
        """Phân tích nội dung email và trả về từ điển lệnh."""
        json_obj = None
        if isinstance(raw_body, dict):
            import json
            json_obj = raw_body
            cleaned = json.dumps(raw_body, ensure_ascii=False)
        else:
            cleaned = cls.clean_v3_email(raw_body, html_body)
            if not cleaned and not subject:
                return None

        # LOOP GUARD: Phát hiện và chặn đứng tin nhắn vòng lặp từ Bot tự phản hồi
        bot_echo_markers = [
            "sếp ơi, em tra được rồi",
            "sep oi, em tra duoc roi",
            "kết quả của sếp đây",
            "ket qua cua sep day",
            "vị trí lưu kho & giờ gr",
            "vi tri luu kho & gio gr",
            "mã tham chiếu:",
            "ma tham chieu:",
            "trợ lý sap:",
            "tro ly sap:",
            "kết quả tra cứu",
            "ket qua tra cuu"
        ]
        norm_check = cls.remove_accents(str(cleaned or "") + " " + str(subject or "")).lower()
        if any(m in norm_check for m in bot_echo_markers):
            return {
                "valid": False,
                "isBotEcho": True,
                "command": "LOOP_GUARD",
                "detectedIntent": "BOT_ECHO",
                "messageId": "",
                "reason": "Phát hiện tin nhắn phản hồi của Bot (Loop Guard). Bỏ qua để chống vòng lặp vô hạn."
            }

        # 1. Hỗ trợ bóc tách JSON trực tiếp từ email (cả dạng thuần JSON lẫn có marker)
        marker_m = re.search(r"SAP_REQUEST_BEGIN\s*(\{.*?\})\s*SAP_REQUEST_END", cleaned, re.DOTALL | re.IGNORECASE)
        if marker_m:
            try:
                import json
                json_obj = json.loads(marker_m.group(1))
            except Exception:
                pass

        if not json_obj:
            first_brace = cleaned.find("{")
            last_brace = cleaned.rfind("}")
            if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
                cand_str = cleaned[first_brace:last_brace+1]
                try:
                    import json
                    parsed_json = json.loads(cand_str)
                    if isinstance(parsed_json, dict) and any(k in parsed_json for k in ["command", "requestId", "schemaVersion", "materials", "RAW_TEXT"]):
                        json_obj = parsed_json
                except Exception:
                    pass

        SUPPORTED_JSON_COMMANDS = {
            "PING", "VITRI", "BIN", "POST", "SCWM", "/SCWM/POST",
            "TON", "TON_KHO", "TONKHO", "MB52",
            "LICHSU", "MB51",
            "IB_TREO", "IB", "INBOUND", "VL06I",
            "STO", "ME2N", "MB5T", "DIEUCHUYEN",
            "TON_THAP", "CHECK_OB", "LICHSU_BIN",
            "TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR",
            "CHECK_IB", "IB_DETAIL", "KIEMTRA_IB",
            "TIM_IB", "FIND_IB",
            "CHECK_BOM", "BOM", "CS03",
            "CHECK_QUEUE", "QUEUE",
            "CHECK_PO", "PO_STATUS", "KIEMTRA_PO",
            "CHECK_FG_OB", "DIEN_OB",
            "CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL", "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING",
            "ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO",
            "HELP"
        }

        if json_obj:
            cmd = str(json_obj.get("command", "")).strip().upper()
            if cmd in ["TON_KHO", "TONKHO"]:
                cmd = "TON_KHO"
            req_id = str(json_obj.get("requestId", "") or json_obj.get("messageId", "")).strip()
            plant = str(json_obj.get("plant", default_plant)).strip().upper() or default_plant
            batch = str(json_obj.get("batch", "")).strip().upper()

            days_raw = json_obj.get("days", 7)
            try:
                days = int(days_raw)
            except (ValueError, TypeError):
                days = 7
            days = max(1, min(days, 365))

            raw_mats = json_obj.get("materials", [])
            valid_mats = []
            if isinstance(raw_mats, (str, int)):
                candidates = re.split(r"[,;\s]+", str(raw_mats))
                for c in candidates:
                    c_clean = c.strip()
                    if c_clean and re.match(r"^[\w\-./]{2,40}$", c_clean):
                        valid_mats.append(c_clean)
            elif isinstance(raw_mats, list):
                for item in raw_mats:
                    if isinstance(item, (str, int)) and not isinstance(item, bool):
                        c_clean = str(item).strip()
                        if c_clean and re.match(r"^[\w\-./]{2,40}$", c_clean):
                            valid_mats.append(c_clean)
            mats = valid_mats

            raw_text = str(json_obj.get("rawText", "") or json_obj.get("RAW_TEXT", "") or json_obj.get("note", "")).strip()
            sender_email = str(json_obj.get("senderEmail", "") or json_obj.get("SENDER_EMAIL", "") or "PowerAutomate").strip()
            sender_name = str(json_obj.get("senderName", "") or json_obj.get("SENDER_NAME", "") or "Hệ thống").strip()
            msg_id = str(json_obj.get("messageId", "") or json_obj.get("MSG_ID", "") or req_id).strip()
            note = str(json_obj.get("note", "")).strip()
            dry_run = bool(json_obj.get("dryRun", False))

            if cmd and cmd not in SUPPORTED_JSON_COMMANDS:
                return {
                    "valid": False,
                    "needHelpReply": True,
                    "messageId": msg_id or req_id,
                    "senderEmail": sender_email,
                    "senderName": sender_name,
                    "detectedIntent": "UNKNOWN",
                    "reason": f"Lệnh '{cmd}' không được hỗ trợ trong hệ thống."
                }

            if cmd == "TON_THAP":
                model_name = str(json_obj.get("model", "")).strip().upper() or "LIMO"
                shop = str(json_obj.get("shop", "")).strip().upper() or ("XMD" if "XMD" in model_name else "BODY")
                plants = json_obj.get("plants", ["3093", "3001"] if shop == "BODY" else ["3093", "2001", "3001"])
                limit = int(json_obj.get("limit", 20))
                return {
                    "valid": True,
                    "command": "TON_THAP",
                    "intentReason": f"JSON Command TON_THAP (Model={model_name}, Shop={shop})",
                    "materials": mats,
                    "model": model_name,
                    "shop": shop,
                    "plants": plants,
                    "limit": limit,
                    "plant": plant,
                    "batch": batch,
                    "days": days,
                    "requestId": req_id,
                    "messageId": msg_id or req_id,
                    "senderEmail": sender_email,
                    "senderName": sender_name,
                    "rawText": raw_text or f"TON_THAP {model_name}",
                    "dryRun": dry_run,
                    "note": note
                }

            if cmd in ("VITRI", "BIN", "TON", "MB52", "LICHSU", "MB51", "IB_TREO", "VL06I", "STO", "ME2N") and not mats and not dry_run:
                return {
                    "valid": False,
                    "needHelpReply": True,
                    "messageId": msg_id or req_id,
                    "senderEmail": sender_email,
                    "senderName": sender_name,
                    "detectedIntent": cmd,
                    "reason": f"Không tìm thấy danh sách mã vật tư (materials) hợp lệ trong yêu cầu."
                }

            if cmd == "PING":
                return {
                    "valid": True,
                    "command": "PING",
                    "intentReason": "JSON Command PING",
                    "materials": [],
                    "plant": plant,
                    "batch": batch,
                    "days": days,
                    "requestId": req_id,
                    "messageId": msg_id or req_id,
                    "senderEmail": sender_email,
                    "senderName": sender_name,
                    "rawText": raw_text or "PING test",
                    "dryRun": dry_run,
                    "note": note
                }

            if mats or cmd in (
                "CHECK_OB", "LICHSU_BIN",
                "TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR",
                "CHECK_IB", "IB_DETAIL", "KIEMTRA_IB",
                "TIM_IB", "FIND_IB",
                "CHECK_PO", "PO_STATUS", "KIEMTRA_PO"
            ):
                canonical_cmd = cmd
                if cmd in ("TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR"):
                    canonical_cmd = "TIM_GR_DOC"
                elif cmd in ("CHECK_IB", "IB_DETAIL", "KIEMTRA_IB"):
                    canonical_cmd = "CHECK_IB"
                elif cmd in ("TIM_IB", "FIND_IB"):
                    canonical_cmd = "TIM_IB"
                elif cmd in ("CHECK_PO", "PO_STATUS", "KIEMTRA_PO"):
                    canonical_cmd = "CHECK_PO"
                elif cmd in ("CHECK_FG_OB", "DIEN_OB"):
                    canonical_cmd = "CHECK_FG_OB"
                elif cmd in ("CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL", "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING"):
                    canonical_cmd = "CHECK_NVL_MIGO"
                elif cmd in ("CHECK_BOM", "BOM", "CS03"):
                    canonical_cmd = "CHECK_BOM"
                elif cmd in ("CHECK_QUEUE", "QUEUE"):
                    canonical_cmd = "CHECK_QUEUE"

                expected_q = 0.0
                try:
                    expected_q = float(json_obj.get("expectedQty", 0.0) or 0.0)
                except Exception:
                    expected_q = 0.0

                return {
                    "valid": True,
                    "command": canonical_cmd or "VITRI",
                    "intentReason": f"JSON Payload ({canonical_cmd or 'VITRI'})",
                    "materials": mats,
                    "plant": plant,
                    "batch": batch,
                    "days": days,
                    "requestId": req_id,
                    "messageId": msg_id or req_id,
                    "senderEmail": sender_email,
                    "senderName": sender_name,
                    "rawText": raw_text,
                    "dryRun": dry_run,
                    "note": note,
                    "obNo": str(json_obj.get("obNo", "")).strip(),
                    "bin": str(json_obj.get("bin", "")).strip(),
                    "model": str(json_obj.get("model", "")).strip(),
                    "shop": str(json_obj.get("shop", "")).strip(),
                    "plants": json_obj.get("plants", []),
                    "limit": int(json_obj.get("limit", 20)),
                    "inboundNo": str(json_obj.get("inboundNo", "") or json_obj.get("inbound", "")).strip(),
                    "poNo": str(json_obj.get("poNo", "") or json_obj.get("po", "")).strip(),
                    "poItem": str(json_obj.get("poItem", "") or json_obj.get("item", "")).strip(),
                    "expectedQty": expected_q,
                    "refNo": str(json_obj.get("refNo", "") or json_obj.get("ref", "") or json_obj.get("bill", "")).strip(),
                    "materialDocument": str(json_obj.get("materialDocument", "") or json_obj.get("docNo", "")).strip(),
                    "folderPath": str(json_obj.get("folderPath", "") or json_obj.get("folder", "")).strip(),
                    "attachmentDocs": json_obj.get("attachmentDocs", []) or json_obj.get("docs", [])
                }

            if raw_text:
                cleaned = raw_text

        meta = {
            "messageId": "",
            "senderEmail": "Unknown",
            "senderName": "Đồng nghiệp",
            "senderId": "",
            "chatId": "",
            "rawText": ""
        }

        # Bắt MSG_ID từ Subject nếu có (Dự phòng an toàn)
        subj_id = re.search(r"\[SAP-REQUEST\]\s*(\d+)", subject or "")
        if subj_id:
            meta["messageId"] = subj_id.group(1).strip()

        # Bóc tách các trường từ Body của Mail V3 (hỗ trợ cả SENDER_ID, CHAT_ID)
        TAG_STOP = r"(?=\s*(?:MSG_ID|SENDER_EMAIL|SENDER_NAME|SENDER_ID|CHAT_ID|CONVERSATION_ID|RAW_TEXT)\s*:|\r|\n|$)"
        msg_m = re.search(r"MSG_ID\s*:\s*([^\r\n]+?)" + TAG_STOP, cleaned, re.I)
        mail_m = re.search(r"SENDER_EMAIL\s*:\s*([^\r\n]+?)" + TAG_STOP, cleaned, re.I)
        name_m = re.search(r"SENDER_NAME\s*:\s*([^\r\n]+?)" + TAG_STOP, cleaned, re.I)
        sid_m = re.search(r"SENDER_ID\s*:\s*([^\r\n]+?)" + TAG_STOP, cleaned, re.I)
        chat_m = re.search(r"(?:CHAT_ID|CONVERSATION_ID)\s*:\s*([^\r\n]+?)" + TAG_STOP, cleaned, re.I)
        raw_m = re.search(r"RAW_TEXT\s*:\s*(.+)", cleaned, re.I | re.DOTALL)

        if msg_m and msg_m.group(1).strip():
            meta["messageId"] = msg_m.group(1).strip()
        if mail_m:
            meta["senderEmail"] = mail_m.group(1).strip()
        if name_m:
            meta["senderName"] = name_m.group(1).strip()
        if sid_m:
            meta["senderId"] = sid_m.group(1).strip()
        if chat_m:
            meta["chatId"] = chat_m.group(1).strip()

        command_text = raw_m.group(1).strip() if raw_m else cleaned
        meta["rawText"] = command_text
        norm_text = cls.remove_accents(command_text)

        # 1. Chấm điểm Ý định (Intent Scoring)
        scores = {k: 0 for k in cls.INTENT_SCORES.keys()}
        matched_keywords = []

        for intent, rules in cls.INTENT_SCORES.items():
            for pat, weight in rules:
                m = re.search(pat, norm_text)
                if m:
                    scores[intent] += weight
                    matched_keywords.append(f"{intent}:'{m.group(0)}' (+{weight})")

        best_intent = max(scores, key=scores.get)
        best_score = scores[best_intent]

        # Phân biệt TỒN KHO (Chỉ kiểm tra Kho 1005 & 3093) vs TỒN TỔNG HỢP (Toàn bộ SLoc)
        if best_intent == "TON":
            if re.search(r"\b(?:kho|warehouse)\b", norm_text):
                best_intent = "TON_KHO"
                matched_keywords.append("TON_KHO: Lọc riêng Kho NVL 1005 & Kho FG 3093")
            else:
                matched_keywords.append("TON: Tra cứu toàn bộ SLoc plant")

        # 2. Xử lý các thực thể đặc biệt: MODEL, SHOP, PLANTS, LIMIT, OB, BIN
        from app.material_resolver import MaterialResolver
        model_name, shop_name, plants_list = MaterialResolver.detect_model_and_shop(command_text)

        limit_val = 20
        top_m = re.search(r"\b(?:top\s*|check\s*(?:cho\s*toi\s*)?)?(\d{1,3})\s*(?:ma|dong)\b|\btop\s*(\d{1,3})\b", norm_text)
        if top_m:
            val = top_m.group(1) or top_m.group(2)
            if val and val.isdigit():
                limit_val = max(1, min(int(val), 100))

        ob_no = ""
        ob_m = re.search(r"\b(?:ob\s*[:=]?\s*)?(80\d{6,8}|08\d{6,8}|10\d{8})\b", command_text, re.I)
        if ob_m:
            ob_no = ob_m.group(1)

        bin_val = ""
        bin_m = re.search(r"\b(?:bin|ke|\bo\b)\s*[:=]?\s*([a-z0-9_-]{3,15})\b|\b(\d{2}-[a-z0-9]+-[a-z0-9]+)\b", norm_text)
        if bin_m:
            bin_cand = (bin_m.group(1) or bin_m.group(2)).upper()
            if bin_cand not in cls.STOP_WORDS:
                bin_val = bin_cand

        inbound_no = ""
        ib_m = re.search(r"\b(?:inbound|ib)\s*[:=]?\s*([0-9]{8,10})\b|\b(18\d{8})\b", command_text, re.I)
        if ib_m:
            inbound_no = (ib_m.group(1) or ib_m.group(2)).strip()

        po_no = ""
        po_m = re.search(r"\b(?:po|sa|ebeln|don\s*hang)\s*[:=]?\s*([0-9]{10})\b|\b((?:45|55)\d{8})\b", command_text, re.I)
        if po_m:
            po_no = (po_m.group(1) or po_m.group(2)).strip()

        po_item = ""
        po_item_m = re.search(r"\b(?:item|po\s*item|line)\s*[:=]?\s*(\d{1,5})\b", norm_text)
        if po_item_m:
            po_item = po_item_m.group(1).strip()

        expected_qty = 0.0
        eqty_m = re.search(r"\b(?:sl\s*yeu\s*cau|sl\s*dat|sl|so\s*luong|expected\s*qty|qty)\s*[:=]?\s*(\d+(?:[.,]\d+)?)\b", norm_text)
        if eqty_m:
            try:
                expected_qty = float(eqty_m.group(1).replace(",", "."))
            except Exception:
                expected_qty = 0.0

        ref_no = ""
        ref_m = re.search(r"\b(?:bill|van\s*don|ref|so\s*bill|tham\s*chieu|lifex)\s*[:=]?\s*([a-z0-9_-]{3,25})\b", norm_text)
        if ref_m:
            ref_no = ref_m.group(1).strip().upper()

        mat_doc = ""
        doc_m = re.search(r"\b(?:ct|chung\s*tu|mblnr|mat\s*doc)\s*[:=]?\s*(\d{10})\b", norm_text)
        if doc_m:
            mat_doc = doc_m.group(1).strip()

        # Bóc tách đường dẫn thư mục và danh sách chứng từ đính kèm (nếu có lệnh ATT_MIGO)
        from app.sap_attachment import clean_folder_path, extract_folder_path_multiline, DEFAULT_MIGO_ATT_FOLDER
        raw_folder = extract_folder_path_multiline(command_text)
        folder_path = ""
        if raw_folder:
            folder_path = clean_folder_path(raw_folder)
        elif re.search(r"\b(?:att\s*mail|attach\s*mail|dinh\s*kem\s*mail|att)\b", norm_text):
            # Nếu người dùng gửi lệnh đính kèm 'SAP att mail' mà không dán link dài:
            # Tự động lấy thư mục mặc định từ ô B3 của file 261002_MIGO_ATT_V3_4_3.xlsm
            folder_path = DEFAULT_MIGO_ATT_FOLDER

        att_items = []
        try:
            from app.sap_attachment import parse_attachment_items_from_text
            att_items = parse_attachment_items_from_text(command_text)
        except Exception:
            att_items = []

        att_docs = [item["doc"] for item in att_items] if att_items else re.findall(r"\b(?:49\d{8}|500\d{7}|\d{10})\b", command_text)

        # NẾU CÓ FOLDER PATH HOẶC YÊU CẦU ĐIỀN SỐ PHIẾU / ATT_MIGO:
        has_phieu_intent = bool(re.search(r"\b(?:dien\s*so\s*phieu|dien\s*phieu|so\s*phieu|headtext|bktxt|att)\b", norm_text))
        if (folder_path or has_phieu_intent or best_intent == "ATT_MIGO") and (att_items or att_docs):
            best_intent = "ATT_MIGO"
            best_score = max(best_score, 35)
            matched_keywords.append("ATT_MIGO: Tự động nhận diện lệnh điền số phiếu HeadText / đính kèm MIGO")

        # 3. Bóc tách các tham số phụ: PLANT, BATCH, SỐ NGÀY
        explicit_plant = False
        plant = default_plant
        plant_m = re.search(r"\b(?:plant|nha\s*may|werks)\s*[:=]?\s*(\d{4})\b", norm_text)
        if plant_m:
            plant = plant_m.group(1)
            explicit_plant = True

        # Kiểm tra nếu người dùng yêu cầu nguyên list / toàn bộ cho Model Tồn thấp
        is_full_model_list = bool(re.search(r"\b(?:nguyen\s*list|toan\s*bo|full\s*list|tat\s*ca|tatca|full)\b", norm_text))
        if is_full_model_list and (best_intent == "TON_THAP" or "body" in norm_text or "limo" in norm_text):
            if best_intent != "TON_THAP" and ("limo" in norm_text or "vf" in norm_text or "body" in norm_text):
                best_intent = "TON_THAP"
                best_score = max(best_score, 25)
            limit_val = 100

        batch = ""
        batch_m = re.search(r"\b(?:batch|so\s*lo|\blo\b|charg)\s*[:=]?\s*([a-z0-9_-]{3,15})\b", norm_text)
        if batch_m:
            cand = batch_m.group(1).upper()
            if cand not in cls.STOP_WORDS:
                batch = cand

        days = 7
        days_m = re.search(r"\b(?:ngay|days|day)\s*[:=]?\s*(\d{1,3})\b|\b(\d{1,3})\s*(?:ngay|days|day|hom)\b", norm_text)
        if days_m:
            d_val = days_m.group(1) or days_m.group(2)
            days = max(1, min(int(d_val), 365))
            if best_score == 0:
                best_intent = "LICHSU"
                best_score = 5

        # 4. Bóc tách danh sách Mã hàng (Materials)
        work_str = command_text.upper()
        if meta["messageId"]:
            work_str = work_str.replace(meta["messageId"].upper(), " ")
        if folder_path:
            work_str = work_str.replace(folder_path.upper(), " ")
        # Xóa sạch mọi dạng đường dẫn UNC hoặc thư mục còn sót lại
        work_str = re.sub(r"(?:\\\\[^\r\n\"\'<>|]+|[A-Za-z]:\\[^\r\n\"\'<>|]+)", " ", work_str)
        # Xóa địa chỉ IP (ví dụ: 10.128.10.72) để không bị nhận nhầm thành mã vật tư
        work_str = re.sub(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b", " ", work_str)
        if best_intent == "ATT_MIGO":
            for d in att_docs:
                work_str = work_str.replace(d, " ")
            for item in att_items:
                if item.get("doc"):
                    work_str = work_str.replace(item["doc"], " ")
                if item.get("ticket_no"):
                    work_str = work_str.replace(item["ticket_no"].upper(), " ")
        if batch_m and batch:
            work_str = re.sub(r"\b(?:BATCH|SO\s*LO|LO|CHARG|LÔ|SỐ\s*LÔ)\s*[:=]?\s*" + re.escape(batch), " ", work_str, flags=re.I)
        if plant_m:
            work_str = re.sub(r"\b(?:PLANT|WERKS|NHA\s*MAY|NHÀ\s*MÁY)\s*[:=]?\s*" + re.escape(plant), " ", work_str, flags=re.I)
        if days_m:
            work_str = re.sub(r"\b\d{1,3}\s*(?:NGAY|NGÀY|DAYS|DAY|HOM|HÔM)\b", " ", work_str, flags=re.I)
        if inbound_no:
            work_str = work_str.replace(inbound_no, " ")
        if po_no:
            work_str = work_str.replace(po_no, " ")
        if mat_doc:
            work_str = work_str.replace(mat_doc, " ")
        if ob_no:
            work_str = work_str.replace(ob_no, " ")

        tokens = re.findall(r"\b[A-Z0-9._-]{5,18}\b", work_str)
        materials = []
        for tk in tokens:
            clean_tk = tk.strip(".-_")
            if len(clean_tk) < 5 or clean_tk in cls.STOP_WORDS:
                continue
            if re.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$", clean_tk):
                continue
            if any(ch.isdigit() for ch in clean_tk) and clean_tk != plant and clean_tk != batch:
                if clean_tk not in materials:
                    materials.append(clean_tk)

        # Nếu chưa có batch nhưng phần tử cuối trong materials trông giống số lô (VD: B2610, L2024, LOT01)
        if not batch and len(materials) > 1:
            last_mat = materials[-1]
            if re.fullmatch(r"[A-Z]\d{3,7}|LOT[A-Z0-9]{1,6}", last_mat):
                batch = last_mat
                materials.pop()

        # Nếu không có từ khóa batch rõ ràng, kiểm tra token ngay sau mã hàng
        if not batch and len(materials) == 1:
            after_mat_tokens = [t.strip().upper() for t in re.split(r"[,;\s]+", command_text) if t.strip()]
            for idx, t in enumerate(after_mat_tokens):
                if t == materials[0] and idx + 1 < len(after_mat_tokens):
                    candidate = after_mat_tokens[idx + 1]
                    if re.fullmatch(r"[A-Z0-9_-]{4,12}", candidate) and candidate not in cls.STOP_WORDS and candidate != plant:
                        batch = candidate
                        break

        # 5. Xử lý Thông minh (Smart Fallback & Coaching):
        if best_score == 0 and att_items and all(item.get("ticket_no") for item in att_items):
            best_intent = "ATT_MIGO"
            best_score = 25
            matched_keywords.append("Mặc định thông minh (Có cặp Chứng từ và Số phiếu -> Điền số phiếu MIGO)")

        is_no_mat_allowed = (
            best_score > 0 and (
                best_intent == "TON_THAP"
                or (best_intent == "CHECK_OB" and ob_no)
                or (best_intent == "LICHSU_BIN" and bin_val)
                or (best_intent == "VITRI" and bin_val)
                or (best_intent == "CHECK_IB" and inbound_no)
                or (best_intent == "TIM_IB" and (po_no or ref_no))
                or (best_intent == "CHECK_PO" and po_no)
                or (best_intent == "TIM_GR_DOC" and (inbound_no or ref_no or mat_doc))
                or (best_intent == "CHECK_QUEUE")
                or (best_intent == "ATT_MIGO")
            )
        ) or (best_score == 0 and bool(bin_val))

        if (best_intent == "HELP" or not materials) and not is_no_mat_allowed:
            if not meta["messageId"] and best_intent != "HELP":
                return None
            meta.update({
                "valid": False,
                "needHelpReply": bool(meta["messageId"]),
                "detectedIntent": best_intent if best_score > 0 else "UNKNOWN",
                "reason": "Người dùng yêu cầu hướng dẫn hoặc chưa nhập Mã hàng"
            })
            return meta

        if best_score == 0 and materials:
            best_intent = "VITRI"
            matched_keywords.append("Mặc định thông minh (Có mã hàng -> Check Vị trí & GR)")
        elif best_score == 0 and bin_val:
            best_intent = "VITRI"
            matched_keywords.append("Mặc định thông minh (Có tên Bin -> Check Vị trí Bin)")

        if best_intent == "ATT_MIGO":
            materials = []

        meta.update({
            "valid": True,
            "command": best_intent,
            "intentReason": ", ".join(matched_keywords) if matched_keywords else f"Trọng số {best_score}",
            "materials": materials,
            "plant": plant,
            "batch": batch,
            "days": days,
            "model": model_name or ("LIMO" if best_intent == "TON_THAP" else ""),
            "shop": shop_name,
            "plants": plants_list,
            "limit": limit_val,
            "obNo": ob_no,
            "bin": bin_val,
            "inboundNo": inbound_no,
            "poNo": po_no,
            "poItem": po_item,
            "expectedQty": expected_qty,
            "refNo": ref_no,
            "materialDocument": mat_doc,
            "folderPath": folder_path,
            "attachmentDocs": att_docs,
            "attachmentItems": att_items,
            "explicitPlant": explicit_plant,
            "fullModelList": is_full_model_list
        })
        return meta

    @classmethod
    def parse_from_mail(cls, raw_body: str, html_body: str = "", default_plant: str = "3093") -> Optional[Dict[str, Any]]:
        """Hàm alias tương thích ngược."""
        return cls.parse(raw_body, html_body=html_body, subject="", default_plant=default_plant)


# Alias tương thích ngược
EmailCommandParser = SmartCommandParser
NaturalCommandParser = SmartCommandParser
