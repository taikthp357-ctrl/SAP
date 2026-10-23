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
        "VITRI": [
            (r"\b(?:vi\s*tri|vitri|vtri|bin|posting|scwm/post|scwm)\b", 10),
            (r"\b(?:ngay\s*gr|gio\s*gr|check\s*gr|kiem\s*tra\s*gr)\b", 10),
            (r"\b(?:tim\s*hang|o\s*dau|ke\s*nao|nam\s*o|khu\s*vuc|zone)\b", 8),
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
        "NAM", "DAU", "NAO", "NHE", "NHA", "THANKS", "CAM", "ON", "DUM", "HO", "STATUS"
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
        source = html_body if (not raw_body or len(raw_body.strip()) < 5) else raw_body
        text = re.sub(r"<br\s*/?>|</p>|</div>|</tr>|</li>", "\n", source, flags=re.IGNORECASE)
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
    def parse(cls, raw_body: str, html_body: str = "", subject: str = "", default_plant: str = "3093") -> Optional[Dict[str, Any]]:
        """Phân tích nội dung email và trả về từ điển lệnh."""
        cleaned = cls.clean_v3_email(raw_body, html_body)
        if not cleaned and not subject:
            return None

        # 1. Hỗ trợ bóc tách JSON trực tiếp từ email (cả dạng thuần JSON lẫn có marker)
        json_obj = None
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

        if json_obj:
            cmd = str(json_obj.get("command", "")).strip().upper()
            req_id = str(json_obj.get("requestId", "") or json_obj.get("messageId", "")).strip()
            plant = str(json_obj.get("plant", default_plant)).strip().upper() or default_plant
            batch = str(json_obj.get("batch", "")).strip().upper()
            days = int(json_obj.get("days", 7))
            mats = json_obj.get("materials", [])
            if isinstance(mats, str):
                mats = [m.strip() for m in re.split(r"[,;\s]+", mats) if m.strip()]

            raw_text = str(json_obj.get("rawText", "") or json_obj.get("RAW_TEXT", "") or json_obj.get("note", "")).strip()
            sender_email = str(json_obj.get("senderEmail", "") or json_obj.get("SENDER_EMAIL", "") or "PowerAutomate").strip()
            sender_name = str(json_obj.get("senderName", "") or json_obj.get("SENDER_NAME", "") or "Hệ thống").strip()
            msg_id = str(json_obj.get("messageId", "") or json_obj.get("MSG_ID", "") or req_id).strip()
            note = str(json_obj.get("note", "")).strip()
            dry_run = bool(json_obj.get("dryRun", False))

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

            if mats:
                return {
                    "valid": True,
                    "command": cmd or "VITRI",
                    "intentReason": f"JSON Payload ({cmd or 'VITRI'})",
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
                    "note": note
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
        scores = {"VITRI": 0, "LICHSU": 0, "TON": 0, "HELP": 0}
        matched_keywords = []

        for intent, rules in cls.INTENT_SCORES.items():
            for pat, weight in rules:
                m = re.search(pat, norm_text)
                if m:
                    scores[intent] += weight
                    matched_keywords.append(f"{intent}:'{m.group(0)}' (+{weight})")

        best_intent = max(scores, key=scores.get)
        best_score = scores[best_intent]

        # 2. Bóc tách các tham số phụ: PLANT, BATCH, SỐ NGÀY
        plant = default_plant
        plant_m = re.search(r"\b(?:plant|nha\s*may|werks)\s*[:=]?\s*(\d{4})\b", norm_text)
        if plant_m:
            plant = plant_m.group(1)

        batch = ""
        batch_m = re.search(r"\b(?:batch|so\s*lo|lo|charg)\s*[:=]?\s*([a-z0-9_-]{3,15})\b", norm_text)
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

        # 3. Bóc tách danh sách Mã hàng (Materials)
        work_str = command_text.upper()
        if meta["messageId"]:
            work_str = work_str.replace(meta["messageId"].upper(), " ")
        if batch_m and batch:
            work_str = re.sub(r"\b(?:BATCH|SO\s*LO|LO|CHARG|LÔ|SỐ\s*LÔ)\s*[:=]?\s*" + re.escape(batch), " ", work_str, flags=re.I)
        if plant_m:
            work_str = re.sub(r"\b(?:PLANT|WERKS|NHA\s*MAY|NHÀ\s*MÁY)\s*[:=]?\s*" + re.escape(plant), " ", work_str, flags=re.I)
        if days_m:
            work_str = re.sub(r"\b\d{1,3}\s*(?:NGAY|NGÀY|DAYS|DAY|HOM|HÔM)\b", " ", work_str, flags=re.I)

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

        # 4. Xử lý Thông minh (Smart Fallback & Coaching):
        if best_intent == "HELP" or not materials:
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

        meta.update({
            "valid": True,
            "command": best_intent,
            "intentReason": ", ".join(matched_keywords) if matched_keywords else f"Trọng số {best_score}",
            "materials": materials,
            "plant": plant,
            "batch": batch,
            "days": days
        })
        return meta

    @classmethod
    def parse_from_mail(cls, raw_body: str, html_body: str = "", default_plant: str = "3093") -> Optional[Dict[str, Any]]:
        """Hàm alias tương thích ngược."""
        return cls.parse(raw_body, html_body=html_body, subject="", default_plant=default_plant)


# Alias tương thích ngược
EmailCommandParser = SmartCommandParser
NaturalCommandParser = SmartCommandParser
