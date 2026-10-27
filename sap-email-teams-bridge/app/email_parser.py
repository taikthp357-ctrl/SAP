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
        "WHO_IS_TAI": [
            (r"\b(?:anh\s+tai|sap\s+tai|ong\s+tai|tai)\s+(?:la\s+ai|nao\s+day|la\s+ai\s+day)\b", 100),
            (r"\b(?:ai\s+tao\s+ra\s+bot|ai\s+viet\s+bot|tac\s+gia\s+bot|pham\s+tien\s+tai|ai\s+lam\s+ra\s+bot)\b", 100),
            (r"\b(?:who\s+is\s+tai|author|tac\s*gia|cha\s+de\s+bot)\b", 100),
        ],
        "CHECK_ATT": [
            (r"\b(?:check\s*att|kiem\s*tra\s*att|tra\s*cuu\s*att|soi\s*att|check\s*dinh\s*kem|kiem\s*tra\s*dinh\s*kem)\b", 97),
            (r"\b(?:co\s*bbbg\s*(?:va\s*)?vat\s*chua|co\s*vat\s*(?:va\s*)?bbbg\s*chua|da\s*co\s*bbbg|da\s*co\s*vat|da\s*att\s*chua)\b", 96),
            (r"\b(?:check\s*bbbg|check\s*vat|kiem\s*tra\s*bbbg|kiem\s*tra\s*vat|tra\s*cuu\s*bbbg|tra\s*cuu\s*vat)\b", 96),
            (r"\b(?:check\s*att\s*doc|check\s*att\s*migo|kt\s*att|kt\s*dinh\s*kem|zfir266)\b", 97),
            (r"\b(?:doc\s*nay\s*co\s*bbbg|doc\s*nay\s*da\s*att|doc\s*co\s*att\s*chua)\b", 96),
        ],
        "ATT_MIGO": [
            (r"\b(?:dien\s*so\s*phieu|dien\s*phieu|dien\s*headtext|cap\s*nhat\s*headtext|dien\s*bktxt)\b", 95),
            (r"\b(?:att\s*mail|attach\s*mail|att\s*email|attach\s*email|dinh\s*kem\s*mail|dinh\s*kem\s*email)\b", 95),
            (r"\b(?:att\s*vat|att\s*bbbg|dinh\s*kem\s*vat|dinh\s*kem\s*bbbg|migo\s*att|att\s*migo|dinh\s*kem\s*migo)\b", 90),
            (r"\b(?:att\s*file|dinh\s*kem\s*file|gos\s*att|att\s*chung\s*tu)\b", 85),
            (r"\b(?:att\s*vat\s*bbbg|att\s*bbbg\s*vat)\b", 92),
        ],
        "CHECK_BOM": [
            (r"\b(?:bom|check\s*bom|kiem\s*tra\s*bom|co\s*bom|bom\s*co\s*khong|dinh\s*muc|cs03)\b", 90),
            (r"\b(?:linh\s*kien|linh\s*kien\s*con|ma\s*con|thanh\s*phan|subcon)\b", 88),
        ],
        "CHECK_QUEUE": [
            (r"\b(?:queue|check\s*queue|kiem\s*tra\s*queue|ket\s*queue|loi\s*queue|queue\s*ewm|mqueue|scwm/mon\s*queue|mqueue0001)\b", 85),
        ],
        "CHECK_OB": [
            (r"\b(?:ob|outbound|con\s*ob|check\s*ob|kiem\s*tra\s*ob|kt\s*ob|don\s*xuat|don\s*ob|vl06o|prdo|ob\s*treo|cho\s*xuat)\b", 80),
            (r"\b(?:ob\s*nao\s*dang\s*cho|ob\s*mo|ob\s*chua\s*xuat|ob\s*ma|con\s*ob\s*ma)\b", 78),
        ],
        "TIM_GR_DOC": [
            (r"\b(?:tim\s*gr|chung\s*tu\s*gr|gr\s*doc|mb51\s*gr|lich\s*su\s*gr|so\s*gr|tra\s*gr|chung\s*tu\s*nhap)\b", 84),
            (r"\b(?:tra\s*cuu\s*gr|tim\s*chung\s*tu|xem\s*gr|chung\s*tu\s*mb51)\b", 84),
        ],
        "TIM_IB": [
            (r"\b(?:tim\s*ib|tim\s*inbound|find\s*ib|find\s*inbound|inbound\s*theo\s*po|inbound\s*theo\s*bill)\b", 83),
        ],
        "IB_TREO": [
            (r"\b(?:ib\s*treo|ibtreo|don\s*treo|inbound\s*treo|chua\s*gr|chua\s*nhap|hang\s*chua\s*ve|cho\s*nhap|treo)\b", 82),
        ],
        "CHECK_IB": [
            (r"\b(?:chi\s*tiet\s*ib|chi\s*tiet\s*inbound|ib\s*detail|check\s*inbound\s*detail|doi\s*chieu\s*ib|so\s*khop\s*ib|kiem\s*tra\s*inbound)\b", 80),
            (r"\b(?:check\s*ib\s*18\d{6,8}|inbound\s*18\d{6,8}|ib\s*18\d{6,8})\b", 80),
            (r"\b(?:ib|inbound|con\s*ib|còn\s*ib|check\s*ib|kiem\s*tra\s*ib|hang\s*ve|hàng\s*về|lich\s*ve|lịch\s*về|asn|vl06i|zmme240)\b", 78),
        ],
        "TON_THAP": [
            (r"\b(?:s[1-7]|ca\s*[s0-9]|[1-7])\s+(?:nay\s+)?ch[aạ]y\b|\bch[aạ]y\s+(?:ca\s*)?(?:s[1-7]|[1-7])\b|\b(?:tinh|check)\s+doh\b", 92),
            (r"\b(?:canh\s*bao\s*ton\s*ca|ton\s*thap\s*ca|doh\s*ca|thieu\s*doh|chua\s*du\s*doh)\b", 88),
            (r"\b(?:ton\s*thap|thap|canh\s*bao\s*ton|low\s*stock|doh\s*thap|sap\s*het|thieu\s*hang)\b", 78),
            (r"\b(?:top\s*\d+\s*(?:ma)?\s*ton\s*thap|check\s*(?:cho\s*toi\s*)?\d+\s*ma\s*.*ton\s*thap)\b", 82),
            (r"\b(?:ton\s*thap\s*(?:body|xmd|limo|vf\d+))\b", 80),
        ],
        "STO": [
            (r"\b(?:sto|check\s*sto|po\s*mo|check\s*po|check\s*sa|me2n|don\s*mua|con\s*no\s*po|mb5t|dieu\s*chuyen)\b", 75),
            (r"\b(?:po\s*dieu\s*chuyen|don\s*dieu\s*chuyen)\b", 72),
        ],
        "LICHSU": [
            (r"\b(?:lich\s*su|lichsu|mb51|xuat\s*nhap|nhap\s*xuat|giao\s*dich|chung\s*tu\s*nhap|gr\s*doc)\b", 70),
            (r"\b\d+\s*(?:ngay|hom)\s*(?:gan\s*nhat|qua)?\b", 70),
        ],
        "CHECK_RES": [
            (r"\b(?:res|reservation|mb22|mb23|mb25|phieu\s*linh)\b", 68),
        ],
        "TON": [
            (r"\b(?:check\s*ton|check\s*ton|ton\s*kho|mb52|check\s*sloc|ton\s*sloc|sloc)\b", 65),
            (r"\b(?:so\s*luong\s*ton|con\s*ton|con\s*hang|stock|ton)\b", 62),
        ],
        "VITRI": [
            (r"\b(?:vi\s*tri|vtri|check\s*bin|ton\s*bin|scwm/post|scwm|post|o\s*bin|ke\s*bin|tim\s*hang|ngay\s*gr|gio\s*gr)\b", 60),
            (r"\b(?:check\s*bin|kiem\s*tra\s*bin|bin\s*nao|o\s*ke|tra\s*bin|bin)\b", 60),
        ],
        "LICHSU_BIN": [
            (r"\b(?:lich\s*su\s*bin|lichsu\s*bin|bin\s*to\s*bin|bintobin|chuyen\s*bin|cancel\s*wht|wht|scwm/mon|/scwm/mon)\b", 72),
        ],
        "CHECK_FG_OB": [
            (r"\b(?:dien\s*ob|di\s*en\s*ob|check\s*ob\s*fg|kha\s*dung\s*ob|ob\s*fg|ton\s*ob)\b", 82),
        ],
        "CHECK_NVL_MIGO": [
            (r"\b(?:posting\s*change|postingchange|check\s*posting|tai\s*posting|ra\s*don\s*migo|migo\s*311|migo|cap\s*migo|chuyen\s*kho\s*311|can\s*doi\s*nvl|check\s*nvl|kiem\s*tra\s*nvl|nvl\s*migo|nvl\s*1005|kho\s*1005)\b", 84),
        ],
        "CHECK_PO": [
            (r"\b(?:check\s*po|kiem\s*tra\s*po|tinh\s*trang\s*po|po\s*status|don\s*dat\s*hang|po\s*con\s*no|po\s*da\s*giao)\b", 74),
            (r"\b(?:po\s*45\d{8}|po\s*55\d{8}|me2n\s*po)\b", 76),
        ],
        "TAO_BBBG": [
            (r"\b(?:tao\s*bbbg|doi\s*soat\s*bbbg|xuat\s*bbbg|bbbg\s*fg|tao\s*bien\s*ban\s*ban\s*giao|bbbg\s*new|lap\s*bbbg)\b", 92),
        ],
        "XEM_BBBG": [
            (r"\b(?:xem\s*bbbg|preview\s*bbbg|xem\s*truoc\s*bbbg|xac\s*nhan\s*gui\s*bbbg)\b", 92),
        ],
        "RA_DON_NVL": [
            (r"\b(?:ra\s*don\s*nvl|phan\s*bo\s*nvl|ke\s*hoach\s*nvl|nvl\s*v31|nvl\s*zone|len\s*don\s*nvl)\b", 92),
        ],
        "DIEU_PHOI_OB": [
            (r"\b(?:dieu\s*phoi\s*ob|rc65|ob\s*tool|xuat\s*don\s*ob|gate\s*ob|whole\s*batch)\b", 92),
        ],
        "TRA_CUU_ASN": [
            (r"\b(?:tra\s*cuu\s*asn|xuat\s*xu\s*asn|tim\s*asn|zmme240|container\s*asn|bill\s*asn)\b", 88),
        ],
        "TRA_CUU_CAP_NVL": [
            (r"\b(?:cap\s*(?:cho\s*)?(?:ca|model|xe)\s*nao|dung\s*(?:cho\s*)?(?:ca|model|xe)\s*nao|dung\s*chung)\b", 98),
            (r"\b(?:cap\s*du\s*(?:chua|hay\s*thieu)|cap\s*thieu\s*(?:khsx|chay\s*not)?|ra\s*don\s*cap\s*not|cap\s*giup\s*e|cap\s*cho\s*a\s*con)\b", 96),
            (r"\b(?:tinh\s*trang\s*cap\s*(?:ma|vat\s*tu)|tra\s*cuu\s*cap\s*nvl|check\s*cap\s*ma|ho\s*tro\s*em\s*may\s*ma)\b", 95),
            (r"\b(?:cap\s*not|cap\s*bu|xin\s*cap)\b", 92),
        ],
        "TINH_TRANG_NVL": [
            (r"\b(?:tinh\s*trang\s*cap\s*(?:hang\s*)?nvl|tinh\s*trang\s*nvl|tinh\s*trang\s*cap\s*hang|tien\s*do\s*cap\s*nvl|tien\s*do\s*nvl)\b", 94),
            (r"\b(?:bao\s*cao\s*cap\s*(?:hang\s*)?nvl|bao\s*cao\s*nvl|bao\s*cao\s*cap\s*hang|ket\s*qua\s*cap\s*nvl)\b", 93),
            (r"\b(?:tinh\s*trang\s*(?:don\s*)?nvl\s*(?:ca|model)|tien\s*do\s*don\s*nvl)\b", 92),
            (r"\b(?:tinh\s*trang\s*cap\s*don|kiem\s*tra\s*cap\s*nvl|check\s*cap\s*nvl)\b", 91),
        ],
        "IN_WHT": [
            (r"\b(?:in\s*wht|in\s*phieu\s*wht|in\s*lenh\s*wht|print\s*wht)\b", 96),
            (r"\b(?:tcode\s*in\s*wht|huong\s*dan\s*in\s*wht|cach\s*in\s*wht|in\s*warehouse\s*task)\b", 95),
            (r"\b(?:in\s*to|in\s*phieu\s*xuat|mb90|lt31|in\s*phieu\s*cap)\b", 90),
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
        "NHẤT", "CỦA", "HOẶC", "HOAC", "XMD", "LIMO", "BODY", "XMĐ", "XE", "MAY", "DIEN", "OB", "WHT", "POSTING", "POSTINGCHANGE", "MIGO", "NVL", "CHANGE",
        "PLAN", "PLAT", "PL", "NHAMAY", "NHA", "MAY"
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

        # 1. Bảo toàn cấu trúc Table khi copy từ Excel/Teams (loại bỏ p/div/span lồng trong ô)
        def _clean_table_row(match):
            row_html = match.group(1)
            cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", row_html, flags=re.IGNORECASE | re.DOTALL)
            cleaned_cells = []
            for c in cells:
                c_text = re.sub(r"<[^>]+>", " ", c)
                c_text = re.sub(r"\s+", " ", c_text).strip()
                cleaned_cells.append(c_text)
            return "\t".join(cleaned_cells) + "\n"

        text = re.sub(r"<tr[^>]*>(.*?)</tr>", _clean_table_row, text, flags=re.IGNORECASE | re.DOTALL)

        # 2. Xử lý các thẻ ngoài bảng
        text = re.sub(r"<br\s*/?>|</p>|</div>|</li>", "\n", text, flags=re.IGNORECASE)
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
    def parse_nvl_table_rows(cls, text: str) -> List[Dict[str, Any]]:
        """
        Bóc tách danh sách các dòng yêu cầu ra đơn NVL (dán từ Excel, Teams bảng hoặc tab-separated).
        Mỗi dòng gồm: Mã NVL (cột 1), Số lượng (cột 2), Trạm cấp (cột 3).
        Ví dụ: BIW75111244AA    100    KV08-Thức/Vũ
        Hỗ trợ cả trường hợp có cột STT đứng đầu: 1  BIW75111244AA  100  KV08-Thức/Vũ
        """
        if not text:
            return []
        rows = []
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            # Bỏ qua tiền tố lệnh ở đầu dòng nếu người dùng gõ liền (VD: "SAP ra đơn nvl BIW... 128 KV-1")
            m_pref = re.match(r"^\s*(?:sap\s+)?(?:ra\s+[dđ][oơ][n]\s+nvl|k[eê]\s+ho[aạ]ch\s+nvl|[dđ][oơ][n]\s+nvl)\s*[:=-]?\s*", line, re.I)
            if m_pref and m_pref.end() < len(line):
                line = line[m_pref.end():].strip()

            norm_line = cls.remove_accents(line).lower()
            # Bỏ qua các dòng chỉ chứa tiêu đề thuần túy
            if any(h in norm_line for h in ["ma nvl", "ma hang", "material", "so luong", "sl", "qty", "tram", "station", "ra don nvl"]):
                if not re.search(r"\b\d+\b", line):
                    continue

            # Phân tách cột theo tab, pipe hoặc khoảng cách rộng (>= 2 spaces)
            cols = []
            if "\t" in line:
                cols = [c.strip() for c in line.split("\t") if c.strip()]
            elif "|" in line:
                cols = [c.strip() for c in line.split("|") if c.strip()]
            else:
                parts = re.split(r"\s+", line)
                known_units = {"EA", "PC", "KG", "M", "L", "SET", "ROL", "CAR", "THÙNG", "BOX", "BAG", "PALLET"}
                if len(parts) >= 4 and parts[2].upper() in known_units:
                    cols = [parts[0], parts[1], parts[2], " ".join(parts[3:])]
                elif len(parts) >= 3:
                    cols = [parts[0], parts[1], " ".join(parts[2:])]
                else:
                    cols = parts

            # Nếu có cột STT hoặc OB ở đầu (VD: 1, 01, OB1, #1) và cột kế tiếp là mã vật tư
            if len(cols) >= 3 and (re.fullmatch(r"#?\d{1,3}|OB\d{1,3}", cols[0], re.I)):
                cols = cols[1:]

            if len(cols) >= 2:
                mat_cand = cols[0].strip().upper()
                if len(mat_cand) < 3 or mat_cand in cls.STOP_WORDS:
                    continue
                qty_str = cols[1].replace(",", ".").strip()
                try:
                    qty_val = float(qty_str)
                    if qty_val <= 0:
                        continue
                except ValueError:
                    continue

                # Hàng rào kiểm tra hình thái dữ liệu (Layer 2 Data Shape Shield):
                # Số chứng từ DOC SAP (10 chữ số: 50xxxxxxxx, 49xxxxxxxx, 10xxxxxxxx) tuyệt đối KHÔNG phải mã NVL!
                # Bất kể cột thứ 2 là năm, số phiếu hay số lượng, DOC number không bao giờ được coi là mã vật tư.
                if re.fullmatch(r"(?:49|50|10)\d{8}", mat_cand):
                    continue

                unit_cand = "PC"
                station_cand = ""
                known_units = {"EA", "PC", "KG", "M", "L", "SET", "ROL", "CAR", "THÙNG", "BOX", "BAG", "PALLET"}
                if len(cols) >= 4 and cols[2].strip().upper() in known_units:
                    unit_cand = cols[2].strip().upper()
                    station_cand = " ".join(cols[3:]).strip()
                elif len(cols) >= 3:
                    station_cand = " ".join(cols[2:]).strip()

                int_or_float = int(qty_val) if qty_val.is_integer() else qty_val
                rows.append({
                    "material": mat_cand,
                    "qty": int_or_float,
                    "unit": unit_cand,
                    "station": station_cand
                })
        return rows

    JSON_COMMAND_ALIASES = {
        "GR_DOC": "TIM_GR_DOC", "CHUNGTU_GR": "TIM_GR_DOC",
        "IB_DETAIL": "CHECK_IB", "KIEMTRA_IB": "CHECK_IB", "FIND_IB": "TIM_IB",
        "PO_STATUS": "CHECK_PO", "KIEMTRA_PO": "CHECK_PO", "DIEN_OB": "CHECK_FG_OB",
        "RA_DON_MIGO": "CHECK_NVL_MIGO", "MIGO_311": "CHECK_NVL_MIGO",
        "CHECK_NVL": "CHECK_NVL_MIGO", "CHECK_POSTING_CHANGE": "CHECK_NVL_MIGO",
        "POSTING_CHANGE": "CHECK_NVL_MIGO", "POSTING": "CHECK_NVL_MIGO",
        "BOM": "CHECK_BOM", "CS03": "CHECK_BOM", "QUEUE": "CHECK_QUEUE",
        "ATTACH_VAT_BBBG": "ATT_MIGO", "MIGO_ATT": "ATT_MIGO", "ATTACH_MIGO": "ATT_MIGO",
        "CHECK_ATT": "CHECK_ATT", "CHECK_ATT_DOC": "CHECK_ATT", "CHECK_DINH_KEM": "CHECK_ATT", "ZFIR266": "CHECK_ATT",
        "KE_HOACH_NVL": "RA_DON_NVL", "TONKHO": "TON_KHO",
    }
    JSON_COMMANDS = frozenset(INTENT_SCORES) | {
        "PING", "BIN", "POST", "SCWM", "/SCWM/POST", "TON_KHO", "MB52", "MB51",
        "IB", "INBOUND", "VL06I", "ME2N", "MB5T", "DIEUCHUYEN"
    }

    @classmethod
    def _parse_json_request(cls, obj: Dict[str, Any], default_plant: str) -> Dict[str, Any]:
        from decimal import Decimal, InvalidOperation
        command = str(obj.get("command", "")).strip().upper()
        command = cls.JSON_COMMAND_ALIASES.get(command, command)
        base = {"requestId": str(obj.get("requestId") or obj.get("messageId") or ""),
                "messageId": str(obj.get("messageId") or obj.get("MSG_ID") or obj.get("requestId") or ""),
                "senderEmail": str(obj.get("senderEmail") or obj.get("SENDER_EMAIL") or "PowerAutomate"),
                "senderName": str(obj.get("senderName") or obj.get("SENDER_NAME") or "Hệ thống"),
                "senderId": str(obj.get("senderId") or obj.get("SENDER_ID") or ""),
                "chatId": str(obj.get("chatId") or obj.get("CHAT_ID") or ""),
                "command": command, "explicitWriteIntent": command == "ATT_MIGO",
                "executionMode": "WRITE_PREVIEW" if command == "ATT_MIGO" else "READ_ONLY"}
        def invalid(reason):
            return {**base, "valid": False, "needHelpReply": True,
                    "detectedIntent": command or "UNKNOWN", "reason": reason}
        if command not in cls.JSON_COMMANDS:
            return invalid(f"Lệnh '{command}' không được hỗ trợ trong hệ thống.")
        if command == "HELP":
            return invalid("Người dùng yêu cầu hướng dẫn cú pháp.")
        dry_run = obj.get("dryRun", False)
        if not isinstance(dry_run, bool):
            return invalid("dryRun phải là true hoặc false trong JSON.")
        raw_mats = obj.get("materials", [])
        if isinstance(raw_mats, (str, int)) and not isinstance(raw_mats, bool):
            raw_mats = re.split(r"[,;\s]+", str(raw_mats))
        mats = []
        for value in raw_mats if isinstance(raw_mats, list) else []:
            if isinstance(value, (str, int)) and not isinstance(value, bool):
                text = str(value).strip()
                if re.fullmatch(r"[\w\-./*?]{2,40}", text) and text not in mats:
                    mats.append(text)
        try:
            days = max(1, min(int(obj.get("days", 7)), 365))
        except (ValueError, TypeError, OverflowError):
            days = 7
        try:
            limit = max(1, min(int(obj.get("limit", 20)), 100))
        except (ValueError, TypeError, OverflowError):
            return invalid("limit phải là số nguyên.")
        qty = obj.get("qty", obj.get("quantity", obj.get("expectedQty")))
        if qty is not None:
            try:
                if isinstance(qty, bool):
                    raise ValueError()
                quantity = Decimal(str(qty))
                if not quantity.is_finite() or quantity < 0:
                    raise ValueError()
                if command == "RA_DON_NVL" and quantity <= 0:
                    raise ValueError()
                qty = float(quantity)
            except (ValueError, InvalidOperation):
                return invalid("Số lượng phải là số JSON không âm; đơn NVL phải lớn hơn 0.")
        items = obj.get("attachmentItems", [])
        docs = obj.get("attachmentDocs", obj.get("docs", []))
        year = str(obj.get("year", "") or "").strip()
        if year and not re.fullmatch(r"(?:19|20|21)\d{2}", year):
            return invalid("year phải là năm chứng từ gồm 4 chữ số.")
        normalized_items = []
        if command == "ATT_MIGO":
            if not isinstance(items, list) or not isinstance(docs, list):
                return invalid("attachmentItems và attachmentDocs phải là danh sách.")
            if not items:
                items = [{"doc": doc} for doc in docs]
            for item in items:
                if not isinstance(item, dict):
                    return invalid("Mỗi attachmentItems phải là object có doc và ticket_no.")
                doc = str(item.get("doc", "")).strip()
                item_year = str(item.get("year") or year)
                ticket = str(item.get("ticket_no", item.get("ticketNo", "")) or "").strip()
                if not re.fullmatch(r"\d{10}", doc) or (item_year and not re.fullmatch(r"(?:19|20|21)\d{2}", item_year)):
                    return invalid("Mỗi chứng từ cần doc 10 chữ số và year hợp lệ.")
                normalized_items.append({"doc": doc, "ticket_no": ticket, "year": item_year})
            if not normalized_items:
                return invalid("ATT_MIGO cần danh sách attachmentItems; không suy đoán từ rawText.")
            docs = [item["doc"] for item in normalized_items]
        if command in ("CHECK_ATT", "CHECK_ATT_DOC"):
            if not isinstance(docs, list):
                if isinstance(docs, str):
                    docs = [d.strip() for d in re.split(r"[,;\s]+", docs) if d.strip()]
                else:
                    docs = []
            if not docs and items:
                docs = [item["doc"] if isinstance(item, dict) else str(item) for item in items]
            if not docs and mats:
                docs = [m for m in mats if m.isdigit() and len(m) >= 8]
            if not docs:
                return invalid("CHECK_ATT cần danh sách số chứng từ attachmentDocs hoặc docs.")

        bin_value = str(obj.get("bin", "") or "").strip()
        if command in {"VITRI", "BIN", "POST", "SCWM", "/SCWM/POST", "TON", "TON_KHO", "MB52", "LICHSU", "MB51", "IB_TREO", "IB", "INBOUND", "VL06I", "STO", "ME2N", "MB5T", "DIEUCHUYEN", "CHECK_BOM", "CHECK_FG_OB", "CHECK_NVL_MIGO", "RA_DON_NVL"} and not mats and not (command in {"VITRI", "BIN"} and bin_value) and not dry_run:
            return invalid("Không tìm thấy danh sách mã vật tư (materials) hợp lệ trong yêu cầu.")
        model = str(obj.get("model", "") or ("LIMO" if command == "TON_THAP" else "")).upper()
        shop = str(obj.get("shop", "") or ("XMD" if "XMD" in model else "BODY" if command == "TON_THAP" else "")).upper()
        result = {**base, "valid": True, "materials": mats, "plant": str(obj.get("plant") or default_plant).upper(),
                  "batch": str(obj.get("batch") or "").upper(), "days": days, "dryRun": dry_run,
                  "intentReason": f"JSON Payload ({command})", "rawText": str(obj.get("rawText") or obj.get("RAW_TEXT") or ""),
                  "note": str(obj.get("note") or ""), "limit": limit, "model": model, "shop": shop,
                  "plants": obj.get("plants", ["3093", "3001"] if shop == "BODY" else ["3093", "2001", "3001"] if shop == "XMD" else []),
                  "qty": qty, "expectedQty": qty if qty is not None else 0.0,
                  "station": str(obj.get("station") or "").strip().upper(),
                  "uom": str(obj.get("uom") or obj.get("unit") or "").strip().upper(),
                  "year": year, "bin": bin_value, "obNo": str(obj.get("obNo") or ""),
                  "inboundNo": str(obj.get("inboundNo") or obj.get("inbound") or ""),
                  "poNo": str(obj.get("poNo") or obj.get("po") or ""), "poItem": str(obj.get("poItem") or obj.get("item") or ""),
                  "refNo": str(obj.get("refNo") or obj.get("ref") or obj.get("bill") or ""),
                  "materialDocument": str(obj.get("materialDocument") or obj.get("docNo") or ""),
                  "folderPath": str(obj.get("folderPath") or obj.get("folder") or ""),
                  "attachmentItems": normalized_items, "attachmentDocs": docs,
                  "explicitPlant": bool(obj.get("plant"))}
        # Preserve structured order inputs; their service performs business validation.
        for key in ("orders", "nvlOrders", "demands", "items", "containerNo", "batchId", "planId", "warehouse", "sloc"):
            if key in obj:
                result[key] = obj[key]
        if "orders" in obj and "nvlOrders" not in result:
            result["nvlOrders"] = obj["orders"]
        elif "nvlOrders" in obj and "orders" not in result:
            result["orders"] = obj["nvlOrders"]
        return result

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

        # Structured requests are parsed once and never enter the natural-language
        # heuristic path. A missing/malformed command must not become a SAP write.
        import json
        marker_m = re.search(r"SAP_REQUEST_BEGIN\s*(.*?)\s*SAP_REQUEST_END", cleaned, re.DOTALL | re.IGNORECASE)
        looks_structured = isinstance(raw_body, dict) or bool(marker_m) or cleaned.lstrip().startswith(("{", "["))
        candidate = marker_m.group(1) if marker_m else cleaned.strip()
        if json_obj is None and not looks_structured:
            first, last = cleaned.find("{"), cleaned.rfind("}")
            if first >= 0 and last > first and re.search(r'"(?:command|requestId|schemaVersion|materials|RAW_TEXT)"\s*:', cleaned[first:last+1]):
                looks_structured = True
                candidate = cleaned[first:last+1]
        if looks_structured:
            try:
                if json_obj is None:
                    json_obj = json.loads(candidate)
                if not isinstance(json_obj, dict):
                    raise ValueError("Yêu cầu JSON phải là một object")
                return cls._parse_json_request(json_obj, default_plant)
            except (ValueError, TypeError, OverflowError) as exc:
                return {"valid": False, "needHelpReply": True, "command": "INVALID",
                        "detectedIntent": "INVALID_JSON", "reason": str(exc),
                        "explicitWriteIntent": False, "executionMode": "READ_ONLY"}

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

        # Nhận diện bảng yêu cầu cấp NVL (nhiều dòng: Mã NVL, Số lượng, Trạm)
        nvl_table_rows = cls.parse_nvl_table_rows(command_text)
        has_explicit_ra_don = any(k in norm_text for k in ["ra don", "don nvl", "ke hoach nvl", "phan bo nvl", "nvl v31", "nvl zone"])
        is_att_or_doc_intent = best_intent in ("CHECK_ATT", "CHECK_ATT_DOC", "ATT_MIGO")

        # BẢO VỆ CHẶN NHẦM Ý ĐỊNH (Layer 1 Intent Lock):
        # Nếu câu lệnh đã được nhận diện là kiểm tra/đính kèm file (CHECK_ATT, ZFIR266, ATT_MIGO)
        # hoặc có từ khóa đính kèm (att mail, headtext, bktxt, dinh kem), TUYỆT ĐỐI KHÔNG để nvl_table_rows ghi đè sang RA_DON_NVL
        # trừ phi có từ khóa ra đơn tường minh (has_explicit_ra_don).
        has_att_keywords = bool(re.search(r"\b(?:att\s*mail|attach\s*mail|dinh\s*kem|headtext|bktxt|zfir266|bbbg|vat)\b", norm_text))
        if (is_att_or_doc_intent or has_att_keywords) and not has_explicit_ra_don:
            nvl_table_rows = []

        if nvl_table_rows:
            if has_explicit_ra_don or (best_intent in ("RA_DON_NVL", "KE_HOACH_NVL") and not is_att_or_doc_intent):
                best_intent = "RA_DON_NVL"
                best_score = max(best_score, 95)
                matched_keywords.append(f"RA_DON_NVL: Bảng yêu cầu cấp NVL ({len(nvl_table_rows)} dòng)")
            elif best_score == 0 and len(nvl_table_rows) >= 2 and not has_att_keywords:
                best_intent = "RA_DON_NVL"
                best_score = 90
                matched_keywords.append(f"RA_DON_NVL: Tự động nhận diện bảng NVL ({len(nvl_table_rows)} dòng)")

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

        # Hỗ trợ nhận diện nhiều Model cùng chạy trong ca (ví dụ: "S6 nay chạy VF6 Limo")
        known_models_list = ["LIMO", "VF3", "VF5", "VF6", "VF7", "VF8", "VF9", "VFE34", "WILD", "KLARA", "FELIZ", "VENTO", "THEON", "EVO"]
        detected_multi_models = []
        for km in known_models_list:
            if re.search(r"\b" + km + r"\b", norm_text.upper()):
                detected_multi_models.append(km)
        if len(detected_multi_models) > 1:
            model_name = ", ".join(detected_multi_models)

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

        shift_val = ""
        shift_m = re.search(r"\b(?:ca\s*|shift\s*)(s[1-7]|[1-7])\b|\b(s[1-7])\b", norm_text)
        if shift_m:
            shift_val = (shift_m.group(1) or shift_m.group(2)).upper()
            if not shift_val.startswith("S") and shift_val.isdigit():
                shift_val = f"S{shift_val}"

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

        date_val = ""
        date_m = re.search(r"\b(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?)\b", command_text)
        if date_m:
            date_val = date_m.group(1).strip()
        elif re.search(r"\b(?:hom\s*nay|nay|today)\b", norm_text):
            import datetime
            date_val = datetime.date.today().strftime("%Y-%m-%d")

        if best_intent == "IN_WHT" and not ob_no and not mat_doc:
            wht_doc_m = re.search(r"\b(10\d{8}|50\d{8})\b", command_text)
            if wht_doc_m:
                ob_no = wht_doc_m.group(1).strip()

        # Bóc tách đường dẫn thư mục và danh sách chứng từ đính kèm (nếu có lệnh ATT_MIGO)
        from app.sap_attachment import clean_folder_path, extract_folder_path_multiline, DEFAULT_MIGO_ATT_FOLDER
        raw_folder = extract_folder_path_multiline(command_text)
        folder_path = ""
        if raw_folder:
            folder_path = clean_folder_path(raw_folder)
        elif re.search(r"\b(?:att\s*mail|attach\s*mail|dinh\s*kem\s*mail)\b", norm_text):
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
        has_phieu_intent = bool(re.search(r"\b(?:dien\s*so\s*phieu|dien\s*phieu|dien\s*headtext|cap\s*nhat\s*headtext|dien\s*bktxt|att\s*mail|attach\s*mail|dinh\s*kem\s*mail)\b", norm_text))
        if best_intent not in ("CHECK_ATT", "CHECK_ATT_DOC") and (has_phieu_intent or best_intent == "ATT_MIGO") and (att_items or att_docs):
            best_intent = "ATT_MIGO"
            best_score = max(best_score, 35)
            matched_keywords.append("ATT_MIGO: Tự động nhận diện lệnh điền số phiếu HeadText / đính kèm MIGO")

        # 3. Bóc tách các tham số phụ: PLANT, BATCH, SỐ NGÀY
        explicit_plant = False
        plant = default_plant
        # 1. Bắt theo từ khóa: chấp nhận plant, plan (gõ thiếu t), plat, pl, werks, nha may, kho
        plant_m = re.search(r"\b(?:plant|plan|plat|pl|werks|nha\s*may|kho)\s*[:=-]?\s*(\d{4})\b", norm_text)
        if plant_m:
            plant = plant_m.group(1)
            explicit_plant = True
        else:
            # 2. Dự phòng: Nếu trong câu có đứng độc lập một mã Plant 4 số thuộc whitelist (3093, 4007, 4004, 2001, 3001, 2301, 4001)
            known_plants = {"3093", "4007", "4004", "2001", "3001", "2301", "4001"}
            for token in re.findall(r"\b(\d{4})\b", norm_text):
                if token in known_plants:
                    plant = token
                    explicit_plant = True
                    break

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
        if plant_m or explicit_plant:
            work_str = re.sub(r"\b(?:PLANT|PLAN|PLAT|PL|WERKS|NHA\s*MAY|NHÀ\s*MÁY|KHO)\s*[:=-]?\s*" + re.escape(plant), " ", work_str, flags=re.I)
            work_str = re.sub(r"\b" + re.escape(plant) + r"\b", " ", work_str)
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

        tokens = re.findall(r"[A-Z0-9._*-]{3,25}", work_str)
        materials = []
        for tk in tokens:
            clean_tk = tk.strip(".-_ ")
            if not clean_tk or set(clean_tk) <= {"*"}:
                continue
            has_wildcard = "*" in clean_tk or "?" in clean_tk
            min_len = 3 if has_wildcard else 5
            if len(clean_tk) < min_len or clean_tk in cls.STOP_WORDS:
                continue
            if re.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$", clean_tk):
                continue
            if (any(ch.isdigit() for ch in clean_tk) or (has_wildcard and len(clean_tk) >= 3)) and clean_tk != plant and clean_tk != batch:
                if clean_tk not in materials:
                    materials.append(clean_tk)

        # Nếu có bảng đơn hàng NVL đa dòng, ưu tiên mã lấy từ bảng (tránh ô Trạm bị nhận nhầm thành mã)
        if nvl_table_rows and best_intent in ("RA_DON_NVL", "KE_HOACH_NVL"):
            materials = list(dict.fromkeys(r["material"] for r in nvl_table_rows))

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

        # Bảo vệ: Nếu là lệnh CHECK_ATT / CHECK_ATT_DOC, không nạp folderPath để tránh nhầm với tác vụ upload
        if best_intent in ("CHECK_ATT", "CHECK_ATT_DOC"):
            folder_path = ""

        # 5. Xử lý Thông minh (Smart Fallback & Coaching):
        if best_score == 0 and att_items and all(item.get("ticket_no") for item in att_items):
            return {**meta, "valid": False, "needHelpReply": True,
                    "detectedIntent": "AMBIGUOUS_DOCUMENT_LIST", "explicitWriteIntent": False,
                    "reason": "Cần ghi rõ ATT MIGO hoặc điền số phiếu để yêu cầu thay đổi chứng từ."}

        is_no_mat_allowed = (
            best_score > 0 and (
                best_intent == "WHO_IS_TAI"
                or best_intent == "CHECK_RES"
                or best_intent == "TON_THAP"
                or (best_intent == "CHECK_OB" and ob_no)
                or (best_intent == "LICHSU_BIN" and bin_val)
                or (best_intent == "VITRI" and bin_val)
                or (best_intent == "CHECK_IB" and inbound_no) or best_intent == "IB_TREO"
                or (best_intent == "TIM_IB" and (po_no or ref_no))
                or (best_intent == "CHECK_PO" and po_no)
                or (best_intent == "TIM_GR_DOC" and (inbound_no or ref_no or mat_doc))
                or (best_intent == "CHECK_QUEUE")
                or (best_intent == "ATT_MIGO")
                or (best_intent in ("CHECK_ATT", "CHECK_ATT_DOC") and (att_docs or raw_mats))
                or (best_intent in ("TAO_BBBG", "XEM_BBBG", "DIEU_PHOI_OB", "TRA_CUU_ASN", "TINH_TRANG_NVL", "IN_WHT", "TRA_CUU_CAP_NVL"))
                or (best_intent in ("RA_DON_NVL", "KE_HOACH_NVL") and bool(nvl_table_rows))
            )
        ) or (best_score == 0 and bool(bin_val))

        if best_intent in ("CHECK_ATT", "CHECK_ATT_DOC") and not (att_docs or raw_mats):
            meta.update({
                "valid": False,
                "needHelpReply": bool(meta["messageId"]),
                "detectedIntent": best_intent,
                "reason": "Vui lòng nhập ít nhất một số chứng từ DOC (10 chữ số) để kiểm tra đính kèm."
            })
            return meta

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

        if best_intent in ("ATT_MIGO", "WHO_IS_TAI", "CHECK_ATT", "CHECK_ATT_DOC"):
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
            "shift": shift_val,
            "date": date_val,
            "targetDate": date_val,
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
            "fullModelList": is_full_model_list,
            "nvlOrders": nvl_table_rows,
            "orders": nvl_table_rows,
            "qty": (sum(r["qty"] for r in nvl_table_rows) if nvl_table_rows else (expected_qty if eqty_m else None)),
            "station": (
                (list(dict.fromkeys(r["station"] for r in nvl_table_rows if r.get("station")))[0]
                 if len(dict.fromkeys(r["station"] for r in nvl_table_rows if r.get("station"))) == 1
                 else ("MULTI_STATION" if any(r.get("station") for r in nvl_table_rows) else ""))
                if nvl_table_rows
                else (re.search(r"\b(?:station|tram)\s*[:=]?\s*([a-z0-9_-]+)", norm_text).group(1).upper()
                      if re.search(r"\b(?:station|tram)\s*[:=]?\s*([a-z0-9_-]+)", norm_text) else "")
            ),
            "uom": (
                "PC" if nvl_table_rows
                else (re.search(r"\b(?:uom|don\s*vi)\s*[:=]?\s*([a-z0-9]+)", norm_text).group(1).upper()
                      if re.search(r"\b(?:uom|don\s*vi)\s*[:=]?\s*([a-z0-9]+)", norm_text) else "")
            ),
            "year": (
                re.search(r"\b(?:year|nam)\s*[:=]?\s*(\d{4})\b", norm_text).group(1)
                if re.search(r"\b(?:year|nam)\s*[:=]?\s*(\d{4})\b", norm_text)
                else (re.search(r"\b(202[0-9])\b", command_text).group(1) if (best_intent in ("CHECK_ATT", "CHECK_ATT_DOC") and re.search(r"\b(202[0-9])\b", command_text)) else "")
            ),
            "dryRun": bool(re.search(r"\b(?:xem\s*truoc|dry\s*run|chay\s*thu|simulate)\b", norm_text)),
            "explicitWriteIntent": best_intent in ("ATT_MIGO", "RA_DON_NVL"),
            "executionMode": (
                "WRITE_PREVIEW" if (best_intent in ("ATT_MIGO", "RA_DON_NVL") and bool(re.search(r"\b(?:xem\s*truoc|dry\s*run|chay\s*thu|simulate)\b", norm_text)))
                else ("WRITE_LIVE" if best_intent in ("ATT_MIGO", "RA_DON_NVL") else "READ_ONLY")
            )
        })
        return meta

    @classmethod
    def parse_from_mail(cls, raw_body: str, html_body: str = "", default_plant: str = "3093") -> Optional[Dict[str, Any]]:
        """Hàm alias tương thích ngược."""
        return cls.parse(raw_body, html_body=html_body, subject="", default_plant=default_plant)


# Alias tương thích ngược
EmailCommandParser = SmartCommandParser
NaturalCommandParser = SmartCommandParser


def classify_sap_query(raw_text: str) -> dict:
    """
    Phan tich cau hoi tieng Viet tu nhien cua nguoi dung va tra ve:
    - intent: Nghiep vu can goi
    - tcode: T-Code SAP tuong ung
    - days: So ngay tra cuu (neu co, mac dinh 7 ngay)
    """
    parsed = SmartCommandParser.parse(raw_text)
    if not parsed or not parsed.get("valid"):
        return {"intent": "UNKNOWN", "tcode": "/n", "days": 7}
    tcode_map = {
        "WHO_IS_TAI": "NONE (Easter Egg)",
        "CHECK_ATT": "/nZFIR266 (Fallback /nMIGO)",
        "CHECK_ATT_DOC": "/nZFIR266 (Fallback /nMIGO)",
        "ATT_MIGO": "/nMIGO (A12 Change / GOS)",
        "CHECK_BOM": "/nCS03",
        "CHECK_QUEUE": "/n/SCWM/MON",
        "CHECK_OB": "/nVL06O",
        "CHECK_IB": "/nVL06I -> /nZMME240",
        "STO": "/nME2N",
        "LICHSU": "/nMB51",
        "MB52": "/nMB52",
        "TON": "/nMB52",
        "TON_KHO": "/nMB52",
        "VITRI": "/n/SCWM/POST",
        "CHECK_RES": "/nMB25"
    }
    cmd = parsed.get("command", "VITRI")
    return {
        "intent": cmd,
        "tcode": tcode_map.get(cmd, "/n/SCWM/POST"),
        "days": parsed.get("days", 7)
    }
