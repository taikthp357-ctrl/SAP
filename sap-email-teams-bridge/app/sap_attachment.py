# -*- coding: utf-8 -*-
"""Mô đun tự động đính kèm hóa đơn VAT, Biên bản bàn giao (BBBG), và Email (.msg/.eml) vào MIGO qua GOS,
đồng thời tự động điền số phiếu vào trường Header Text (txtGOHEAD-BKTXT) trên SAP MIGO.

Kế thừa và tối ưu hóa vượt bậc từ bộ công cụ 261002_MIGO_ATT_V3_4_3 & 260801 Tools Attach Mail:
1. Tự động điền số phiếu (HeadText GOHEAD-BKTXT): Kiểm tra, điền mới và lưu an toàn qua phím SendVKey(11) (Ctrl+S).
2. Đính kèm đa định dạng: Hỗ trợ BBBG, VAT, và Email trao đổi / xác nhận (.msg, .eml, .pdf).
3. Tối ưu quét thư mục UNC / Local: Lập chỉ mục trong RAM 1 lần duy nhất bằng os.scandir (O(1)).
4. Bóc tách thông minh từ Email / Tin nhắn: Tự động trích xuất danh sách cặp (DOC, Số phiếu) từ email nhận được.
5. Độc lập phiên SAP (Multi-Session): Không chiếm dụng Tab 0 của người dùng; tự động dùng Session riêng.
6. Báo cáo trực quan tương thích 100% với định dạng thông báo Teams/Email (Ảnh 2).
"""

import os
from pathlib import Path
import re
import time
import unicodedata
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple, Set, Callable


DEFAULT_MIGO_ATT_FOLDER = (
    r"\\10.128.10.72\fvn$\03.Khối sản xuất\01.Nhà máy VinFast Hải Phòng"
    r"\04.Xưởng lắp ráp kim loại - cơ khí NMHP\01. Lắp ráp kim loại"
    r"\5.LOGISTIC\III. Internal\III. Kho P part\3. Theo doi hang nhap P"
    r"\240101_FT-P\260802_ File ATT BBGH"
)


def clean_folder_path(raw_path: str) -> str:
    """Chuẩn hóa đường dẫn thư mục UNC / Local: xóa ký tự ẩn NBSP, zero-width, nháy, và chuẩn hóa Unicode NFC.
    Tự động hoàn thiện đường dẫn nếu bị Outlook ngắt dòng ở ký tự thứ 73-76.
    """
    if not raw_path:
        return ""
    p = str(raw_path)
    p = p.replace("\xa0", " ").replace("\u200b", "").replace("\u3000", " ")
    p = p.strip(" \t\r\n\"\'<>|")
    p = unicodedata.normalize("NFC", p)
    if len(p) > 3 and p.endswith(("\\", "/")):
        p = p.rstrip("\\/")
    # Tự động bù đắp nếu bị ngắt dòng cụt do Outlook (ví dụ: '... \01.Nhà máy VinFast Hải')
    if len(p) >= 20 and DEFAULT_MIGO_ATT_FOLDER.startswith(p) and len(p) < len(DEFAULT_MIGO_ATT_FOLDER):
        return DEFAULT_MIGO_ATT_FOLDER
    return p


def extract_folder_path_multiline(text: str) -> str:
    """Bóc tách đường dẫn thư mục UNC hoặc Local Drive, tự động ghép các dòng bị ngắt bởi Outlook/Teams."""
    if not text:
        return ""
    m = re.search(r"(?:\\\\|[A-Za-z]:\\)", text)
    if not m:
        return ""
    start_pos = m.start()
    remaining = text[start_pos:]
    lines = remaining.splitlines()
    segments = []
    for line in lines:
        l = line.strip(" \t\r\n\"\'<>|")
        if not l:
            break
        # Dừng nếu dòng bắt đầu bằng 10 chữ số chứng từ (bước sang danh sách)
        if re.match(r"^(?:49\d{8}|500\d{7}|\d{10})\b", l):
            break
        # Cắt nếu trên cùng dòng có số chứng từ ở phía sau
        doc_b = re.search(r"[\s\t]+(?=(?:49\d{8}|500\d{7}|\d{10})\b)", l)
        if doc_b:
            segments.append(l[:doc_b.start()].strip(" \t\r\n\"\'<>|"))
            break
        segments.append(l)
        if len(segments) >= 15:
            break

    if not segments:
        return ""

    full = segments[0]
    for seg in segments[1:]:
        if seg.startswith("\\") or full.endswith("\\"):
            full = full.rstrip("\\") + "\\" + seg.lstrip("\\")
        else:
            full = full + " " + seg
    return full


def normalize_text(text: str) -> str:
    """Chuẩn hóa chuỗi văn bản: bỏ dấu tiếng Việt, chuyển chữ thường, loại bỏ ký tự lạ."""
    if not text:
        return ""
    text_str = str(text).replace("đ", "d").replace("Đ", "d")
    nfkd = unicodedata.normalize("NFKD", text_str)
    ascii_text = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return ascii_text.lower().strip()


# Regex phát hiện loại file chứng từ
BBBG_REGEX = re.compile(r"(?:^|[^a-z0-9])(bbbg|bbgh|bbg|bb|bien[ _-]*ban)(?:$|[^a-z0-9])", re.I)
VAT_REGEX = re.compile(r"(?:^|[^a-z0-9])(vat|invoice|inv|hoa[ _-]*don|hd|hddt|gtgt|tax|ct|chung[ _-]*tu)(?:$|[^a-z0-9])", re.I)
MAIL_REGEX = re.compile(r"(?:^|[^a-z0-9])(mail|email|thu|xac[ _-]*nhan|trao[ _-]*doi|order)(?:$|[^a-z0-9])", re.I)


def extract_search_tokens(text: str) -> Tuple[List[str], str]:
    """Bóc tách các token tìm kiếm từ số phiếu hoặc số chứng từ.
    
    Ví dụ: 'Thành Đạt in 6157' -> nums: ['6157'], compact: 'thanhdatin6157'
    '5006995954' -> nums: ['5006995954'], compact: '5006995954'
    """
    norm = normalize_text(text)
    nums = re.findall(r"\d{3,10}", norm)
    compact = re.sub(r"[^a-z0-9]+", "", norm)
    return nums, compact


def format_status_label(ticket_no: str, is_success: bool) -> str:
    """Định dạng chuỗi trạng thái thành công/thất bại: 'DONE' nếu thành công, 'FILE_NOT_FOUND' nếu thiếu file."""
    if not is_success:
        return "FILE_NOT_FOUND"
    return "DONE"


def parse_attachment_items_from_text(text: str) -> List[Dict[str, str]]:
    """Bóc tách danh sách chứng từ và số phiếu từ nội dung email hoặc tin nhắn Teams.
    
    Hỗ trợ mọi trường hợp:
    - Danh sách nhiều dòng từ Teams / Outlook
    - Bảng copy-paste từ Excel (phân cách bằng tab, dấu cách, hoặc dấu gạch |)
    - Định dạng nhiều cặp chứng từ trên cùng một dòng hoặc toàn bộ chuỗi
    - Bảng HTML từ Teams/Outlook
    """
    if not text:
        return []

    # 1. Làm sạch tag HTML nếu có
    clean_text = re.sub(r"</?(?:tr|div|p|br\s*/?|li)>", "\n", str(text), flags=re.I)
    clean_text = re.sub(r"</td>", "\t", clean_text, flags=re.I)
    clean_text = re.sub(r"<[^>]+>", " ", clean_text)

    # 2. Xóa các đường dẫn UNC hoặc thư mục cục bộ để số trong đường dẫn không bị nhận nhầm thành số DOC
    clean_text = re.sub(r"(?:\\\\[^\r\n\"\'<>|]+|[A-Za-z]:\\[^\r\n\"\'<>|]+)", " ", clean_text)
    clean_text = re.sub(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b", " ", clean_text)

    # 3. Quét tất cả các số chứng từ 10 số (bắt đầu bằng 49, 500 hoặc 10 chữ số)
    doc_matches = list(re.finditer(r"\b(49\d{8}|500\d{7}|\d{10})\b", clean_text))
    if not doc_matches:
        return []

    items = []
    seen_docs = set()

    for idx, m in enumerate(doc_matches):
        doc = m.group(1)
        if doc in seen_docs:
            continue
        seen_docs.add(doc)

        start_pos = m.end()
        # Số phiếu của chứng từ này nằm từ sau doc hiện tại tới vị trí bắt đầu doc tiếp theo
        if idx + 1 < len(doc_matches):
            raw_cand = clean_text[start_pos:doc_matches[idx + 1].start()]
        else:
            raw_cand = clean_text[start_pos:]
            # Nếu là phần tử cuối cùng, loại bỏ phần chữ ký / footer email nếu có
            sig_match = re.search(
                r"(?i)\n\s*(?:người yêu cầu|người gửi|trân trọng|thân ái|regards|best regards|thanks|sent from|sender_|from:|to:|subject:|date:|\-{3,}|\_{3,})",
                raw_cand
            )
            if sig_match:
                raw_cand = raw_cand[:sig_match.start()]

        # Outlook / Power Automate thường tự động ngắt dòng ở cột 76-80 ký tự (hard wrap)
        # Nối tất cả các dòng bị ngắt giữa doc[i] và doc[i+1] thành một chuỗi liên tục
        clean_ticket = " ".join(raw_cand.split()).strip(" :-\t,|\\/")
        for _ in range(4):
            clean_ticket = re.sub(
                r"^(?:SAP|BOT|LENH|LỆNH|ĐIỀN\s*SỐ\s*PHIẾU|DIEN\s*SO\s*PHIEU|ĐIỀN\s*PHIẾU|DIEN\s*PHIEU|SỐ\s*PHIẾU|SO\s*PHIEU|PHIẾU|PHIEU|HEADTEXT|BKTXT|ATTACH|ATT|MAIL|EMAIL)\s*[:=]?\s*",
                "",
                clean_ticket,
                flags=re.I
            ).strip(" :-\t,|\\/")

        items.append({"doc": doc, "ticket_no": clean_ticket})

    return items


class MigoAttachmentRunner:
    """Điều khiển tự động hóa điền số phiếu và đính kèm BBBG / VAT / Email vào SAP MIGO qua GOS."""

    def __init__(self, sap_engine=None):
        self.sap_engine = sap_engine

    @staticmethod
    def index_source_folder(folder_path: str, recursive: bool = True) -> Dict[str, Any]:
        """Quét và lập chỉ mục toàn bộ file PDF/Ảnh/Email (.msg/.eml) trong thư mục nguồn vào RAM (O(1))."""
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            if str(folder_path).startswith("\\\\"):
                raise FileNotFoundError(
                    f"Thư mục chia sẻ mạng UNC '{folder_path}' không tồn tại hoặc máy trạm chưa kết nối mạng nội bộ/VPN (chưa xác thực quyền truy cập Windows)."
                )
            raise FileNotFoundError(f"Thư mục nguồn không tồn tại hoặc không thể truy cập: {folder_path}")

        files_indexed: List[Dict[str, Any]] = []
        token_to_files: Dict[str, List[int]] = {}

        valid_exts = {".pdf", ".jpg", ".jpeg", ".png", ".xlsx", ".docx", ".tif", ".tiff", ".msg", ".eml"}

        def scan_dir(dir_path: Path, depth: int = 0):
            if depth > 4:
                return
            try:
                with os.scandir(dir_path) as it:
                    for entry in it:
                        try:
                            if entry.is_file() and Path(entry.name).suffix.lower() in valid_exts:
                                fpath = Path(entry.path)
                                norm_name = normalize_text(fpath.stem)
                                ext = fpath.suffix.lower()

                                is_bbbg = bool(BBBG_REGEX.search(norm_name))
                                is_vat = bool(VAT_REGEX.search(norm_name))
                                is_mail = ext in {".msg", ".eml"} or bool(MAIL_REGEX.search(norm_name))

                                nums = re.findall(r"\d{3,10}", norm_name)
                                compact = re.sub(r"[^a-z0-9]+", "", norm_name)

                                file_idx = len(files_indexed)
                                files_indexed.append({
                                    "path": fpath,
                                    "name": fpath.name,
                                    "norm_name": norm_name,
                                    "compact": compact,
                                    "nums": set(nums),
                                    "is_bbbg": is_bbbg,
                                    "is_vat": is_vat,
                                    "is_mail": is_mail
                                })

                                for num in nums:
                                    token_to_files.setdefault(num, []).append(file_idx)

                            elif entry.is_dir() and recursive:
                                scan_dir(Path(entry.path), depth + 1)
                        except Exception:
                            continue
            except Exception as e:
                print(f"[INDEX WARN] Không thể quét thư mục '{dir_path}': {e}")

        scan_dir(folder, depth=0)
        return {
            "folder": str(folder),
            "files": files_indexed,
            "token_map": token_to_files,
            "total_files": len(files_indexed)
        }

    @classmethod
    def match_all_files(cls, doc: str, ticket_no: str, index: Dict[str, Any]) -> Dict[str, List[Path]]:
        """Tìm tất cả các file BBBG, VAT (hỗ trợ nhiều file VAT) và Mail phù hợp."""
        files = index.get("files", [])
        token_map = index.get("token_map", {})

        doc_clean = str(doc).strip()
        ticket_clean = str(ticket_no).strip()

        doc_nums, doc_compact = extract_search_tokens(doc_clean)
        ticket_nums, ticket_compact = extract_search_tokens(ticket_clean)

        search_nums = set(doc_nums + ticket_nums)

        cand_indices: Set[int] = set()
        for num in search_nums:
            if num in token_map:
                cand_indices.update(token_map[num])

        if not cand_indices:
            for idx, f in enumerate(files):
                for num in search_nums:
                    if num in f["compact"]:
                        cand_indices.add(idx)

        cand_list = [files[i] for i in cand_indices]

        # 1. Tìm các file BBBG
        bbbg_files = [f["path"] for f in cand_list if f["is_bbbg"]]

        # 2. Tìm các file VAT
        vat_files = [f["path"] for f in cand_list if f["is_vat"]]

        # 3. Tìm các file Mail
        mail_files = [f["path"] for f in cand_list if f["is_mail"]]

        # Fallback VAT: Nếu không có file nào mang từ khóa VAT rõ ràng,
        # nhưng có file khớp số phiếu và không phải BBBG / Mail -> coi đó là hóa đơn VAT (ký hiệu riêng)
        if not vat_files:
            fallback = [f["path"] for f in cand_list if not f["is_bbbg"] and not f["is_mail"]]
            if fallback:
                vat_files = fallback

        return {
            "bbbg": bbbg_files,
            "vat": vat_files,
            "mail": mail_files
        }

    @classmethod
    def match_files(cls, doc: str, ticket_no: str, index: Dict[str, Any]) -> Tuple[Optional[Path], Optional[Path], Optional[Path], str]:
        """Tìm file BBBG, VAT, và Email phù hợp nhất cho số chứng từ và số phiếu đã cho (tương thích ngược)."""
        all_matches = cls.match_all_files(doc, ticket_no, index)
        bbbg_match = all_matches["bbbg"][0] if all_matches["bbbg"] else None
        vat_match = all_matches["vat"][0] if all_matches["vat"] else None
        mail_match = all_matches["mail"][0] if all_matches["mail"] else None

        notes = []
        if all_matches["bbbg"]:
            notes.append(f"BBBG: {len(all_matches['bbbg'])} file ({', '.join(p.name for p in all_matches['bbbg'][:2])})")
        else:
            notes.append("BBBG: FILE_NOT_FOUND")

        if all_matches["vat"]:
            notes.append(f"VAT: {len(all_matches['vat'])} file ({', '.join(p.name for p in all_matches['vat'][:2])})")
        else:
            notes.append("VAT: FILE_NOT_FOUND")

        if mail_match:
            notes.append(f"Mail: {mail_match.name}")

        return bbbg_match, vat_match, mail_match, " | ".join(notes)

    def handle_migo_headtext(self, session, ticket_no: str) -> Tuple[str, str]:
        """Kiểm tra và tự động điền số phiếu vào trường Header Text (txtGOHEAD-BKTXT) trên SAP MIGO."""
        if not ticket_no or not str(ticket_no).strip():
            return "SKIP_NO_TICKET", "Không có số phiếu cần điền"

        ticket_clean = str(ticket_no).strip()

        try:
            # 1. Bảo đảm phần Header MIGO được mở
            try:
                session.SendCommand("MIGO_OK_HEADER_OPEN")
                time.sleep(0.3)
            except Exception:
                pass

            # 2. Tìm trường GOHEAD-BKTXT (findByName trước, sau đó thử các fallback ID)
            head_field = None
            try:
                head_field = session.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
            except Exception:
                pass
            if head_field is None:
                for cand_id in [
                    "wnd[0]/usr/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-BKTXT",
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-BKTXT",
                    "wnd[0]/usr/txtGOHEAD-BKTXT"
                ]:
                    try:
                        head_field = session.findById(cand_id)
                        if head_field:
                            break
                    except Exception:
                        pass

            if head_field is None:
                return "HEADTEXT_NOT_FOUND", "Không tìm thấy trường GOHEAD-BKTXT trên MIGO"

            current_text = str(getattr(head_field, "Text", "")).strip()

            # Trường hợp 1: Đã có sẵn HeadText và trùng khớp số phiếu
            if current_text:
                curr_tokens, curr_compact = extract_search_tokens(current_text)
                tick_tokens, tick_compact = extract_search_tokens(ticket_clean)
                if curr_compact == tick_compact or (curr_tokens and set(curr_tokens).intersection(set(tick_tokens))):
                    return "HEADTEXT_OK", f"Đã có đúng số phiếu ('{current_text}')"
                else:
                    return "HEADTEXT_DIFF", f"SAP đã có HeadText khác: '{current_text}'"

            # Trường hợp 2: HeadText đang rỗng -> Điền mới và bấm Lưu (Save / vKey 11)
            try:
                head_field.Text = ticket_clean
                session.findById("wnd[0]").SendVKey(11)  # Ctrl+S / Save
                time.sleep(0.6)

                sbar = session.findById("wnd[0]/sbar")
                s_type = getattr(sbar, "MessageType", "")
                s_msg = getattr(sbar, "Text", "")
                if s_type in ("E", "A"):
                    return "HEADTEXT_FAIL", f"Lỗi lưu: {s_msg}"
                return "HEADTEXT_DONE", f"Đã điền và lưu mới HeadText thành công: '{ticket_clean}'"
            except Exception as write_err:
                return "HEADTEXT_WRITE_ERROR", str(write_err)

        except Exception as e:
            return "HEADTEXT_EXCEPTION", str(e)

    @staticmethod
    def check_existing_sap_attachments(titles: List[str], ticket_no: str) -> Tuple[Optional[str], Optional[str]]:
        """Kiểm tra danh sách tiêu đề file đính kèm đã có trên SAP MIGO (từ popup Attachment list).
        
        Đối chiếu với số phiếu đã cho:
        - Tìm file BBBG đã có (VD: '3874 BB', 'BBBG 3874', 'in 3874 BBBG')
        - Tìm file VAT đã có (VD: 'VAT 3874', '3874 VAT', 'in 3874 VAT')
        Trả về: (sap_bbbg_title, sap_vat_title)
        """
        if not titles:
            return None, None

        search_nums, search_compact = extract_search_tokens(ticket_no)
        bbbg_found = None
        vat_found = None

        for title in titles:
            norm = normalize_text(title)
            title_nums, title_compact = extract_search_tokens(title)

            # Kiểm tra xem title trên SAP có chứa số phiếu của chứng từ này không
            is_ticket_match = False
            if search_nums and title_nums:
                if set(search_nums).intersection(set(title_nums)):
                    is_ticket_match = True
            elif search_compact and (search_compact in title_compact or title_compact in search_compact):
                is_ticket_match = True
            elif not search_nums:
                is_ticket_match = True

            if is_ticket_match:
                if not bbbg_found and BBBG_REGEX.search(norm):
                    bbbg_found = title
                if not vat_found and VAT_REGEX.search(norm):
                    vat_found = title

        return bbbg_found, vat_found

    def read_sap_attachment_titles(self, session) -> List[str]:
        """Kiểm tra và đọc danh sách tiêu đề các file đính kèm đã có sẵn trên SAP MIGO qua nút OK_REF_ATTA / Attachment list."""
        titles: List[str] = []
        if session is None:
            return titles

        try:
            # 1. Tìm nút OK_REF_ATTA (Display Attachment) trên màn hình MIGO
            button = None
            for cand_id in [
                "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/btnOK_REF_ATTA",
                "wnd[0]/usr/btnOK_REF_ATTA"
            ]:
                try:
                    button = session.findById(cand_id)
                    if button:
                        break
                except Exception:
                    pass

            if button is None:
                try:
                    button = session.findById("wnd[0]").findByName("OK_REF_ATTA", "GuiButton")
                except Exception:
                    pass

            has_opened = False

            # Nếu tìm thấy nút OK_REF_ATTA và nút đang sáng (Enabled)
            if button is not None:
                enabled = getattr(button, "Enabled", True)
                if enabled is False or str(enabled).upper() == "FALSE":
                    # Nút bị mờ / disabled -> SAP chưa có đính kèm nào trên chứng từ này
                    return []
                try:
                    button.Press()
                    has_opened = True
                except Exception:
                    pass

            # Fallback nếu không bấm được nút OK_REF_ATTA: thử mở qua GOS shell %GOS_VIEW_ATTA
            if not has_opened:
                for shell_id in ["wnd[0]/titl/shellcont/shell", "wnd[0]/titl/shell"]:
                    try:
                        shell = session.findById(shell_id)
                        if shell:
                            try:
                                shell.PressContextButton("%GOS_TOOLBOX")
                                shell.SelectContextMenuItem("%GOS_VIEW_ATTA")
                                has_opened = True
                                break
                            except Exception:
                                pass
                    except Exception:
                        pass

            if not has_opened:
                return []

            time.sleep(0.5)

            # 2. Đọc popup Attachment list (wnd[1])
            popup = None
            try:
                popup = session.findById("wnd[1]")
            except Exception:
                pass

            if popup is not None:
                grid = None
                for cand_grid_id in [
                    "wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell",
                    "wnd[1]/usr/cntlCONTAINER/shellcont/shell",
                    "wnd[1]/usr/cntlCONTAINER_0100/shellcont",
                    "wnd[1]/usr/shell"
                ]:
                    try:
                        g = session.findById(cand_grid_id)
                        if g:
                            grid = g
                            break
                    except Exception:
                        pass

                if grid is None:
                    def find_grid(obj, depth=0):
                        if depth > 10 or obj is None:
                            return None
                        try:
                            c_type = str(getattr(obj, "Type", ""))
                            c_sub = str(getattr(obj, "SubType", ""))
                            if "Grid" in c_type or "Grid" in c_sub or c_type in ("GuiGridView", "GuiTableControl"):
                                return obj
                            children = getattr(obj, "Children", None)
                            if children:
                                for idx in range(children.Count):
                                    res = find_grid(children.Item(idx), depth + 1)
                                    if res:
                                        return res
                        except Exception:
                            pass
                        return None
                    grid = find_grid(popup)

                if grid:
                    try:
                        row_count = int(getattr(grid, "RowCount", 0))
                        for r in range(row_count):
                            title_val = ""
                            for col_name in ["BITM_DESCR", "OBJDES", "TITLE", "DESCRIPT", "FILENAME"]:
                                try:
                                    val = grid.GetCellValue(r, col_name)
                                    if val and str(val).strip():
                                        title_val = str(val).strip()
                                        break
                                except Exception:
                                    continue
                            if title_val:
                                titles.append(title_val)
                    except Exception as e:
                        print(f"[SAP ATT LIST WARN] Lỗi đọc grid: {e}")

                # 3. Đóng popup wnd[1]
                try:
                    popup.Close()
                except Exception:
                    try:
                        session.findById("wnd[1]").SendVKey(12)
                    except Exception:
                        pass

                time.sleep(0.3)
        except Exception as ex:
            print(f"[SAP ATT LIST ERROR] {ex}")

        return titles

    def attach_file_to_current_migo(self, session, file_path: Path) -> Tuple[bool, str]:
        """Đính kèm 1 file vật lý vào chứng từ MIGO đang mở trên SAP qua GOS (%GOS_TOOLBOX)."""
        abs_path = file_path.resolve()
        if not abs_path.exists():
            return False, f"FILE_NOT_FOUND: {abs_path}"

        parent_folder = str(abs_path.parent)
        filename = abs_path.name

        try:
            # 1. Nhấn nút GOS Toolbox
            shell = None
            for shell_id in ["wnd[0]/titl/shellcont/shell", "wnd[0]/titl/shell"]:
                try:
                    shell = session.findById(shell_id)
                    if shell:
                        break
                except Exception:
                    pass

            if shell is None:
                return False, "GOS_SHELL_NOT_FOUND"

            try:
                shell.PressContextButton("%GOS_TOOLBOX")
            except Exception:
                try:
                    shell.PressButton("%GOS_TOOLBOX")
                except Exception:
                    pass

            # 2. Chọn menu Create Attachment (%GOS_PCATTA_CREA)
            try:
                shell.SelectContextMenuItem("%GOS_PCATTA_CREA")
            except Exception:
                shell.SelectMenuItem("%GOS_PCATTA_CREA")

            time.sleep(0.5)

            # 3. Điền đường dẫn và tên file vào popup wnd[1]
            try:
                session.findById("wnd[1]/usr/ctxtDY_PATH").Text = parent_folder
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
                session.findById("wnd[1]/tbar[0]/btn[0]").Press()
            except Exception:
                pass

            time.sleep(0.6)

            # 4. Kiểm tra thông báo trạng thái tại thanh sbar
            sbar = session.findById("wnd[0]/sbar")
            msg = getattr(sbar, "Text", "")
            msg_type = getattr(sbar, "MessageType", "")

            success_keywords = ["SUCCESS", "CREATED", "SAVED", "STORED", "THANH CONG", "DA TAO", "EXISTING"]
            if any(k in msg.upper() for k in success_keywords) or msg_type == "S":
                return True, f"DONE_SAP_ACK: {msg or 'Created successfully'}"
            elif any(k in msg.upper() for k in ["NOT CREATED", "FAILED", "ERROR", "CANCELLED", "CANCELED"]):
                return False, f"SAP_ERROR: {msg}"
            else:
                return True, f"DONE_ACK (msg: {msg})"

        except Exception as e:
            return False, f"GOS_EXCEPTION: {e}"

    def open_migo_document(self, session, doc: str, year: str = "2026", mode: str = "DISPLAY") -> Tuple[bool, str]:
        """Điều hướng SAP tới chứng từ MIGO cần xem (A04 Display) hoặc sửa (A12 Change), R02 Material Document."""
        try:
            try:
                session.findById("wnd[0]").maximize()
            except Exception:
                pass

            wnd_text = getattr(session.findById("wnd[0]"), "Text", "")
            if "Material Document" in wnd_text or "MIGO" in wnd_text:
                try:
                    session.findById("wnd[0]").SendVKey(12)  # F12 Cancel/Back để đóng chứng từ cũ
                    time.sleep(0.3)
                except Exception:
                    pass
            else:
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIGO"
                session.findById("wnd[0]").SendVKey(0)
                time.sleep(0.5)

            # 1. Thiết lập ACTION (A12 Change / A04 Display)
            target_act = "A12" if mode == "CHANGE" else "A04"
            act_cb = None
            try:
                act_cb = session.findById("wnd[0]").findByName("GODYNPRO-ACTION", "GuiComboBox")
            except Exception:
                pass
            if act_cb is None:
                for cand_id in [
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-ACTION",
                    "wnd[0]/usr/cmbGODYNPRO-ACTION"
                ]:
                    try:
                        act_cb = session.findById(cand_id)
                        if act_cb:
                            break
                    except Exception:
                        pass
            if act_cb:
                try:
                    if getattr(act_cb, "Key", "") != target_act:
                        act_cb.Key = target_act
                        time.sleep(0.3)
                except Exception:
                    if mode == "CHANGE":
                        try:
                            act_cb.Key = "A02"
                            time.sleep(0.3)
                        except Exception:
                            pass

            # 2. Thiết lập REFDOC (R02 Material Document)
            ref_cb = None
            try:
                ref_cb = session.findById("wnd[0]").findByName("GODYNPRO-REFDOC", "GuiComboBox")
            except Exception:
                pass
            if ref_cb is None:
                for cand_id in [
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC",
                    "wnd[0]/usr/cmbGODYNPRO-REFDOC"
                ]:
                    try:
                        ref_cb = session.findById(cand_id)
                        if ref_cb:
                            break
                    except Exception:
                        pass
            if ref_cb:
                try:
                    if getattr(ref_cb, "Key", "") != "R02":
                        ref_cb.Key = "R02"
                        time.sleep(0.3)
                except Exception:
                    pass

            # 3. Điền số chứng từ vào GODYNPRO-MAT_DOC
            doc_fld = None
            try:
                doc_fld = session.findById("wnd[0]").findByName("GODYNPRO-MAT_DOC")
            except Exception:
                pass
            if doc_fld is None:
                for cand_id in [
                    "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC",
                    "wnd[0]/usr/txtGODYNPRO-MAT_DOC",
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC",
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/ctxtGODYNPRO-MAT_DOC"
                ]:
                    try:
                        doc_fld = session.findById(cand_id)
                        if doc_fld:
                            break
                    except Exception:
                        pass

            if doc_fld is None:
                return False, "CONTROL_NOT_FOUND: Không tìm thấy ô nhập số chứng từ GODYNPRO-MAT_DOC trên MIGO"

            try:
                doc_fld.SetFocus()
            except Exception:
                pass
            doc_fld.Text = str(doc).strip()

            # 4. Điền DOC_YEAR nếu có
            if year:
                yr_fld = None
                try:
                    yr_fld = session.findById("wnd[0]").findByName("GODYNPRO-DOC_YEAR")
                except Exception:
                    pass
                if yr_fld is None:
                    for cand_id in [
                        "wnd[0]/usr/ctxtGODYNPRO-DOC_YEAR",
                        "wnd[0]/usr/txtGODYNPRO-DOC_YEAR",
                        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-DOC_YEAR"
                    ]:
                        try:
                            yr_fld = session.findById(cand_id)
                            if yr_fld:
                                break
                        except Exception:
                            pass
                if yr_fld:
                    try:
                        yr_fld.Text = str(year).strip()
                    except Exception:
                        pass

            # 5. Gửi phím Enter (VKey 0) để mở chứng từ
            session.findById("wnd[0]").SendVKey(0)
            time.sleep(0.6)

            # 6. Kiểm tra thông báo trạng thái
            try:
                sbar = session.findById("wnd[0]/sbar")
                msg = getattr(sbar, "Text", "")
                msg_type = getattr(sbar, "MessageType", "")
                if msg_type in ("E", "A"):
                    return False, f"DOC_NOT_FOUND: {msg}"
            except Exception:
                pass

            return True, "LOADED_OK"
        except Exception as e:
            return False, f"NAV_ERROR: {e}"

    def process_attachment_batch(
        self,
        folder_path: str,
        items: List[Dict[str, str]],
        year: str = "2026",
        fill_headtext: bool = True,
        dry_run: bool = False,
        progress_callback: Optional[Callable[[int, int, str, int], None]] = None
    ) -> Dict[str, Any]:
        """Xử lý hàng loạt danh sách chứng từ: Điền số phiếu vào HeadText và Đính kèm file BBBG/VAT."""
        clean_folder = clean_folder_path(folder_path)
        is_folder_requested = bool(clean_folder)
        has_folder = False
        folder_error = ""

        if is_folder_requested:
            try:
                folder_obj = Path(clean_folder)
                if folder_obj.exists() and folder_obj.is_dir():
                    has_folder = True
                else:
                    folder_error = f"Thư mục không tồn tại hoặc không thể truy cập: {clean_folder}"
            except Exception as e:
                folder_error = f"Lỗi truy cập thư mục: {e}"

        if has_folder:
            index_data = self.index_source_folder(clean_folder, recursive=True)
            print(f"[ATT RUNNER] Đã quét thư mục '{clean_folder}': Tìm thấy {index_data['total_files']} files hợp lệ.")
        elif is_folder_requested:
            index_data = {"total_files": 0, "files": [], "token_map": {}}
            print(f"[ATT WARN] Không thể truy cập thư mục '{clean_folder}': {folder_error}")
            print(f"[ATT WARN] Vui lòng kiểm tra lại quyền truy cập mạng LAN/VPN VinFast hoặc mở thư mục trong File Explorer.")
        else:
            index_data = {"total_files": 0, "files": [], "token_map": {}}
            print(f"[ATT RUNNER] Chế độ điền số phiếu HeadText (Không đính kèm file).")

        session = None
        if not dry_run and self.sap_engine:
            session = self.sap_engine.ensure_session()

        results = []
        total_success = 0
        total_missing = 0

        total_items = len(items)
        last_progress_time = time.time()
        PROGRESS_INTERVAL_SEC = 180.0  # Cập nhật tiến độ lên Teams mỗi 3 phút

        for idx, item in enumerate(items, 1):
            doc = str(item.get("doc", "")).strip()
            ticket_no = str(item.get("ticket_no", "")).strip()

            # Thông báo tiến độ định kỳ lên Teams sau mỗi 3 phút khi xử lý danh sách lớn (100 - 300 doc)
            now = time.time()
            if progress_callback and total_items > 20 and (now - last_progress_time >= PROGRESS_INTERVAL_SEC):
                last_progress_time = now
                try:
                    progress_callback(idx - 1, total_items, doc, total_success)
                except Exception as ex:
                    print(f"[PROGRESS WARN] Callback error: {ex}")
            if not doc:
                continue

            # =========================================================================
            # TRƯỜNG HỢP 1: NGƯỜI DÙNG KHÔNG YÊU CẦU ĐÍNH KÈM (Chỉ điền số phiếu HeadText)
            # =========================================================================
            if not is_folder_requested:
                headtext_res = "SKIP"
                headtext_msg = "Không có số phiếu cần điền"
                overall_status = "CHƯA_XỬ_LÝ"

                if dry_run or session is None:
                    if fill_headtext and ticket_no:
                        headtext_res = "HEADTEXT_DONE"
                        headtext_msg = f"Mô phỏng: Điền '{ticket_no}'"
                        overall_status = format_status_label(ticket_no, True)
                        total_success += 1
                    else:
                        headtext_res = "SKIP_NO_TICKET"
                        headtext_msg = "Thiếu số phiếu"
                        overall_status = "THIẾU_SỐ_PHIẾU"
                        total_missing += 1
                else:
                    loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE")
                    if not loaded_ok:
                        headtext_res = "OPEN_FAIL"
                        headtext_msg = nav_msg
                        overall_status = f"Lỗi mở DOC: {nav_msg}"
                        total_missing += 1
                    else:
                        if fill_headtext and ticket_no:
                            headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no)
                            if headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK"):
                                overall_status = format_status_label(ticket_no, True)
                                total_success += 1
                            else:
                                overall_status = f"Lỗi HeadText: {headtext_res}"
                                total_missing += 1
                        else:
                            headtext_res = "SKIP_NO_TICKET"
                            headtext_msg = "Không có số phiếu để điền"
                            overall_status = "THIẾU_SỐ_PHIẾU"
                            total_missing += 1

                results.append({
                    "doc": doc,
                    "ticket_no": ticket_no,
                    "headtext_status": headtext_res,
                    "headtext_msg": headtext_msg,
                    "overall_status": overall_status,
                    "bbbg_status": "SKIP_NO_ATTACH",
                    "vat_status": "SKIP_NO_ATTACH",
                    "mail_status": "SKIP_NO_ATTACH",
                    "bbbg_file": None,
                    "vat_file": None,
                    "mail_file": None,
                    "match_notes": headtext_msg
                })
                continue

            # =========================================================================
            # TRƯỜNG HỢP 2: CÓ YÊU CẦU ĐÍNH KÈM NHƯNG THƯ MỤC LỖI TRUY CẬP (Mất mạng / VPN)
            # =========================================================================
            if not has_folder:
                headtext_res = "SKIP"
                headtext_msg = "Không điền HeadText"

                if fill_headtext and ticket_no:
                    if dry_run or session is None:
                        headtext_res = "HEADTEXT_DONE"
                        headtext_msg = f"Mô phỏng: Điền '{ticket_no}'"
                    else:
                        loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE")
                        if loaded_ok:
                            headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no)
                        else:
                            headtext_res = "OPEN_FAIL"
                            headtext_msg = nav_msg

                bbbg_res = "LỖI_THƯ_MỤC"
                vat_res = "LỖI_THƯ_MỤC"
                if headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK"):
                    overall_status = f"{format_status_label(ticket_no, True)} (Lỗi link thư mục)"
                else:
                    overall_status = "LỖI_THƯ_MỤC"
                total_missing += 1

                results.append({
                    "doc": doc,
                    "ticket_no": ticket_no,
                    "headtext_status": headtext_res,
                    "headtext_msg": headtext_msg,
                    "bbbg_file": None,
                    "vat_file": None,
                    "mail_file": None,
                    "bbbg_status": bbbg_res,
                    "vat_status": vat_res,
                    "mail_status": "LỖI_THƯ_MỤC",
                    "overall_status": overall_status,
                    "match_notes": folder_error or "Không thể truy cập thư mục đính kèm"
                })
                continue

            # =========================================================================
            # TRƯỜNG HỢP 3: CÓ THƯ MỤC HỢP LỆ -> QUÉT & ĐÍNH KÈM FILE BBBG, VAT VÀO MIGO
            # =========================================================================
            all_matches = self.match_all_files(doc, ticket_no, index_data)
            bbbg_paths = all_matches["bbbg"]
            vat_paths = all_matches["vat"]
            mail_paths = all_matches["mail"]

            match_notes_list = []
            if bbbg_paths:
                match_notes_list.append(f"BBBG: {len(bbbg_paths)} file")
            else:
                match_notes_list.append("BBBG: FILE_NOT_FOUND")
            if vat_paths:
                match_notes_list.append(f"VAT: {len(vat_paths)} file")
            else:
                match_notes_list.append("VAT: FILE_NOT_FOUND")
            match_notes = " | ".join(match_notes_list)

            headtext_res = "SKIP"
            headtext_msg = "Không điền HeadText"
            bbbg_res = "FILE_NOT_FOUND"
            vat_res = "FILE_NOT_FOUND"
            mail_res = "FILE_NOT_FOUND" if mail_paths else None
            overall_status = "FILE_NOT_FOUND"

            if dry_run or session is None:
                if fill_headtext and ticket_no:
                    headtext_res = "HEADTEXT_DONE"
                    headtext_msg = f"Mô phỏng: Điền '{ticket_no}'"
                bbbg_res = "DONE" if bbbg_paths else "FILE_NOT_FOUND"
                vat_res = f"DONE ({len(vat_paths)})" if len(vat_paths) > 1 else ("DONE" if vat_paths else "FILE_NOT_FOUND")
                if mail_paths:
                    mail_res = "DONE"
                if bbbg_paths or vat_paths or mail_paths:
                    overall_status = "DONE"
                    total_success += 1
                else:
                    overall_status = "FILE_NOT_FOUND"
                    total_missing += 1
            else:
                loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE" if (fill_headtext and ticket_no) else "DISPLAY")
                if not loaded_ok:
                    bbbg_res = nav_msg
                    vat_res = nav_msg
                    overall_status = nav_msg
                    total_missing += 1
                else:
                    # 0. KIỂM TRA ĐÍNH KÈM ĐÃ CÓ SẴN TRÊN SAP MIGO (Display Attachment / OK_REF_ATTA)
                    sap_titles = self.read_sap_attachment_titles(session)
                    sap_bbbg, sap_vat = self.check_existing_sap_attachments(sap_titles, ticket_no)

                    # 1. Điền số phiếu vào Header Text nếu được yêu cầu
                    if fill_headtext and ticket_no:
                        headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no)

                    # 2. Đính kèm file BBBG
                    b_ok = False
                    if sap_bbbg:
                        b_ok = True
                        bbbg_res = "DONE"
                        print(f"[ATT RUNNER] DOC {doc}: BBBG đã có sẵn trên SAP ('{sap_bbbg}'), bỏ qua upload lại.")
                    elif bbbg_paths:
                        succ_bb = []
                        fail_bb = []
                        for bp in bbbg_paths:
                            ok_b, msg_b = self.attach_file_to_current_migo(session, bp)
                            if ok_b:
                                succ_bb.append(bp.name)
                            else:
                                fail_bb.append(f"{bp.name}: {msg_b}")
                        if succ_bb:
                            b_ok = True
                            bbbg_res = f"DONE ({len(succ_bb)})" if len(succ_bb) > 1 else "DONE"
                        else:
                            bbbg_res = f"Lỗi: {'; '.join(fail_bb)}"
                    else:
                        bbbg_res = "FILE_NOT_FOUND"

                    # 3. Đính kèm file VAT (hỗ trợ 1 hoặc nhiều file VAT)
                    v_ok = False
                    if sap_vat:
                        v_ok = True
                        vat_res = "DONE"
                        print(f"[ATT RUNNER] DOC {doc}: VAT đã có sẵn trên SAP ('{sap_vat}'), bỏ qua upload lại.")
                    elif vat_paths:
                        succ_vat = []
                        fail_vat = []
                        for vp in vat_paths:
                            ok_v, msg_v = self.attach_file_to_current_migo(session, vp)
                            if ok_v:
                                succ_vat.append(vp.name)
                            else:
                                fail_vat.append(f"{vp.name}: {msg_v}")
                        if succ_vat:
                            v_ok = True
                            vat_res = f"DONE ({len(succ_vat)})" if len(succ_vat) > 1 else "DONE"
                        else:
                            vat_res = f"Lỗi: {'; '.join(fail_vat)}"
                    else:
                        vat_res = "FILE_NOT_FOUND"

                    # 4. Đính kèm file Mail (nếu có)
                    m_ok = False
                    if mail_paths:
                        ok_m, msg_m = self.attach_file_to_current_migo(session, mail_paths[0])
                        if ok_m:
                            m_ok = True
                            mail_res = "DONE"
                        else:
                            mail_res = f"Lỗi: {msg_m}"

                    # Thành công nếu BBBG, VAT, Mail hoặc HeadText thành công
                    if b_ok or v_ok or m_ok or (headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK")):
                        overall_status = "DONE"
                        total_success += 1
                    else:
                        overall_status = "INCOMPLETE"
                        total_missing += 1

            results.append({
                "doc": doc,
                "ticket_no": ticket_no,
                "headtext_status": headtext_res,
                "headtext_msg": headtext_msg,
                "bbbg_file": ", ".join(p.name for p in bbbg_paths) if bbbg_paths else None,
                "vat_file": ", ".join(p.name for p in vat_paths) if vat_paths else None,
                "mail_file": mail_paths[0].name if mail_paths else None,
                "bbbg_status": bbbg_res,
                "vat_status": vat_res,
                "mail_status": mail_res,
                "overall_status": overall_status,
                "match_notes": match_notes
            })

        return {
            "folder_path": clean_folder if is_folder_requested else "",
            "folder_error": folder_error,
            "total_items": len(items),
            "total_success": total_success,
            "total_missing": total_missing,
            "results": results
        }

    @staticmethod
    def format_teams_report(batch_summary: Dict[str, Any]) -> str:
        """Định dạng báo cáo kết quả đính kèm và điền số phiếu gửi về Teams / Email."""
        results = batch_summary.get("results", [])
        total = batch_summary.get("total_items", len(results))
        success = batch_summary.get("total_success", 0)
        missing = batch_summary.get("total_missing", 0)
        folder_path = batch_summary.get("folder_path", "")
        folder_error = batch_summary.get("folder_error", "")

        lines = []
        if not folder_path:
            lines.append(f"📝 **KẾT QUẢ ĐIỀN SỐ PHIẾU HEADTEXT MIGO**")
            lines.append(f"• Tiến độ: Hoàn tất **{success}/{total}** chứng từ (Lỗi/Thiếu: {missing})")
            lines.append("")
            lines.append("| Số DOC | Số phiếu | Điền HeadText | Trạng thái |")
            lines.append("| :--- | :--- | :--- | :--- |")
            for r in results:
                doc = r["doc"]
                ticket = r["ticket_no"] or "-"
                head = r.get("headtext_status", "-")
                status = r.get("overall_status", "-")
                lines.append(f"| {doc} | {ticket} | {head} | {status} |")
        else:
            lines.append(f"📎 **KẾT QUẢ ĐÍNH KÈM BBBG, VAT & ĐIỀN SỐ PHIẾU MIGO**")
            lines.append(f"• Thư mục nguồn: `{folder_path}`")
            if folder_error:
                lines.append(f"⚠️ **Cảnh báo thư mục:** `{folder_error}`")
            lines.append(f"• Tiến độ: Hoàn tất **{success}/{total}** chứng từ (Thiếu/Lỗi: {missing})")
            lines.append("")
            lines.append("**Danh sách trạng thái từng chứng từ:**")
            for r in results:
                doc = r["doc"]
                status = r["overall_status"]
                lines.append(f"- `{doc}`: {status}")

            lines.append("")
            lines.append("| Số DOC | Số phiếu | Điền HeadText | BBBG | VAT | Trạng thái |")
            lines.append("| :--- | :--- | :--- | :--- | :--- | :--- |")
            for r in results:
                doc = r["doc"]
                ticket = r["ticket_no"] or "-"
                head = r.get("headtext_status", "-")
                bbbg = r.get("bbbg_status", "-")
                vat = r.get("vat_status", "-")
                status = r.get("overall_status", "-")
                lines.append(f"| {doc} | {ticket} | {head} | {bbbg} | {vat} | {status} |")

        return "\n".join(lines)
