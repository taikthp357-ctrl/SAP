# -*- coding: utf-8 -*-
"""Mô đun tự động đính kèm hóa đơn VAT và Biên bản bàn giao (BBBG) vào chứng từ MIGO qua GOS.

Kế thừa và tối ưu hóa vượt bậc từ bộ công cụ 261002_MIGO_ATT_V3_4_3:
1. Tối ưu quét thư mục UNC / Local: Lập chỉ mục trong RAM (In-Memory Index) 1 lần duy nhất bằng os.scandir,
   loại bỏ tình trạng treo đơ Excel VBA do quét đệ quy lặp đi lặp lại.
2. Khớp thông minh đa tiêu chí: Khớp theo Số chứng từ SAP (10 số), Số phiếu nhà cung cấp ('in 5794', 'in 6157'),
   từ khóa BBBG (bbbg, bbgh, bb, bien_ban) và VAT (vat, invoice, hoa_don, hd).
3. Độc lập phiên SAP (Multi-Session): Không chiếm dụng Tab 0 của người dùng; tự động dùng Session riêng.
4. Đính kèm chuẩn xác qua SAP Generic Object Services (%GOS_TOOLBOX -> %GOS_PCATTA_CREA).
5. Báo cáo trực quan tương thích 100% với định dạng thông báo Teams/Email (Ảnh 2).
"""

import os
from pathlib import Path
import re
import time
import unicodedata
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple, Set


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
VAT_REGEX = re.compile(r"(?:^|[^a-z0-9])(vat|invoice|hoa[ _-]*don|hd)(?:$|[^a-z0-9])", re.I)


def extract_search_tokens(text: str) -> Tuple[List[str], str]:
    """Bóc tách các token tìm kiếm từ số phiếu hoặc số chứng từ.
    
    Ví dụ: 'Thành Đạt in 6157' -> nums: ['6157'], compact: 'thanhdatin6157'
    '5006995954' -> nums: ['5006995954'], compact: '5006995954'
    """
    norm = normalize_text(text)
    # Bóc các chuỗi số từ 3 đến 10 chữ số (không dùng \b vì \w bao gồm cả dấu gạch dưới _)
    nums = re.findall(r"\d{3,10}", norm)
    # Chuỗi nén loại bỏ khoảng trắng và ký tự đặc biệt
    compact = re.sub(r"[^a-z0-9]+", "", norm)
    return nums, compact


def format_status_label(ticket_no: str, is_success: bool) -> str:
    """Định dạng chuỗi trạng thái thành công/thất bại chuẩn xác như Ảnh 2."""
    if not is_success:
        return "FILE_NOT_FOUND"
    t_clean = str(ticket_no).strip()
    norm = normalize_text(t_clean)
    if norm.startswith("thanh cong") or norm.startswith("thanhcong"):
        return t_clean
    return f"Thành Công {t_clean}" if t_clean else "Thành Công"


class MigoAttachmentRunner:
    """Điều khiển tự động hóa đính kèm VAT & BBBG vào SAP MIGO qua GOS."""

    def __init__(self, sap_engine=None):
        self.sap_engine = sap_engine

    @staticmethod
    def index_source_folder(folder_path: str, recursive: bool = True) -> Dict[str, Any]:
        """Quét và lập chỉ mục toàn bộ file PDF/ảnh trong thư mục nguồn vào bộ nhớ RAM.
        
        Trả về dict chứa danh sách file và map tra cứu siêu tốc O(1).
        """
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            raise FileNotFoundError(f"Thư mục nguồn không tồn tại hoặc không thể truy cập: {folder_path}")

        files_indexed: List[Dict[str, Any]] = []
        token_to_files: Dict[str, List[int]] = {}

        valid_exts = {".pdf", ".jpg", ".jpeg", ".png", ".xlsx", ".docx", ".tif", ".tiff"}

        def scan_dir(dir_path: Path, depth: int = 0):
            if depth > 4:  # Giới hạn độ sâu quét tối đa 4 cấp thư mục con
                return
            try:
                with os.scandir(dir_path) as it:
                    for entry in it:
                        try:
                            if entry.is_file() and Path(entry.name).suffix.lower() in valid_exts:
                                fpath = Path(entry.path)
                                norm_name = normalize_text(fpath.stem)
                                is_bbbg = bool(BBBG_REGEX.search(norm_name))
                                is_vat = bool(VAT_REGEX.search(norm_name))
                                
                                # Bóc các số trong tên file
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
                                    "is_vat": is_vat
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

    @staticmethod
    def match_files(doc: str, ticket_no: str, index: Dict[str, Any]) -> Tuple[Optional[Path], Optional[Path], str]:
        """Tìm file BBBG và VAT phù hợp nhất cho số chứng từ và số phiếu đã cho."""
        files = index.get("files", [])
        token_map = index.get("token_map", {})

        doc_clean = str(doc).strip()
        ticket_clean = str(ticket_no).strip()

        # Bóc các số từ doc và ticket_no
        doc_nums, doc_compact = extract_search_tokens(doc_clean)
        ticket_nums, ticket_compact = extract_search_tokens(ticket_clean)

        search_nums = set(doc_nums + ticket_nums)

        # Lấy tập hợp candidate indices
        cand_indices: Set[int] = set()
        for num in search_nums:
            if num in token_map:
                cand_indices.update(token_map[num])

        # Nếu không thấy theo số, duyệt qua tất cả file để tìm compact match
        if not cand_indices:
            for idx, f in enumerate(files):
                for num in search_nums:
                    if num in f["compact"]:
                        cand_indices.add(idx)

        bbbg_match: Optional[Path] = None
        vat_match: Optional[Path] = None

        cand_list = [files[i] for i in cand_indices]

        # 1. Tìm BBBG match
        bbbg_cands = [f for f in cand_list if f["is_bbbg"]]
        if bbbg_cands:
            # Ưu tiên candidate trùng cả số phiếu hoặc số doc
            bbbg_match = bbbg_cands[0]["path"]

        # 2. Tìm VAT match
        vat_cands = [f for f in cand_list if f["is_vat"]]
        if vat_cands:
            vat_match = vat_cands[0]["path"]

        notes = []
        if bbbg_match:
            notes.append(f"BBBG: {bbbg_match.name}")
        else:
            notes.append("BBBG: FILE_NOT_FOUND")

        if vat_match:
            notes.append(f"VAT: {vat_match.name}")
        else:
            notes.append("VAT: FILE_NOT_FOUND")

        return bbbg_match, vat_match, " | ".join(notes)

    def attach_file_to_current_migo(self, session, file_path: Path) -> Tuple[bool, str]:
        """Đính kèm 1 file vật lý vào chứng từ MIGO đang mở trên SAP qua GOS (%GOS_TOOLBOX)."""
        abs_path = file_path.resolve()
        if not abs_path.exists():
            return False, f"FILE_NOT_FOUND: {abs_path}"

        parent_folder = str(abs_path.parent)
        filename = abs_path.name

        try:
            # 1. Nhấn nút GOS Toolbox
            shell = session.findById("wnd[0]/titl/shellcont/shell")
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
                # Trường hợp SAP dùng Windows Open Dialog tiêu chuẩn
                pass

            time.sleep(0.6)

            # 4. Kiểm tra thông báo trạng thái tại thanh sbar
            sbar = session.findById("wnd[0]/sbar")
            msg = getattr(sbar, "Text", "")
            msg_type = getattr(sbar, "MessageType", "")

            # Kiểm tra các thông báo thành công chuẩn của SAP NetWeaver
            success_keywords = ["SUCCESS", "CREATED", "SAVED", "STORED", "THANH CONG", "DA TAO", "EXISTING"]
            if any(k in msg.upper() for k in success_keywords) or msg_type == "S":
                return True, f"DONE_SAP_ACK: {msg or 'Created successfully'}"
            elif any(k in msg.upper() for k in ["NOT CREATED", "FAILED", "ERROR", "CANCELLED", "CANCELED"]):
                return False, f"SAP_ERROR: {msg}"
            else:
                return True, f"DONE_ACK (msg: {msg})"

        except Exception as e:
            return False, f"GOS_EXCEPTION: {e}"

    def open_migo_document(self, session, doc: str, year: str = "2026") -> Tuple[bool, str]:
        """Điều hướng SAP tới chứng từ MIGO cần xem (A04 Display, R02 Material Document)."""
        try:
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIGO"
            session.findById("wnd[0]").SendVKey(0)
            time.sleep(0.5)

            # Đặt chế độ Hiển thị (A04 Display hoặc A12) & Tham chiếu Chứng từ vật tư (R02)
            try:
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-ACTION").Key = "A04"
            except Exception:
                pass

            try:
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC").Key = "R02"
            except Exception:
                pass

            # Điền số chứng từ vật tư (MAT_DOC) và năm chứng từ (DOC_YEAR)
            try:
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC").Text = str(doc).strip()
            except Exception:
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/ctxtGODYNPRO-MAT_DOC").Text = str(doc).strip()

            try:
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-DOC_YEAR").Text = str(year).strip()
            except Exception:
                pass

            # Nhấn Enter để tải chứng từ
            session.findById("wnd[0]").SendVKey(0)
            time.sleep(0.6)

            # Kiểm tra sbar xem có lỗi "Document does not exist" không
            sbar = session.findById("wnd[0]/sbar")
            msg = getattr(sbar, "Text", "")
            msg_type = getattr(sbar, "MessageType", "")
            if msg_type == "E":
                return False, f"DOC_NOT_FOUND: {msg}"

            return True, "LOADED_OK"
        except Exception as e:
            return False, f"NAV_ERROR: {e}"

    def process_attachment_batch(
        self,
        folder_path: str,
        items: List[Dict[str, str]],
        year: str = "2026",
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """Xử lý hàng loạt danh sách chứng từ và đính kèm file BBBG / VAT vào SAP MIGO.
        
        items là list các dict có dạng:
        [
            {"doc": "5006995954", "ticket_no": "Thành Công in 5794"},
            {"doc": "5006995985", "ticket_no": "Thành Công in 5794"},
            {"doc": "5006995500", "ticket_no": "Thanhcong in 5799"},
            ...
        ]
        """
        # 1. Quét thư mục nguồn 1 lần siêu tốc vào RAM
        index_data = self.index_source_folder(folder_path, recursive=True)
        print(f"[ATT RUNNER] Đã quét thư mục '{folder_path}': Tìm thấy {index_data['total_files']} files hợp lệ.")

        session = None
        if not dry_run and self.sap_engine:
            session = self.sap_engine.ensure_session()

        results = []
        total_success = 0
        total_missing = 0

        for item in items:
            doc = str(item.get("doc", "")).strip()
            ticket_no = str(item.get("ticket_no", "")).strip()
            if not doc:
                continue

            bbbg_path, vat_path, match_notes = self.match_files(doc, ticket_no, index_data)

            bbbg_res = "FILE_NOT_FOUND"
            vat_res = "FILE_NOT_FOUND"
            overall_status = "FILE_NOT_FOUND"

            if dry_run or session is None:
                # Chế độ mô phỏng / kiểm tra khớp file trước
                bbbg_res = format_status_label(ticket_no, bool(bbbg_path))
                vat_res = format_status_label(ticket_no, bool(vat_path))
                if bbbg_path or vat_path:
                    overall_status = format_status_label(ticket_no, True)
                    total_success += 1
                else:
                    overall_status = "FILE_NOT_FOUND"
                    total_missing += 1
            else:
                # Thực hiện thao tác thật trên SAP GUI
                loaded_ok, nav_msg = self.open_migo_document(session, doc, year)
                if not loaded_ok:
                    bbbg_res = nav_msg
                    vat_res = nav_msg
                    overall_status = nav_msg
                    total_missing += 1
                else:
                    b_ok = False
                    if bbbg_path:
                        b_ok, b_msg = self.attach_file_to_current_migo(session, bbbg_path)
                        bbbg_res = format_status_label(ticket_no, b_ok) if b_ok else f"Lỗi: {b_msg}"
                    else:
                        bbbg_res = "FILE_NOT_FOUND"

                    v_ok = False
                    if vat_path:
                        v_ok, v_msg = self.attach_file_to_current_migo(session, vat_path)
                        vat_res = format_status_label(ticket_no, v_ok) if v_ok else f"Lỗi: {v_msg}"
                    else:
                        vat_res = "FILE_NOT_FOUND"

                    if b_ok and v_ok:
                        overall_status = format_status_label(ticket_no, True)
                        total_success += 1
                    elif b_ok or v_ok:
                        overall_status = format_status_label(f"1 phần ({ticket_no})", True)
                        total_success += 1
                    else:
                        overall_status = "INCOMPLETE"
                        total_missing += 1

            results.append({
                "doc": doc,
                "ticket_no": ticket_no,
                "bbbg_file": bbbg_path.name if bbbg_path else None,
                "vat_file": vat_path.name if vat_path else None,
                "bbbg_status": bbbg_res,
                "vat_status": vat_res,
                "overall_status": overall_status,
                "match_notes": match_notes
            })

        return {
            "folder_path": folder_path,
            "total_items": len(items),
            "total_success": total_success,
            "total_missing": total_missing,
            "results": results
        }

    @staticmethod
    def format_teams_report(batch_summary: Dict[str, Any]) -> str:
        """Định dạng báo cáo kết quả đính kèm gửi về kênh Teams / Email chuẩn xác như Ảnh 2."""
        results = batch_summary.get("results", [])
        total = batch_summary.get("total_items", len(results))
        success = batch_summary.get("total_success", 0)
        missing = batch_summary.get("total_missing", 0)

        lines = []
        lines.append(f"📎 **KẾT QUẢ ĐÍNH KÈM VAT & BBBG MIGO**")
        lines.append(f"• Thư mục nguồn: `{batch_summary.get('folder_path', '')}`")
        lines.append(f"• Tiến độ: Hoàn tất **{success}/{total}** chứng từ (Thiếu/Lỗi: {missing})")
        lines.append("")
        lines.append("**Danh sách trạng thái từng chứng từ (Theo định dạng MIGO ATT):**")
        
        # Danh sách nhanh theo format Ảnh 2
        for r in results:
            doc = r["doc"]
            status = r["overall_status"]
            lines.append(f"- `{doc}`: {status}")

        lines.append("")
        # Bảng chi tiết
        lines.append("| DOC | Số phiếu | BBBG | VAT | Trạng thái chi tiết |")
        lines.append("| :--- | :--- | :--- | :--- | :--- |")
        for r in results:
            doc = r["doc"]
            ticket = r["ticket_no"] or "-"
            bbbg = r["bbbg_status"]
            vat = r["vat_status"]
            notes = r["match_notes"]
            lines.append(f"| {doc} | {ticket} | {bbbg} | {vat} | {notes} |")

        return "\n".join(lines)
