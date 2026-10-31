# -*- coding: utf-8 -*-
"""Worker chuyên trách GHI DỮ LIỆU: Điền số phiếu vào Header Text MIGO (A12 Change).
TUYỆT ĐỐI KHÔNG ĐỤNG TỚI GOS / WIN32 MUTEX / ZFIR266.
ĐẢM BẢO 100% LUÔN Ở CHẾ ĐỘ CHANGE (A12), KHÔNG BAO GIỜ BỊ KẸT Ở DISPLAY (A04).
"""
import re
import time
from typing import Tuple
from app.core.sap_primitives import (
    find_by_id_safe,
    wait_idle,
    dismiss_popups,
    open_transaction,
    smart_find,
    get_status,
)


class MigoHeadtextWriter:
    """Chuyên trách nghiệp vụ điền số phiếu vào trường Header Text trên MIGO Change (A12)."""

    @classmethod
    def ensure_change_mode(cls, session) -> bool:
        """Đảm bảo MIGO đang ở đúng Action=A12 (Change) và RefDoc=R02 (Material Document)."""
        if session is None:
            return False

        w0 = find_by_id_safe(session, "wnd[0]")
        titl = find_by_id_safe(session, "wnd[0]/titl")
        titl_text = str(getattr(titl, "Text", getattr(titl, "text", "")) or "").upper()

        # Nếu đang ở Display hoặc màn hình đang mở chứng từ cũ làm khóa ComboBox -> Esc trước
        if "DISPLAY" in titl_text or (titl_text and "CHANGE" not in titl_text and "MATERIAL" in titl_text):
            if w0:
                try:
                    w0.SendVKey(12)  # Esc (Cancel)
                    wait_idle(session, 1.0)
                    dismiss_popups(session)
                except Exception:
                    pass

        # Tìm và chỉnh ComboBox Action GODYNPRO-ACTION về A12
        cand_action_ids = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
            "wnd[0]/usr/cmbGODYNPRO-ACTION",
        ]
        cmb_action = None
        for cid in cand_action_ids:
            cmb_action = find_by_id_safe(session, cid)
            if cmb_action is not None:
                break
        if cmb_action is None:
            cmb_action = smart_find(session, "GODYNPRO-ACTION", "GuiComboBox")

        if cmb_action is not None:
            cur_key = str(getattr(cmb_action, "Key", getattr(cmb_action, "key", "")) or "").strip()
            cur_txt = str(getattr(cmb_action, "Text", getattr(cmb_action, "text", "")) or "").upper()
            if cur_key != "A12" and "CHANGE" not in cur_txt:
                try:
                    cmb_action.Key = "A12"
                    wait_idle(session, 2.0)
                except Exception:
                    try:
                        cmb_action.key = "A12"
                        wait_idle(session, 2.0)
                    except Exception:
                        pass

        # Tìm và chỉnh ComboBox RefDoc GODYNPRO-REFDOC về R02
        cand_ref_ids = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC",
            "wnd[0]/usr/cmbGODYNPRO-REFDOC",
        ]
        cmb_ref = None
        for cid in cand_ref_ids:
            cmb_ref = find_by_id_safe(session, cid)
            if cmb_ref is not None:
                break
        if cmb_ref is None:
            cmb_ref = smart_find(session, "GODYNPRO-REFDOC", "GuiComboBox")

        if cmb_ref is not None:
            cur_ref_key = str(getattr(cmb_ref, "Key", getattr(cmb_ref, "key", "")) or "").strip()
            if cur_ref_key != "R02":
                try:
                    cmb_ref.Key = "R02"
                    wait_idle(session, 1.5)
                except Exception:
                    try:
                        cmb_ref.key = "R02"
                        wait_idle(session, 1.5)
                    except Exception:
                        pass

        return True

    @classmethod
    def fill(
        cls,
        session,
        doc: str,
        year: str = "",
        ticket_no: str = "",
        allow_overwrite: bool = False
    ) -> Tuple[bool, str, str]:
        """Quy trình chuẩn 5 bước nạp MIGO Change và điền số phiếu vào Header Text.
        Trả về (success: bool, status_code: str, message: str).
        """
        if session is None:
            return False, "NO_SESSION", "Chưa kết nối SAP"

        if not ticket_no or not str(ticket_no).strip():
            return False, "SKIP_NO_TICKET", "Không có số phiếu cần điền"

        ticket_clean = str(ticket_no).strip()
        doc_clean = re.sub(r"\D", "", str(doc)).strip()
        year_clean = str(year).strip() or str(time.localtime().tm_year)

        w0 = find_by_id_safe(session, "wnd[0]")
        if not w0:
            return False, "NO_WINDOW", "Không tìm thấy cửa sổ SAP wnd[0]"

        # Bước 1: Mở /nMIGO nếu chưa ở MIGO
        trans = str(getattr(session.Info, "Transaction", "")).upper()
        if trans != "MIGO":
            open_transaction(session, "MIGO")
            wait_idle(session, 3.0)

        # Bước 2: Ép đúng chế độ Change (A12) / Material Document (R02)
        cls.ensure_change_mode(session)

        # Bước 3: Tìm ô nhập số DOC và Năm
        doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")
        # Nếu ô DOC bị khóa hoặc không thấy, Esc để mở khóa
        if doc_fld is None or not getattr(doc_fld, "Changeable", True):
            w0.SendVKey(12)  # Esc
            wait_idle(session, 1.0)
            cls.ensure_change_mode(session)
            doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")

        if doc_fld is None:
            return False, "CONTROL_NOT_FOUND", "Không tìm thấy ô nhập số chứng từ GODYNPRO-MAT_DOC"

        for _ in range(3):
            try:
                doc_fld.Text = doc_clean
                break
            except Exception:
                try:
                    doc_fld.text = doc_clean
                    break
                except Exception:
                    time.sleep(0.2)

        year_fld = smart_find(session, "GODYNPRO-DOC_YEAR", "GuiTextField") or smart_find(session, "GODYNPRO-DOC_YEAR", "GuiCTextField")
        if year_fld is not None:
            for _ in range(3):
                try:
                    year_fld.Text = year_clean
                    break
                except Exception:
                    try:
                        year_fld.text = year_clean
                        break
                    except Exception:
                        time.sleep(0.2)

        # Enter nạp chứng từ
        w0.SendVKey(0)
        wait_idle(session, 4.0)
        dismiss_popups(session)

        # Kiểm tra thông báo lỗi trên thanh trạng thái sbar
        s_txt, s_type = get_status(session)
        if s_type in ("E", "A"):
            w0.SendVKey(12)  # Esc
            return False, "DOC_ERROR", f"Lỗi SAP khi nạp DOC {doc_clean}: {s_txt}"

        # Kiểm tra tiêu đề cửa sổ: PHẢI là Change Material Document (Ảnh 3)
        titl = find_by_id_safe(session, "wnd[0]/titl")
        titl_text = str(getattr(titl, "Text", getattr(titl, "text", "")) or "").upper()
        if "DISPLAY" in titl_text:
            w0.SendVKey(12)  # Esc
            return False, "DISPLAY_MODE_BLOCKED", f"MIGO nạp nhầm vào mục Display ('{titl_text}'). Không thể sửa Header Text."

        # Bước 4: Mở Header và điền số phiếu vào GOHEAD-BKTXT
        try:
            session.SendCommand("MIGO_OK_HEADER_OPEN")
            wait_idle(session, 1.0)
        except Exception:
            pass

        head_fld = smart_find(session, "GOHEAD-BKTXT", "GuiTextField")
        if head_fld is None:
            w0.SendVKey(12)
            return False, "HEADTEXT_NOT_FOUND", "Không tìm thấy trường GOHEAD-BKTXT trên MIGO"

        current_text = str(getattr(head_fld, "Text", getattr(head_fld, "text", "")) or "").strip()
        if current_text:
            curr_compact = re.sub(r"[^A-Za-z0-9]+", "", current_text).upper()
            tick_compact = re.sub(r"[^A-Za-z0-9]+", "", ticket_clean).upper()
            if curr_compact == tick_compact or ticket_clean in current_text:
                w0.SendVKey(12)  # Esc giải phóng chứng từ
                return True, "ALREADY_FILLED", f"Header Text đã có sẵn số phiếu khớp ({current_text})"

            if not allow_overwrite:
                w0.SendVKey(12)  # Esc giải phóng chứng từ
                return False, "HEADTEXT_EXISTS", f"Header Text đã có số phiếu khác ({current_text}), không ghi đè."

        # Điền số phiếu vào ô
        for _ in range(3):
            try:
                head_fld.Text = ticket_clean
                break
            except Exception:
                try:
                    head_fld.text = ticket_clean
                    break
                except Exception:
                    time.sleep(0.2)

        # Bước 5: Lưu chứng từ (Ctrl+S / btn[11] / btn[23])
        w0.SendVKey(11)  # Save
        wait_idle(session, 5.0)
        dismiss_popups(session)

        post_txt, post_type = get_status(session)
        if post_type in ("E", "A"):
            w0.SendVKey(12)
            return False, "SAVE_ERROR", f"Lỗi khi lưu Header Text DOC {doc_clean}: {post_txt}"

        print(f"[MIGO HEADTEXT SUCCESS] Đã lưu thành công số phiếu '{ticket_clean}' cho DOC {doc_clean} ({post_txt}).")
        # Esc thoát chứng từ sạch sẽ cho lần chạy tiếp theo
        try:
            w0.SendVKey(12)
            wait_idle(session, 0.5)
        except Exception:
            pass

        return True, "FILLED", f"Đã điền số phiếu '{ticket_clean}' và lưu thành công ({post_txt or 'OK'})"
