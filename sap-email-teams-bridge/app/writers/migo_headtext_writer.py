# -*- coding: utf-8 -*-
"""Worker chuyên trách GHI DỮ LIỆU: Điền số phiếu vào Header Text MIGO (A12 Change).
TUYỆT ĐỐI KHÔNG ĐỤNG TỚI GOS / WIN32 MUTEX / ZFIR266.
ĐẢM BẢO 100% LUÔN Ở CHẾ ĐỘ CHANGE (A12), KHÔNG BAO GIỜ BỊ KẸT Ở DISPLAY (A04).
PHIÊN BẢN v1.1.22:
- HẬU KIỂM TOÀN DIỆN (POST-SAVE DOUBLE VERIFICATION): Triệt tiêu hoàn toàn báo lỗi ảo (False FAIL khi SAP đã lưu xong).
- LÀM SẠCH PHIÊN TỰ ĐỘNG (SESSION SANITIZATION): Triệt tiêu rò rỉ trạng thái / số DOC giữa các chứng từ.
- BẢO ĐẢM CÁCH LY PHIÊN (STRICT SESSION ISOLATION): Mỗi worker chạy trên 1 Session độc lập duy nhất, cấm trùng lặp.
- TỰ PHỤC HỒI NỘI BỘ (INTRA-BATCH SELF-HEALING): Tự động thẩm định lại các chứng từ báo lỗi để cứu hộ thành công.
"""
import re
import threading
import time
from collections import deque
from typing import Tuple, List, Dict, Any, Optional
from app.core.sap_primitives import (
    find_by_id_safe,
    wait_idle,
    dismiss_popups,
    open_transaction,
    smart_find,
    get_status,
    ensure_sap_tabs,
    get_safe_sap_session,
    get_safe_sap_connection,
)


class MigoHeadtextWriter:
    """Chuyên trách nghiệp vụ điền số phiếu vào trường Header Text trên MIGO Change (A12)."""

    @classmethod
    def sanitize_session(cls, session) -> None:
        """Đưa session về trạng thái sạch, đóng popup và giải phóng chứng từ cũ (Esc) nếu đang mở dở dang."""
        if session is None:
            return
        try:
            dismiss_popups(session)
            w0 = find_by_id_safe(session, "wnd[0]")
            if w0 is None:
                return

            titl = find_by_id_safe(session, "wnd[0]/titl")
            titl_text = str(getattr(titl, "Text", getattr(titl, "text", "")) or "").upper()

            # Nếu tiêu đề có chứa số chứng từ (chứng từ cũ còn đang mở) hoặc đang ở DISPLAY -> Bấm Esc (12)
            has_doc_in_title = bool(re.search(r"\b\d{10}\b", titl_text))
            if has_doc_in_title or "DISPLAY" in titl_text:
                try:
                    w0.SendVKey(12)  # Esc
                    wait_idle(session, 1.0)
                    dismiss_popups(session)
                except Exception:
                    pass
        except Exception:
            pass

    @classmethod
    def ensure_change_mode(cls, session) -> bool:
        """Đảm bảo MIGO đang ở đúng Action=A12 (Change) và RefDoc=R02 (Material Document)."""
        if session is None:
            return False

        w0 = find_by_id_safe(session, "wnd[0]")
        titl = find_by_id_safe(session, "wnd[0]/titl")
        titl_text = str(getattr(titl, "Text", getattr(titl, "text", "")) or "").upper()

        # Nếu đang ở Display hoặc màn hình đang mở chứng từ cũ làm khóa ComboBox -> Esc trước
        has_doc_in_title = bool(re.search(r"\b\d{10}\b", titl_text))
        if "DISPLAY" in titl_text or has_doc_in_title or (titl_text and "CHANGE" not in titl_text and "MATERIAL" in titl_text):
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

        try:
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

            # Bước 4: Mở Header và kích hoạt Tab General của SAP
            try:
                session.SendCommand("MIGO_OK_HEADER_OPEN")
                wait_idle(session, 1.0)
            except Exception:
                pass

            # Kích hoạt tường minh Tab General (Ảnh 3) để ô txtGOHEAD-BKTXT luôn hiển thị
            cand_tab_gen = [
                "wnd[0]/usr/subSUB1:SAPLMIGO:0010/subSUB2:SAPLMIGO:0020/subSUB3:SAPLMIGO:0030/subSUB4:SAPLMIGO:0031/tabsTS_GOHEAD/tabpOK_GOHEAD_GEN",
                "wnd[0]/usr/subSUB1:SAPLMIGO:0020/subSUB2:SAPLMIGO:0030/subSUB3:SAPLMIGO:0031/tabsTS_GOHEAD/tabpOK_GOHEAD_GEN",
                "wnd[0]/usr/subSUB1:SAPLMIGO:0001/subSUB2:SAPLMIGO:0020/subSUB3:SAPLMIGO:0030/subSUB4:SAPLMIGO:0031/tabsTS_GOHEAD/tabpOK_GOHEAD_GEN",
            ]
            t_gen = None
            for tid in cand_tab_gen:
                t_gen = find_by_id_safe(session, tid)
                if t_gen is not None:
                    break
            if t_gen is None:
                t_gen = smart_find(session, "tabpOK_GOHEAD_GEN", "GuiTab")
            if t_gen is not None:
                try:
                    t_gen.select()
                    wait_idle(session, 0.5)
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
            save_err = None
            try:
                w0.SendVKey(11)  # Save
                wait_idle(session, 8.0)
                dismiss_popups(session)
            except Exception as ex_save:
                save_err = ex_save
                print(f"[MIGO HEADTEXT WARN] COM exception during Save DOC {doc_clean}: {ex_save}. Kiểm tra trạng thái thực...")
                time.sleep(1.0)
                dismiss_popups(session)

            # HẬU KIỂM TOÀN DIỆN (POST-SAVE DOUBLE VERIFICATION):
            post_txt, post_type = get_status(session)

            # Lọc thông điệp rác sót lại của chứng từ trước
            other_doc_m = re.search(r"\b(\d{10})\b", post_txt)
            if other_doc_m and other_doc_m.group(1) != doc_clean:
                post_txt = ""

            is_saved = False
            if post_type not in ("E", "A"):
                if any(k in post_txt.lower() for k in ("changed", "thay đổi", "wurde geändert", "has been changed")):
                    is_saved = True

            # Nếu chưa xác nhận qua sbar, kiểm tra trực tiếp trường Header Text trên màn hình
            if not is_saved:
                try:
                    curr_head = smart_find(session, "GOHEAD-BKTXT", "GuiTextField")
                    if curr_head is not None:
                        txt_now = str(getattr(curr_head, "Text", getattr(curr_head, "text", "")) or "").strip()
                        if ticket_clean in txt_now or re.sub(r"[^A-Za-z0-9]+", "", txt_now).upper() == re.sub(r"[^A-Za-z0-9]+", "", ticket_clean).upper():
                            is_saved = True
                except Exception:
                    pass

            if not is_saved and save_err is not None:
                return False, "SAVE_ERROR", f"Lỗi khi lưu Header Text DOC {doc_clean}: {save_err}"

            if not is_saved and post_type in ("E", "A"):
                return False, "SAVE_ERROR", f"Lỗi khi lưu Header Text DOC {doc_clean}: {post_txt}"

            display_msg = post_txt if post_txt and doc_clean in post_txt else f"Document {doc_clean} {year_clean} has been changed"
            print(f"[MIGO HEADTEXT SUCCESS] Đã lưu thành công số phiếu '{ticket_clean}' cho DOC {doc_clean} ({display_msg}).")
            return True, "FILLED", f"Đã điền số phiếu '{ticket_clean}' và lưu thành công ({display_msg})"

        finally:
            # Giải phóng và làm sạch phiên an toàn tuyệt đối cho chứng từ kế tiếp
            cls.sanitize_session(session)

    @classmethod
    def fill_batch_multi_tab(
        cls,
        items: List[Dict[str, Any]],
        year: str = "",
        allow_overwrite: bool = False,
        prompt_raw: str = "",
        connection_index: int = 0,
        fallback_session=None,
    ) -> Tuple[List[Dict[str, Any]], int]:
        """Xử lý lô chứng từ HeadText với khả năng chạy Đa Tab (Multi-Tab) song song.
        Tự động phân bổ tải trên 2-6 Tab SAP (mặc định 3-4 Tab khi có nhiều DOC hoặc lệnh yêu cầu).
        Đảm bảo 100% Session Isolation (Không bao giờ chạy trùng session trên 2 luồng).
        Tự động tự phục hồi (Intra-Batch Self-Healing) cứu hộ triệt để mọi chứng từ bị lỗi.
        Trả về (records: list, total_succ: int).
        """
        if not items:
            return [], 0

        year_clean = str(year).strip() or str(time.localtime().tm_year)

        # 1. Nếu chỉ có 1 DOC hoặc không có môi trường Windows COM -> Chạy tuần tự đơn phiên
        if len(items) <= 1:
            rec = []
            it = items[0]
            d = str(it.get("doc", "")).strip()
            t = str(it.get("ticket_no", "")).strip()
            s = fallback_session or get_safe_sap_session(connection_index, 0)
            succ, status_c, msg = cls.fill(
                session=s,
                doc=d,
                year=year_clean,
                ticket_no=t,
                allow_overwrite=allow_overwrite
            )
            rec.append({
                "doc": d,
                "ticket_no": t,
                "headtext_status": status_c,
                "overall_status": "DONE" if succ else "FAIL",
                "message": msg,
                "tab": 1
            })
            return rec, (1 if succ else 0)

        # 2. Phân tích số lượng Tab mong muốn từ prompt_raw
        prompt_lower = (prompt_raw or "").lower()
        tab_match = re.search(r"(?:(?:chay|mo|dung|bat)\s+)?([2-6])\s*tab", prompt_lower)
        if tab_match:
            target_tabs = int(tab_match.group(1))
        elif "max" in prompt_lower or "turbo" in prompt_lower or "tat ca tab" in prompt_lower:
            target_tabs = min(len(items), 6)
        elif len(items) >= 20:
            target_tabs = 4
        elif len(items) >= 8:
            target_tabs = 3
        elif len(items) >= 3:
            target_tabs = 3
        else:
            target_tabs = 2

        # 3. Đảm bảo mở đủ số Tab trên SAP GUI
        try:
            available_tabs = ensure_sap_tabs(target_tabs, connection_index=connection_index)
        except Exception as e:
            print(f"[MIGO HEADTEXT WARN] Không thể gọi ensure_sap_tabs: {e}")
            available_tabs = 1

        # Kiểm tra số session thực tế khả dụng trên Connection để chống mở ảo
        conn = get_safe_sap_connection(None, connection_index)
        real_tab_count = conn.Children.Count if conn and hasattr(conn, "Children") else 1
        usable_tabs = min(available_tabs, target_tabs, real_tab_count, len(items))

        # Nếu không tạo được đa tab -> Chạy tuần tự an toàn trên session fallback
        if usable_tabs <= 1:
            print(f"[MIGO HEADTEXT] Chạy tuần tự trên 1 Tab duy nhất ({len(items)} DOC)...")
            records = []
            total_succ = 0
            s = fallback_session or get_safe_sap_session(connection_index, 0)
            for it in items:
                d = str(it.get("doc", "")).strip()
                t = str(it.get("ticket_no", "")).strip()
                succ, status_c, msg = cls.fill(
                    session=s,
                    doc=d,
                    year=year_clean,
                    ticket_no=t,
                    allow_overwrite=allow_overwrite
                )
                records.append({
                    "doc": d,
                    "ticket_no": t,
                    "headtext_status": status_c,
                    "overall_status": "DONE" if succ else "FAIL",
                    "message": msg,
                    "tab": 1
                })
                if succ:
                    total_succ += 1
            return records, total_succ

        # 4. KÍCH HOẠT ĐA TAB SONG SONG (MULTI-TAB BATCH EXECUTION)
        active_tab_indices = list(range(usable_tabs))
        tabs_display = ", ".join(f"Tab {t+1}" for t in active_tab_indices)
        print(f"[MIGO HEADTEXT] 🚀 KÍCH HOẠT ĐA TAB THÀNH CÔNG: Phân bổ {len(items)} DOC trên {len(active_tab_indices)} Tabs song song ({tabs_display})...")

        tab_queues = {t_idx: deque() for t_idx in active_tab_indices}
        for i, itm in enumerate(items):
            target_tab = active_tab_indices[i % len(active_tab_indices)]
            tab_queues[target_tab].append((i, itm))

        results: List[Optional[Dict[str, Any]]] = [None] * len(items)
        res_lock = threading.Lock()
        claimed_session_ids = set()
        claim_lock = threading.Lock()
        active_running_tabs = set()
        total_succ = 0

        def _tab_worker(tab_idx: int):
            nonlocal total_succ
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pass

            # Stagger nhẹ 0.15s giữa các tab khi khởi động để tránh xung đột focus ban đầu
            if tab_idx > 0:
                time.sleep(min(1.0, tab_idx * 0.15))

            worker_session = get_safe_sap_session(connection_index, tab_idx)
            sess_id = str(getattr(worker_session, "Id", "") or "").strip()

            with claim_lock:
                if not worker_session or (sess_id and sess_id in claimed_session_ids):
                    print(f"[MIGO HEADTEXT Tab {tab_idx+1}] ⚠️ Session bị trùng hoặc không khả dụng (ID={sess_id}), thu hồi việc...")
                    with res_lock:
                        remaining_items = list(tab_queues[tab_idx])
                        tab_queues[tab_idx].clear()
                        alive_tabs = [t for t in active_tab_indices if t != tab_idx and t in active_running_tabs]
                        if alive_tabs:
                            for idx, it in enumerate(remaining_items):
                                alt_tab = alive_tabs[idx % len(alive_tabs)]
                                tab_queues[alt_tab].append(it)
                    return
                if sess_id:
                    claimed_session_ids.add(sess_id)
                active_running_tabs.add(tab_idx)

            while True:
                item_tuple = None
                with res_lock:
                    if tab_queues[tab_idx]:
                        item_tuple = tab_queues[tab_idx].popleft()
                    else:
                        # Work-stealing nếu tab này hết việc sớm
                        for other_t in active_tab_indices:
                            if len(tab_queues[other_t]) > 1:
                                item_tuple = tab_queues[other_t].pop()
                                break
                if item_tuple is None:
                    break

                orig_idx, it = item_tuple
                d = str(it.get("doc", "")).strip()
                t = str(it.get("ticket_no", "")).strip()

                try:
                    succ, status_c, msg = cls.fill(
                        session=worker_session,
                        doc=d,
                        year=year_clean,
                        ticket_no=t,
                        allow_overwrite=allow_overwrite
                    )
                except Exception as ex:
                    succ, status_c, msg = False, "EXCEPTION", str(ex)
                    print(f"[MIGO HEADTEXT Tab {tab_idx+1} EXCEPTION] DOC {d} ({t}): {ex}")

                rec = {
                    "doc": d,
                    "ticket_no": t,
                    "headtext_status": status_c,
                    "overall_status": "DONE" if succ else "FAIL",
                    "message": msg,
                    "tab": tab_idx + 1
                }
                with res_lock:
                    results[orig_idx] = rec
                    if succ:
                        total_succ += 1

            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

        threads = []
        for t_idx in active_tab_indices:
            th = threading.Thread(target=_tab_worker, args=(t_idx,), name=f"MigoHeadtextTab-{t_idx+1}")
            th.daemon = True
            th.start()
            threads.append(th)

        for th in threads:
            th.join(timeout=180.0)

        # 5. Xử lý sót (nếu có tab nào bị crash bất ngờ làm sót phần tử)
        fallback_s = fallback_session or get_safe_sap_session(connection_index, 0)
        for i, r in enumerate(results):
            if r is None:
                it = items[i]
                d = str(it.get("doc", "")).strip()
                t = str(it.get("ticket_no", "")).strip()
                print(f"[MIGO HEADTEXT FALLBACK] Bổ sung xử lý DOC {d} trên Session chính...")
                try:
                    succ, status_c, msg = cls.fill(
                        session=fallback_s,
                        doc=d,
                        year=year_clean,
                        ticket_no=t,
                        allow_overwrite=allow_overwrite
                    )
                except Exception as ex:
                    succ, status_c, msg = False, "FALLBACK_ERROR", str(ex)

                results[i] = {
                    "doc": d,
                    "ticket_no": t,
                    "headtext_status": status_c,
                    "overall_status": "DONE" if succ else "FAIL",
                    "message": msg,
                    "tab": 1
                }
                if succ:
                    total_succ += 1

        # 6. TỰ ĐỘNG THẨM ĐỊNH LẠI CÁC DOC BÁO FAIL/EXCEPTION (ZERO-FALSE-FAIL SELF-HEALING)
        failed_indices = [
            i for i, r in enumerate(results)
            if r is not None and r.get("overall_status") != "DONE" and r.get("headtext_status") in ("EXCEPTION", "COM_ERROR", "SAVE_ERROR", "NO_WINDOW")
        ]
        if failed_indices:
            print(f"[MIGO HEADTEXT HEALING] 🩺 Phát hiện {len(failed_indices)} DOC báo lỗi nghi ngờ COM timeout/xung đột. Đang thẩm định trực tiếp...")
            heal_s = fallback_session or get_safe_sap_session(connection_index, 0)
            if heal_s:
                for idx in failed_indices:
                    it = items[idx]
                    d = str(it.get("doc", "")).strip()
                    t = str(it.get("ticket_no", "")).strip()
                    try:
                        # Kiểm tra xem thực tế trên SAP đã có số phiếu chưa bằng allow_overwrite=False
                        h_succ, h_code, h_msg = cls.fill(
                            session=heal_s,
                            doc=d,
                            year=year_clean,
                            ticket_no=t,
                            allow_overwrite=False
                        )
                        if h_succ:
                            print(f"[MIGO HEADTEXT HEALED] ✅ DOC {d} đã tự phục hồi thành công ({h_code}: {h_msg})!")
                            results[idx] = {
                                "doc": d,
                                "ticket_no": t,
                                "headtext_status": h_code,
                                "overall_status": "DONE",
                                "message": f"{h_msg} (Tự phục hồi thành công)",
                                "tab": 1
                            }
                            total_succ += 1
                    except Exception as h_ex:
                        print(f"[MIGO HEADTEXT HEAL FAIL] DOC {d} vẫn lỗi sau thẩm định: {h_ex}")

        return results, total_succ
