# -*- coding: utf-8 -*-
"""Worker chuyên trách CHỈ ĐỌC giao dịch ZFIR266 (Báo cáo ALV Grid file đính kèm).
TUYỆT ĐỐI KHÔNG GỌI /nMIGO HOẶC MỞ CHẾ ĐỘ DISPLAY (A04) ĐỂ TRÁNH LÀM BẨN SESSION SAP.
"""
import re
from datetime import datetime
from typing import Dict, Any, List
from app.core.sap_primitives import (
    find_by_id_safe,
    wait_idle,
    dismiss_popups,
    open_transaction,
    reset_to_session_manager,
)
from app.core.sap_alv import fill_multiple_selection, get_column_caption
from app.sap_attachment import classify_sap_attachments


class Zfir266Reader:
    """Đọc dữ liệu file đính kèm BBBG & VAT qua ZFIR266 ALV Grid."""

    @classmethod
    def read(
        cls,
        session,
        docs: List[str],
        year: str = "",
        company_code: str = "2301",
        ticket_no: str = ""
    ) -> List[Dict[str, Any]]:
        clean_docs = []
        for d in docs:
            cleaned = re.sub(r"\D", "", str(d)).strip()
            if cleaned and len(cleaned) >= 8 and cleaned not in clean_docs:
                clean_docs.append(cleaned)

        if not clean_docs:
            return []

        if not year or not re.fullmatch(r"(?:19|20|21)\d{2}", str(year)):
            year = str(datetime.now().year)

        if session is None:
            return [
                {
                    "doc": d,
                    "company_code": company_code,
                    "year": year,
                    "bbbg_status": "N/A",
                    "vat_status": "N/A",
                    "overall_status": "Chưa kết nối SAP (NO_SESSION)",
                    "status_code": "NO_SESSION",
                    "files_found": "N/A",
                    "bbbg_file": None,
                    "vat_file": None,
                    "ticket_no": ticket_no
                }
                for d in clean_docs
            ]

        zfir_results = {d: [] for d in clean_docs}
        try:
            open_transaction(session, "ZFIR266")
            wait_idle(session, 3.0)

            sbar = find_by_id_safe(session, "wnd[0]/sbar")
            sbar_text = str(getattr(sbar, "Text", "")).strip().lower() if sbar else ""
            trans_name = str(getattr(session.Info, "Transaction", "")).upper()

            if trans_name != "ZFIR266" or "does not exist" in sbar_text or "không tồn tại" in sbar_text or "not authorized" in sbar_text:
                print(f"[ZFIR266 WARN] ZFIR266 không khả dụng ({sbar_text or trans_name}).")
            else:
                # 1. Thử chạy BATCH toàn bộ DOCs qua Multiple Selection (1 lượt duy nhất F8)
                multi_btn_id = "wnd[0]/usr/btn%_S_MBELN_%_APP_%-VALU_PUSH"
                has_multi_btn = find_by_id_safe(session, multi_btn_id) is not None
                batch_ran_ok = False

                if has_multi_btn and len(clean_docs) > 1:
                    try:
                        print(f"[ZFIR266 BATCH] Đang đưa đồng thời {len(clean_docs)} DOC vào Multiple Selection của ZFIR266...")
                        b_f = find_by_id_safe(session, "wnd[0]/usr/ctxtP_BUKRS") or find_by_id_safe(session, "wnd[0]/usr/txtP_BUKRS")
                        if b_f:
                            b_f.text = str(company_code)
                        g_f = find_by_id_safe(session, "wnd[0]/usr/txtP_GJAHR") or find_by_id_safe(session, "wnd[0]/usr/ctxtP_GJAHR")
                        if g_f:
                            g_f.text = str(year)
                        m_f = find_by_id_safe(session, "wnd[0]/usr/txtP_MAX")
                        if m_f:
                            m_f.text = "1000"

                        # Xóa trắng giới hạn Posting Date / Entry Date để không sót DOC khác ngày/tháng
                        for d_id in [
                            "wnd[0]/usr/ctxtS_BUDAT-LOW", "wnd[0]/usr/ctxtS_BUDAT-HIGH",
                            "wnd[0]/usr/ctxtS_BLDAT-LOW", "wnd[0]/usr/ctxtS_BLDAT-HIGH",
                            "wnd[0]/usr/usr/ctxtS_BUDAT-LOW", "wnd[0]/usr/usr/ctxtS_BUDAT-HIGH",
                            "wnd[0]/usr/ctxtSO_BUDAT-LOW", "wnd[0]/usr/ctxtSO_BUDAT-HIGH",
                            "wnd[0]/usr/ctxtP_BUDAT", "wnd[0]/usr/ctxtS_CPUDT-LOW", "wnd[0]/usr/ctxtS_CPUDT-HIGH"
                        ]:
                            df = find_by_id_safe(session, d_id)
                            if df:
                                try:
                                    df.text = ""
                                except Exception:
                                    pass

                        # Điền toàn bộ danh sách DOC vào Multiple Selection
                        fill_multiple_selection(session, multi_btn_id, clean_docs)

                        # Bấm F8 Execute
                        w0 = find_by_id_safe(session, "wnd[0]")
                        if w0:
                            w0.sendVKey(8)
                        wait_idle(session, 3.0)
                        dismiss_popups(session)

                        # Trích xuất ALV Grid chứa toàn bộ các DOC
                        grid = find_by_id_safe(session, "wnd[0]/usr/cntlGRID1/shellcont/shell") or find_by_id_safe(session, "wnd[0]/usr/shell")
                        has_grid = grid is not None and hasattr(grid, "RowCount") and grid.RowCount > 0
                        if has_grid:
                            resolved_doc_col = None
                            resolved_file_col = None
                            try:
                                col_order = grid.ColumnOrder
                                for c_idx in range(col_order.Count):
                                    c_name = str(col_order.ElementAt(c_idx)).strip()
                                    c_up = c_name.upper()
                                    c_cap = ""
                                    try:
                                        c_cap = str(get_column_caption(grid, c_name)).strip().upper()
                                    except Exception:
                                        pass
                                    if not resolved_doc_col and any(k in c_up or k in c_cap for k in ["MBELN", "MBLNR", "MAT_DOC", "DOCUMENTNO", "DOCUMENT_NO", "DOCNO", "BELNR", "DOCUMENT", "CHỨNG TỪ"]):
                                        resolved_doc_col = c_name
                                    if not resolved_file_col and (
                                        any(k in c_up or k in c_cap for k in ["FILE_NAME", "FILENAME", "FILE", "TITLE", "DESCRIPT", "DESCR", "ZFILE", "LFNAM", "TẬP TIN", "TÊN TỆP"])
                                        or c_up == "NAME"
                                    ):
                                        resolved_file_col = c_name
                            except Exception:
                                pass

                            for r in range(grid.RowCount):
                                if r % 20 == 0:
                                    try:
                                        grid.firstVisibleRow = r
                                    except Exception:
                                        pass
                                r_doc = ""
                                r_file = ""
                                if resolved_doc_col:
                                    try:
                                        r_doc = str(grid.GetCellValue(r, resolved_doc_col)).strip()
                                    except Exception:
                                        pass
                                if resolved_file_col:
                                    try:
                                        r_file = str(grid.GetCellValue(r, resolved_file_col)).strip()
                                    except Exception:
                                        pass

                                if not r_file:
                                    for col_cand in ["FILE_NAME", "FILENAME", "FILE", "TITLE", "DESCRIPT", "DESCR", "NAME", "ZFILE", "LFNAM"]:
                                        try:
                                            v = grid.GetCellValue(r, col_cand)
                                            if v:
                                                r_file = str(v).strip()
                                                break
                                        except Exception:
                                            pass

                                if not r_file:
                                    try:
                                        col_order = grid.ColumnOrder
                                        for c_idx in range(col_order.Count):
                                            c_name = str(col_order.ElementAt(c_idx)).strip()
                                            try:
                                                v = str(grid.GetCellValue(r, c_name)).strip()
                                                if re.search(r"\.(pdf|png|jpg|jpeg|docx?|xlsx?|zip|rar|msg|txt)$", v, re.I):
                                                    r_file = v
                                                    break
                                            except Exception:
                                                pass
                                    except Exception:
                                        pass

                                matched_doc = None
                                if r_doc and r_doc in zfir_results:
                                    matched_doc = r_doc
                                else:
                                    for k_cand in zfir_results.keys():
                                        for c_cand in ["MBELN", "MBLNR", "MAT_DOC", "DOCUMENTNO", "DOCUMENT_NO", "BELNR", "DOCNO"]:
                                            try:
                                                if str(grid.GetCellValue(r, c_cand)).strip() == k_cand:
                                                    matched_doc = k_cand
                                                    break
                                            except Exception:
                                                pass
                                        if matched_doc:
                                            break

                                if matched_doc and r_file:
                                    if r_file not in zfir_results[matched_doc]:
                                        zfir_results[matched_doc].append(r_file)

                            print(f"[ZFIR266 BATCH SUCCESS] Đã trích xuất xong {grid.RowCount} dòng từ ZFIR266.")
                            batch_ran_ok = True
                    except Exception as b_err:
                        print(f"[ZFIR266 BATCH WARN] Batch query lỗi ({b_err}), chuyển sang quét từng chứng từ...")
                        open_transaction(session, "ZFIR266")
                        wait_idle(session, 2.0)

                # 2. Nếu chỉ có 1 DOC hoặc BATCH không thành công: duyệt từng DOC
                if not batch_ran_ok:
                    for doc in clean_docs:
                        titles_found = []
                        try:
                            btn8_test = find_by_id_safe(session, "wnd[0]/tbar[1]/btn[8]")
                            if btn8_test is None:
                                open_transaction(session, "ZFIR266")
                                wait_idle(session, 2.0)

                            w0 = find_by_id_safe(session, "wnd[0]")
                            bukrs_field = find_by_id_safe(session, "wnd[0]/usr/ctxtP_BUKRS") or find_by_id_safe(session, "wnd[0]/usr/txtP_BUKRS")
                            gjahr_f = find_by_id_safe(session, "wnd[0]/usr/txtP_GJAHR") or find_by_id_safe(session, "wnd[0]/usr/ctxtP_GJAHR")
                            doc_f = find_by_id_safe(session, "wnd[0]/usr/ctxtS_MBELN-LOW") or find_by_id_safe(session, "wnd[0]/usr/txtS_MBELN-LOW")
                            max_f = find_by_id_safe(session, "wnd[0]/usr/txtP_MAX")

                            if bukrs_field:
                                bukrs_field.text = str(company_code)
                            if gjahr_f:
                                gjahr_f.text = str(year)
                            if doc_f:
                                doc_f.text = str(doc)
                            if max_f:
                                max_f.text = "500"

                            for d_id in [
                                "wnd[0]/usr/ctxtS_BUDAT-LOW", "wnd[0]/usr/ctxtS_BUDAT-HIGH",
                                "wnd[0]/usr/ctxtS_BLDAT-LOW", "wnd[0]/usr/ctxtS_BLDAT-HIGH",
                                "wnd[0]/usr/ctxtSO_BUDAT-LOW", "wnd[0]/usr/ctxtSO_BUDAT-HIGH",
                                "wnd[0]/usr/ctxtP_BUDAT"
                            ]:
                                df = find_by_id_safe(session, d_id)
                                if df:
                                    try:
                                        df.text = ""
                                    except Exception:
                                        pass

                            if w0:
                                w0.sendVKey(8)
                            wait_idle(session, 2.5)
                            dismiss_popups(session)

                            grid = find_by_id_safe(session, "wnd[0]/usr/cntlGRID1/shellcont/shell") or find_by_id_safe(session, "wnd[0]/usr/shell")
                            has_grid = grid is not None and hasattr(grid, "RowCount") and grid.RowCount > 0
                            if has_grid:
                                resolved_file_col = None
                                try:
                                    col_order = grid.ColumnOrder
                                    for c_idx in range(col_order.Count):
                                        c_name = str(col_order.ElementAt(c_idx)).strip()
                                        c_up = c_name.upper()
                                        c_cap = ""
                                        try:
                                            c_cap = str(get_column_caption(grid, c_name)).strip().upper()
                                        except Exception:
                                            pass
                                        if any(k in c_up or k in c_cap for k in ["FILE_NAME", "FILENAME", "FILE", "TITLE", "DESCRIPT", "DESCR", "NAME", "ZFILE", "LFNAM"]):
                                            resolved_file_col = c_name
                                            break
                                except Exception:
                                    pass

                                for r in range(grid.RowCount):
                                    if r % 20 == 0:
                                        try:
                                            grid.firstVisibleRow = r
                                        except Exception:
                                            pass
                                    fname = ""
                                    if resolved_file_col:
                                        try:
                                            v = grid.GetCellValue(r, resolved_file_col)
                                            if v:
                                                fname = str(v).strip()
                                        except Exception:
                                            pass
                                    if not fname:
                                        for col_cand in ["FILE_NAME", "FILENAME", "FILE", "TITLE", "DESCRIPT", "DESCR", "NAME", "ZFILE", "LFNAM"]:
                                            try:
                                                v = grid.GetCellValue(r, col_cand)
                                                if v:
                                                    fname = str(v).strip()
                                                    break
                                            except Exception:
                                                pass
                                    if not fname:
                                        try:
                                            col_order = grid.ColumnOrder
                                            for c_idx in range(col_order.Count):
                                                c_name = str(col_order.ElementAt(c_idx)).strip()
                                                try:
                                                    v = str(grid.GetCellValue(r, c_name)).strip()
                                                    if re.search(r"\.(pdf|png|jpg|jpeg|docx?|xlsx?|zip|rar|msg|txt)$", v, re.I):
                                                        fname = v
                                                        break
                                                except Exception:
                                                    pass
                                        except Exception:
                                            pass
                                    if fname and fname not in titles_found:
                                        titles_found.append(fname)

                                w0_back = find_by_id_safe(session, "wnd[0]")
                                if w0_back:
                                    w0_back.sendVKey(3)
                                    wait_idle(session, 1.0)
                                    dismiss_popups(session)
                        except Exception as e:
                            print(f"[ZFIR266 DOC WARN] Lỗi đọc DOC {doc}: {e}")
                            try:
                                open_transaction(session, "ZFIR266")
                                wait_idle(session, 2.0)
                            except Exception:
                                pass
                        zfir_results[doc] = titles_found
        except Exception as zfir_err:
            print(f"[ZFIR266 EXCEPTION] {zfir_err}")

        # Đưa session về /n an toàn, TUYỆT ĐỐI KHÔNG MỞ MIGO DISPLAY FALLBACK!
        reset_to_session_manager(session)

        records = []
        for doc in clean_docs:
            titles = zfir_results.get(doc, [])
            info = classify_sap_attachments(titles, ticket_no=ticket_no)
            found_in_zfir = doc in zfir_results and bool(zfir_results[doc])
            need_bbbg = not info["has_bbbg"]
            need_vat = not info["has_vat"]
            need_ht = bool(ticket_no and not info.get("has_ticket"))
            records.append({
                "doc": doc,
                "company_code": company_code,
                "sap_titles": list(titles),
                "year": year,
                "bbbg_status": "CÓ" if info["has_bbbg"] else "CHƯA",
                "vat_status": "CÓ" if info["has_vat"] else "CHƯA",
                "overall_status": info["overall_status"],
                "status_code": info["status_code"],
                "files_found": ", ".join(titles) if titles else "Không có file",
                "bbbg_file": info["bbbg_file"],
                "vat_file": info["vat_file"],
                "ticket_no": ticket_no,
                "found_in_zfir": found_in_zfir,
                "need_bbbg": need_bbbg,
                "need_vat": need_vat,
                "need_header_text": need_ht
            })

        return records
