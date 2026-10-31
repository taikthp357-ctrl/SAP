# -*- coding: utf-8 -*-
"""Xử lý bảng báo cáo dạng lưới ALV Grid và Multiple Selection trên SAP GUI.
TUYỆT ĐỐI KHÔNG CHỨA LOGIC NGHIỆP VỤ RIÊNG BIỆT (Không gán cứng T-Code).
"""
import re
import time
import unicodedata
from typing import Dict, List, Optional
from app.core.sap_primitives import (
    find_by_id_safe,
    wait_idle,
    copy_to_clipboard_safely,
    get_clipboard_safely,
)


def find_largest_grid(session, max_wait_sec: float = 10.0):
    """Dò tìm ALV Grid lớn nhất trên màn hình SAP một cách đệ quy."""
    start_t = time.time()
    best_grid = None
    while time.time() - start_t < max_wait_sec:
        best_grid = None
        best_rows = -1

        def _collect(obj, depth=0):
            nonlocal best_grid, best_rows
            if obj is None or depth > 14:
                return
            try:
                obj_type = getattr(obj, "Type", "")
                if obj_type in ("GuiGridView", "GuiGridControl", "GuiShell"):
                    try:
                        rc = int(obj.RowCount)
                        if rc > best_rows:
                            best_rows = rc
                            best_grid = obj
                    except Exception:
                        pass
                children = getattr(obj, "Children", None)
                if children is not None:
                    for idx in range(children.Count):
                        _collect(children.ElementAt(idx), depth + 1)
            except Exception:
                pass

        try:
            wnd0 = find_by_id_safe(session, "wnd[0]")
            if wnd0 is not None:
                _collect(wnd0, 0)
        except Exception:
            pass

        if best_grid is not None and best_rows > 0:
            return best_grid
        time.sleep(0.3)

    return best_grid


def normalize_header(s: str) -> str:
    """Chuẩn hóa chuỗi tiêu đề/mã cột: loại bỏ dấu, ký tự đặc biệt, chuyển chữ hoa."""
    if not s:
        return ""
    nfd = unicodedata.normalize("NFD", str(s))
    no_acc = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    return re.sub(r"[^A-Za-z0-9]", "", no_acc).upper()


def get_column_caption(alv_obj, col_id: str) -> str:
    """Lấy tiêu đề hiển thị đầy đủ của cột qua GetDisplayedColumnTitle và GetColumnTitles."""
    parts = []
    try:
        disp = alv_obj.GetDisplayedColumnTitle(col_id)
        if isinstance(disp, str) and disp:
            parts.append(disp.strip())
    except Exception:
        pass

    try:
        titles = alv_obj.GetColumnTitles(col_id)
        if isinstance(titles, (list, tuple)):
            for t in titles:
                if isinstance(t, str) and t.strip():
                    parts.append(t.strip())
        elif isinstance(titles, str) and titles.strip():
            parts.append(titles.strip())
        elif hasattr(titles, "Count") and isinstance(getattr(titles, "Count", None), int):
            for i in range(titles.Count):
                t = str(titles.ElementAt(i)).strip()
                if t:
                    parts.append(t)
    except Exception:
        pass

    try:
        tooltip = alv_obj.GetColumnTooltip(col_id)
        if isinstance(tooltip, str) and tooltip:
            parts.append(tooltip.strip())
    except Exception:
        pass

    return " ".join(parts)


def resolve_grid_columns(alv_obj, column_mapping: Dict[str, List[str]]) -> Dict[str, str]:
    """Phân giải tên cột ALV Grid theo thuật toán 2 lượt:
    - Lượt 0: Khóa cứng theo mã kỹ thuật (khớp chính xác).
    - Lượt 1: Khớp chính xác mã kỹ thuật hoặc Caption đã chuẩn hóa.
    - Lượt 2: Khớp mờ (Fuzzy match) cho các chuỗi có độ dài >= 4 ký tự.
    """
    resolved = {}
    try:
        col_collection = alv_obj.ColumnOrder
        col_count = col_collection.Count
    except Exception:
        return resolved

    tech_cols = []
    col_captions = {}
    for c_idx in range(col_count):
        try:
            col_name = str(col_collection.ElementAt(c_idx)).strip()
            tech_cols.append(col_name)
            col_captions[col_name] = get_column_caption(alv_obj, col_name)
        except Exception:
            pass

    print(f"[SAP ALV] Bảng có {col_count} cột. Đang phân giải tự động...")
    for out_key, candidates in column_mapping.items():
        matched_col = None

        # Lượt 0: Khóa cứng theo mã kỹ thuật
        for cand in candidates:
            cand_clean = str(cand).strip().upper()
            for tc in tech_cols:
                if tc.upper() == cand_clean:
                    matched_col = tc
                    break
            if matched_col:
                break

        cand_norms = [normalize_header(c) for c in candidates if normalize_header(c)]

        # Lượt 1: Khớp chính xác mã kỹ thuật hoặc Caption
        if not matched_col:
            for tc in tech_cols:
                tc_norm = normalize_header(tc)
                cap_norm = normalize_header(col_captions.get(tc, ""))
                for cn in cand_norms:
                    if cn == tc_norm or cn == cap_norm:
                        matched_col = tc
                        break
                if matched_col:
                    break

        # Lượt 2: Khớp mờ cho tiêu đề có ý nghĩa (>= 4 ký tự)
        if not matched_col:
            for tc in tech_cols:
                cap_full_norm = normalize_header(f"{tc} {col_captions.get(tc, '')}")
                for cn in cand_norms:
                    if len(cn) >= 4 and cn in cap_full_norm:
                        matched_col = tc
                        break
                if matched_col:
                    break

        if matched_col:
            resolved[out_key] = matched_col
            cap_display = col_captions.get(matched_col, "")
            print(f"  ✓ [{out_key:8}] -> Cột SAP '{matched_col}' ('{cap_display}')")
        else:
            print(f"  ✗ [{out_key:8}] -> Chưa tìm thấy cột phù hợp!")

    return resolved


def clear_multiple_selection(session, push_button_id: str):
    """Mở popup Multiple Selection và xóa sạch lựa chọn bằng btn[16], sau đó bấm F8 để lưu."""
    push_btn = find_by_id_safe(session, push_button_id)
    if push_btn is None:
        return
    try:
        push_btn.press()
        wait_idle(session, 2.0)
        btn16 = find_by_id_safe(session, "wnd[1]/tbar[0]/btn[16]")
        if btn16 is not None:
            try:
                btn16.press()
                wait_idle(session, 1.0)
            except Exception:
                pass
        btn8 = find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
        if btn8 is not None:
            btn8.press()
        else:
            wnd1 = find_by_id_safe(session, "wnd[1]")
            if wnd1 is not None:
                wnd1.sendVKey(8)
        wait_idle(session, 2.0)
    except Exception as e:
        print(f"[SAP WARN] Lỗi xóa multiple selection ({push_button_id}): {e}")


def fill_multiple_selection(session, push_button_id: str, values: List[str]):
    """Điền danh sách nhiều giá trị vào Multiple Selection.
    - Nếu <= 5 giá trị: điền trực tiếp vào các ô mà không cần clipboard.
    - Nếu > 5 giá trị: sao chép và bấm nút btn[24] (Upload from Clipboard), khôi phục clipboard cũ.
    """
    clean_vals = [v.strip().upper() for v in values if v.strip()]
    if not clean_vals:
        return

    push_btn = find_by_id_safe(session, push_button_id)
    if push_btn is None:
        return
    push_btn.press()
    wait_idle(session, 3.0)

    btn16 = find_by_id_safe(session, "wnd[1]/tbar[0]/btn[16]")
    if btn16 is not None:
        try:
            btn16.press()
            wait_idle(session, 1.5)
        except Exception:
            pass

    prev_clip = None
    used_clipboard = False
    btn24 = find_by_id_safe(session, "wnd[1]/tbar[0]/btn[24]")
    if len(clean_vals) > 5 and btn24 is not None:
        try:
            prev_clip = get_clipboard_safely()
            clip_text = "\r\n".join(clean_vals)
            if copy_to_clipboard_safely(clip_text):
                btn24.press()
                wait_idle(session, 2.0)
                cell0 = find_by_id_safe(session, "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]")
                if cell0 is not None and cell0.text.strip().upper() == clean_vals[0]:
                    used_clipboard = True
                    print(f"[SAP MULTI-SELECT] Đã paste thành công {len(clean_vals)} giá trị từ clipboard vào Multiple Selection.")
                else:
                    print("[SAP MULTI-SELECT WARN] Paste clipboard chưa thấy ăn dòng đầu, thử lại...")
                    btn24.press()
                    wait_idle(session, 2.0)
                    used_clipboard = True
        except Exception as clip_err:
            print(f"[SAP WARN] Lỗi paste clipboard vào Multiple Selection: {clip_err}")

    if not used_clipboard:
        tbl_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
        for idx, val in enumerate(clean_vals[:8]):
            cell_id = f"{tbl_id}/ctxtRSCSEL_255-SLOW_I[1,{idx}]"
            cell = find_by_id_safe(session, cell_id)
            if cell is not None:
                try:
                    cell.text = val
                except Exception:
                    alt_cell = find_by_id_safe(session, cell_id.replace("ctxtRSCSEL_255", "txtRSCSEL_255"))
                    if alt_cell is not None:
                        alt_cell.text = val

    btn8 = find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
    if btn8 is not None:
        btn8.press()
    else:
        wnd1 = find_by_id_safe(session, "wnd[1]")
        if wnd1 is not None:
            wnd1.sendVKey(8)

    wait_idle(session, 3.0)

    if prev_clip is not None:
        try:
            copy_to_clipboard_safely(prev_clip)
        except Exception:
            pass


def extract_alv_grid(
    alv_obj,
    column_mapping: Dict[str, List[str]],
    max_rows: Optional[int] = 15000
) -> List[Dict[str, str]]:
    """Trích xuất ALV Grid kèm cơ chế forward-fill mã hàng và lọc bỏ dòng tổng cộng (Summary rows)."""
    try:
        row_count = int(alv_obj.RowCount)
    except Exception as exc:
        raise RuntimeError("Không đọc được RowCount; dữ liệu nguồn chưa xác định") from exc

    if row_count == 0:
        return []

    resolved_cols = resolve_grid_columns(alv_obj, column_mapping)
    records = []
    effective_max = max_rows if max_rows is not None else 15000
    if row_count > effective_max:
        raise RuntimeError(f"Nguồn có {row_count} dòng, vượt giới hạn {effective_max}; thu hẹp bộ lọc")
    if not resolved_cols:
        raise RuntimeError("Không nhận diện được cột dữ liệu SAP")

    limit = row_count
    last_product = ""

    for r in range(limit):
        item = {}
        for out_key, sap_col in resolved_cols.items():
            try:
                val = alv_obj.GetCellValue(r, sap_col)
                item[out_key] = str(val).strip()
            except Exception:
                try:
                    alv_obj.firstVisibleRow = r
                    val = alv_obj.GetCellValue(r, sap_col)
                    item[out_key] = str(val).strip()
                except Exception as exc:
                    raise RuntimeError(f"Không đọc được ô ALV {r}/{out_key}") from exc

        curr_prod = item.get("product", "")
        if curr_prod:
            last_product = curr_prod
        elif last_product and "product" in item:
            item["product"] = last_product

        p_val = str(item.get("product", "")).strip()
        if p_val == "*" or p_val.lower().startswith("total") or p_val.lower().startswith("tổng"):
            continue

        try:
            row_type = str(alv_obj.GetRowType(r)).strip().upper()
            if row_type and ("TOTAL" in row_type or row_type in ("1", "2")):
                continue
        except Exception:
            pass

        has_pos_or_doc = bool(
            item.get("sloc") or item.get("bin") or item.get("batch") or
            item.get("docNo") or item.get("obNo") or item.get("ibNo") or
            item.get("stoNo") or item.get("poNo") or item.get("wtNo") or
            item.get("item") or item.get("mvtType") or item.get("refNo") or
            item.get("component") or item.get("queue_name") or item.get("document") or item.get("error_message") or
            (item.get("transit") and str(item.get("transit")).strip() not in ("", "0", "0.0", "0,0"))
        )
        if not has_pos_or_doc:
            continue

        records.append(item)

    return records
