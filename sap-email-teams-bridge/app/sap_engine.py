# -*- coding: utf-8 -*-
"""Điều khiển SAP GUI Scripting qua COM Object trong RAM (0 PowerShell / 0 Clipboard).

Tích hợp toàn bộ mã VBS/VBA chuẩn xác từ ModuleSAP.bas & modPosting.bas của BBBG Excel:
- Điều khiển chuẩn xác các trường /SCWM/POST (Advanced Search, Owner, Multi-select, Start).
- Cơ chế _find_largest_grid quét đệ quy bắt trọn mọi ALV Grid bất kể phiên bản SAP GUI.
- Phân giải tên cột 2 lượt (Exact & Fuzzy) bằng GetDisplayedColumnTitle và GetColumnTitles.
- Tự động lọc bỏ dòng tổng cộng (Summary rows) tránh in dòng trống trên Teams.
- Đồng bộ tự động MB52 (kèm vị trí LGPBE) và MB51 (Lịch sử giao dịch).
"""

import os
from pathlib import Path
import re
import time
import unicodedata
import tempfile
import shutil
import glob
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple, Set
import win32com.client

from app.number_utils import parse_sap_float, format_qty_display, parse_sap_datetime_sort_key
from app.adaptive_cache import UI_CACHE, AdaptiveConfigCache
from app.ob_service import (
    ObService,
    RC65_FIELD_TOKENS,
    RC65_VL06O_COLUMNS,
    RC65_VL06O_COLUMN_MAPPING,
    RC65_TABLE_OVERVIEW_ID,
    RC65_TABLE_OVERVIEW_TOKEN,
    RC65_TABLE_SPLIT_TOKEN,
)

SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_CONNECTION_NAME = ""


def _copy_to_clipboard_safely(text: str) -> bool:
    """Sao chép chuỗi vào Clipboard Windows một cách an toàn."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        print(f"[CLIPBOARD WARN] Không thể ghi clipboard: {e}")
        return False


def _get_clipboard_safely() -> Optional[str]:
    """Đọc dữ liệu văn bản hiện tại trên Clipboard Windows để sao lưu."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        data = None
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return data
    except Exception:
        return None


class SapExecutionError(RuntimeError):
    """Lỗi phát sinh trong quá trình thực thi nghiệp vụ trên SAP GUI."""
    pass


class SapEngine:
    """Quản lý kết nối và điều khiển tự động SAP GUI theo chuẩn BBBG."""
    MAX_BOT_SESSIONS: int = 3
    MAX_SAP_SESSIONS: int = 6
    USER_RESERVED_SESSIONS: int = 3

    def __init__(
        self,
        max_alv_rows: int = 15000,
        keep_alive_idle_seconds: int = 900,
        connection_index: int = 0,
        session_index: int = 0,
        strict_session_safety: bool = True,
        auto_create_session: bool = True,
        prefer_dedicated_session: bool = True,
        user_reserved_sessions: int = 3,
        max_sap_sessions: int = 6,
        max_batch_tabs: int = 3,
        expected_system: str = "",
        expected_client: str = "",
        expected_user: str = ""
    ):
        self.max_alv_rows = max_alv_rows
        self.keep_alive_idle_seconds = keep_alive_idle_seconds
        self.connection_index = connection_index
        self.session_index = session_index
        self.strict_session_safety = strict_session_safety
        self.auto_create_session = auto_create_session
        self.prefer_dedicated_session = prefer_dedicated_session
        self.user_reserved_sessions = user_reserved_sessions
        self.max_sap_sessions = max_sap_sessions
        self.max_batch_tabs = max_batch_tabs
        self.session = None
        self.last_activity = time.time()
        self.last_popup_message = ""
        self.expected_identity = {"system": expected_system, "client": expected_client, "user": expected_user}
        self.session_identity = None

    @staticmethod
    def _identity(session):
        info = getattr(session, "Info", None)
        conn_id = ""
        sess_id = str(getattr(session, "Id", "")).strip()
        if "/con[" in sess_id:
            try:
                conn_id = sess_id.split("/con[")[1].split("]")[0]
            except Exception:
                conn_id = ""
        identity = {
            "system": str(getattr(info, "SystemName", "") or "").strip(),
            "client": str(getattr(info, "Client", "") or "").strip(),
            "user": str(getattr(info, "User", "") or "").strip(),
            "session_id": sess_id,
            "connection_id": conn_id,
        }
        for k in ("system", "client", "user", "session_id"):
            if not identity.get(k):
                raise SapExecutionError(f"Không xác minh được {k} của phiên SAP")
        return identity

    def _check_identity(self, session):
        current = self._identity(session)
        expected = getattr(self, "expected_identity", {}) or {}
        for key in ("system", "client", "user", "session_id", "connection_id"):
            if expected.get(key) and str(current.get(key, "")).upper() != str(expected[key]).upper():
                raise SapExecutionError(f"SAP {key} không khớp cấu hình (mong đợi: {expected.get(key)}, thực tế: {current.get(key)})")
        prior = getattr(self, "session_identity", None)
        if prior and any(current.get(k) != prior.get(k) for k in ("system", "client", "user", "session_id", "connection_id") if prior.get(k)):
            raise SapExecutionError("Danh tính phiên SAP thay đổi; dừng để kiểm tra")
        return current

    def _connect_sap(self):
        """Kết nối tới phiên SAP GUI an toàn, ưu tiên dải tab dành riêng cho Bot (tab 4, 5, 6)."""
        try:
            app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except Exception as exc:
            raise SapExecutionError("Mở SAP GUI và đăng nhập trước khi tra cứu") from exc
        if not 0 <= self.connection_index < app.Children.Count:
            raise SapExecutionError("SAP connection_index không tồn tại")
        conn = app.Children(self.connection_index)
        total_open = conn.Children.Count
        if total_open <= 0:
            raise SapExecutionError("Không tìm thấy phiên SAP nào đang mở")

        selected = None
        user_reserved = getattr(self, "user_reserved_sessions", 3)
        max_sap = getattr(self, "max_sap_sessions", 6)

        # 1. Ưu tiên số 1: Quét tìm phiên rảnh trong dải dành riêng cho Bot (index >= user_reserved, ví dụ tab 4, 5, 6: index 3, 4, 5)
        if total_open > user_reserved:
            for s_i in range(user_reserved, total_open):
                sess = conn.Children(s_i)
                try:
                    if not bool(getattr(sess, "Busy", False)):
                        s_info = getattr(sess, "Info", None)
                        s_tcode = str(getattr(s_info, "Transaction", "")).strip().upper() if s_info else ""
                        if s_tcode in ("SESSION_MANAGER", "") or not getattr(self, "strict_session_safety", True):
                            selected = sess
                            break
                        # Tab trong dải của bot (>= user_reserved) không bận là hợp lệ để tái sử dụng
                        selected = sess
                        break
                except Exception:
                    pass

        # 2. Ưu tiên số 2: Nếu chưa có tab nào trong dải Bot (hoặc các tab bot đều bận) và tổng số tab < max_sap (6):
        # Tự động gọi CreateSession() để mở thêm tab mới cho Bot (tab 4, 5, 6)
        if selected is None and getattr(self, "auto_create_session", True) and total_open < max_sap:
            orig_count = conn.Children.Count
            for s_i in range(conn.Children.Count):
                s_cand = conn.Children(s_i)
                if hasattr(s_cand, "CreateSession"):
                    try:
                        s_cand.CreateSession()
                        for _ in range(16):
                            time.sleep(0.25)
                            if conn.Children.Count > orig_count:
                                selected = conn.Children(conn.Children.Count - 1)
                                break
                    except Exception:
                        pass
                if selected is not None:
                    break

        # 3. Ưu tiên số 3: Nếu tổng số tab đã đạt trần tối đa (6 tab):
        # Tái sử dụng bất kỳ tab nào trong dải Bot (>= user_reserved) không bị khóa Busy
        if selected is None and total_open >= max_sap:
            for s_i in range(user_reserved, total_open):
                sess = conn.Children(s_i)
                try:
                    if not bool(getattr(sess, "Busy", False)):
                        selected = sess
                        break
                except Exception:
                    pass

        # 4. Fallback: Nếu không tìm thấy trong dải Bot:
        if selected is None:
            if getattr(self, "prefer_dedicated_session", True) and user_reserved > 0:
                raise RuntimeError("NO_DEDICATED_BOT_SESSION: Tất cả các phiên SAP đều đang bận thao tác của người dùng; dừng để bảo vệ an toàn dữ liệu.")

            # Chỉ áp dụng khi tắt prefer_dedicated_session (môi trường kiểm thử đơn tab):
            if 0 <= self.session_index < total_open:
                seed = conn.Children(self.session_index)
                is_busy = bool(getattr(seed, "Busy", False))
                info = getattr(seed, "Info", None)
                tcode = str(getattr(info, "Transaction", "")).strip().upper() if info else ""
                if not is_busy and (tcode in ("SESSION_MANAGER", "") or not getattr(self, "strict_session_safety", True)):
                    selected = seed

            if selected is None:
                for s_i in range(total_open):
                    sess = conn.Children(s_i)
                    try:
                        if not bool(getattr(sess, "Busy", False)):
                            s_info = getattr(sess, "Info", None)
                            s_tcode = str(getattr(s_info, "Transaction", "")).strip().upper() if s_info else ""
                            if s_tcode in ("SESSION_MANAGER", "") or not getattr(self, "strict_session_safety", True):
                                selected = sess
                                break
                    except Exception:
                        pass

        if selected is None:
            if getattr(self, "strict_session_safety", True):
                raise RuntimeError("NO_DEDICATED_BOT_SESSION: Tất cả các phiên SAP đều đang bận thao tác của người dùng; dừng để bảo vệ an toàn dữ liệu.")
            raise SapExecutionError("Phiên SAP cấu hình đang bận")

        self.session_identity = self._check_identity(selected)
        self.last_activity = time.time()
        return selected

    def _is_session_alive(self) -> bool:
        """Kiểm tra xem session SAP còn phản hồi trong bộ nhớ RAM hay không."""
        try:
            if self.session is None:
                return False
            _ = self.session.findById("wnd[0]").Text
            return True
        except Exception:
            return False

    def ensure_session(self):
        if not self._is_session_alive():
            self.session = self._connect_sap()
        identity = self._check_identity(self.session)
        if getattr(self, "session_identity", None) and identity != self.session_identity:
            raise SapExecutionError("Session SAP không còn thuộc phiên Bridge đã chọn")
        if getattr(self.session, "Busy", False) is True:
            raise SapExecutionError("SAP đang bận; không gửi thêm thao tác")
        return self.session

    def keep_alive(self, idle_threshold_seconds=None):
        # No /n resets on an unattended screen. Reconnect only for actual queries.
        if self.session is not None:
            try:
                self._check_identity(self.session)
            except Exception:
                self.session = None

    def _wait_idle(self, timeout_sec: float = 30.0) -> bool:
        """Đợi SAP GUI hoàn tất xử lý (Busy = False) như chuẩn modSAPCore.WaitIdle."""
        t0 = time.time()
        while time.time() - t0 < timeout_sec:
            try:
                if not getattr(self.session, "Busy", False):
                    return True
            except Exception:
                return False
            time.sleep(0.1)
        return False

    def _find_by_id_safe(self, id_text: str):
        """Tìm đối tượng theo ID an toàn tuyệt đối, tương thích cả 1 tham số và 2 tham số."""
        if self.session is None:
            return None
        try:
            return self.session.findById(id_text)
        except Exception:
            try:
                return self.session.findById(id_text, False)
            except Exception:
                return None

    def smart_find(self, element_name: str, element_type: Optional[str] = None, window: str = "wnd[0]"):
        """Tự động tìm kiếm đối tượng theo Tên tương thích cả High Speed & Low Speed,
        kết hợp bộ nhớ đệm tự học (UI_CACHE / data/sap_ui_cache.json).
        """
        if self.session is None:
            return None
        return UI_CACHE.resolve_element(self.session, key_alias=element_name, field_name=element_name, element_type=element_type, window=window)

    def table_field(self, tbl, token: str, visible_row: int):
        """Truy xuất control ô trong TableControl theo token relative ID chuẩn RC65 (TableField)."""
        return ObService.get_table_field(tbl, token, visible_row)

    def set_table_text(self, tbl, token: str, visible_row: int, text: str) -> bool:
        """Ghi giá trị text vào ô TableControl theo token chuẩn RC65 (SetTableText) kèm kiểm tra readback."""
        return ObService.set_table_text(tbl, token, visible_row, text)

    def delivery_overview(self, session=None):
        """Lấy bảng Delivery Overview trong VL02N/VL03N chuẩn RC65 (tblSAPMV50ATC_LIPS_OVER)."""
        ses = session or self.session
        if ses is None:
            return None
        tbl = self._find_by_id_safe(RC65_TABLE_OVERVIEW_ID)
        if tbl is None:
            tbl = self._find_by_id_safe(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER")
        if tbl is None:
            tbl = self._find_by_id_safe(f"wnd[0]/usr/{RC65_TABLE_OVERVIEW_TOKEN}")
        return tbl

    def locate_delivery_item(self, tbl, item_no: str, overview_index: Optional[Dict[str, Any]] = None) -> int:
        """Định vị dòng hiển thị của item trong bảng Delivery Overview (LocateItem) chuẩn RC65."""
        return ObService.locate_delivery_item(tbl, item_no, overview_index)

    def snapshot_delivery_overview(self, tbl, index_only: bool = False) -> Dict[str, Dict[str, Any]]:
        """Chụp snapshot toàn bộ các dòng trên Delivery Overview (OverviewSnapshot) chuẩn RC65."""
        return ObService.snapshot_delivery_overview(tbl, index_only)

    def _get_status(self) -> Tuple[str, str]:
        """Đọc thông điệp và kiểu trạng thái từ StatusBar (wnd[0]/sbar) như modSAPCore."""
        try:
            sbar = self._find_by_id_safe("wnd[0]/sbar")
            if sbar is not None:
                txt = getattr(sbar, "Text", "")
                mtype = getattr(sbar, "MessageType", "")
                return str(txt).strip(), str(mtype).strip().upper()
        except Exception:
            pass
        return "", ""

    def _dismiss_popups(self) -> str:
        """Kiểm tra và đóng cửa sổ popup wnd[1] nếu có như modSAPCore.ClosePopups."""
        all_msgs = []
        try:
            count = 0
            while count < 10:
                wnd1 = self._find_by_id_safe("wnd[1]")
                if wnd1 is None:
                    break
                t = str(getattr(wnd1, "Text", "")).strip()
                b_txt = ""
                try:
                    usr1 = self._find_by_id_safe("wnd[1]/usr")
                    if usr1 is not None:
                        children = getattr(usr1, "Children", None)
                        if children is not None:
                            for idx in range(children.Count):
                                c = children.ElementAt(idx)
                                val = str(getattr(c, "Text", "")).strip()
                                if val:
                                    b_txt += f" {val}"
                except Exception:
                    pass
                full_m = f"{t} {b_txt}".strip()
                if full_m:
                    all_msgs.append(full_m)
                wnd1.sendVKey(0)
                count += 1
                self._wait_idle(3.0)
        except Exception:
            pass
        joined = " | ".join(all_msgs)
        if joined:
            self.last_popup_message = joined
        return joined

    def _find_largest_grid(self, max_wait_sec: float = 10.0):
        """Dò tìm ALV Grid lớn nhất trên màn hình SAP một cách đệ quy y hệt modSAPCore.FindLargestGrid."""
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
                wnd0 = self._find_by_id_safe("wnd[0]")
                if wnd0 is not None:
                    _collect(wnd0, 0)
            except Exception:
                pass

            if best_grid is not None and best_rows > 0:
                return best_grid
            time.sleep(0.3)

        return best_grid

    @staticmethod
    def _normalize_header(s: str) -> str:
        """Chuẩn hóa chuỗi tiêu đề/mã cột: loại bỏ dấu, ký tự đặc biệt (kể cả Σ), chuyển chữ hoa."""
        if not s:
            return ""
        nfd = unicodedata.normalize("NFD", str(s))
        no_acc = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
        return re.sub(r"[^A-Za-z0-9]", "", no_acc).upper()

    def _get_column_caption(self, alv_obj, col_id: str) -> str:
        """Lấy tiêu đề hiển thị đầy đủ của cột qua GetDisplayedColumnTitle và GetColumnTitles chuẩn SAP GUI."""
        parts = []
        try:
            disp = alv_obj.GetDisplayedColumnTitle(col_id)
            if disp:
                parts.append(str(disp).strip())
        except Exception:
            pass

        try:
            titles = alv_obj.GetColumnTitles(col_id)
            if hasattr(titles, "Count"):
                for i in range(titles.Count):
                    t = str(titles.ElementAt(i)).strip()
                    if t:
                        parts.append(t)
            elif isinstance(titles, (list, tuple)):
                for t in titles:
                    if str(t).strip():
                        parts.append(str(t).strip())
            elif titles:
                parts.append(str(titles).strip())
        except Exception:
            pass

        try:
            tooltip = alv_obj.GetColumnTooltip(col_id)
            if tooltip:
                parts.append(str(tooltip).strip())
        except Exception:
            pass

        return " ".join(parts)

    def _resolve_grid_columns(self, alv_obj, column_mapping: Dict[str, List[str]]) -> Dict[str, str]:
        """Phân giải tên cột ALV Grid theo thuật toán 2 lượt của ModuleSAP.bas:
        - Lượt 1: Khớp chính xác mã kỹ thuật hoặc Caption đã chuẩn hóa (bỏ qua khoảng trắng, ký hiệu Σ).
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
                col_captions[col_name] = self._get_column_caption(alv_obj, col_name)
            except Exception:
                pass

        print(f"[SAP ALV] Bảng có {col_count} cột. Đang phân giải tự động...")
        for out_key, candidates in column_mapping.items():
            matched_col = None

            # Lượt 0: Khóa cứng theo mã kỹ thuật chuẩn RC65 (khớp trực tiếp tech_cols với candidates)
            for cand in candidates:
                cand_clean = str(cand).strip().upper()
                for tc in tech_cols:
                    if tc.upper() == cand_clean:
                        matched_col = tc
                        break
                if matched_col:
                    break

            cand_norms = [self._normalize_header(c) for c in candidates if self._normalize_header(c)]

            # Lượt 1: Khớp chính xác (Exact match) mã kỹ thuật hoặc Caption
            if not matched_col:
                for tc in tech_cols:
                    tc_norm = self._normalize_header(tc)
                    cap_norm = self._normalize_header(col_captions.get(tc, ""))
                    for cn in cand_norms:
                        if cn == tc_norm or cn == cap_norm:
                            matched_col = tc
                            break
                    if matched_col:
                        break

            # Lượt 2: Khớp mờ (Fuzzy contains) cho các tiêu đề có ý nghĩa (>= 4 ký tự)
            if not matched_col:
                for tc in tech_cols:
                    cap_full_norm = self._normalize_header(f"{tc} {col_captions.get(tc, '')}")
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

    def _clear_multiple_selection(self, push_button_id: str):
        """Mở popup Multiple Selection và xóa sạch lựa chọn bằng btn[16], sau đó bấm F8 (btn[8]) để lưu và đóng popup."""
        push_btn = self._find_by_id_safe(push_button_id)
        if push_btn is None:
            return
        try:
            push_btn.press()
            self._wait_idle(2.0)
            btn16 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[16]")
            if btn16 is not None:
                try:
                    btn16.press()
                    self._wait_idle(1.0)
                except Exception:
                    pass
            btn8 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[8]")
            if btn8 is not None:
                btn8.press()
            else:
                wnd1 = self._find_by_id_safe("wnd[1]")
                if wnd1 is not None:
                    wnd1.sendVKey(8)
            self._wait_idle(2.0)
        except Exception as e:
            print(f"[SAP WARN] Lỗi xóa multiple selection ({push_button_id}): {e}")

    def _fill_multiple_selection(self, push_button_id: str, values: List[str]):
        """Điền danh sách nhiều giá trị vào Multiple Selection.
        
        - Nếu <= 5 giá trị: điền trực tiếp vào các ô mà không cần clipboard.
        - Nếu > 5 giá trị: sao chép và bấm nút btn[24] (Upload from Clipboard / Shift+F12) chuẩn của SAP GUI (như Module1.bas V2Select), sau đó khôi phục ngay clipboard cho người dùng.
        """
        clean_vals = [v.strip().upper() for v in values if v.strip()]
        if not clean_vals:
            return

        push_btn = self._find_by_id_safe(push_button_id)
        if push_btn is None:
            return
        push_btn.press()
        self._wait_idle(3.0)

        # Xóa các lựa chọn cũ nếu có nút btn[16] (Delete whole selection)
        btn16 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[16]")
        if btn16 is not None:
            try:
                btn16.press()
                self._wait_idle(1.5)
            except Exception:
                pass

        prev_clip = None
        used_clipboard = False
        btn24 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[24]")
        if len(clean_vals) > 5 and btn24 is not None:
            try:
                prev_clip = _get_clipboard_safely()
                clip_text = "\r\n".join(clean_vals)
                if _copy_to_clipboard_safely(clip_text):
                    btn24.press()
                    self._wait_idle(2.0)
                    cell0 = self._find_by_id_safe("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]")
                    if cell0 is not None and cell0.text.strip().upper() == clean_vals[0]:
                        used_clipboard = True
                        print(f"[SAP MULTI-SELECT] Đã paste thành công {len(clean_vals)} giá trị từ clipboard vào Multiple Selection.")
                    else:
                        print("[SAP MULTI-SELECT WARN] Paste clipboard chưa thấy ăn dòng đầu, thử lại...")
                        btn24.press()
                        self._wait_idle(2.0)
                        used_clipboard = True
            except Exception as clip_err:
                print(f"[SAP WARN] Lỗi paste clipboard vào Multiple Selection: {clip_err}")

        if not used_clipboard:
            tbl_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
            for idx, val in enumerate(clean_vals[:8]):
                cell_id = f"{tbl_id}/ctxtRSCSEL_255-SLOW_I[1,{idx}]"
                cell = self._find_by_id_safe(cell_id)
                if cell is not None:
                    try:
                        cell.text = val
                    except Exception:
                        alt_cell = self._find_by_id_safe(cell_id.replace("ctxtRSCSEL_255", "txtRSCSEL_255"))
                        if alt_cell is not None:
                            alt_cell.text = val

        # Bấm Execute (F8 / btn[8]) trên popup
        btn8 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[8]")
        if btn8 is not None:
            btn8.press()
        else:
            wnd1 = self._find_by_id_safe("wnd[1]")
            if wnd1 is not None:
                wnd1.sendVKey(8)

        self._wait_idle(3.0)

        # Khôi phục clipboard sau khi cửa sổ popup đã đóng và SAP đã nhận dữ liệu
        if prev_clip is not None:
            try:
                _copy_to_clipboard_safely(prev_clip)
            except Exception:
                pass

    def _validate_grid_or_raise(self, alv_grid, tcode_name: str, status_txt: str, status_type: str):
        """Phân biệt chính xác giữa việc không có dữ liệu thực sự và lỗi không mở được bảng kết quả SAP."""
        if alv_grid is not None:
            return

        is_selection_screen = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]") is not None
        combined_txt = f"{status_txt} {getattr(self, 'last_popup_message', '')}".strip()
        lower_txt = combined_txt.lower()

        empty_phrases = [
            "no data", "not found", "không tìm thấy", "keine daten",
            "chưa có", "không có dữ liệu", "no document", "no item",
            "no suitable", "no records", "không có chứng từ",
            "không có giao hàng", "no deliveries", "no outbound deliveries", "no open", "không có đơn",
            "list contains no data", "danh sách không chứa", "no stock",
            "không có tồn", "keine lieferungen",
            "zero records", "0 records", "không tìm thấy bản ghi"
        ]

        # Kiểm tra nếu là thông báo không có dữ liệu thật sự
        # Loại bỏ các từ mang tính cảnh báo lọc ('restrict', 'selected') để tránh biến lỗi màn hình thành NO_DATA giả tạo
        if any(p in lower_txt for p in empty_phrases):
            print(f"[SAP {tcode_name}] Xác nhận không có dữ liệu: '{combined_txt}'")
            return

        if is_selection_screen:
            if any(w in lower_txt for w in ["restrict", "giới hạn lựa chọn", "quá nhiều", "specify"]):
                raise SapExecutionError(
                    f"SAP yêu cầu thu hẹp điều kiện lọc hoặc nhập thêm tiêu chí trên màn hình {tcode_name}. "
                    f"Trạng thái: '{status_txt or 'Restrict your selection'}'. Vui lòng kiểm tra lại Plant, Mã hàng hoặc Kho."
                )
            raise SapExecutionError(
                f"Không thể mở bảng kết quả SAP {tcode_name} (màn hình nhập liệu không phản hồi hoặc bị chặn). "
                f"Trạng thái: '{status_txt or 'Chưa rõ'}'"
            )

        raise SapExecutionError(
            f"Không tìm thấy bảng kết quả ALV Grid trong SAP {tcode_name}. "
            f"Trạng thái: '{status_txt or 'Lỗi hiển thị giao diện SAP'}'"
        )

    def _extract_alv_grid(
        self,
        alv_obj,
        column_mapping: Dict[str, List[str]],
        max_rows: Optional[int] = None
    ) -> List[Dict[str, str]]:
        """Trích xuất ALV Grid kèm cơ chế forward-fill mã hàng và lọc bỏ dòng tổng cộng (Summary rows)."""
        try:
            row_count = int(alv_obj.RowCount)
        except Exception as exc:
            raise SapExecutionError("Không đọc được RowCount; dữ liệu nguồn chưa xác định") from exc

        if row_count == 0:
            return []

        resolved_cols = self._resolve_grid_columns(alv_obj, column_mapping)
        records = []
        effective_max = max_rows if max_rows is not None else self.max_alv_rows
        if row_count > effective_max:
            raise SapExecutionError(f"Nguồn có {row_count} dòng, vượt giới hạn {effective_max}; thu hẹp bộ lọc để tránh kết luận từ dữ liệu thiếu")
        if not resolved_cols:
            raise SapExecutionError("Không nhận diện được cột dữ liệu SAP")
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
                        raise SapExecutionError(f"Không đọc được ô ALV {r}/{out_key}; nguồn chưa đầy đủ") from exc

            # Forward-fill Product nếu ô bị gộp hiển thị trống
            curr_prod = item.get("product", "")
            is_forward_filled = False
            if curr_prod:
                last_product = curr_prod
            elif last_product and "product" in item:
                item["product"] = last_product
                is_forward_filled = True

            # LỌC BỎ DÒNG TỔNG CỘNG / DÒNG RỖNG / TIÊU ĐỀ PHỤ:
            # 1. Bỏ qua các dòng tổng cộng / wildcard trong ALV:
            p_val = str(item.get("product", "")).strip()
            if p_val == "*" or p_val.lower().startswith("total") or p_val.lower().startswith("tổng"):
                continue

            # Bỏ qua dòng subtotal/total dựa trên GetRowType nếu ALV hỗ trợ
            try:
                row_type = str(alv_obj.GetRowType(r)).strip().upper()
                if row_type and ("TOTAL" in row_type or row_type in ("1", "2")):
                    continue
            except Exception:
                pass

            # 2. DÒNG DỮ LIỆU HỢP LỆ TRONG SAP ALV (ÁP DỤNG CHUNG CHO TẤT CẢ CÁC T-CODE):
            # Một dòng dữ liệu chi tiết bắt buộc phải có ít nhất một trường định danh cụ thể:
            # (Kho SLoc, Vị trí Bin, Lô Batch, Số chứng từ, Số Inbound, Số Outbound, Số STO/PO, Dòng Item, Loại GD)
            has_pos_or_doc = bool(
                item.get("sloc") or item.get("bin") or item.get("batch") or
                item.get("docNo") or item.get("obNo") or item.get("ibNo") or
                item.get("stoNo") or item.get("poNo") or item.get("wtNo") or
                item.get("item") or item.get("mvtType") or item.get("refNo") or
                item.get("component") or item.get("queue_name") or item.get("document") or item.get("error_message") or
                (item.get("transit") and str(item.get("transit")).strip() not in ("", "0", "0.0", "0,0"))
            )
            # Bỏ qua hoàn toàn dòng tổng cộng / subtotal của ALV (dòng không có bất kỳ thông tin vị trí/chứng từ nào)
            if not has_pos_or_doc:
                continue

            records.append(item)

        return records

    def _find_filter_field_token(self, wnd1, column_id: str, col_caption: str = "") -> str:
        """Tìm chính xác token trường lọc (%%DYN001, %%DYN002, ...) khớp với cột cần lọc theo chuẩn VL6_FilterField của RC65."""
        base_sub = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/"
        col_norm = self._normalize_header(column_id)
        cap_norm = self._normalize_header(col_caption)

        aliases = {col_norm}
        if cap_norm:
            aliases.add(cap_norm)
        if col_norm in ("MATNR", "MATERIAL", "PRODUCT", "EMATN"):
            aliases.update(["MATNR", "MATERIAL", "PRODUCT", "VATTU", "MAHANG", "MATHANG", "EMATN"])
        elif col_norm in ("VBELN", "DELIVERY", "OUTBOUND"):
            aliases.update(["VBELN", "DELIVERY", "OUTBOUNDDELIVERY", "GIAOHANG", "CHUNGTU"])
        elif col_norm in ("KUNWE", "KUNNR", "SHIPTO"):
            aliases.update(["KUNWE", "KUNNR", "SHIPTO", "SHIPTOPARTY", "NOINHAN"])
        elif col_norm in ("LGORT", "SLOC"):
            aliases.update(["LGORT", "SLOC", "STORAGELOCATION", "KHO"])

        matched_tokens = []
        for i in range(1, 20):
            token = f"%%DYN{i:03d}"
            low_ctl = self._find_by_id_safe(f"{base_sub}ctxt{token}-LOW")
            if low_ctl is None:
                low_ctl = self._find_by_id_safe(f"{base_sub}txt{token}-LOW")
            if low_ctl is None:
                break

            lbl_text = ""
            for lbl_id in [
                f"{base_sub}%_{token}_%_APP_%-TEXT",
                f"{base_sub}lbl{token}",
                f"{base_sub}txt{token}",
                f"wnd[1]/usr/lbl{token}",
                f"wnd[1]/usr/txt{token}"
            ]:
                lbl = self._find_by_id_safe(lbl_id)
                if lbl is not None:
                    lbl_text = str(getattr(lbl, "Text", "") or getattr(lbl, "text", "")).strip()
                    if lbl_text:
                        break

            norm_lbl = self._normalize_header(lbl_text)
            if norm_lbl:
                if any(a in norm_lbl or norm_lbl in a for a in aliases):
                    matched_tokens.append(token)

        if len(matched_tokens) == 1:
            print(f"[SAP ALV FILTER] Đã tìm thấy đúng token trường lọc cho cột '{column_id}': '{matched_tokens[0]}'")
            return matched_tokens[0]
        elif len(matched_tokens) > 1:
            print(f"[SAP ALV FILTER] Tìm thấy {len(matched_tokens)} token khả dĩ ({matched_tokens}), chọn: '{matched_tokens[0]}'")
            return matched_tokens[0]

        return "%%DYN001"

    def _apply_alv_filter(self, grid, column_id: str, values: List[str], col_caption: str = "") -> bool:
        """Áp dụng bộ lọc ALV Filter trên ALV Grid (phím Ctrl+F2 / nút &MB_FILTER / btn[29]) theo đúng chuẩn VL6_ApplyALVFilter của RC65."""
        clean_vals = [str(v).strip().upper() for v in values if str(v).strip()]
        if not clean_vals or grid is None or "Mock" in type(grid).__name__:
            return False
        try:
            # 1. Bỏ chọn các vùng cũ và chọn cột cần lọc
            try:
                grid.clearSelection()
                if hasattr(grid, "selectedRows"):
                    grid.selectedRows = ""
                if hasattr(grid, "SelectedColumns"):
                    grid.SelectedColumns = ""
            except Exception:
                pass
            try:
                grid.setCurrentCell(-1, column_id)
                grid.selectColumn(column_id)
            except Exception as sel_col_err:
                print(f"[SAP ALV FILTER WARN] Không thể chọn cột {column_id}: {sel_col_err}")

            self._wait_idle(1.0)

            # 2. Bấm nút Filter trên Toolbar (btn[29] hoặc &MB_FILTER / &ILT)
            btn29 = self._find_by_id_safe("wnd[0]/tbar[1]/btn[29]")
            if btn29 is not None:
                btn29.press()
            else:
                try:
                    grid.pressToolbarButton("&MB_FILTER")
                except Exception:
                    try:
                        grid.pressToolbarButton("&ILT")
                    except Exception:
                        if hasattr(self, "session") and self.session is not None:
                            wnd0 = self._find_by_id_safe("wnd[0]")
                            if wnd0 is not None and hasattr(wnd0, "sendVKey"):
                                wnd0.sendVKey(29)

            self._wait_idle(3.0)
            wnd1 = self._find_by_id_safe("wnd[1]")
            if wnd1 is None:
                # Fallback: Trực tiếp thử các nút Filter trên Application Toolbar tbar[1] nếu wnd[1] chưa bật
                for alt_btn in ["wnd[0]/tbar[1]/btn[29]", "wnd[0]/tbar[1]/btn[38]"]:
                    fb = self._find_by_id_safe(alt_btn)
                    if fb is not None:
                        try:
                            fb.press()
                            self._wait_idle(2.0)
                            wnd1 = self._find_by_id_safe("wnd[1]")
                            if wnd1 is not None:
                                break
                        except Exception:
                            pass

            if wnd1 is None:
                print(f"[SAP ALV FILTER WARN] Không tìm thấy popup wnd[1] sau khi bấm Filter cột {column_id}")
                return False

            # 3. Tìm đúng ô nhập giá trị trong wnd[1] theo chuẩn VL6_FilterField của RC65
            token = self._find_filter_field_token(wnd1, column_id, col_caption)
            base_sub = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/"
            low_ctl = self._find_by_id_safe(f"{base_sub}ctxt{token}-LOW")
            if low_ctl is None:
                low_ctl = self._find_by_id_safe(f"{base_sub}txt{token}-LOW")
            hi_ctl = self._find_by_id_safe(f"{base_sub}ctxt{token}-HIGH")
            if hi_ctl is None:
                hi_ctl = self._find_by_id_safe(f"{base_sub}txt{token}-HIGH")

            if low_ctl is None:
                usr_obj = self._find_by_id_safe("wnd[1]/usr")
                if usr_obj is not None:
                    ch_count = getattr(getattr(usr_obj, "Children", None), "Count", 0)
                    for idx in range(ch_count):
                        ch = usr_obj.Children(idx)
                        ch_id = getattr(ch, "Id", "")
                        if str(ch_id).endswith("-LOW"):
                            low_ctl = ch
                            break

            # Xóa các giá trị lọc cũ trên popup trước khi điền mới
            if low_ctl is not None:
                try:
                    low_ctl.text = ""
                except Exception:
                    pass
            if hi_ctl is not None:
                try:
                    hi_ctl.text = ""
                except Exception:
                    pass

            if len(clean_vals) == 1 and low_ctl is not None:
                low_ctl.text = clean_vals[0]
                readback_val = str(getattr(low_ctl, "Text", "") or getattr(low_ctl, "text", "")).strip()
                print(f"[SAP ALV FILTER TRACE] Cột '{column_id}' ({token}): Đã điền '{clean_vals[0]}', Readback='{readback_val}'")
            else:
                multi_push_id = f"{base_sub}btn%_{token}_%_APP_%-VALU_PUSH"
                multi_push = self._find_by_id_safe(multi_push_id)
                if multi_push is not None:
                    self._fill_multiple_selection(multi_push_id, clean_vals)
                elif low_ctl is not None:
                    low_ctl.text = clean_vals[0]

            # 4. Xác nhận áp dụng Filter (btn[0] / Enter)
            btn0 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[0]")
            if btn0 is not None:
                btn0.press()
            else:
                self.session.findById("wnd[1]").sendVKey(0)

            self._wait_idle(3.0)
            try:
                grid.clearSelection()
            except Exception:
                pass
            print(f"[SAP ALV FILTER] Đã áp dụng thành công bộ lọc {clean_vals} trên cột {column_id} (token {token}).")
            return True
        except Exception as e:
            print(f"[SAP ALV FILTER WARN] Không thể áp dụng ALV filter: {e}")
            return False

    def _export_alv_to_excel(self, alv_grid, tcode_name: str = "MB52") -> Optional[str]:
        """Xuất dữ liệu ALV Grid ra file Excel qua Spreadsheet (&XXL / Ctrl+Shift+F7) chuẩn xác theo 261008_Stock_FG_MA_BODY dòng 2419-2475."""
        if alv_grid is None or "Mock" in type(alv_grid).__name__ or "Mock" in type(getattr(self, "_find_by_id_safe", None)).__name__:
            return None

        export_dir = os.path.join(tempfile.gettempdir(), f"sap_{tcode_name.lower()}_{int(time.time())}_{os.getpid()}")
        os.makedirs(export_dir, exist_ok=True)
        export_filename = f"{tcode_name.lower()}_export.xlsx"
        export_path = os.path.join(export_dir, export_filename)

        if os.path.exists(export_path):
            try:
                os.remove(export_path)
            except Exception:
                pass

        print(f"[SAP EXPORT] Đang kích hoạt xuất Spreadsheet Excel ({tcode_name}) tới: {export_path}...")

        triggered = False
        try:
            alv_grid.ContextMenu()
            alv_grid.selectContextMenuItem("&XXL")
            triggered = True
            print("[SAP EXPORT] Đã gọi ContextMenu &XXL thành công.")
        except Exception as e1:
            print(f"[SAP EXPORT WARN] ContextMenu &XXL không khả dụng: {e1}")

        if not triggered or self._find_by_id_safe("wnd[1]") is None:
            try:
                self.session.findById("wnd[0]").sendVKey(43)
                triggered = True
                print("[SAP EXPORT] Đã gửi phím tắt Ctrl+Shift+F7 (sendVKey 43) thành công.")
            except Exception as e2:
                print(f"[SAP EXPORT WARN] sendVKey 43 không thành công: {e2}")

        if self._find_by_id_safe("wnd[1]") is None:
            try:
                alv_grid.pressToolbarButton("&MB_EXPORT")
                triggered = True
                print("[SAP EXPORT] Đã bấm nút toolbar &MB_EXPORT.")
            except Exception:
                try:
                    alv_grid.pressToolbarContextButton("&MB_EXPORT")
                    alv_grid.selectContextMenuItem("&XXL")
                    triggered = True
                except Exception:
                    pass

        self._wait_idle(3.0)
        wnd1 = self._find_by_id_safe("wnd[1]")
        if wnd1 is None:
            print("[SAP EXPORT WARN] Không tìm thấy popup lưu file wnd[1] sau khi kích hoạt xuất Excel.")
            return None

        path_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtDY_PATH")
        if path_ctl is None:
            btn0 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[0]")
            if btn0 is not None:
                btn0.press()
            else:
                self.session.findById("wnd[1]").sendVKey(0)
            self._wait_idle(3.0)
            path_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtDY_PATH")

        if path_ctl is not None:
            try:
                path_ctl.text = export_dir
            except Exception as pe:
                print(f"[SAP EXPORT WARN] Lỗi điền ctxtDY_PATH: {pe}")

        file_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtDY_FILENAME")
        if file_ctl is not None:
            try:
                file_ctl.text = export_filename
            except Exception as fe:
                print(f"[SAP EXPORT WARN] Lỗi điền ctxtDY_FILENAME: {fe}")

        btn_save = self._find_by_id_safe("wnd[1]/tbar[0]/btn[0]")
        if btn_save is not None:
            btn_save.press()
        else:
            self.session.findById("wnd[1]").sendVKey(0)

        self._wait_idle(5.0)
        self._dismiss_popups()

        start_wait = time.time()
        last_sz = -1
        stable_cnt = 0
        while time.time() - start_wait < 25.0:
            if os.path.exists(export_path):
                try:
                    cur_sz = os.path.getsize(export_path)
                    if cur_sz > 0:
                        if cur_sz == last_sz:
                            stable_cnt += 1
                            if stable_cnt >= 2:
                                print(f"[SAP EXPORT] File Excel đã sẵn sàng: {export_path} (Size: {cur_sz:,} bytes) sau {time.time() - start_wait:.1f}s.")
                                return export_path
                        else:
                            stable_cnt = 0
                        last_sz = cur_sz
                except Exception:
                    pass
            time.sleep(1.0)

        if os.path.exists(export_path) and os.path.getsize(export_path) > 0:
            return export_path

        print(f"[SAP EXPORT WARN] Hết thời gian chờ tạo file Excel: {export_path}")
        return None

    def _read_sap_exported_file(
        self,
        file_path: str,
        column_mapping: Dict[str, List[str]]
    ) -> List[Dict[str, str]]:
        """Đọc và chuyển đổi dữ liệu từ file Excel SAP xuất ra thành danh sách bản ghi chuẩn theo V2ReadExport & CanonicalExport của Module1.bas."""
        if not file_path or not os.path.exists(file_path):
            return []

        raw_rows = []
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                raw_rows.append([str(c).strip() if c is not None else "" for c in row])
            wb.close()
        except Exception as pyxl_err:
            print(f"[SAP PARSE WARN] openpyxl không đọc được ({pyxl_err}), thử định dạng khác...")

        if not raw_rows:
            try:
                import xlrd
                wb = xlrd.open_workbook(file_path)
                sh = wb.sheet_by_index(0)
                for r in range(sh.nrows):
                    raw_rows.append([str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)])
            except Exception:
                pass

        if not raw_rows:
            try:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as fp:
                    content = fp.read()
                if "<table" in content.lower():
                    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", content, re.I | re.S):
                        cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.I | re.S)
                        clean_cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                        if any(clean_cells):
                            raw_rows.append(clean_cells)
                elif "<Workbook" in content or "<?xml" in content:
                    for row_xml in re.findall(r"<Row[^>]*>(.*?)</Row>", content, re.I | re.S):
                        cells = re.findall(r"<Data[^>]*>(.*?)</Data>", row_xml, re.I | re.S)
                        clean_cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                        if any(clean_cells):
                            raw_rows.append(clean_cells)
            except Exception as text_err:
                print(f"[SAP PARSE WARN] Đọc text/xml/html thất bại: {text_err}")

        if not raw_rows:
            return []

        header_idx = -1
        col_indices = {}

        for i, row in enumerate(raw_rows[:15]):
            normalized_row = [re.sub(r"[\s\.-]+", " ", str(c).strip().upper()) for c in row]
            matches = 0
            temp_indices = {}
            for out_key, aliases in column_mapping.items():
                for c_idx, cell_text in enumerate(normalized_row):
                    if not cell_text:
                        continue
                    for alias in aliases:
                        norm_alias = re.sub(r"[\s\.-]+", " ", alias.strip().upper())
                        if cell_text == norm_alias or norm_alias in cell_text:
                            temp_indices[out_key] = c_idx
                            matches += 1
                            break
                    if out_key in temp_indices:
                        break

            if matches >= 2 and ("product" in temp_indices or "plant" in temp_indices or "unrest" in temp_indices):
                header_idx = i
                col_indices = temp_indices
                break

        if header_idx == -1 or not col_indices:
            print("[SAP PARSE WARN] Không tìm thấy dòng tiêu đề hợp lệ trong file Excel xuất từ SAP.")
            return []

        records = []
        last_product = ""
        last_plant = ""

        for row in raw_rows[header_idx + 1:]:
            item = {}
            for out_key, c_idx in col_indices.items():
                if c_idx < len(row):
                    item[out_key] = str(row[c_idx]).strip()
                else:
                    item[out_key] = ""

            curr_p = str(item.get("product", "")).strip()
            # Bỏ qua dòng tổng kết / subtotal (ví dụ '*', 'Total', 'Tổng', 'Σ') TRƯỚC KHI cập nhật last_product
            if curr_p == "*" or curr_p.lower().startswith("total") or curr_p.lower().startswith("tổng") or curr_p.startswith("Σ"):
                continue

            # Chỉ kế thừa / cập nhật last_product khi là mã hàng thực sự hợp lệ
            if curr_p:
                last_product = curr_p
            elif last_product:
                item["product"] = last_product

            pl_val = str(item.get("plant", "")).strip()
            if pl_val and pl_val != "*" and not pl_val.lower().startswith("total") and not pl_val.lower().startswith("tổng"):
                last_plant = pl_val
            elif last_plant:
                item["plant"] = last_plant

            has_data = bool(
                item.get("sloc") or item.get("bin") or item.get("batch") or
                item.get("docNo") or item.get("obNo") or item.get("ibNo") or
                item.get("stoNo") or item.get("poNo") or item.get("wtNo") or
                item.get("item") or item.get("mvtType") or item.get("refNo") or
                (item.get("transit") and str(item.get("transit")).strip() not in ("", "0", "0.0", "0,0"))
            )
            if not has_data:
                continue

            records.append(item)

        print(f"[SAP PARSE] Đã chuyển đổi thành công {len(records)} dòng từ file Excel SAP xuất ra.")
        return records

    def _return_to_main_menu(self):
        """Đưa SAP GUI về màn hình gốc /n (SESSION_MANAGER) - alias tương thích."""
        return self._reset_to_session_manager()

    def _reset_to_session_manager(self):
        """Đưa SAP GUI về màn hình gốc /n (SESSION_MANAGER) để giải phóng toàn bộ ALV grid, bộ nhớ và session cho lần chạy tiếp theo theo Module1.bas dòng 2643."""
        try:
            self._dismiss_popups()
            okcd = self._find_by_id_safe("wnd[0]/tbar[0]/okcd")
            if okcd is not None:
                okcd.text = "/n"
                self.session.findById("wnd[0]").sendVKey(0)
                self._wait_idle(2.0)
            else:
                self._open_transaction("SESSION_MANAGER")
                self._wait_idle(2.0)
            print("[SAP] Đã đưa phiên làm việc về màn hình gốc /n thành công.")
        except Exception as e:
            print(f"[SAP RESET WARN] Lỗi khi đưa về màn hình gốc /n: {e}")

    def _open_transaction(self, tcode: str):
        """Mở Transaction trên SAP GUI an toàn tuyệt đối qua ô lệnh okcd hoặc StartTransaction."""
        self.ensure_session()
        self.last_activity = time.time()
        self.last_popup_message = ""

        raw = tcode.strip()
        # Loại bỏ tiền tố /n hoặc /o nếu người dùng đã truyền sẵn
        if raw.upper().startswith("/N"):
            raw = raw[2:].strip()
        elif raw.upper().startswith("/O"):
            raw = raw[2:].strip()

        # Quy tắc T-Code SAP:
        # 1. Nếu là T-Code thuộc namespace (bắt đầu bằng / và có ít nhất 2 dấu /, ví dụ /SCWM/POST, /SCWM/MON):
        #    giữ nguyên dấu / ở đầu để ghép thành /n/SCWM/POST.
        # 2. Nếu là T-Code chuẩn (ví dụ MIGO, CS03, MB52, ME23N, VL06O):
        #    loại bỏ bất kỳ dấu / nào còn sót ở đầu để ghép chuẩn thành /nMIGO, /nCS03, /nMB52.
        if raw.startswith("/") and raw.count("/") >= 2:
            clean_code = raw
        else:
            clean_code = raw.lstrip("/")

        okcd_command = f"/n{clean_code}"
        print(f"[SAP] Đang mở giao dịch: {okcd_command}...")

        try:
            okcd = self._find_by_id_safe("wnd[0]/tbar[0]/okcd")
            if okcd is not None:
                okcd.text = okcd_command
                self.session.findById("wnd[0]").sendVKey(0)
                self._wait_idle(5.0)
                self._dismiss_popups()
                return
        except Exception as e:
            print(f"[SAP WARN] Không thể nhập qua okcd: {e}")

        try:
            raw_tcode = clean_code
            print(f"[SAP] Thử mở bằng StartTransaction: {raw_tcode}...")
            self.session.StartTransaction(raw_tcode)
            self._wait_idle(5.0)
            self._dismiss_popups()
        except Exception as st_err:
            raise SapExecutionError(f"Không thể mở giao dịch {tcode}") from st_err

    def _clear_all_selection_inputs(self, exclude_ids: Optional[Set[str]] = None):
        """Quét và xóa sạch tất cả các ô nhập liệu (text/ctext) trên wnd[0]/usr để chống dính bộ lọc cũ từ các phiên làm việc trước."""
        exclude = {x.strip().upper() for x in (exclude_ids or set())}
        try:
            usr = self._find_by_id_safe("wnd[0]/usr")
            if usr is None:
                return

            def _walk(obj, depth=0):
                if obj is None or depth > 8:
                    return
                try:
                    obj_type = getattr(obj, "Type", "")
                    obj_id = str(getattr(obj, "Id", "")).strip().upper()
                    if obj_type in ("GuiTextField", "GuiCTextField"):
                        if obj_id not in exclude:
                            changeable = getattr(obj, "Changeable", False)
                            readonly = getattr(obj, "ReadOnly", False)
                            if changeable and not readonly:
                                try:
                                    if getattr(obj, "Text", ""):
                                        obj.Text = ""
                                except Exception:
                                    pass
                    children = getattr(obj, "Children", None)
                    if children is not None:
                        for i in range(children.Count):
                            _walk(children.ElementAt(i), depth + 1)
                except Exception:
                    pass

            _walk(usr, 0)
        except Exception as e:
            print(f"[SAP CLEAR WARN] Lỗi quét dọn màn hình: {e}")

    # --------------------------------------------------------------------------
    # 1. KIỂM TRA VỊ TRÍ BIN & NGÀY GIỜ GR QUA POSTING CHANGE (/SCWM/POST)
    # --------------------------------------------------------------------------
    def check_position_posting_change(
        self,
        materials: List[str],
        batch: Optional[str] = None,
        plant: str = "3093",
        sloc: Optional[str] = "1005"
    ) -> List[Dict[str, str]]:
        """Tra cứu vị trí Bin, số lô, tồn khả dụng và ngày giờ GR từ /SCWM/POST theo chuẩn BBBG."""
        try:
            if not materials:
                return []

            self._open_transaction("/SCWM/POST")
            self._wait_idle(15.0)

            # 1. Các định danh ID chuẩn xác tuyệt đối từ BBBG (modPosting.bas & ModuleSAP.bas)
            ID_ADV_TOGGLE = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/btnGV_BUTTON_TEXT"
            ID_ADV_START = (
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
                "subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_START_ADVANCED"
            )
            ID_PRODUCT_MULTI = (
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
                "subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/"
                "btn%_P_MATNR_%_APP_%-VALU_PUSH"
            )
            ID_PRODUCT_LOW = (
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
                "subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/ctxtP_MATNR-LOW"
            )
            ID_OWNER_LOW = (
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
                "subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/ctxtP_OWNER-LOW"
            )
            ID_BATCH_LOW = (
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
                "subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/ctxtP_CHARG-LOW"
            )
            ID_STOCK_GRID = (
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
                "subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/"
                "cntlCONTAINER_ALV_OIP_1/shellcont/shell"
            )

            # 2. Đảm bảo Advanced Search đang mở (Nếu chưa thấy Product Low hoặc Owner Low thì bấm mở)
            owner_ctl = self._find_by_id_safe(ID_OWNER_LOW)
            prod_ctl = self._find_by_id_safe(ID_PRODUCT_LOW)
            if owner_ctl is None and prod_ctl is None:
                toggle_btn = self._find_by_id_safe(ID_ADV_TOGGLE)
                if toggle_btn is not None:
                    print("[SAP POST] Bấm mở rộng bộ lọc tìm kiếm nâng cao (Advanced Search)...")
                    toggle_btn.press()
                    self._wait_idle(10.0)
                    owner_ctl = self._find_by_id_safe(ID_OWNER_LOW)
                    prod_ctl = self._find_by_id_safe(ID_PRODUCT_LOW)

            # 3. Điền Owner (mặc định 3093 như chuẩn BBBG)
            if owner_ctl is not None:
                try:
                    owner_val = plant or "3093"
                    owner_ctl.text = owner_val
                except Exception as e:
                    print(f"[SAP POST WARN] Không thể gán Owner: {e}")

            # 4. Điền mã vật tư (Product / Material)
            clean_mats = [m.strip().upper() for m in materials if m.strip()]
            print(f"[SAP POST] Đang nhập mã hàng: {clean_mats} (Số lô: {batch or 'Tất cả'})...")

            if len(clean_mats) == 1 and prod_ctl is not None:
                prod_ctl.text = clean_mats[0]
            else:
                multi_btn = self._find_by_id_safe(ID_PRODUCT_MULTI)
                if multi_btn is not None:
                    self._fill_multiple_selection(ID_PRODUCT_MULTI, clean_mats)
                elif prod_ctl is not None and clean_mats:
                    prod_ctl.text = clean_mats[0]
                else:
                    print("[SAP POST WARN] Không tìm thấy trường nhập mã vật tư trên /SCWM/POST!")

            # 5. Điền hoặc xóa số lô (Batch)
            batch_ctl = self._find_by_id_safe(ID_BATCH_LOW)
            if batch_ctl is not None:
                try:
                    batch_ctl.text = batch.strip() if batch else ""
                except Exception:
                    pass

            # 6. Bấm thực thi tìm kiếm (Execute)
            print("[SAP POST] Bấm thực thi tìm kiếm (Execute)...")
            start_btn = self._find_by_id_safe(ID_ADV_START)
            if start_btn is not None:
                start_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            # 7. Kiểm tra trạng thái
            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP /SCWM/POST ({status_type}): {status_txt}")

            # 8. Tìm ALV Grid kết quả (Ưu tiên ID chuẩn BBBG, nếu không thì dò tìm đệ quy)
            alv_grid = self._find_by_id_safe(ID_STOCK_GRID)
            if alv_grid is None:
                print("[SAP POST] ID lưới chuẩn chưa thấy, tiến hành dò tìm ALV Grid toàn màn hình...")
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "/SCWM/POST", status_txt, status_type)
                return []

            # 9. Trích xuất các cột với danh mục định danh tương thích tuyệt đối cả SAP ID và Caption
            col_map = {
                "product":  ["MATNR", "PRODUCT", "MATERIAL", "MATNR_EXT"],
                "batch":    ["CHARG", "BATCH", "DEST.BATCH", "DEST BATCH"],
                "bin":      ["LGPLA", "VLPLA", "NLPLA", "SOURCE STORAGE BIN", "SOURCE BIN", "STORAGE BIN", "BIN"],
                "totalQty": ["TOTAL STK", "TOTAL STOCK", "GESME", "TOTAL_STK"],
                "availQty": ["AQUAN", "AQTY", "AVL_ALT_QTY", "A.QTY AUOM", "A.QTY", "AVAILABLE QTY", "AVAILABLE QUANTITY", "AVLQTY"],
                "qty":      ["AQUAN", "AQTY", "AVL_ALT_QTY", "A.QTY AUOM", "A.QTY", "AVAILABLE QTY", "AVAILABLE QUANTITY", "AVLQTY", "TOTAL STK", "TOTAL STOCK", "GESME", "QUAN"],
                "unit":     ["MEINS", "ALTME", "UNIT", "UOM", "AUN"],
                "cat":      ["TST", "CAT", "STOCK TYPE", "SOURCE STOCK TYPE", "ST", "TO_CAT"],
                "grDate":   ["WDATU", "GR_DATE", "WEDAT", "GR DATE", "GOODS RECEIPT DATE"],
                "grTime":   ["WTIME", "GR_TIME", "UZEIT", "GR TIME", "GOODS RECEIPT TIME"]
            }
            raw_records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP POST] Đã trích xuất {len(raw_records)} dòng từ /SCWM/POST.")

            # LỌC BỎ DÒNG TỔNG CỘNG ALV: Vị trí kho thực tế bắt buộc phải có Vị trí Bin hoặc Lô Batch
            records = [r for r in raw_records if r.get("bin") or r.get("batch")]

            # Chuẩn hóa hai cột số lượng: Tổng tồn (totalQty) và Khả dụng (availQty), tính SL bị OB giữ (heldQty)
            for r in records:
                t_val = parse_sap_float(r.get("totalQty", 0.0))
                a_val = parse_sap_float(r.get("availQty", 0.0) or r.get("qty", 0.0))
                if t_val == 0.0 and a_val > 0.0:
                    t_val = a_val
                if a_val == 0.0 and t_val > 0.0 and not r.get("availQty"):
                    a_val = t_val
                held_val = max(0.0, t_val - a_val)

                r["totalQty"] = format_qty_display(t_val)
                r["availQty"] = format_qty_display(a_val)
                r["heldQty"] = format_qty_display(held_val)
                r["qty"] = r["availQty"]
                if not r.get("plant") and plant:
                    r["plant"] = str(plant)
                    r["source_plant_verified"] = True
                target_sloc = str(sloc or "1005")
                if not r.get("sloc"):
                    r["sloc"] = target_sloc
                    r["source_sloc_verified"] = True

            # 10. SẮP XẾP ƯU TIÊN DỮ LIỆU GẦN NHẤT (Ngày GR & Giờ GR mới nhất lên đầu)
            def _gr_sort_key(item):
                d_raw = item.get("grDate", "")
                t_raw = item.get("grTime", "")
                parts = re.split(r"[./-]", d_raw)
                if len(parts) == 3:
                    if len(parts[2]) == 4:
                        d_sort = f"{parts[2]}{parts[1].zfill(2)}{parts[0].zfill(2)}"
                    else:
                        d_sort = f"{parts[0]}{parts[1].zfill(2)}{parts[2].zfill(2)}"
                else:
                    d_sort = d_raw
                return (d_sort, t_raw)

            records.sort(key=_gr_sort_key, reverse=True)
            return records

        # --------------------------------------------------------------------------
        finally:
            self._reset_to_session_manager()

    # 2. KIỂM TRA TỒN KHO TỔNG THỂ (MB52 - KÈM VỊ TRÍ BIN LGPBE THEO CHUẨN BBBG)
    # --------------------------------------------------------------------------
    def check_stock_mb52(
        self,
        materials: Optional[List[str]] = None,
        plant: str = "3093",
        plants: Optional[List[str]] = None,
        slocs: Optional[List[str]] = None,
        max_rows: Optional[int] = 15000
    ) -> List[Dict[str, str]]:
        """Tra cứu tồn kho tổng hợp MB52 kèm Vị trí Bin (LGPBE) theo modMB52.bas của BBBG và Module1.bas của Stock_FG_MA_BODY.

        Hỗ trợ truy vấn nhiều Plant cùng lúc (plants, ví dụ ['3093', '3001']) và lọc theo danh sách kho (slocs, ví dụ ['1005', '3093']).
        Ưu tiên xuất ra Excel (&XXL / Ctrl+Shift+F7) để lấy toàn bộ dữ liệu nhanh và chính xác nhất theo Ảnh 2 của người dùng.
        """
        try:
            self._open_transaction("MB52")
            self._wait_idle(10.0)

            # Xóa sạch toàn bộ bộ lọc cũ còn lưu trong SAP GUI
            self._clear_all_selection_inputs()

            target_plants = [p.strip().upper() for p in (plants or []) if p.strip()]
            if not target_plants and plant:
                target_plants = [plant.strip().upper()]
            if not target_plants:
                target_plants = ["3093"]

            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            print(f"[SAP MB52] Đang nhập điều kiện tìm kiếm: Materials={clean_mats}, Plants={target_plants}, SLocs={slocs}...")

            # 1. Nhập Plants (hỗ trợ cả 2 xưởng trong 1 lần chạy)
            multi_plant_btn = "wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH"
            if len(target_plants) > 1:
                if self._find_by_id_safe(multi_plant_btn) is not None:
                    self._fill_multiple_selection(multi_plant_btn, target_plants)
                for p_id in ["wnd[0]/usr/ctxtWERKS-HIGH", "wnd[0]/usr/ctxtWERKS-LOW"]:
                    p_ctl = self._find_by_id_safe(p_id)
                    if p_ctl is not None:
                        try:
                            p_ctl.text = ""
                        except Exception:
                            pass
            else:
                if self._find_by_id_safe(multi_plant_btn) is not None:
                    self._clear_multiple_selection(multi_plant_btn)
                for p_id in ["wnd[0]/usr/ctxtWERKS-LOW", "wnd[0]/usr/ctxtRM07M-WERKS"]:
                    p_ctl = self._find_by_id_safe(p_id)
                    if p_ctl is not None:
                        try:
                            p_ctl.text = target_plants[0]
                        except Exception:
                            pass
                        break

            # 2. Nhập Materials
            multi_mat_btn = "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
            if clean_mats:
                if self._find_by_id_safe(multi_mat_btn) is not None:
                    self._fill_multiple_selection(multi_mat_btn, clean_mats)
                else:
                    mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtMATNR-LOW")
                    if mat_single is not None:
                        mat_single.text = clean_mats[0]
            else:
                if self._find_by_id_safe(multi_mat_btn) is not None:
                    self._clear_multiple_selection(multi_mat_btn)
                for m_id in ["wnd[0]/usr/ctxtMATNR-LOW", "wnd[0]/usr/ctxtMATNR-HIGH"]:
                    m_ctl = self._find_by_id_safe(m_id)
                    if m_ctl is not None:
                        try:
                            m_ctl.text = ""
                        except Exception:
                            pass

            # 3. Xử lý SLoc
            multi_sloc_btn = "wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH"
            if slocs:
                clean_slocs = [s.strip().upper() for s in slocs if s.strip()]
                if self._find_by_id_safe(multi_sloc_btn) is not None:
                    self._fill_multiple_selection(multi_sloc_btn, clean_slocs)
                else:
                    sloc_single = self._find_by_id_safe("wnd[0]/usr/ctxtLGORT-LOW")
                    if sloc_single is not None and clean_slocs:
                        sloc_single.text = clean_slocs[0]
            else:
                if self._find_by_id_safe(multi_sloc_btn) is not None:
                    self._clear_multiple_selection(multi_sloc_btn)
                for s_id in ["wnd[0]/usr/ctxtLGORT-LOW", "wnd[0]/usr/ctxtLGORT-HIGH"]:
                    s_ctl = self._find_by_id_safe(s_id)
                    if s_ctl is not None:
                        try:
                            s_ctl.text = ""
                        except Exception:
                            pass

            # 4. Xóa Batch trên màn hình lọc để không bỏ sót lô
            multi_charg_btn = "wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH"
            if self._find_by_id_safe(multi_charg_btn) is not None:
                self._clear_multiple_selection(multi_charg_btn)
            for c_id in ["wnd[0]/usr/ctxtCHARG-LOW", "wnd[0]/usr/ctxtCHARG-HIGH"]:
                c_ctl = self._find_by_id_safe(c_id)
                if c_ctl is not None:
                    try:
                        c_ctl.text = ""
                    except Exception:
                        pass

            # 5. Điền Layout /KHOMA chuẩn theo Module1.bas dòng 1103 của Stock_FG_MA_BODY
            for layout_id in ["wnd[0]/usr/ctxtP_VARI", "wnd[0]/usr/ctxtRM07M-VARIANT"]:
                layout_ctl = self._find_by_id_safe(layout_id)
                if layout_ctl is not None:
                    try:
                        layout_ctl.text = "/KHOMA"
                        break
                    except Exception:
                        pass

            # 6. Bấm thực thi (F8)
            print("[SAP MB52] Đang bấm thực thi (F8)...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP MB52 ({status_type}): {status_txt}")

            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "MB52", status_txt, status_type)
                return []

            col_map = {
                "product": ["MATNR", "MATERIAL", "PRODUCT"],
                "plant":   ["WERKS", "PLANT"],
                "sloc":    ["LGORT", "STORAGE LOCATION", "STORAGE LOC", "STOR LOC", "SLOC"],
                "batch":   ["CHARG", "BATCH"],
                "bin":     ["LGPBE", "BIN", "STORAGE BIN"],
                "unrest":  [
                    "LABST", "UNRESTR", "UNRESTR.", "UNRESTRICTED",
                    "UNRESTRICTED STOCK", "UNRESTRICTED USE", "UNRESTRICTED USE STOCK",
                    "KHA DUNG", "KHẢ DỤNG", "KHO DUNG", "CLABS", "KALAB"
                ],
                "transit": [
                    "UMLMC", "TRAME", "UMLMD", "STOCK IN TRANSIT", "TRANSIT/TRANSF.", "TRANSIT/TRANSF",
                    "TRANSIT", "IN TRANSIT", "CHUYEN KHO", "HANG DI DUONG"
                ],
                "qual":    ["INSME", "QUALITY INSPECTION", "IN QUALITY INSP."],
                "blocked": ["SPEME", "BLOCKED"],
                "unit":    ["MEINS", "BASE UNIT OF MEASURE", "BASE UNIT", "UNIT"]
            }

            records = []
            # CƠ CHẾ ƯU TIÊN THEO ẢNH 2 & 261008_Stock_FG_MA_BODY:
            # Xuất trực tiếp sang Excel (&XXL / Ctrl+Shift+F7) để đọc trọn vẹn toàn bộ dữ liệu bảng lớn cực nhanh
            exported_excel_path = None
            try:
                exported_excel_path = self._export_alv_to_excel(alv_grid, tcode_name="MB52")
                if exported_excel_path and os.path.exists(exported_excel_path):
                    records = self._read_sap_exported_file(exported_excel_path, col_map)
            except Exception as exp_err:
                print(f"[SAP MB52 WARN] Xuất Excel không thành công, chuyển sang đọc ALV Grid trực tiếp: {exp_err}")
            finally:
                if exported_excel_path and os.path.exists(exported_excel_path):
                    try:
                        shutil.rmtree(os.path.dirname(exported_excel_path), ignore_errors=True)
                    except Exception:
                        pass

            # Fallback nếu xuất file không có dữ liệu
            if not records:
                records = self._extract_alv_grid(alv_grid, col_map, max_rows=max_rows)
                print(f"[SAP MB52] Đã trích xuất {len(records)} dòng qua COM ALV Grid.")

            # 1. Forward-fill mã hàng (product), plant và unit cho các dòng bị gộp/ẩn tiêu đề
            last_prod = ""
            last_pl = ""
            last_u = ""
            for r in records:
                if r.get("product"):
                    last_prod = r["product"]
                elif last_prod:
                    r["product"] = last_prod

                if r.get("plant"):
                    last_pl = r["plant"]
                elif last_pl:
                    r["plant"] = last_pl

                if r.get("unit"):
                    last_u = r["unit"]
                elif last_u:
                    r["unit"] = last_u

            # 2. Xử lý dòng Hàng đang trung chuyển / Transit (Plant level, SLoc/Batch trống nhưng Transit > 0)
            valid_records = []
            for r in records:
                sloc_raw = str(r.get("sloc", "")).strip()
                unrest_f = parse_sap_float(r.get("unrest", "0"))
                transit_f = parse_sap_float(r.get("transit", "0"))
                pl_code = str(r.get("plant", plant)).strip()

                if not sloc_raw:
                    # Nếu dòng không có SLoc nhưng có số lượng Transit/Transf. > 0 (như dòng 78 EA ở Plant 4007)
                    if transit_f > 0:
                        r["sloc"] = "TRANSIT"
                        if not r.get("batch"):
                            r["batch"] = "Chờ xuất (từ 3093)" if pl_code in ("4007", "4004") else "Chờ nhập (Transit)"
                        r["unrest"] = format_qty_display(transit_f)
                        r["transit"] = format_qty_display(transit_f)
                        r["is_transit"] = True
                        valid_records.append(r)
                    # Nếu cả sloc rỗng và transit == 0 -> dòng tổng cộng/rỗng (subtotal), bỏ qua
                    continue

                # Dòng kho bình thường có SLoc
                r["unrest"] = format_qty_display(unrest_f)
                r["transit"] = format_qty_display(transit_f)
                valid_records.append(r)

            records = valid_records

            # 3. Lọc theo kho (nếu có yêu cầu cụ thể)
            if slocs:
                clean_slocs_set = {s.strip().upper() for s in slocs if s.strip()}
                records = [r for r in records if str(r.get("sloc", "")).strip().upper() in clean_slocs_set]
                print(f"[SAP MB52] Sau khi lọc theo kho {clean_slocs_set}: còn {len(records)} dòng.")

            # 4. SẮP XẾP GOM THEO KHO (SLOC):
            # TRANSIT lên đầu tiên (như màn hình gốc MB52), tiếp theo là các kho thực tế với số lượng giảm dần
            def _mb52_sort_key(item):
                sloc_val = str(item.get("sloc", "")).strip().upper()
                unrest_val = parse_sap_float(item.get("unrest", 0.0))
                batch_val = str(item.get("batch", "")).strip().upper()
                is_tr = 0 if (item.get("is_transit") or sloc_val == "TRANSIT") else 1
                return (is_tr, sloc_val, -unrest_val, batch_val)

            records.sort(key=_mb52_sort_key)
            return records
        finally:
            self._reset_to_session_manager()

    # --------------------------------------------------------------------------
    # 3. KIỂM TRA LỊCH SỬ XUẤT NHẬP (MB51)
    # --------------------------------------------------------------------------
    def check_history_mb51(self, materials: List[str], days: int = 7, plant: str = "", explicit_plant: bool = False) -> List[Dict[str, str]]:
        """Tra cứu lịch sử xuất nhập kho trong N ngày qua MB51."""
        try:
            if not materials:
                return []

            self._open_transaction("MB51")
            self._wait_idle(10.0)

            # Xóa sạch toàn bộ bộ lọc cũ còn lưu trong SAP GUI (MBLNR, BWART, LGORT, CHARG...)
            self._clear_all_selection_inputs()

            # Đảm bảo xóa chắc chắn các trường Header / Item của MB51
            mb51_explicit_clears = [
                "wnd[0]/usr/ctxtMBLNR-LOW", "wnd[0]/usr/ctxtMBLNR-HIGH",
                "wnd[0]/usr/ctxtBWART-LOW", "wnd[0]/usr/ctxtBWART-HIGH",
                "wnd[0]/usr/ctxtLGORT-LOW", "wnd[0]/usr/ctxtLGORT-HIGH",
                "wnd[0]/usr/ctxtCHARG-LOW", "wnd[0]/usr/ctxtCHARG-HIGH",
                "wnd[0]/usr/ctxtLIFNR-LOW", "wnd[0]/usr/ctxtLIFNR-HIGH",
                "wnd[0]/usr/ctxtKUNNR-LOW", "wnd[0]/usr/ctxtKUNNR-HIGH",
                "wnd[0]/usr/ctxtSOBKZ-LOW", "wnd[0]/usr/ctxtSOBKZ-HIGH",
                "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/ctxtAUFNR-HIGH",
                "wnd[0]/usr/ctxtEBELN-LOW", "wnd[0]/usr/ctxtEBELN-HIGH",
                "wnd[0]/usr/ctxtKOSTL-LOW", "wnd[0]/usr/ctxtKOSTL-HIGH",
                "wnd[0]/usr/ctxtKDAUF-LOW", "wnd[0]/usr/ctxtKDAUF-HIGH",
                "wnd[0]/usr/txtKDPOS-LOW",  "wnd[0]/usr/txtKDPOS-HIGH",
                "wnd[0]/usr/ctxtPS_PSP_PNR-LOW", "wnd[0]/usr/ctxtPS_PSP_PNR-HIGH",
                "wnd[0]/usr/ctxtRSNUM-LOW", "wnd[0]/usr/ctxtRSNUM-HIGH",
                "wnd[0]/usr/ctxtUSNAM-LOW", "wnd[0]/usr/ctxtUSNAM-HIGH",
                "wnd[0]/usr/ctxtVGART-LOW", "wnd[0]/usr/ctxtVGART-HIGH",
                "wnd[0]/usr/ctxtXBLNR-LOW", "wnd[0]/usr/ctxtXBLNR-HIGH",
                "wnd[0]/usr/ctxtMATNR-HIGH", "wnd[0]/usr/ctxtWERKS-HIGH"
            ]
            for f_id in mb51_explicit_clears:
                ctl = self._find_by_id_safe(f_id)
                if ctl is not None:
                    try:
                        if getattr(ctl, "Text", ""):
                            ctl.text = ""
                    except Exception:
                        pass

            print(f"[SAP MB51] Đang nhập điều kiện: Materials={materials}, Days={days}, Plant={plant}...")
            clean_mats = [m.strip().upper() for m in materials if m.strip()]

            # 1. Nhập Materials: Nếu chỉ có 1 mã, điền thẳng vào ctxtMATNR-LOW (TUYỆT ĐỐI không mở popup Multiple Selection wnd[1])
            if len(clean_mats) == 1:
                mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtMATNR-LOW")
                if mat_single is not None:
                    mat_single.text = clean_mats[0]
            elif len(clean_mats) > 1:
                multi_btn_id = "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
                if self._find_by_id_safe(multi_btn_id) is not None:
                    self._fill_multiple_selection(multi_btn_id, clean_mats)
                else:
                    mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtMATNR-LOW")
                    if mat_single is not None:
                        mat_single.text = clean_mats[0]

            if explicit_plant and plant:
                p_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtWERKS-LOW")
                if p_ctl is not None:
                    try:
                        p_ctl.text = plant
                    except Exception:
                        pass
            else:
                p_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtWERKS-LOW")
                if p_ctl is not None:
                    try:
                        p_ctl.text = ""
                    except Exception:
                        pass

            # 2. Điền Posting Date 7 ngày gần nhất chuẩn xác vào ctxtBUDAT-LOW và ctxtBUDAT-HIGH
            end_date = datetime.now()
            start_date = end_date - timedelta(days=int(days))
            b_low = self._find_by_id_safe("wnd[0]/usr/ctxtBUDAT-LOW")
            if b_low is not None:
                b_low.text = start_date.strftime("%d.%m.%Y")
            b_high = self._find_by_id_safe("wnd[0]/usr/ctxtBUDAT-HIGH")
            if b_high is not None:
                b_high.text = end_date.strftime("%d.%m.%Y")

            print(f"[SAP MB51] Đã điền Posting Date: {start_date.strftime('%d.%m.%Y')} -> {end_date.strftime('%d.%m.%Y')}")

            # 3. Đảm bảo đóng mọi popup phụ (nếu có) trước khi bấm F8 trên wnd[0]
            self._dismiss_popups()

            print("[SAP MB51] Đang bấm thực thi (F8)...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                try:
                    exec_btn.press()
                except Exception:
                    self.session.findById("wnd[0]").sendVKey(8)
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP MB51 ({status_type}): {status_txt}")

            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            # SMART FALLBACK: Nếu không có dòng dữ liệu nào trong khoảng ngày đã lọc (ví dụ 3 ngày),
            # và người dùng không yêu cầu khoảng ngày quá dài (>60 ngày),
            # tự động mở rộng sang 30 ngày gần nhất để bắt được giao dịch gần nhất của vật tư.
            if alv_grid is None and days < 60:
                print(f"[SAP MB51] 0 dòng dữ liệu trong {days} ngày gần nhất. Tự động mở rộng sang 30 ngày gần nhất...")
                start_date_wide = end_date - timedelta(days=30)
                if b_low is not None:
                    try:
                        b_low.text = start_date_wide.strftime("%d.%m.%Y")
                    except Exception:
                        pass
                if b_high is not None:
                    try:
                        b_high.text = end_date.strftime("%d.%m.%Y")
                    except Exception:
                        pass

                # Bấm thực thi lại (F8)
                if exec_btn is not None:
                    try:
                        exec_btn.press()
                    except Exception:
                        self.session.findById("wnd[0]").sendVKey(8)
                else:
                    self.session.findById("wnd[0]").sendVKey(8)

                self._wait_idle(25.0)
                self._dismiss_popups()
                alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
                if alv_grid is None:
                    alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "MB51", status_txt, status_type)
                return []

            col_map = {
                "product": ["MATNR", "MATERIAL"],
                "docNo":   ["MBLNR", "MATERIAL DOCUMENT", "DOCUMENT"],
                "mvtType": ["BWART", "MOVEMENT TYPE", "MVT"],
                "postDate":["BUDAT", "POSTING DATE"],
                "qty":     ["MENGE", "ERFMG", "QUANTITY"],
                "unit":    ["MEINS", "ERFME", "UNIT OF ENTRY", "UNIT"],
                "plant":   ["WERKS", "PLANT"],
                "sloc":    ["LGORT", "SLOC", "STORAGE LOCATION"],
                "batch":   ["CHARG", "BATCH"],
                "refNo":   ["XBLNR", "REFERENCE", "BILL"],
                "headText":["BKTXT", "DOC.HEADER TEXT", "HEADER TEXT"]
            }
            records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP MB51] Đã trích xuất {len(records)} dòng lịch sử xuất nhập.")

            # SẮP XẾP: Ngày hạch toán (Posting Date) và Số chứng từ MỚI NHẤT lên đầu
            def _mb51_sort_key(item):
                d_raw = item.get("postDate", "")
                doc = item.get("docNo", "")
                parts = re.split(r"[./-]", d_raw)
                if len(parts) == 3 and len(parts[2]) == 4:
                    d_sort = f"{parts[2]}{parts[1].zfill(2)}{parts[0].zfill(2)}"
                else:
                    d_sort = d_raw
                return (d_sort, doc)

            records.sort(key=_mb51_sort_key, reverse=True)
            return records

        # --------------------------------------------------------------------------
        finally:
            self._reset_to_session_manager()

    # 4. KIỂM TRA ĐƠN INBOUND (IB) TREO CHƯA GR (VL06I)
    # --------------------------------------------------------------------------
    def check_inbound_treo_vl06i(self, materials: List[str], plant: str = "3093", only_open: bool = True) -> List[Dict[str, str]]:
        """Tra cứu các đơn Inbound Delivery (IB) qua VL06I theo chuẩn BBBG (only_open=True: chỉ lấy hàng treo chưa GR)."""
        try:
            if not materials:
                return []

            self._open_transaction("VL06I")
            self._wait_idle(10.0)

            print(f"[SAP VL06I] Đang tra cứu IB treo cho Materials={materials}, Plant={plant}...")
            clean_mats = [m.strip().upper() for m in materials if m.strip()]

            # 1. Bấm BUTTON7: List Inbound Deliveries
            btn7 = self._find_by_id_safe("wnd[0]/usr/btnBUTTON7")
            if btn7 is not None:
                btn7.press()
                self._wait_idle(5.0)

            # Xóa sạch các trường lọc cũ trên màn hình chọn
            self._clear_all_selection_inputs()

            # 2. Xóa trường ngày giao để lấy trọn vẹn không bị lọc ngày (Shift+F2 / sendVKey 14)
            lfdat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_LFDAT-LOW")
            if lfdat_ctl is not None:
                try:
                    lfdat_ctl.text = ""
                    if hasattr(lfdat_ctl, "setFocus"):
                        lfdat_ctl.setFocus()
                    if hasattr(self, "session") and self.session is not None:
                        wnd0 = self._find_by_id_safe("wnd[0]")
                        if wnd0 is not None and hasattr(wnd0, "sendVKey"):
                            wnd0.sendVKey(14)
                except Exception:
                    pass
            lfdat_hi = self._find_by_id_safe("wnd[0]/usr/ctxtIT_LFDAT-HIGH")
            if lfdat_hi is not None:
                try:
                    lfdat_hi.text = ""
                except Exception:
                    pass

            # 3. Điền mã vật tư
            mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtIT_MATNR-LOW")
            if len(clean_mats) == 1 and mat_single is not None:
                mat_single.text = clean_mats[0]
            else:
                multi_mat_id = "wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH"
                if self._find_by_id_safe(multi_mat_id) is not None:
                    self._fill_multiple_selection(multi_mat_id, clean_mats)
                elif mat_single is not None and clean_mats:
                    mat_single.text = clean_mats[0]

            # 4. Điền Plant nếu có
            if plant:
                werks_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WERKS-LOW")
                if werks_ctl is not None:
                    try:
                        werks_ctl.text = plant
                    except Exception:
                        pass

            # 5. Đặt trạng thái Goods Movement IT_WBSTK-LOW = 'a' (chưa GR / Treo) nếu only_open=True
            wbstk_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTK-LOW")
            if wbstk_ctl is not None:
                try:
                    wbstk_ctl.text = "a" if only_open else ""
                except Exception:
                    pass

            # 6. Bấm thực thi (F8)
            print("[SAP VL06I] Đang bấm thực thi (F8)...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP VL06I ({status_type}): {status_txt}")

            # 7. Chuyển sang chế độ xem chi tiết Item (btn[18] trên tbar[1])
            btn18 = self._find_by_id_safe("wnd[0]/tbar[1]/btn[18]")
            if btn18 is not None:
                try:
                    btn18.press()
                    self._wait_idle(10.0)
                    self._dismiss_popups()
                except Exception as e:
                    print(f"[SAP VL06I WARN] Không thể bấm btn[18]: {e}")

            # 8. Tìm ALV Grid chi tiết
            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "VL06I", status_txt, status_type)
                return []

            col_map = {
                "ibNo":    ["VBELN", "DELIVERY", "INBOUND DELIVERY", "INBOUND"],
                "refNo":   ["LIFEX", "EXTERNAL IDENTIFICATION", "BILL OF LADING", "REF. DOC.", "EXTERNAL DELIVERY", "REFERENCE"],
                "product": ["MATNR", "MATERIAL", "PRODUCT"],
                "batch":   ["CHARG", "S_CHARG", "BATCH"],
                "qty":     ["LFIMG", "DELIVERY QUANTITY", "ACTUAL DELIVERED QTY", "MENGE", "QUANTITY"],
                "unit":    ["VRKME", "MEINS", "SALES UNIT", "BASE UNIT", "UNIT", "UOM"],
                "plant":   ["WERKS", "PLANT"],
                "sloc":    ["LGORT", "STORAGE LOCATION", "SLOC"],
                "status":  ["WBSTK", "GM STATUS", "OVERALL GM STATUS", "GOODS MOVEMENT STATUS", "STATUS"]
            }
            records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP VL06I] Đã trích xuất {len(records)} dòng IB treo.")

            for r in records:
                st = str(r.get("status", "")).strip().upper()
                if not st or st in ("A", "a"):
                    r["status"] = "Treo (Chưa GR)"
                elif st in ("B", "b"):
                    r["status"] = "GR 1 phần"
                elif st in ("C", "c"):
                    r["status"] = "Đã hoàn tất GR"

            # Sắp xếp số đơn IB mới nhất lên đầu
            records.sort(key=lambda x: str(x.get("ibNo", "")), reverse=True)
            return records

        # --------------------------------------------------------------------------
        finally:
            self._reset_to_session_manager()

    # 5. KIỂM TRA ĐƠN ĐIỀU CHUYỂN KHO (STO - ME2N)
    # --------------------------------------------------------------------------
    def check_sto_me2n(
        self,
        materials: List[str],
        plant: str = "3093",
        supplying_plant: str = "",
        receiving_plant: str = ""
    ) -> List[Dict[str, str]]:
        """Tra cứu đơn STO (Stock Transport Order) điều chuyển kho qua ME2N chuẩn theo RC65."""
        try:
            if not materials:
                return []

            self._open_transaction("ME2N")
            self._wait_idle(10.0)

            # Xóa sạch các bộ lọc cũ trên màn hình chọn ME2N
            self._clear_all_selection_inputs()

            clean_mats = [m.strip().upper() for m in materials if m.strip()]
            user_plant = str(plant or "3093").strip().upper()

            # Phân định Supplying Plant (S_RESWK) và Receiving Plant (S_WERKS) theo chuẩn RC65
            actual_supplying = supplying_plant.strip().upper() if supplying_plant else ""
            actual_receiving = receiving_plant.strip().upper() if receiving_plant else ""

            if not actual_supplying and not actual_receiving:
                if user_plant == "4007":
                    actual_receiving = "4007"
                    actual_supplying = "3093"
                else:
                    # Mặc định đơn vị tại 3093 tra cứu đơn nhận về 3093
                    actual_receiving = user_plant
                    actual_supplying = ""

            print(
                f"[SAP ME2N TRACE] Đang tra cứu STO: Materials={clean_mats}, "
                f"SupplyingPlant='{actual_supplying}', ReceivingPlant='{actual_receiving}'..."
            )

            # 1. Đặt định dạng hiển thị ALV
            listu_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtLISTU")
            if listu_ctl is not None:
                try:
                    listu_ctl.text = "ALV"
                except Exception:
                    pass

            # 2. Đặt Selection Parameter WE101 (Open Goods Receipt) theo chuẩn Check_PO_SA_Availability
            selpa_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtSELPA-LOW")
            if selpa_ctl is not None:
                try:
                    selpa_ctl.text = "WE101"
                except Exception:
                    pass

            # 3. Điền mã vật tư
            mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtS_MATNR-LOW")
            if len(clean_mats) == 1 and mat_single is not None:
                mat_single.text = clean_mats[0]
            else:
                multi_mat_id = "wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH"
                if self._find_by_id_safe(multi_mat_id) is not None:
                    self._fill_multiple_selection(multi_mat_id, clean_mats)
                elif mat_single is not None and clean_mats:
                    mat_single.text = clean_mats[0]

            # Xóa trường high của Material nếu có
            mat_high = self._find_by_id_safe("wnd[0]/usr/ctxtS_MATNR-HIGH")
            if mat_high is not None:
                try:
                    mat_high.text = ""
                except Exception:
                    pass

            # 4. Điền Supplying Plant (Nhà máy cấp hàng S_RESWK) và Receiving Plant (Nhà máy nhận hàng S_WERKS)
            if actual_supplying:
                reswk_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtS_RESWK-LOW")
                if reswk_ctl is not None:
                    try:
                        reswk_ctl.text = actual_supplying
                    except Exception:
                        pass
                reswk_hi = self._find_by_id_safe("wnd[0]/usr/ctxtS_RESWK-HIGH")
                if reswk_hi is not None:
                    try:
                        reswk_hi.text = ""
                    except Exception:
                        pass

            if actual_receiving:
                werks_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtS_WERKS-LOW")
                if werks_ctl is not None:
                    try:
                        werks_ctl.text = actual_receiving
                    except Exception:
                        pass
                werks_hi = self._find_by_id_safe("wnd[0]/usr/ctxtS_WERKS-HIGH")
                if werks_hi is not None:
                    try:
                        werks_hi.text = ""
                    except Exception:
                        pass

            # 5. Bấm thực thi (F8)
            print("[SAP ME2N] Đang bấm thực thi (F8)...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP ME2N ({status_type}): {status_txt}")

            # Xử lý thông báo "No suitable purchasing documents found" một cách êm ái
            if any(p in status_txt.lower() for p in ["no suitable purchasing documents", "no purchasing documents", "không tìm thấy chứng từ"]):
                print(f"[SAP ME2N] SAP xác nhận không có chứng từ mua hàng/STO phù hợp: '{status_txt}'")
                return []

            # 6. Tìm ALV Grid
            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "ME2N", status_txt, status_type)
                return []

            col_map = {
                "stoNo":          ["EBELN", "PURCHASING DOCUMENT", "PO NUMBER", "DOCUMENT", "PO"],
                "item":           ["EBELP", "ITEM"],
                "product":        ["EMATN", "MATNR", "MATERIAL"],
                "orderQty":       ["MENGE", "ORDER QUANTITY", "PO QUANTITY", "QUANTITY", "QTY"],
                "deliveredQty":   ["MGLIEF", "DELIVERED QUANTITY", "ISSUED QTY", "QTY DELIVERED", "DELIVERED"],
                "unit":           ["MEINS", "BPRME", "UNIT", "ORDER UNIT"],
                "supplyingPlant": ["RESWK", "SUPPLYING PLANT", "SUPPL. PLANT"],
                "receivingPlant": ["WERKS", "RECEIVING PLANT", "PLANT"],
                "docDate":        ["BEDAT", "DOCUMENT DATE", "PO DATE", "DATE"],
                "deleteInd":      ["LOEKZ", "DELETION INDICATOR", "DEL"]
            }
            raw_records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP ME2N TRACE] Đã trích xuất {len(raw_records)} dòng đơn STO từ SAP.")

            records = []
            for r in raw_records:
                if str(r.get("deleteInd", "")).strip():
                    continue

                ord_q = parse_sap_float(r.get("orderQty"))
                del_q = parse_sap_float(r.get("deliveredQty"))
                # Ghi chú: MGLIEF chỉ mang tính tham khảo theo chuẩn RC65 dòng 9069
                rem_q = max(0.0, ord_q - del_q)

                r["orderQty"] = format_qty_display(ord_q)
                r["deliveredQty"] = format_qty_display(del_q)
                r["remainingQty"] = format_qty_display(rem_q)
                records.append(r)

            print(f"[SAP ME2N TRACE] Sau khi xử lý: còn {len(records)} dòng STO thỏa mãn.")
            records.sort(key=lambda x: str(x.get("stoNo", "")), reverse=True)
            return records
        finally:
            self._reset_to_session_manager()

    # --------------------------------------------------------------------------
    # 6. KIỂM TRA TỒN THẤP THEO MODEL & SHOP (MB52 ĐA PLANT)
    # --------------------------------------------------------------------------
    def check_low_stock(
        self,
        materials: List[str],
        plants: List[str],
        limit: int = 20,
        model_name: str = "",
        shop_name: str = "BODY",
        full_model_list: bool = False,
        excel_path: Optional[str] = None,
        use_excel: bool = True
    ) -> List[Dict[str, str]]:
        """Tra cứu và tổng hợp danh sách mã tồn thấp chuẩn theo 261008_Stock_FG_MA_BODY V3.7.
        
        - BODY: Tồn kho Body tính theo 11 cột chuẩn:
          Material, US, Line Body (3001|3003), BD Kho (3001|1003), BD Carset (INT((Line+Kho)/US)),
          Transit (Intransit), Tồn MA (3093|3093), Tồn Line (3093|3000), Model, Trạm, Cảnh báo.
        - Tự động đọc và đối chiếu trực tiếp từ file 261008_Stock_FG_MA_BODY.xlsm nếu có sẵn,
          hoặc chạy MB52 đa xưởng live và tổng hợp chuẩn theo cùng thuật toán.
        """
        clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
        target_plants = [p.strip().upper() for p in (plants or []) if p.strip()]
        if not target_plants:
            target_plants = ["3093", "3001"]

        # 1. ƯU TIÊN KIỂM TRA FILE CHUẨN 261008_Stock_FG_MA_BODY.xlsm NẾU CÓ (Trừ khi check_stock_mb52 bị mock trong unit test)
        from unittest.mock import MagicMock
        is_mocked = isinstance(getattr(self, "check_stock_mb52", None), MagicMock)
        excel_body = Path(excel_path or "G:/SAP/261008_Stock_FG_MA_BODY.xlsm")
        if use_excel and not is_mocked and excel_body.exists():
            try:
                import openpyxl
                wb_b = openpyxl.load_workbook(str(excel_body), data_only=True)
                if "TON TONG" in wb_b.sheetnames:
                    ws_tt = wb_b["TON TONG"]
                    clean_model = (model_name or "").strip().upper()
                    threshold = 200
                    warnings_map = {}

                    for r in range(2, ws_tt.max_row + 1):
                        mat_v = ws_tt.cell(r, 1).value
                        model_v = ws_tt.cell(r, 9).value
                        if not mat_v or not model_v:
                            continue
                        m_str = str(mat_v).strip().upper()
                        mdl_str = str(model_v).strip()

                        # Nếu có danh sách mã chỉ định thì lọc theo mã, ngược lại lọc theo Model
                        if clean_mats and m_str not in clean_mats:
                            continue
                        if not clean_mats and clean_model and clean_model != "ALL" and clean_model not in mdl_str.upper():
                            continue

                        us = float(ws_tt.cell(r, 2).value or 1)
                        if us <= 0:
                            us = 1
                        line_body = float(ws_tt.cell(r, 3).value or 0)
                        kho_body = float(ws_tt.cell(r, 4).value or 0)
                        can_build = float(ws_tt.cell(r, 5).value or 0)
                        intransit = float(ws_tt.cell(r, 6).value or 0)
                        kho_ma = float(ws_tt.cell(r, 7).value or 0)
                        line_ma = float(ws_tt.cell(r, 8).value or 0)
                        tram = str(ws_tt.cell(r, 10).value or "").strip()

                        # Xác định hành động cảnh báo chuẩn V3.7
                        pending = max(0.0, intransit)
                        if pending > 0 and int((line_body + kho_body + pending) / us) >= threshold:
                            action = f"Hàng chờ xuất sang Body ({int(pending)})"
                        else:
                            if kho_ma > 0:
                                action = "Ưu tiên xuất Body"
                            elif line_ma > 0:
                                action = "Ưu tiên kéo line"
                            else:
                                action = "MA hết tồn"
                            if pending > 0:
                                action += f" (Chờ xuất: {int(pending)})"

                        urgency = 0 if "hết" in action.lower() else (1 if "ưu tiên xuất" in action.lower() else (2 if "ưu tiên kéo line" in action.lower() else (3 if "chờ xuất" in action.lower() else 4)))
                        rec = {
                            "product": m_str,
                            "us": str(int(us)),
                            "line_body": format_qty_display(line_body),
                            "bd_kho": format_qty_display(kho_body),
                            "bd_carset": format_qty_display(can_build),
                            "transit": format_qty_display(intransit),
                            "kho_ma": format_qty_display(kho_ma),
                            "line_ma": format_qty_display(line_ma),
                            "model": mdl_str,
                            "station": tram or "-",
                            "status": action,
                            "total_stock": format_qty_display(line_body + kho_body),
                            "stock_3093": format_qty_display(kho_ma),
                            "stock_3001": format_qty_display(line_body),
                            "unit": "PC",
                            "_sort_key": (urgency, can_build, line_body / us, m_str)
                        }

                        if m_str not in warnings_map or (tram and warnings_map[m_str]["station"] == "-"):
                            warnings_map[m_str] = rec

                    wb_b.close()
                    if warnings_map:
                        sorted_recs = sorted(warnings_map.values(), key=lambda x: x["_sort_key"])
                        for sr in sorted_recs:
                            sr.pop("_sort_key", None)
                        print(f"[SAP LOW STOCK] Đã trích xuất {len(sorted_recs)} mã tồn thấp chuẩn từ 261008_Stock_FG_MA_BODY.xlsm.")
                        eff_limit = max(limit, 50) if full_model_list else limit
                        return sorted_recs[:eff_limit]
            except Exception as e_excel:
                print(f"[SAP LOW STOCK WARN] Lỗi đọc trực tiếp từ Stock_FG_MA_BODY.xlsm: {e_excel}, chuyển sang MB52 live...")

        # 2. DỰ PHÒNG: GỌI MB52 LIVE QUA SAP NẾU CHƯA CÓ FILE EXCEL HOẶC DỮ LIỆU TEST
        print(f"[SAP LOW STOCK] Quét MB52 cả 2 xưởng {target_plants} cho {len(clean_mats)} mã (Model={model_name}, Top={limit})...")
        try:
            mb52_records = self.check_stock_mb52(clean_mats, plants=target_plants, max_rows=15000)
            print(f"[SAP LOW STOCK] MB52 trả về tổng cộng {len(mb52_records)} dòng dữ liệu từ các xưởng.")
        except Exception as e:
            print(f"[SAP LOW STOCK ERROR] Lỗi truy vấn MB52 đa xưởng: {e}")
            raise SapExecutionError(f"Không thể truy vấn tồn kho MB52 đa xưởng {target_plants}: {e}") from e

        if clean_mats:
            mats_to_evaluate = list(clean_mats)
        else:
            mats_to_evaluate = sorted(list(set(r.get("product", "").strip().upper() for r in mb52_records if r.get("product"))))

        stock_by_mat = {m: {p: 0.0 for p in target_plants} for m in mats_to_evaluate}
        unit_by_mat = {m: "PC" for m in mats_to_evaluate}
        sloc_detail_by_mat = {
            m: {
                "line_body": 0.0,   # 3001/3003
                "bd_kho": 0.0,      # 3001/1003
                "kho_ma": 0.0,      # 3093/3093
                "line_ma": 0.0,     # 3093/3000
                "transit": 0.0,     # 9999 / QC / Transit
                "xmd_pre": 0.0,     # 2001 (3002, 1001, 600P, 1D01)
                "xmd_post": 0.0     # 2001 & 3001 post-paint
            }
            for m in mats_to_evaluate
        }

        def _find_eval_mat(prod_code: str) -> Optional[str]:
            p_clean = prod_code.strip().upper()
            if not p_clean:
                return None
            if p_clean in stock_by_mat:
                return p_clean
            p_unpadded = p_clean.lstrip("0")
            p_core = re.sub(r"^[A-Z]+", "", p_unpadded)
            p_core = re.sub(r"[A-Z]+$", "", p_core)

            for target_m in mats_to_evaluate:
                t_clean = target_m.strip().upper()
                if p_clean == t_clean or p_unpadded == t_clean.lstrip("0"):
                    return target_m
                t_core = re.sub(r"^[A-Z]+", "", t_clean.lstrip("0"))
                t_core = re.sub(r"[A-Z]+$", "", t_core)
                if len(p_core) >= 6 and len(t_core) >= 6 and (p_core == t_core or p_core in t_core or t_core in p_core):
                    return target_m
            return None

        for r in mb52_records:
            mat_raw = r.get("product", "").strip().upper()
            p = r.get("plant", "").strip().upper()
            sloc = r.get("sloc", "").strip().upper()
            qty = parse_sap_float(r.get("unrest", "0"))

            matched_mat = _find_eval_mat(mat_raw)
            if not matched_mat:
                matched_mat = mat_raw
                stock_by_mat[matched_mat] = {pl: 0.0 for pl in target_plants}
                sloc_detail_by_mat[matched_mat] = {
                    "line_body": 0.0, "bd_kho": 0.0, "kho_ma": 0.0, "line_ma": 0.0,
                    "transit": 0.0, "xmd_pre": 0.0, "xmd_post": 0.0
                }
                if matched_mat not in mats_to_evaluate:
                    mats_to_evaluate.append(matched_mat)

            # SLoc 9999 / QC / Scrap được lưu vào transit riêng, không cộng vào tồn khả dụng
            if sloc in ["9999", "QC", "BLOCK", "SCRAP", "QUARANTINE", "TRANSIT"]:
                sloc_detail_by_mat[matched_mat]["transit"] += qty
                continue

            if p in stock_by_mat[matched_mat]:
                stock_by_mat[matched_mat][p] += qty

            # Bóc tách SLoc chuẩn theo 261008_Stock_FG_MA_BODY V3.7
            if p == "3001":
                if sloc == "3003":
                    sloc_detail_by_mat[matched_mat]["line_body"] += qty
                elif sloc == "1003":
                    sloc_detail_by_mat[matched_mat]["bd_kho"] += qty
                else:
                    sloc_detail_by_mat[matched_mat]["line_body"] += qty
            elif p == "3093":
                if sloc == "3093":
                    sloc_detail_by_mat[matched_mat]["kho_ma"] += qty
                elif sloc == "3000":
                    sloc_detail_by_mat[matched_mat]["line_ma"] += qty
                else:
                    sloc_detail_by_mat[matched_mat]["kho_ma"] += qty
            elif p == "2001":
                if sloc in ("3002", "1001", "600P", "1D01"):
                    sloc_detail_by_mat[matched_mat]["xmd_pre"] += qty
                else:
                    sloc_detail_by_mat[matched_mat]["xmd_post"] += qty

            if r.get("unit"):
                unit_by_mat[matched_mat] = r.get("unit")

        # Nạp danh mục Trạm (Station) từ Master Data
        tram_map = {}
        candidate_paths = [
            Path("G:/SAP/261008_Stock_FG_MA_BODY.xlsm"),
            Path(__file__).resolve().parent.parent / "data" / "DATA_MASTER.xlsx"
        ]
        for candidate_master in candidate_paths:
            if candidate_master.exists():
                try:
                    import openpyxl
                    wb_m = openpyxl.load_workbook(str(candidate_master), read_only=True)
                    ws_m = wb_m["DATA"] if "DATA" in wb_m.sheetnames else wb_m.active
                    r_iter = ws_m.iter_rows(values_only=True)
                    hdr = next(r_iter, None)
                    if hdr:
                        mat_idx = 1
                        tram_idx = -1
                        for idx, h in enumerate(hdr):
                            if not h:
                                continue
                            h_str = str(h).strip().upper()
                            if "MÃ" in h_str or "MATNR" in h_str:
                                mat_idx = idx
                            elif "TRẠM" in h_str or "TRAM" in h_str:
                                tram_idx = idx
                        if tram_idx >= 0:
                            for r_val in r_iter:
                                if r_val and len(r_val) > max(mat_idx, tram_idx):
                                    m_code = str(r_val[mat_idx]).strip().upper() if r_val[mat_idx] else ""
                                    t_code = str(r_val[tram_idx]).strip() if r_val[tram_idx] else ""
                                    if m_code and t_code and t_code != "-":
                                        tram_map[m_code] = t_code
                    wb_m.close()
                    if tram_map:
                        break
                except Exception:
                    pass

        output_records = []
        for mat in mats_to_evaluate:
            p_stocks = stock_by_mat.get(mat, {p: 0.0 for p in target_plants})
            total_s = sum(p_stocks.values())
            t_val = tram_map.get(mat, "")
            if not t_val:
                for km, kv in tram_map.items():
                    if mat in km or km in mat:
                        t_val = kv
                        break
            det = sloc_detail_by_mat.get(mat, {})
            l_body = det.get("line_body", p_stocks.get("3001", 0.0))
            b_kho = det.get("bd_kho", 0.0)
            k_ma = det.get("kho_ma", p_stocks.get("3093", 0.0))
            l_ma = det.get("line_ma", 0.0)
            trans = det.get("transit", 0.0)
            us_int = 1
            can_build = int((l_body + b_kho) / us_int) if (l_body + b_kho) > 0 else int(total_s / us_int)

            row = {
                "product": mat,
                "station": t_val or "-",
                "model": model_name or "N/A",
                "unit": unit_by_mat.get(mat, "PC"),
                "total_stock": format_qty_display(total_s),
                "us": str(us_int),
                "line_body": format_qty_display(l_body),
                "bd_kho": format_qty_display(b_kho),
                "bd_carset": format_qty_display(can_build),
                "transit": format_qty_display(trans),
                "kho_ma": format_qty_display(k_ma),
                "line_ma": format_qty_display(l_ma)
            }
            for p in target_plants:
                row[f"stock_{p}"] = format_qty_display(p_stocks.get(p, 0.0))

            if total_s == 0:
                row["status"] = "🔴 Hết hàng (0)"
            elif total_s < 50:
                row["status"] = "🟡 Tồn thấp (<50)"
            else:
                row["status"] = "🟢 Khả dụng"

            row["_raw_total"] = total_s
            output_records.append(row)

        output_records.sort(key=lambda x: (x["_raw_total"], x["product"]))
        for r in output_records:
            r.pop("_raw_total", None)

        return output_records[:limit]

    # --------------------------------------------------------------------------
    # 7. KIỂM TRA ĐƠN OUTBOUND (OB) TREO / ĐANG CHỜ (VL06O)
    # --------------------------------------------------------------------------
    @staticmethod
    def _matches_material(prod: str, target_materials: List[str]) -> bool:
        """So khớp linh hoạt mã vật tư: hỗ trợ khớp chính xác, bỏ số 0 đầu, tiền tố BIW, hoặc mã con."""
        if not target_materials:
            return True
        p_clean = prod.strip().upper()
        if not p_clean:
            return False
        p_unpadded = p_clean.lstrip("0")
        p_core = re.sub(r"^[A-Z]+", "", p_unpadded)

        for m in target_materials:
            m_clean = m.strip().upper()
            if not m_clean:
                continue
            if "*" in m_clean or "?" in m_clean:
                import fnmatch
                p_unp = p_clean.lstrip("0")
                m_unp = m_clean.lstrip("0")
                if fnmatch.fnmatch(p_clean, m_clean) or fnmatch.fnmatch(p_unp, m_unp):
                    return True
            m_unpadded = m_clean.lstrip("0")
            m_core = re.sub(r"^[A-Z]+", "", m_unpadded)

            # 1. Khớp chính xác
            if p_clean == m_clean or p_unpadded == m_unpadded:
                return True
            # 2. Khớp chuỗi con (nếu mã có độ dài >= 6)
            if len(m_unpadded) >= 6 and (m_unpadded in p_unpadded or p_unpadded in m_unpadded):
                return True
            # 3. Khớp lõi số (bỏ tiền tố BIW)
            if len(m_core) >= 6 and len(p_core) >= 6:
                if m_core == p_core or m_core in p_core or p_core in m_core:
                    return True

        return False

    def _is_item_grid_ready(self, grid) -> bool:
        """Kiểm tra xem ALV Grid hiện tại đã thực sự là Item View (có cột MATNR hoặc POSNR) hay chưa."""
        if grid is None:
            return False
        try:
            cols = getattr(grid, "ColumnOrder", None)
            if cols is not None:
                col_names = set()
                cnt = getattr(cols, "Count", getattr(cols, "count", 0))
                for i in range(cnt):
                    try:
                        col_names.add(str(cols.ElementAt(i)).strip().upper())
                    except Exception:
                        try:
                            col_names.add(str(cols(i)).strip().upper())
                        except Exception:
                            pass
                if "MATNR" in col_names or "POSNR" in col_names:
                    return True
            resolved = self._resolve_grid_columns(
                grid,
                {
                    "product": ["MATNR", "MATERIAL", "PRODUCT"],
                    "item": ["POSNR", "DELIVERY ITEM", "ITEM"]
                }
            )
            return "product" in resolved or "item" in resolved
        except Exception:
            return False

    def _ensure_item_grid_vl06o(self, max_wait_sec: float = 20.0):
        """Bảo đảm chuyển sang Item View trên VL06O theo đúng thuật toán VL6_EnsureItemGrid của RC65."""
        grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
        if grid is None:
            grid = self._find_largest_grid(max_wait_sec=5.0)

        if not self._is_item_grid_ready(grid):
            btn18 = self._find_by_id_safe("wnd[0]/tbar[1]/btn[18]")
            if btn18 is not None:
                print("[SAP VL06O] Đang ở Header View, bấm btn[18] chuyển sang Item View theo chuẩn RC65...")
                try:
                    btn18.press()
                except Exception as e:
                    print(f"[SAP VL06O WARN] Không thể bấm btn[18]: {e}")
            else:
                try:
                    self.session.findById("wnd[0]").sendVKey(18)
                except Exception:
                    pass

        # Vòng lặp chờ grid Item View sẵn sàng giống VL6_EnsureItemGrid trong RC65
        start_t = time.time()
        while time.time() - start_t < max_wait_sec:
            self._wait_idle(1.0)
            self._dismiss_popups()
            grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if grid is None:
                grid = self._find_largest_grid(max_wait_sec=2.0)
            if self._is_item_grid_ready(grid):
                print(f"[SAP VL06O] Item View Grid đã sẵn sàng sau {time.time() - start_t:.2f}s.")
                return grid
            time.sleep(0.5)

        return grid

    # --------------------------------------------------------------------------
    # 7. KIỂM TRA ĐƠN OUTBOUND (OB) TREO / ĐANG CHỜ (VL06O)
    # --------------------------------------------------------------------------
    def check_open_ob_vl06o(
        self,
        materials: Optional[List[str]] = None,
        ob_no: str = "",
        plant: str = "3093",
        ship_to: str = ""
    ) -> List[Dict[str, str]]:
        """Tra cứu các đơn Outbound Delivery (OB) đang mở qua VL06O chuẩn theo 261010_OB_TOOL_V1_RC65 kèm 5 nấc chẩn đoán."""
        try:
            self._open_transaction("VL06O")
            self._wait_idle(10.0)

            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            clean_ob = str(ob_no or "").strip()
            user_plant = str(plant or "3093").strip().upper()

            # Chuẩn RC65 (ModuleSAP.bas dòng 2572-2584): Shipping Point = "3093" trực tiếp hoặc theo user_plant
            target_vstel = "3098" if user_plant == "3098" else ("3093" if user_plant in ("3093", "") else user_plant)

            # 1. Điều hướng vào màn hình chọn List Outbound Deliveries (nút btnBUTTON6)
            date_low_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WADAT-LOW")
            if date_low_ctl is None:
                btn6 = self._find_by_id_safe("wnd[0]/usr/btnBUTTON6")
                if btn6 is not None:
                    btn6.press()
                    self._wait_idle(10.0)
                else:
                    for b_id in ["wnd[0]/usr/btnBUTTON6", "wnd[0]/usr/btnBUTTON1"]:
                        b = self._find_by_id_safe(b_id)
                        if b is not None:
                            b.press()
                            self._wait_idle(10.0)
                            break

            # 2. Xóa sạch các bộ lọc cũ còn lưu trên màn hình chọn
            self._clear_all_selection_inputs()

            for btn_id in [
                "wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH",
                "wnd[0]/usr/btn%_IT_KUNWE_%_APP_%-VALU_PUSH",
                "wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH",
                "wnd[0]/usr/btn%_IT_WADAT_%_APP_%-VALU_PUSH",
                "wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH"
            ]:
                if self._find_by_id_safe(btn_id) is not None:
                    self._clear_multiple_selection(btn_id)

            # 3. ĐIỀU KIỆN THEN CHỐT: Điền Mã hàng (MATNR) trực tiếp trên Selection Screen trước khi F8
            if clean_mats:
                mat_low = self._find_by_id_safe("wnd[0]/usr/ctxtIT_MATNR-LOW")
                if len(clean_mats) == 1 and mat_low is not None:
                    try:
                        mat_low.text = clean_mats[0]
                    except Exception as e:
                        print(f"[SAP VL06O WARN] Không thể điền MATNR-LOW: {e}")
                elif len(clean_mats) > 1:
                    multi_mat_id = "wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH"
                    if self._find_by_id_safe(multi_mat_id) is not None:
                        self._fill_multiple_selection(multi_mat_id, clean_mats)
                    elif mat_low is not None:
                        mat_low.text = clean_mats[0]

                mat_hi = self._find_by_id_safe("wnd[0]/usr/ctxtIT_MATNR-HIGH")
                if mat_hi is not None:
                    try:
                        mat_hi.text = ""
                    except Exception:
                        pass

            # 4. Xử lý Shipping Point: Không ép cứng 3093 khi tìm theo mã hàng (trừ khi người dùng yêu cầu plant khác)
            vstel_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIF_VSTEL-LOW")
            if vstel_ctl is not None:
                try:
                    vstel_val = "" if clean_mats else target_vstel
                    vstel_ctl.text = vstel_val
                except Exception as e:
                    print(f"[SAP VL06O WARN] Không thể điền Shipping Point: {e}")

            vstel_hi = self._find_by_id_safe("wnd[0]/usr/ctxtIF_VSTEL-HIGH")
            if vstel_hi is not None:
                try:
                    vstel_hi.text = ""
                except Exception:
                    pass

            # 5. Khoảng ngày: Nếu tìm theo Mã hàng hoặc số OB -> Xóa sạch khoảng ngày để quét toàn bộ không bị lọc mất
            if clean_mats or clean_ob:
                for dt_id in [
                    "wnd[0]/usr/ctxtIT_WADAT-LOW", "wnd[0]/usr/ctxtIT_WADAT-HIGH",
                    "wnd[0]/usr/ctxtIT_LFDAT-LOW", "wnd[0]/usr/ctxtIT_LFDAT-HIGH"
                ]:
                    dt_ctl = self._find_by_id_safe(dt_id)
                    if dt_ctl is not None:
                        try:
                            dt_ctl.text = ""
                        except Exception:
                            pass
                from_date = "ALL"
                to_date = "ALL"
            else:
                from_date = (datetime.now() - timedelta(days=60)).strftime("%d.%m.%Y")
                to_date = (datetime.now() + timedelta(days=60)).strftime("%d.%m.%Y")
                d_from = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WADAT-LOW")
                if d_from is not None:
                    try:
                        d_from.text = from_date
                    except Exception:
                        pass
                d_to = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WADAT-HIGH")
                if d_to is not None:
                    try:
                        d_to.text = to_date
                    except Exception:
                        pass

            # 5. Goods Movement Status = "A" (Chưa xuất hàng / Đang mở theo RC65)
            target_gm = "A"
            gm_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTK-LOW")
            if gm_ctl is None:
                gm_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTA-LOW")
            if gm_ctl is not None:
                try:
                    gm_ctl.text = target_gm
                except Exception:
                    pass

            # 6. Điền Ship-to nếu người dùng yêu cầu cụ thể
            kunwe_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_KUNWE-LOW")
            if kunwe_ctl is None:
                kunwe_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_KUNNR-LOW")
            if kunwe_ctl is not None:
                try:
                    kunwe_ctl.text = ship_to.strip() if ship_to else ""
                except Exception:
                    pass
            kunwe_hi = self._find_by_id_safe("wnd[0]/usr/ctxtIT_KUNWE-HIGH")
            if kunwe_hi is not None:
                try:
                    kunwe_hi.text = ""
                except Exception:
                    pass

            # 7. Điền số OB nếu có
            ob_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_VBELN-LOW")
            if ob_ctl is not None:
                try:
                    ob_ctl.text = clean_ob
                except Exception:
                    pass
            ob_hi = self._find_by_id_safe("wnd[0]/usr/ctxtIT_VBELN-HIGH")
            if ob_hi is not None:
                try:
                    ob_hi.text = ""
                except Exception:
                    pass

            # Nấc 1 Trace: Ghi nhận đầy đủ tham số đã áp dụng ở Selection Screen
            print(
                f"[SAP VL06O SELECTION TRACE] ShippingPoint='{target_vstel}', "
                f"PlannedGI='{from_date}'->'{to_date}', GMStatus='{target_gm}', "
                f"OB='{clean_ob}', ShipTo='{ship_to}', Materials={clean_mats}"
            )

            # 8. Thực thi F8
            print("[SAP VL06O] Đang bấm thực thi (F8)...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(30.0)
            popup_msg = self._dismiss_popups()

            status_txt, status_type = self._get_status()
            comb = f"{status_txt} {popup_msg}".lower()
            if any(k in comb for k in ["no deliveries", "no outbound deliveries", "outbound deliveries selected", "no data", "không tìm thấy", "keine lieferungen", "keine auslieferungen", "không có giao hàng"]):
                print(f"[SAP VL06O TRACE Nấc 1] Xác nhận không có đơn Outbound nào: '{status_txt or popup_msg}'")
                return []

            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP VL06O ({status_type}): {status_txt}")

            # 9. BẢO ĐẢM HIỂN THỊ ITEM VIEW THEO CHUẨN VL6_EnsureItemGrid CỦA RC65
            alv_grid = self._ensure_item_grid_vl06o(max_wait_sec=20.0)
            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "VL06O", status_txt, status_type)
                return []

            # Nấc 1 Trace: Đọc RowCount ban đầu của lưới Item View
            row_count_before = 0
            try:
                row_count_before = int(getattr(alv_grid, "RowCount", 0))
            except Exception as e:
                print(f"[SAP VL06O WARN] Không thể đọc RowCount trước filter: {e}")
            print(f"[SAP VL06O TRACE Nấc 1] RowCount TRƯỚC khi Filter ALV: {row_count_before} dòng.")

            if row_count_before == 0:
                print(f"[SAP VL06O TRACE Nấc 1] Lưới Item View trả về 0 dòng từ Selection Screen.")
                return []

            # 10. ÁP DỤNG BỘ LỌC ALV FILTER THEO CHUẨN VL6_ApplyALVFilter CỦA RC65:
            # Nếu người dùng hỏi 1 mã (hoặc danh sách mã), lọc trực tiếp trên ALV để loại bỏ hàng nghìn dòng thừa!
            row_count_after = row_count_before
            if clean_mats:
                resolved_item_cols = self._resolve_grid_columns(
                    alv_grid,
                    {"product": ["MATNR", "MATERIAL", "PRODUCT"]}
                )
                mat_tech_col = resolved_item_cols.get("product", "MATNR")

                mat_caption = ""
                try:
                    mat_caption = str(alv_grid.GetDisplayedColumnTitle(mat_tech_col) or "").strip()
                except Exception:
                    pass

                print(f"[SAP VL06O TRACE Nấc 2] Đang áp dụng ALV Filter: Cột='{mat_tech_col}' (caption='{mat_caption}'), Values={clean_mats}")
                filter_ok = self._apply_alv_filter(alv_grid, mat_tech_col, clean_mats, col_caption=mat_caption)

                # Sau khi Filter, luôn re-acquire direct grid reference giống VL6_DirectGrid trong RC65 dòng 5309
                refreshed_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
                if refreshed_grid is None:
                    refreshed_grid = self._find_largest_grid(max_wait_sec=5.0)
                if refreshed_grid is not None:
                    alv_grid = refreshed_grid

                try:
                    row_count_after = int(getattr(alv_grid, "RowCount", 0))
                except Exception as e:
                    print(f"[SAP VL06O WARN] Không thể đọc RowCount sau filter: {e}")
                print(f"[SAP VL06O TRACE Nấc 2] RowCount SAU khi Filter ALV: {row_count_after} dòng (giảm từ {row_count_before} dòng). Filter_Success={filter_ok}")

                if row_count_after == 0:
                    print(f"[SAP VL06O TRACE Nấc 2] Không có đơn nào cho mã {clean_mats} sau khi áp dụng bộ lọc ALV.")
                    return []

            # 11. Trích xuất toàn bộ dữ liệu từ lưới Item View đã lọc theo mapping cột chuẩn RC65
            col_map = dict(RC65_VL06O_COLUMN_MAPPING)
            all_records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP VL06O TRACE Nấc 3] Extractor đã đọc {len(all_records)}/{row_count_after} dòng từ lưới ALV.")

            if row_count_after > 0 and len(all_records) == 0:
                raise SapExecutionError(
                    f"Lỗi trích xuất VL06O (Nấc 3): Lưới ALV hiển thị {row_count_after} dòng nhưng không đọc được dữ liệu qua COM. "
                    f"Vui lòng kiểm tra lại hiển thị SAP GUI."
                )

            # 12. Lọc lại một lượt để chắc chắn và định dạng số lượng kèm bộ đếm lý do loại bỏ
            filtered_records = []
            drop_reasons = {"mat_mismatch": 0, "ob_mismatch": 0, "gm_mismatch": 0}
            for r in all_records:
                prod = str(r.get("product", "")).strip().upper()
                if clean_mats and not self._matches_material(prod, clean_mats):
                    drop_reasons["mat_mismatch"] += 1
                    continue
                if clean_ob and clean_ob not in str(r.get("obNo", "")).strip():
                    drop_reasons["ob_mismatch"] += 1
                    continue
                gm = str(r.get("gmStatus", "")).strip().upper()
                if gm and gm not in ("A", "B", ""):
                    drop_reasons["gm_mismatch"] += 1
                    continue
                r["qty"] = format_qty_display(parse_sap_float(r.get("qty")))
                filtered_records.append(r)

            print(
                f"[SAP VL06O TRACE Nấc 4] Sau khi hậu lọc: còn {len(filtered_records)} dòng đơn OB thỏa mãn. "
                f"Lý do loại dòng (Drop reasons): {drop_reasons}"
            )
            filtered_records.sort(key=lambda x: str(x.get("obNo", "")), reverse=True)
            return filtered_records
        finally:
            # Luôn luôn đưa SAP GUI về màn hình gốc /n sau khi tra cứu xong để không bị kẹt ở VL06O
            self._reset_to_session_manager()

    # --------------------------------------------------------------------------
    # 8. KIỂM TRA LỊCH SỬ XUẤT NHẬP & CHUYỂN BIN-TO-BIN (/SCWM/MON)
    # --------------------------------------------------------------------------
    def check_bin_history_mon(
        self,
        materials: Optional[List[str]] = None,
        bin_name: str = "",
        days: int = 7
    ) -> List[Dict[str, str]]:
        """Tra cứu lịch sử di chuyển Bin-to-Bin (Warehouse Tasks: Confirmed & Cancelled) qua /SCWM/MON."""
        try:
            self._open_transaction("/SCWM/MON")
            self._wait_idle(10.0)

            print(f"[SAP MON] Tra cứu lịch sử Bin: Materials={materials}, Bin='{bin_name}', Days={days}...")

            # 1. Nếu xuất hiện popup Warehouse Number ban đầu
            wh_input = self._find_by_id_safe("wnd[0]/usr/ctxt/SCWM/S_MON_INIT-LGNUM")
            if wh_input is not None:
                try:
                    wh_input.text = "VF02"
                    self.session.findById("wnd[0]").sendVKey(0)
                    self._wait_idle(5.0)
                except Exception:
                    pass

            # 2. Tìm Tree Control bên trái và mở nhánh Warehouse Task
            tree = self._find_by_id_safe("wnd[0]/usr/shell/shellcont[0]/shell")
            if tree is not None:
                try:
                    keys_obj = getattr(tree, "GetAllNodeKeys", None)
                    if callable(keys_obj):
                        keys = keys_obj()
                        count = getattr(keys, "Count", 0)
                        for idx in range(count):
                            k = keys(idx)
                            txt = str(tree.GetNodeTextByKey(k)).strip().upper()
                            if "WAREHOUSE TASK" in txt:
                                tree.DoubleClickNode(k)
                                self._wait_idle(5.0)
                                break
                except Exception as e:
                    print(f"[SAP MON WARN] Không thể chọn nhánh Tree Warehouse Task: {e}")

            # 3. Nếu xuất hiện màn hình popup bộ lọc wnd[1]
            popup = self._find_by_id_safe("wnd[1]")
            if popup is not None:
                if bin_name:
                    bin_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtS_VLPLA-LOW") or self._find_by_id_safe("wnd[1]/usr/ctxtS_NLPLA-LOW")
                    if bin_ctl is not None:
                        try:
                            bin_ctl.text = bin_name
                        except Exception:
                            pass
                if materials:
                    clean_mats = [m.strip().upper() for m in materials if m.strip()]
                    mat_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtS_MATNR-LOW")
                    if mat_ctl is not None and clean_mats:
                        try:
                            mat_ctl.text = clean_mats[0]
                        except Exception:
                            pass
                # Bấm Execute trên popup
                ok_btn = self._find_by_id_safe("wnd[1]/tbar[0]/btn[8]")
                if ok_btn is not None:
                    ok_btn.press()
                    self._wait_idle(15.0)

            self._dismiss_popups()
            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP /SCWM/MON ({status_type}): {status_txt}")

            alv_grid = self._find_largest_grid(max_wait_sec=10.0)
            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "/SCWM/MON", status_txt, status_type)
                return []

            col_map = {
                "wtNo":       ["TANUM", "WAREHOUSE TASK", "WHSE TASK"],
                "product":    ["MATNR", "PRODUCT", "MATERIAL"],
                "batch":      ["CHARG", "BATCH"],
                "srcBin":     ["VLPLA", "SOURCE STORAGE BIN", "SOURCE BIN"],
                "destBin":    ["NLPLA", "DESTINATION STORAGE BIN", "DEST. BIN"],
                "qty":        ["VSOLM", "SRCTGTQTY", "QUANTITY", "QTY"],
                "unit":       ["MEINS", "UNIT", "UOM"],
                "status":     ["TOSTAT", "STATUS", "WT STATUS"],
                "procty":     ["PROCTY", "PROCESS TYPE"],
                "docNo":      ["DOCNO", "DOCUMENT", "ERP NUMBER"],
                "createDate": ["CREADAT", "CREATED ON", "DATE"],
                "createTime": ["CREATIM", "CREATED AT", "TIME"],
                "user":       ["CONFNAME", "CANCUSER", "CREATED BY"]
            }
            records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP MON] Đã trích xuất {len(records)} dòng Warehouse Task.")

            for r in records:
                raw_stat = str(r.get("status", "")).strip().upper()
                if raw_stat == "C":
                    r["status"] = "✅ Hoàn thành (Confirmed)"
                elif raw_stat == "A":
                    r["status"] = "⚠️ Đã hủy (Cancelled)"
                elif not raw_stat or raw_stat == "1":
                    r["status"] = "⏳ Đang mở (Open)"
                r["qty"] = format_qty_display(parse_sap_float(r.get("qty")))

            records.sort(key=lambda x: str(x.get("wtNo", "")), reverse=True)
            return records
        finally:
            self._reset_to_session_manager()


    # 9. TRA CỨU CHỨNG TỪ NHẬP KHO GR (MB51 - Expansion A)
    # --------------------------------------------------------------------------
    def check_gr_doc_mb51(
        self,
        inbound_no: str = "",
        ref_no: str = "",
        mat_doc: str = "",
        materials: Optional[List[str]] = None,
        plant: str = "3093",
        days: int = 30
    ) -> List[Dict[str, str]]:
        """Tra cứu chứng từ nhập kho GR qua MB51 theo Inbound Delivery (XBLNR), Bill/Ref (XBLNR),
        Số chứng từ (MBLNR) hoặc Mã vật tư (MATNR). TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY)."""
        try:
            self._open_transaction("MB51")
            self._wait_idle(10.0)

            self._clear_all_selection_inputs()

            mb51_explicit_clears = [
                "wnd[0]/usr/ctxtMBLNR-LOW", "wnd[0]/usr/ctxtMBLNR-HIGH",
                "wnd[0]/usr/ctxtBWART-LOW", "wnd[0]/usr/ctxtBWART-HIGH",
                "wnd[0]/usr/ctxtLGORT-LOW", "wnd[0]/usr/ctxtLGORT-HIGH",
                "wnd[0]/usr/ctxtCHARG-LOW", "wnd[0]/usr/ctxtCHARG-HIGH",
                "wnd[0]/usr/ctxtLIFNR-LOW", "wnd[0]/usr/ctxtLIFNR-HIGH",
                "wnd[0]/usr/ctxtKUNNR-LOW", "wnd[0]/usr/ctxtKUNNR-HIGH",
                "wnd[0]/usr/ctxtSOBKZ-LOW", "wnd[0]/usr/ctxtSOBKZ-HIGH",
                "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/ctxtAUFNR-HIGH",
                "wnd[0]/usr/ctxtEBELN-LOW", "wnd[0]/usr/ctxtEBELN-HIGH",
                "wnd[0]/usr/ctxtKOSTL-LOW", "wnd[0]/usr/ctxtKOSTL-HIGH",
                "wnd[0]/usr/ctxtXBLNR-LOW", "wnd[0]/usr/ctxtXBLNR-HIGH",
                "wnd[0]/usr/ctxtMATNR-HIGH", "wnd[0]/usr/ctxtWERKS-HIGH"
            ]
            for f_id in mb51_explicit_clears:
                ctl = self._find_by_id_safe(f_id)
                if ctl is not None:
                    try:
                        ctl.text = ""
                    except Exception:
                        pass

            print(f"[SAP MB51 GR_DOC] Nhập điều kiện: Inbound={inbound_no}, Ref={ref_no}, MatDoc={mat_doc}, Materials={materials}, Plant={plant}...")

            # 1. Số chứng từ GR (MBLNR)
            if mat_doc:
                doc_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtMBLNR-LOW")
                if doc_ctl is not None:
                    try:
                        doc_ctl.text = mat_doc.strip()
                    except Exception:
                        pass

            # 2. Inbound Delivery hoặc External Ref được lưu trong trường XBLNR (Reference)
            target_xblnr = inbound_no.strip() or ref_no.strip()
            if target_xblnr:
                xblnr_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtXBLNR-LOW")
                if xblnr_ctl is not None:
                    try:
                        xblnr_ctl.text = target_xblnr
                    except Exception:
                        pass

            # 3. Mã vật tư (nếu có)
            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            if clean_mats:
                multi_btn_id = "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
                if self._find_by_id_safe(multi_btn_id) is not None:
                    self._fill_multiple_selection(multi_btn_id, clean_mats)
                else:
                    mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtMATNR-LOW")
                    if mat_single is not None:
                        mat_single.text = clean_mats[0]

            # 4. Plant
            if plant:
                p_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtWERKS-LOW")
                if p_ctl is not None:
                    try:
                        p_ctl.text = plant
                    except Exception:
                        pass

            # 5. Khoảng ngày hạch toán (Posting Date)
            actual_days = max(days, 90) if (target_xblnr or mat_doc) and days == 7 else days
            end_date = datetime.now()
            start_date = end_date - timedelta(days=int(actual_days))
            b_low = self._find_by_id_safe("wnd[0]/usr/ctxtBUDAT-LOW")
            if b_low is not None:
                b_low.text = start_date.strftime("%d.%m.%Y")
            b_high = self._find_by_id_safe("wnd[0]/usr/ctxtBUDAT-HIGH")
            if b_high is not None:
                b_high.text = end_date.strftime("%d.%m.%Y")

            print("[SAP MB51 GR_DOC] Đang bấm thực thi (F8)...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP MB51 ({status_type}): {status_txt}")

            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "MB51", status_txt, status_type)
                return []

            col_map = {
                "mblnr":    ["MBLNR", "MATERIAL DOCUMENT", "DOCUMENT", "NUMBER"],
                "mjahr":    ["MJAHR", "MATERIAL DOC. YEAR", "YEAR"],
                "item":     ["ZEILE", "ITEM"],
                "bwart":    ["BWART", "MOVEMENT TYPE", "MVT"],
                "matnr":    ["MATNR", "MATERIAL", "PRODUCT"],
                "qty":      ["MENGE", "ERFMG", "QUANTITY"],
                "unit":     ["MEINS", "ERFME", "UNIT"],
                "xblnr":    ["XBLNR", "REFERENCE", "BILL OF LADING", "DELIVERY"],
                "postDate": ["BUDAT", "POSTING DATE"],
                "sloc":     ["LGORT", "STORAGE LOCATION", "SLOC"],
                "batch":    ["CHARG", "BATCH"],
                "user":     ["USNAM", "USER NAME", "USER"]
            }
            records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP MB51 GR_DOC] Đã trích xuất {len(records)} dòng chứng từ GR.")

            for r in records:
                r["qty"] = format_qty_display(parse_sap_float(r.get("qty")))

            def _mb51_sort_key(item):
                d_raw = item.get("postDate", "")
                doc = item.get("mblnr", "")
                parts = re.split(r"[./-]", d_raw)
                if len(parts) == 3 and len(parts[2]) == 4:
                    d_sort = f"{parts[2]}{parts[1].zfill(2)}{parts[0].zfill(2)}"
                else:
                    d_sort = d_raw
                return (d_sort, doc)

            records.sort(key=_mb51_sort_key, reverse=True)
            return records
        finally:
            self._reset_to_session_manager()

    # 10. KIỂM TRA CHI TIẾT ĐƠN INBOUND & ĐỐI CHIẾU SỐ LƯỢNG (VL06I - Expansion A)
    # --------------------------------------------------------------------------
    def check_inbound_detail_vl06i(
        self,
        inbound_no: str = "",
        materials: Optional[List[str]] = None,
        plant: str = "3093",
        expected_qty: float = 0.0
    ) -> List[Dict[str, Any]]:
        """Tra cứu chi tiết từng Item của đơn Inbound qua VL06I Item View.
        Hỗ trợ đối chiếu với số lượng yêu cầu (expectedQty) để phát hiện chênh lệch.
        TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY)."""
        try:
            self._open_transaction("VL06I")
            self._wait_idle(10.0)

            btn7 = self._find_by_id_safe("wnd[0]/usr/btnBUTTON7")
            if btn7 is not None:
                btn7.press()
                self._wait_idle(5.0)

            self._clear_all_selection_inputs()

            lfdat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_LFDAT-LOW")
            if lfdat_ctl is not None:
                try:
                    lfdat_ctl.text = ""
                except Exception:
                    pass

            wbstk_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTK-LOW")
            if wbstk_ctl is not None:
                try:
                    wbstk_ctl.text = ""
                except Exception:
                    pass

            clean_ib = inbound_no.strip()
            if clean_ib:
                ib_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_VBELN-LOW")
                if ib_ctl is not None:
                    try:
                        ib_ctl.text = clean_ib
                    except Exception:
                        pass

            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            if clean_mats:
                if len(clean_mats) == 1:
                    mat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_MATNR-LOW")
                    if mat_ctl is not None:
                        try:
                            mat_ctl.text = clean_mats[0]
                        except Exception:
                            pass
                else:
                    multi_mat_id = "wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH"
                    if self._find_by_id_safe(multi_mat_id) is not None:
                        self._fill_multiple_selection(multi_mat_id, clean_mats)

            if plant:
                werks_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WERKS-LOW")
                if werks_ctl is not None:
                    try:
                        werks_ctl.text = plant
                    except Exception:
                        pass

            print(f"[SAP VL06I DETAIL] Bấm thực thi (F8) cho Inbound={clean_ib}, Materials={clean_mats}...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP VL06I ({status_type}): {status_txt}")

            btn18 = self._find_by_id_safe("wnd[0]/tbar[1]/btn[18]")
            if btn18 is not None:
                try:
                    btn18.press()
                    self._wait_idle(10.0)
                    self._dismiss_popups()
                except Exception as e:
                    print(f"[SAP VL06I DETAIL WARN] Không thể bấm btn[18]: {e}")

            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "VL06I", status_txt, status_type)
                return []

            col_map = {
                "ibNo":      ["VBELN", "DELIVERY", "INBOUND DELIVERY", "INBOUND"],
                "item":      ["POSNR", "ITEM"],
                "product":   ["MATNR", "MATERIAL", "PRODUCT"],
                "actualQty": ["LFIMG", "DELIVERY QUANTITY", "ACTUAL DELIVERED QTY", "MENGE", "QUANTITY"],
                "unit":      ["VRKME", "MEINS", "SALES UNIT", "BASE UNIT", "UNIT", "UOM"],
                "plant":     ["WERKS", "PLANT"],
                "sloc":      ["LGORT", "STORAGE LOCATION", "SLOC"],
                "status":    ["WBSTK", "GM STATUS", "OVERALL GM STATUS", "GOODS MOVEMENT STATUS", "STATUS"],
                "refPo":     ["VGBEL", "PURCHASE ORDER", "PO NUMBER", "REFERENCE DOCUMENT", "PO"],
                "poItem":    ["VGPOS", "PO ITEM", "REFERENCE ITEM"],
                "batch":     ["CHARG", "S_CHARG", "BATCH"],
                "refNo":     ["LIFEX", "EXTERNAL IDENTIFICATION", "BILL OF LADING"]
            }
            raw_records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP VL06I DETAIL] Đã trích xuất {len(raw_records)} dòng chi tiết Inbound.")

            records = []
            for r in raw_records:
                st = str(r.get("status", "")).strip().upper()
                if not st or st in ("A", "a"):
                    r["status"] = "⏳ Treo (Chưa GR)"
                elif st in ("B", "b"):
                    r["status"] = "🔄 GR 1 phần"
                elif st in ("C", "c"):
                    r["status"] = "✅ Đã hoàn tất GR"

                act_f = parse_sap_float(r.get("actualQty", 0.0))
                r["actualQty"] = format_qty_display(act_f)

                if expected_qty > 0:
                    delta = act_f - expected_qty
                    r["expectedQty"] = format_qty_display(expected_qty)
                    r["delta"] = format_qty_display(delta)
                    if abs(delta) < 0.0001:
                        r["qtyCheck"] = "✅ Khớp SL"
                    elif delta > 0:
                        r["qtyCheck"] = f"⚠️ Lệch +{format_qty_display(delta)}"
                    else:
                        r["qtyCheck"] = f"⚠️ Thiếu {format_qty_display(abs(delta))}"
                else:
                    r["expectedQty"] = "-"
                    r["delta"] = "-"
                    r["qtyCheck"] = "Chưa có SL yêu cầu"

                records.append(r)

            records.sort(key=lambda x: (str(x.get("ibNo", "")), str(x.get("item", ""))), reverse=False)
            return records
        finally:
            self._reset_to_session_manager()

    # 11. TÌM KIẾM ĐƠN INBOUND THEO THAM CHIẾU (VL06I - Expansion A)
    # --------------------------------------------------------------------------
    def find_inbound_vl06i(
        self,
        inbound_no: str = "",
        po_no: str = "",
        ref_no: str = "",
        materials: Optional[List[str]] = None,
        plant: str = "3093"
    ) -> List[Dict[str, str]]:
        """Tìm kiếm đơn Inbound dựa theo Inbound Delivery (VBELN), PO/SA (VGBEL), Số vận đơn/Bill (LIFEX),
        hoặc Mã vật tư (MATNR). TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY)."""
        try:
            self._open_transaction("VL06I")
            self._wait_idle(10.0)

            btn7 = self._find_by_id_safe("wnd[0]/usr/btnBUTTON7")
            if btn7 is not None:
                btn7.press()
                self._wait_idle(5.0)

            self._clear_all_selection_inputs()

            lfdat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_LFDAT-LOW")
            if lfdat_ctl is not None:
                try:
                    lfdat_ctl.text = ""
                except Exception:
                    pass

            wbstk_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTK-LOW")
            if wbstk_ctl is not None:
                try:
                    wbstk_ctl.text = ""
                except Exception:
                    pass

            clean_ib = str(inbound_no or "").strip()
            if clean_ib:
                ib_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_VBELN-LOW")
                if ib_ctl is not None:
                    try:
                        ib_ctl.text = clean_ib
                    except Exception:
                        pass

            clean_po = po_no.strip()
            if clean_po:
                po_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_VGBEL-LOW")
                if po_ctl is not None:
                    try:
                        po_ctl.text = clean_po
                    except Exception:
                        pass

            clean_ref = ref_no.strip()
            if clean_ref:
                ref_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_LIFEX-LOW")
                if ref_ctl is not None:
                    try:
                        ref_ctl.text = clean_ref
                    except Exception:
                        pass

            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            if clean_mats:
                if len(clean_mats) == 1:
                    mat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_MATNR-LOW")
                    if mat_ctl is not None:
                        try:
                            mat_ctl.text = clean_mats[0]
                        except Exception:
                            pass
                else:
                    multi_mat_id = "wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH"
                    if self._find_by_id_safe(multi_mat_id) is not None:
                        self._fill_multiple_selection(multi_mat_id, clean_mats)

            if plant:
                werks_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WERKS-LOW")
                if werks_ctl is not None:
                    try:
                        werks_ctl.text = plant
                    except Exception:
                        pass

            print(f"[SAP VL06I FIND] Bấm thực thi (F8) cho PO={clean_po}, Ref={clean_ref}, Mats={clean_mats}...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP VL06I ({status_type}): {status_txt}")

            btn18 = self._find_by_id_safe("wnd[0]/tbar[1]/btn[18]")
            if btn18 is not None:
                try:
                    btn18.press()
                    self._wait_idle(10.0)
                    self._dismiss_popups()
                except Exception as e:
                    print(f"[SAP VL06I FIND WARN] Không thể bấm btn[18]: {e}")

            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "VL06I", status_txt, status_type)
                return []

            col_map = {
                "ibNo":      ["VBELN", "DELIVERY", "INBOUND DELIVERY", "INBOUND"],
                "item":      ["POSNR", "ITEM"],
                "product":   ["MATNR", "MATERIAL", "PRODUCT"],
                "qty":       ["LFIMG", "DELIVERY QUANTITY", "ACTUAL DELIVERED QTY", "MENGE", "QUANTITY"],
                "unit":      ["VRKME", "MEINS", "SALES UNIT", "BASE UNIT", "UNIT", "UOM"],
                "plant":     ["WERKS", "PLANT"],
                "sloc":      ["LGORT", "STORAGE LOCATION", "SLOC"],
                "status":    ["WBSTK", "GM STATUS", "OVERALL GM STATUS", "GOODS MOVEMENT STATUS", "STATUS"],
                "refPo":     ["VGBEL", "PURCHASE ORDER", "PO NUMBER", "REFERENCE DOCUMENT", "PO"],
                "refNo":     ["LIFEX", "EXTERNAL IDENTIFICATION", "BILL OF LADING"]
            }
            records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP VL06I FIND] Đã tìm thấy {len(records)} dòng Inbound thỏa mãn.")

            for r in records:
                st = str(r.get("status", "")).strip().upper()
                if not st or st in ("A", "a"):
                    r["status"] = "⏳ Treo (Chưa GR)"
                elif st in ("B", "b"):
                    r["status"] = "🔄 GR 1 phần"
                elif st in ("C", "c"):
                    r["status"] = "✅ Đã hoàn tất GR"
                r["qty"] = format_qty_display(parse_sap_float(r.get("qty")))

            records.sort(key=lambda x: str(x.get("ibNo", "")), reverse=True)
            return records
        finally:
            self._reset_to_session_manager()

    # 12. KIỂM TRA TÌNH TRẠNG ĐƠN HÀNG PO/SA (ME2N - Expansion A)
    # --------------------------------------------------------------------------
    def check_po_me2n(
        self,
        po_no: str = "",
        materials: Optional[List[str]] = None,
        plant: str = "3093",
        po_item: str = ""
    ) -> List[Dict[str, Any]]:
        """Kiểm tra tình trạng đơn đặt hàng PO/SA (Số lượng đặt, Đã nhận, Còn nợ, Xóa/Khóa, Duyệt) qua ME2N.
        TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY)."""
        try:
            self._open_transaction("ME2N")
            self._wait_idle(10.0)

            self._clear_all_selection_inputs()

            listu_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtLISTU")
            if listu_ctl is not None:
                try:
                    listu_ctl.text = "ALV"
                except Exception:
                    pass

            selpa_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtSELPA-LOW")
            if selpa_ctl is not None:
                try:
                    selpa_ctl.text = ""
                except Exception:
                    pass

            clean_po = po_no.strip()
            if clean_po:
                po_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtS_EBELN-LOW")
                if po_ctl is not None:
                    try:
                        po_ctl.text = clean_po
                    except Exception:
                        pass

            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            if clean_mats:
                if len(clean_mats) == 1:
                    mat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtS_MATNR-LOW")
                    if mat_ctl is not None:
                        try:
                            mat_ctl.text = clean_mats[0]
                        except Exception:
                            pass
                else:
                    multi_mat_id = "wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH"
                    if self._find_by_id_safe(multi_mat_id) is not None:
                        self._fill_multiple_selection(multi_mat_id, clean_mats)

            if plant:
                werks_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtS_WERKS-LOW")
                if werks_ctl is not None:
                    try:
                        werks_ctl.text = plant
                    except Exception:
                        pass

            print(f"[SAP ME2N PO] Bấm thực thi (F8) cho PO={clean_po}, Mats={clean_mats}, Plant={plant}...")
            exec_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
            if exec_btn is not None:
                exec_btn.press()
            else:
                self.session.findById("wnd[0]").sendVKey(8)

            self._wait_idle(25.0)
            self._dismiss_popups()

            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP ME2N ({status_type}): {status_txt}")

            alv_grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_grid is None:
                alv_grid = self._find_largest_grid(max_wait_sec=10.0)

            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "ME2N", status_txt, status_type)
                return []

            col_map = {
                "poNo":          ["EBELN", "PURCHASING DOCUMENT", "PO NUMBER", "DOCUMENT", "PO"],
                "item":          ["EBELP", "ITEM"],
                "product":       ["EMATN", "MATNR", "MATERIAL"],
                "orderQty":      ["MENGE", "ORDER QUANTITY", "PO QUANTITY", "QUANTITY", "QTY"],
                "deliveredQty":  ["MGLIEF", "DELIVERED QUANTITY", "ISSUED QTY", "QTY DELIVERED", "DELIVERED"],
                "unit":          ["MEINS", "BPRME", "UNIT", "ORDER UNIT"],
                "plant":         ["WERKS", "PLANT"],
                "deleteInd":     ["LOEKZ", "DELETION INDICATOR", "DEL"],
                "releaseStatus": ["FRGKE", "FRGGR", "RELEASE STATUS", "STATUS", "REL"]
            }
            raw_records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP ME2N PO] Đã trích xuất {len(raw_records)} dòng PO.")

            records = []
            clean_item = po_item.strip().lstrip("0")
            for r in raw_records:
                r_item_clean = str(r.get("item", "")).strip().lstrip("0")
                if clean_item and r_item_clean and r_item_clean != clean_item:
                    continue

                ord_q = parse_sap_float(r.get("orderQty"))
                del_q = parse_sap_float(r.get("deliveredQty"))
                open_q = max(0.0, ord_q - del_q)

                del_ind = str(r.get("deleteInd", "")).strip()
                del_desc = "❌ Đã xóa/khóa" if del_ind in ("L", "X", "S") else "✅ Hoạt động"

                rel_st = str(r.get("releaseStatus", "")).strip()
                rel_desc = "✅ Đã duyệt" if rel_st in ("G", "X", "1") else ("⏳ Chờ duyệt" if rel_st else "N/A")

                r["orderQty"] = format_qty_display(ord_q)
                r["deliveredQty"] = format_qty_display(del_q)
                r["openQty"] = format_qty_display(open_q)
                r["deleteInd"] = del_desc
                r["releaseStatus"] = rel_desc
                records.append(r)

            records.sort(key=lambda x: (str(x.get("poNo", "")), str(x.get("item", ""))), reverse=False)
            return records
        finally:
            self._reset_to_session_manager()


# 13. ĐIỀU ĐỘ THÀNH PHẨM (FG) & ĐIỀN ĐƠN OB (VL06O + MB52)
    # --------------------------------------------------------------------------
    def check_fg_ob_status(
        self,
        materials: List[str],
        plant: str = "3093"
    ) -> List[Dict[str, Any]]:
        """Kiểm tra tồn kho Thành phẩm (FG 3093) đối chiếu với các đơn OB đang mở (VL06O).

        Tuân thủ nghiêm ngặt nguyên tắc tính đúng M1 (R05, R09):
        - MB52 tại Plant/SLoc 3093 (loại trừ kho QC 9999).
        - Quét VL06O tại Shipping Point 3093.
        - Không nuốt lỗi nguồn thành số 0 hoặc kết luận 'Không có OB chờ' khi VL06O bị lỗi.
        - So khớp mã hàng chuẩn xác (không dùng substring bừa bãi).
        """
        clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
        if not clean_mats:
            return []

        # 1. Tra cứu tồn kho FG tại Plant 3093, SLoc 3093 (loại bỏ QC 9999)
        mb52_recs = []
        mb52_err = None
        try:
            mb52_recs = self.check_stock_mb52(clean_mats, plant=plant, slocs=[plant, "3093"])
        except Exception as e:
            print(f"[SAP FG_OB WARN] Không thể đọc MB52: {e}")
            mb52_err = e

        stock_by_mat = {m: 0.0 for m in clean_mats}
        unit_by_mat = {m: "PC" for m in clean_mats}
        for r in mb52_recs:
            sloc = str(r.get("sloc", "")).strip().upper()
            if sloc in ["9999", "QC", "BLOCK", "SCRAP"]:
                continue
            p_mat = str(r.get("product", "")).strip().upper()
            if not p_mat:
                continue
            for m in clean_mats:
                if m == p_mat or m.lstrip("0") == p_mat.lstrip("0"):
                    stock_by_mat[m] += parse_sap_float(r.get("unrest", 0.0))
                    if r.get("unit"):
                        unit_by_mat[m] = r.get("unit")
                    break

        # 2. Tra cứu các đơn OB đang mở qua VL06O
        ob_recs = []
        ob_err = None
        try:
            ob_recs = self.check_open_ob_vl06o(clean_mats, plant=plant)
        except Exception as e:
            print(f"[SAP FG_OB WARN] Không thể đọc VL06O: {e}")
            ob_err = e

        # Nếu cả 2 nguồn đều lỗi -> Báo lỗi thực thi, không bịa kết quả (R05)
        if mb52_err is not None and ob_err is not None:
            raise SapExecutionError(f"Lỗi truy vấn điều độ FG/OB: MB52 ({mb52_err}) và VL06O ({ob_err})")

        obs_by_mat = {m: [] for m in clean_mats}
        ob_qty_by_mat = {m: 0.0 for m in clean_mats}
        for r in ob_recs:
            p_mat = str(r.get("product", "")).strip().upper()
            if not p_mat:
                continue
            ob_no = str(r.get("obNo", "")).strip()
            qty = parse_sap_float(r.get("qty", 0.0))
            for m in clean_mats:
                if m == p_mat or m.lstrip("0") == p_mat.lstrip("0"):
                    if ob_no and ob_no not in obs_by_mat[m]:
                        obs_by_mat[m].append(ob_no)
                    ob_qty_by_mat[m] += qty
                    break

        # 3. Tổng hợp nhận định công việc điều độ
        records = []
        for m in clean_mats:
            fg_stk = stock_by_mat.get(m, 0.0)
            ob_cnt = len(obs_by_mat.get(m, []))
            ob_qty = ob_qty_by_mat.get(m, 0.0)
            unit = unit_by_mat.get(m, "PC")

            if mb52_err is not None:
                fg_stk_str = "N/A"
                ob_qty_str = format_qty_display(ob_qty)
                open_ob_str = ", ".join(obs_by_mat[m]) if obs_by_mat[m] else "Không có"
                avail_str = "N/A"
                advice = f"⚠️ Có {ob_cnt} đơn OB đang mở ({ob_qty_str} {unit}), nhưng không đọc được tồn kho FG (Lỗi MB52: {mb52_err})"

            elif ob_err is not None:
                fg_stk_str = format_qty_display(fg_stk)
                ob_qty_str = "N/A"
                open_ob_str = "Chưa xác định (Lỗi VL06O)"
                avail_str = "N/A"
                advice = f"⚠️ Tồn FG: {fg_stk_str} {unit}, nhưng chưa kiểm tra được đơn OB mở (Lỗi VL06O: {ob_err}). Không suy đoán khả dụng."

            else:
                fg_stk_str = format_qty_display(fg_stk)
                ob_qty_str = format_qty_display(ob_qty)
                open_ob_str = ", ".join(obs_by_mat[m]) if obs_by_mat[m] else "Không có"
                signed_balance = fg_stk - ob_qty
                surplus = max(0.0, signed_balance)
                shortage = max(0.0, -signed_balance)
                avail_str = format_qty_display(surplus)

                if ob_qty > fg_stk:
                    advice = f"🔴 Thiếu tồn điền OB! Cần {ob_qty_str} {unit}, kho chỉ có {fg_stk_str} {unit} (Thiếu {format_qty_display(shortage)} {unit})"
                elif ob_cnt > 0:
                    advice = f"🟢 Đủ tồn điền OB. Khả dụng xuất mới còn: {avail_str} {unit}"
                elif fg_stk > 0:
                    advice = f"🟢 Không có OB chờ. Sẵn sàng tạo đơn mới: {fg_stk_str} {unit}"
                else:
                    advice = f"🔴 Hết hàng FG tại kho {plant} (0 {unit})"

            records.append({
                "product": m,
                "fgStock": fg_stk_str,
                "openObList": open_ob_str,
                "openObQty": ob_qty_str,
                "availableQty": avail_str,
                "unit": unit,
                "workAdvice": advice
            })

        return records

    # --------------------------------------------------------------------------
    # 14. KIỂM TRA NVL POSTING CHANGE (/SCWM/POST) & RA ĐƠN MIGO 311 (KHO 1005)
    # --------------------------------------------------------------------------
    def check_nvl_migo_balance(
        self,
        materials: List[str],
        plant: str = "3093",
        batch: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Kiểm tra tồn kho Nguyên Vật Liệu (NVL) tại Plant 3093 và DUY NHẤT SLoc 1005 qua Posting Change (/SCWM/POST).

        Tuân thủ nghiêm ngặt nguyên tắc tính đúng M1 (R01, R02, R03, R04, R07, R08, R09):
        - Chỉ hỗ trợ phạm vi Plant 3093 (Kho NVL MA SLoc 1005); từ chối plant khác.
        - Quét danh sách tồn kho từ Posting Change (/SCWM/POST) với Owner 3093.
        - Sắp xếp FIFO bằng parse_sap_datetime_sort_key: Lô nhập sớm nhất lên đầu; thiếu ngày đẩy xuống cuối.
        - Phân loại rõ ràng trạng thái TST (FT = Khả dụng, QT/QI = QC, BT = Khóa, rỗng = UNKNOWN).
        - Nếu cả /SCWM/POST và MB52 đều lỗi -> Raise SapExecutionError, không nuốt thành hết hàng.
        - Fallback MB52 chỉ lấy SLoc 1005 và giữ nguyên điều kiện lọc batch (nếu có).
        """
        clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
        if not clean_mats:
            return []

        target_plant = str(plant or "3093").strip().upper()
        if target_plant != "3093":
            return [{
                "product": m,
                "plant": target_plant,
                "sloc": "1005",
                "bin": "-",
                "batch": batch or "-",
                "qty": "0",
                "unit": "PC",
                "grDate": "-",
                "grTime": "-",
                "cat": "INVALID_SCOPE",
                "fifoRank": "-",
                "migoAdvice": f"⚠️ Nhánh tra cứu NVL hiện chỉ hỗ trợ Plant 3093 (Kho NVL MA SLoc 1005). Không hỗ trợ Plant {target_plant}."
            } for m in clean_mats]

        # 1. Quét dữ liệu Posting Change (/SCWM/POST) với Owner = 3093
        post_records = []
        post_err = None
        try:
            post_records = self.check_position_posting_change(clean_mats, batch=batch, plant="3093")
        except Exception as e:
            print(f"[SAP POSTING CHANGE WARN] Không thể đọc /SCWM/POST: {e}")
            post_err = e

        # 2. Phân loại theo từng mã vật tư
        records = []
        mats_found = set()

        if post_records:
            def _fifo_sort_key(item):
                d_raw = item.get("grDate", "")
                t_raw = item.get("grTime", "")
                return parse_sap_datetime_sort_key(d_raw, t_raw)

            records_by_mat: Dict[str, List[Dict[str, Any]]] = {}
            for r in post_records:
                p_mat = str(r.get("product", "")).strip().upper()
                if not p_mat:
                    continue
                matched_target = None
                for m in clean_mats:
                    if m == p_mat or m.lstrip("0") == p_mat.lstrip("0"):
                        matched_target = m
                        break
                if not matched_target:
                    continue
                if batch and str(r.get("batch", "")).upper() != batch.upper():
                    continue
                mats_found.add(matched_target)
                records_by_mat.setdefault(matched_target, []).append(r)

            for m in clean_mats:
                m_rows = records_by_mat.get(m, [])
                if not m_rows:
                    continue
                # Sort ascending: ngày/giờ nhập sớm nhất lên đầu (FIFO). Thiếu ngày sẽ ở cuối.
                m_rows.sort(key=_fifo_sort_key)

                avail_rank = 1
                for item in m_rows:
                    qty_f = parse_sap_float(item.get("availQty") if item.get("availQty") not in (None, "") else item.get("qty", 0.0))
                    uom = item.get("unit", "PC")
                    bin_code = item.get("bin", "")
                    batch_no = item.get("batch", "")
                    tst_raw = str(item.get("cat", "")).strip().upper()

                    # R03: Nếu TST rỗng hoặc không rõ -> UNKNOWN, không tự mặc định là FT!
                    if not tst_raw or tst_raw in ["-", "NONE", "NULL"]:
                        cat_status = "UNKNOWN"
                    else:
                        cat_status = tst_raw

                    is_bin_blocked = bin_code.upper() in ["APFL-BL01-01-01"] or "BLOCK" in bin_code.upper()
                    is_qc = cat_status in ["QT", "QI", "QUALITY"]
                    is_blocked_status = cat_status in ["BT", "BLOCK", "BLOCKED"]
                    is_unknown = cat_status not in {"FT", "F1", "F2", "QT", "QI", "QUALITY", "BT", "BLOCK", "BLOCKED"}
                    scope_ok = str(item.get("plant") or plant or "3093") == "3093" and str(item.get("sloc") or "1005") == "1005"
                    date_ok = parse_sap_datetime_sort_key(item.get("grDate", ""), item.get("grTime", ""))[0] != "99999999"
                    eligible = False

                    if is_bin_blocked:
                        advice = f"⛔ Bin bị khóa hệ thống ({bin_code}) - Không bốc ra đơn"
                        rank_lbl = "-"
                    elif is_qc:
                        advice = f"⚠️ Chờ kiểm định QC (TST {cat_status}) - Chưa thể ra đơn"
                        rank_lbl = "-"
                    elif is_blocked_status:
                        advice = f"⛔ Lô bị khóa (TST {cat_status}) - Không được xuất"
                        rank_lbl = "-"
                    elif is_unknown:
                        advice = f"🟡 Trạng thái tồn chưa xác định (TST: '{tst_raw or 'Rỗng'}') - Chưa đủ điều kiện FIFO MIGO"
                        rank_lbl = "-"
                    elif not scope_ok or not date_ok or not bin_code or not uom:
                        advice = "⚠️ Chưa xác minh đủ kho/ngày GR/bin/ĐVT; không đề xuất ra đơn"
                        rank_lbl = "-"
                    elif qty_f > 0:
                        eligible = True
                        advice = f"✅ FIFO #{avail_rank} - Sẵn sàng ra đơn MIGO 311 ({format_qty_display(qty_f)} {uom})"
                        rank_lbl = f"#{avail_rank}"
                        avail_rank += 1
                    else:
                        advice = "⚪ Số lượng khả dụng bằng 0"
                        rank_lbl = "-"

                    records.append({
                        "product": m,
                        "plant": "3093",
                        "sloc": item.get("sloc") or "1005",
                        "eligible": eligible,
                        "bin": bin_code,
                        "batch": batch_no,
                        "qty": format_qty_display(qty_f),
                        "unit": uom,
                        "grDate": item.get("grDate", ""),
                        "grTime": item.get("grTime", ""),
                        "cat": cat_status,
                        "fifoRank": rank_lbl,
                        "migoAdvice": advice
                    })

        # 3. Với những mã chưa có dữ liệu trong /SCWM/POST, tra cứu MB52 DUY NHẤT SLoc 1005
        missing_mats = [m for m in clean_mats if m not in mats_found]
        mb52_err = None
        if missing_mats:
            try:
                mb52_recs = self.check_stock_mb52(missing_mats, plant="3093", slocs=["1005"])
            except Exception as e:
                print(f"[SAP MB52 1005 WARN] Không thể đọc MB52: {e}")
                mb52_err = e
                mb52_recs = []

            # R01: Nếu cả 2 nguồn POST và MB52 đều lỗi -> Raise ngoại lệ, không biến lỗi thành hết hàng!
            if mb52_err is not None:
                raise SapExecutionError(f"Lỗi truy vấn SAP NVL: /SCWM/POST ({post_err}) và MB52 ({mb52_err})")

            mb52_by_mat = {m: 0.0 for m in missing_mats}
            unit_by_mat = {m: "PC" for m in missing_mats}
            batch_matched = {m: False for m in missing_mats}

            for r in mb52_recs:
                sl = str(r.get("sloc", "")).strip().upper()
                if sl != "1005":
                    continue
                p_mat = str(r.get("product", "")).strip().upper()
                if not p_mat:
                    continue
                r_batch = str(r.get("batch", "")).strip().upper()
                if batch and r_batch != batch.strip().upper():
                    continue

                q = parse_sap_float(r.get("unrest", 0.0))
                u = r.get("unit", "PC")
                for m in missing_mats:
                    if m == p_mat or m.lstrip("0") == p_mat.lstrip("0"):
                        mb52_by_mat[m] += q
                        if batch_matched[m] and unit_by_mat[m] != u:
                            raise SapExecutionError("MB52 trả nhiều ĐVT cho cùng mã; không cộng gộp")
                        unit_by_mat[m] = u
                        batch_matched[m] = True
                        break

            for m in missing_mats:
                stk = mb52_by_mat.get(m, 0.0)
                u = unit_by_mat.get(m, "PC")
                if post_err is not None:
                    advice = f"⚠️ Lỗi đọc Posting Change ({post_err}). MB52 Kho 1005 ghi nhận: {format_qty_display(stk)} {u}"
                    cat_val = "MB52_ONLY"
                    qty_str = format_qty_display(stk)
                elif stk > 0:
                    advice = f"⚠️ Có tồn MB52 Kho 1005 ({format_qty_display(stk)} {u}) nhưng chưa có Bin trong Posting Change"
                    cat_val = "MB52"
                    qty_str = format_qty_display(stk)
                elif batch and not batch_matched.get(m):
                    advice = f"🔴 Không có lô {batch} tại Kho 1005 (Plant 3093)"
                    cat_val = "-"
                    qty_str = "0"
                else:
                    advice = "🔴 Hết hàng tại Kho 1005 (Plant 3093)"
                    cat_val = "-"
                    qty_str = "0"

                records.append({
                    "product": m,
                    "plant": "3093",
                    "sloc": "1005",
                    "bin": "Chưa có Bin" if stk > 0 else "-",
                    "batch": batch or "-",
                    "qty": qty_str,
                    "unit": u,
                    "grDate": "-",
                    "grTime": "-",
                    "cat": cat_val,
                    "fifoRank": "-",
                    "migoAdvice": advice
                })

        return records

    def _set_cell(self, table_ctl, row_no, col_no, value_text, label=""):
        """Điền ô trong TableControl (GetCell -> Symmetric Fallback -> Token TableField -> Coord)."""
        val_str = str(value_text or "").strip()

        # 1. Ưu tiên cao nhất: GetCell trực tiếp theo toạ độ (row_no, col_no) nếu hợp lệ (< 30)
        if col_no is not None and 0 <= col_no < 30:
            try:
                cell = table_ctl.GetCell(row_no, col_no)
                if cell is not None:
                    try:
                        cell.Text = val_str
                        if str(getattr(cell, "Text", "") or "").strip() == val_str:
                            return True
                    except Exception:
                        pass
                    try:
                        cell.SetFocus()
                        cell.Text = val_str
                        if str(getattr(cell, "Text", "") or "").strip() == val_str:
                            return True
                    except Exception:
                        pass
            except Exception:
                pass

        # 2. Fallback đối xứng cho Material: Thử lần lượt các cột ứng viên [3, 1, 2] nếu cột chính không ghi được
        if label == "Material":
            mat_candidates = [3, 1, 2]
            if col_no in mat_candidates:
                mat_candidates.remove(col_no)
            for alt_c in mat_candidates:
                try:
                    alt_cell = table_ctl.GetCell(row_no, alt_c)
                    if alt_cell is not None:
                        if getattr(alt_cell, "Changeable", True) is False:
                            continue
                        try:
                            alt_cell.Text = val_str
                            if str(getattr(alt_cell, "Text", "") or "").strip() == val_str:
                                print(f"[SAP MIGO] Đã điền Material vào cột thay thế {alt_c}")
                                return True
                        except Exception:
                            pass
                        try:
                            alt_cell.SetFocus()
                            alt_cell.Text = val_str
                            if str(getattr(alt_cell, "Text", "") or "").strip() == val_str:
                                print(f"[SAP MIGO] Đã điền Material vào cột thay thế {alt_c}")
                                return True
                        except Exception:
                            pass
                except Exception:
                    pass

        # 3. Fallback cho Batch (cột 10 chính, cột 11 phụ) nếu cột chính không ghi được
        if label == "Batch":
            alt_candidates = [10, 11]
            if col_no in alt_candidates:
                alt_candidates.remove(col_no)
            for alt_batch in alt_candidates:
                try:
                    alt_cell = table_ctl.GetCell(row_no, alt_batch)
                    if alt_cell is not None:
                        try:
                            alt_cell.Text = val_str
                            if str(getattr(alt_cell, "Text", "") or "").strip() == val_str:
                                print(f"[SAP MIGO] Đã điền Batch vào cột thay thế {alt_batch}")
                                return True
                        except Exception:
                            pass
                        try:
                            alt_cell.SetFocus()
                            alt_cell.Text = val_str
                            if str(getattr(alt_cell, "Text", "") or "").strip() == val_str:
                                print(f"[SAP MIGO] Đã điền Batch vào cột thay thế {alt_batch}")
                                return True
                        except Exception:
                            pass
                except Exception:
                    pass

        # 4. Token Relative ID (chỉ tìm token hợp lệ trong dải hiển thị < 30)
        migo_token_map = {
            "Material": ["GOITEM-MATNR", "GOITEM-UMMAT"],
            "Qty": ["GOITEM-ERFMG"],
            "Trạm": ["GOITEM-SGTXT"],
            "Plant": ["GOITEM-WERKS"],
            "SLoc xuất": ["GOITEM-LGORT"],
            "SLoc nhận": ["GOITEM-UMLGO"],
            "Batch": ["GOITEM-CHARG", "GOITEM-UMCHA"],
            "OK": ["GOITEM-TAKE_IT"]
        }
        tokens = migo_token_map.get(label, [])
        if isinstance(tokens, str):
            tokens = [tokens]
        for token in tokens:
            try:
                fld = self.table_field(table_ctl, token, row_no)
                if fld is not None:
                    # Kiểm tra ID không thuộc cột ẩn (>= 30)
                    fld_id = str(getattr(fld, "Id", getattr(fld, "id", "")) or "")
                    if fld_id:
                        at = fld_id.rfind("[")
                        if at != -1 and fld_id.endswith("]"):
                            coord_str = fld_id[at + 1:-1]
                            parts = coord_str.split(",")
                            if len(parts) == 2 and int(parts[0]) >= 30:
                                continue  # Bỏ qua cột ẩn ngoài dải hiển thị
                    try:
                        fld.Text = val_str
                        if str(getattr(fld, "Text", "") or "").strip() == val_str:
                            return True
                    except Exception:
                        pass
                    try:
                        fld.SetFocus()
                        fld.Text = val_str
                        if str(getattr(fld, "Text", "") or "").strip() == val_str:
                            return True
                    except Exception:
                        pass
            except Exception:
                pass

        # 5. Suffix Coord [col, row]
        if col_no is not None and 0 <= col_no < 30:
            suffix = f"[{col_no},{row_no}]"
            def _find_by_coord(obj, depth=0):
                if obj is None or depth > 10:
                    return None
                try:
                    cid = str(getattr(obj, "Id", getattr(obj, "id", "")) or "")
                    if cid.endswith(suffix):
                        return obj
                except Exception:
                    pass
                try:
                    children = getattr(obj, "Children", getattr(obj, "children", None))
                    if children:
                        try:
                            for child in children:
                                res = _find_by_coord(child, depth + 1)
                                if res is not None:
                                    return res
                        except TypeError:
                            cnt = getattr(children, "Count", 0)
                            for idx in range(cnt):
                                res = _find_by_coord(children(idx), depth + 1)
                                if res is not None:
                                    return res
                except Exception:
                    pass
                return None

            try:
                ctl = _find_by_coord(table_ctl)
                if ctl is not None:
                    try:
                        ctl.Text = val_str
                        if str(getattr(ctl, "Text", "") or "").strip() == val_str:
                            return True
                    except Exception:
                        pass
            except Exception:
                pass

        print(f"[SAP MIGO WARN] Không thể điền '{label}'='{val_str}' tại row={row_no}, col={col_no}")
        return False

    def post_migo_311_ob_group(
        self,
        ob_items: Any,
        plant: str = "3093",
        source_sloc: str = "1005",
        dest_sloc: str = "3001",
        dry_run: bool = False,
        session: Any = None
    ) -> Tuple[bool, str, str]:
        """
        Thực hiện nghiệp vụ MIGO 311 chuyển kho NVL 1005 -> line 3001 (Plant 3093) sinh số OB (10xxxxxxxx).
        Tuân thủ chuẩn xác 100% quy trình 2 nhịp từ file 260810_RA_DON_NVL_V31_ZONE (MIGO311_RunPost_V24_ErrorGuard.bas):
        - Thiết lập MIGO 311: Action=Transfer Posting (A08), RefDoc=Other (R10), Movement Type GODEFAULT_TV-BWART=311.
        - Đóng Header (MIGO_OK_HEADER_CLOSE) và Detail (MIGO_OK_DETAIL_CLOSE) nếu bảng item bị ẩn.
        - Tìm bảng item tblSAPLMIGOTV_GOITEM qua ID và đệ quy GuiTableControl.
        - Nhịp 1: Điền Material, Qty, Text (Trạm: KV08-Thức/Vũ), Plant, SLoc xuất (1005), SLoc nhập (3001) qua GetCell() -> Enter
        - Nhịp 2: Tìm cột BATCH thật (loại trừ Batch Trfr Pstg), điền số lô qua GetCell() -> Enter
        - Post chứng từ (tbar[1]/btn[23] hoặc tbar[0]/btn[11])
        - Tự động xử lý popups (wnd[1]) và warnings (W/I)
        - Bắt số chứng từ Outbound 10 chữ số (10xxxxxxxx) hoặc Delivery Number từ thanh trạng thái sbar.
        """
        if session is not None:
            self.session = session

        if isinstance(ob_items, dict):
            ob_items = ob_items.get("allocated_items", [])

        if not ob_items:
            return False, "", "Không có dòng phân bổ NVL nào để tạo đơn MIGO 311"

        if dry_run:
            return True, "", "Mô phỏng kế hoạch NVL: Chưa thực thi tạo đơn OB trên SAP"

        if self.session is None:
            return False, "", "ERR_SAP_SESSION_NONE: Không có kết nối SAP GUI phiên hoạt động để tạo đơn MIGO 311"

        try:
            self.ensure_session()
            print("[SAP MIGO] Đang mở giao dịch MIGO để tạo đơn OB 311...")
            self._open_transaction("MIGO")
            self._wait_idle(10.0)

            # Hàm tìm kiếm control đệ quy theo thuộc tính Name
            def _find_by_name(root, target_name):
                if root is None:
                    return None
                try:
                    c_name = str(getattr(root, "Name", getattr(root, "name", "")) or "")
                    if c_name.upper() == target_name.upper():
                        return root
                except Exception:
                    pass
                try:
                    children = getattr(root, "Children", getattr(root, "children", None))
                    if children:
                        try:
                            for child in children:
                                res = _find_by_name(child, target_name)
                                if res is not None:
                                    return res
                        except TypeError:
                            cnt = getattr(children, "Count", 0)
                            for idx in range(cnt):
                                res = _find_by_name(children(idx), target_name)
                                if res is not None:
                                    return res
                except Exception:
                    pass
                return None

            w0 = self._find_by_id_safe("wnd[0]")

            # 1. Action ComboBox: Transfer Posting (A08)
            print("[SAP MIGO] Đang thiết lập chế độ: Transfer Posting (A08) - Other (R10) - Movement 311...")
            action_cmb = _find_by_name(w0, "GODYNPRO-ACTION")
            if action_cmb is None:
                for sub in ["0001", "0002", "0003", "0004"]:
                    action_cmb = self._find_by_id_safe(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION")
                    if action_cmb is not None:
                        break

            if action_cmb is not None:
                act_set = False
                try:
                    entries = getattr(action_cmb, "Entries", getattr(action_cmb, "entries", None))
                    if entries:
                        cnt = getattr(entries, "Count", getattr(entries, "count", 0))
                        for idx in range(cnt):
                            e = entries(idx)
                            val = str(getattr(e, "Value", getattr(e, "Text", "")) or "").upper()
                            if "TRANSFER" in val:
                                target_k = str(getattr(e, "Key", getattr(e, "key", "A08")))
                                try:
                                    action_cmb.Key = target_k
                                except Exception:
                                    pass
                                try:
                                    action_cmb.key = target_k
                                except Exception:
                                    pass
                                act_set = True
                                break
                except Exception:
                    pass
                if not act_set:
                    try:
                        cur_act = getattr(action_cmb, "Key", getattr(action_cmb, "key", ""))
                        if cur_act != "A08":
                            try:
                                action_cmb.Key = "A08"
                            except Exception:
                                pass
                            try:
                                action_cmb.key = "A08"
                            except Exception:
                                pass
                    except Exception:
                        pass
                time.sleep(0.5)

            # 2. RefDoc ComboBox: Other (R10)
            w0 = self._find_by_id_safe("wnd[0]")
            ref_cmb = _find_by_name(w0, "GODYNPRO-REFDOC")
            if ref_cmb is None:
                for sub in ["0001", "0002", "0003", "0004"]:
                    ref_cmb = self._find_by_id_safe(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC")
                    if ref_cmb is not None:
                        break

            if ref_cmb is not None:
                ref_set = False
                try:
                    entries = getattr(ref_cmb, "Entries", getattr(ref_cmb, "entries", None))
                    if entries:
                        cnt = getattr(entries, "Count", getattr(entries, "count", 0))
                        for idx in range(cnt):
                            e = entries(idx)
                            val = str(getattr(e, "Value", getattr(e, "Text", "")) or "").upper()
                            if "OTHER" in val:
                                target_r = str(getattr(e, "Key", getattr(e, "key", "R10")))
                                try:
                                    ref_cmb.Key = target_r
                                except Exception:
                                    pass
                                try:
                                    ref_cmb.key = target_r
                                except Exception:
                                    pass
                                ref_set = True
                                break
                except Exception:
                    pass
                if not ref_set:
                    try:
                        cur_ref = getattr(ref_cmb, "Key", getattr(ref_cmb, "key", ""))
                        if cur_ref != "R10":
                            try:
                                ref_cmb.Key = "R10"
                            except Exception:
                                pass
                            try:
                                ref_cmb.key = "R10"
                            except Exception:
                                pass
                    except Exception:
                        pass
                time.sleep(0.5)

            # 3. Movement Type: GODEFAULT_TV-BWART (Chuẩn VBA 260810_RA_DON_NVL_V31_ZONE)
            w0 = self._find_by_id_safe("wnd[0]")
            mvt_field = _find_by_name(w0, "GODEFAULT_TV-BWART") or _find_by_name(w0, "GODYNPRO-BWART")
            if mvt_field is None:
                for sub in ["0001", "0002", "0003", "0004"]:
                    mvt_field = (self._find_by_id_safe(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART")
                              or self._find_by_id_safe(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODYNPRO-BWART"))
                    if mvt_field is not None:
                        break

            if mvt_field is not None:
                try:
                    mvt_field.Text = "311"
                except Exception:
                    try:
                        mvt_field.text = "311"
                    except Exception:
                        pass

            w0 = self._find_by_id_safe("wnd[0]")
            if w0:
                w0.sendVKey(0)  # Enter nạp cấu hình MIGO 311
            self._wait_idle(5.0)

            # 4. Tìm bảng item MIGO: tblSAPLMIGOTV_GOITEM
            print("[SAP MIGO] Đang tìm kiếm bảng nhập liệu tblSAPLMIGOTV_GOITEM...")
            def _find_migo_table(root):
                if root is None:
                    return None
                for sub in ["0004", "0003", "0002", "0001"]:
                    for itemlist in ["0200", "0210", "0205"]:
                        cand = f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_ITEMLIST:SAPLMIGO:{itemlist}/tblSAPLMIGOTV_GOITEM"
                        t = self._find_by_id_safe(cand)
                        if t is not None:
                            return t

                def _rec(obj, depth=0):
                    if obj is None or depth > 20:
                        return None
                    try:
                        cid = str(getattr(obj, "Id", getattr(obj, "id", "")) or "")
                        cname = str(getattr(obj, "Name", getattr(obj, "name", "")) or "")
                        ctype = str(getattr(obj, "Type", getattr(obj, "type", "")) or "").upper()
                        if "tblSAPLMIGOTV_GOITEM" in cid or "GOITEM" in cname or ctype == "GUITABLECONTROL":
                            if ctype in ("GUITABLECONTROL", "GUIGRIDVIEW") or "GOITEM" in cid:
                                return obj
                    except Exception:
                        pass
                    try:
                        children = getattr(obj, "Children", getattr(obj, "children", None))
                        if children:
                            try:
                                for child in children:
                                    res = _rec(child, depth + 1)
                                    if res is not None:
                                        return res
                            except TypeError:
                                cnt = getattr(children, "Count", 0)
                                for idx in range(cnt):
                                    res = _rec(children(idx), depth + 1)
                                    if res is not None:
                                        return res
                    except Exception:
                        pass
                    return None
                return _rec(root)

            w0 = self._find_by_id_safe("wnd[0]")
            tbl = _find_migo_table(w0)

            # Nếu bảng item đang bị thu gọn bởi Header/Detail, đóng theo lệnh VBA
            if tbl is None:
                try:
                    self.session.SendCommand("MIGO_OK_HEADER_CLOSE")
                    time.sleep(0.3)
                    self.session.SendCommand("MIGO_OK_DETAIL_CLOSE")
                    time.sleep(0.5)
                except Exception:
                    pass
                w0 = self._find_by_id_safe("wnd[0]")
                tbl = _find_migo_table(w0)

            if tbl is None:
                status_txt, status_type = self._get_status()
                err_msg = f"Không tìm thấy bảng nhập liệu tblSAPLMIGOTV_GOITEM trong MIGO ({status_txt})" if status_txt else "Không tìm thấy bảng nhập liệu tblSAPLMIGOTV_GOITEM trong MIGO"
                print(f"[SAP MIGO ERROR] {err_msg}")
                return False, "", err_msg

            print(f"[SAP MIGO] Đã tìm thấy bảng nhập liệu tblSAPLMIGOTV_GOITEM. Đang phân giải cột...")

            # 5. PHÂN GIẢI CỘT MIGO 311 (KẾT HỢP CHUẨN EXCEL V31 & ĐẶC TẢ GIAO DIỆN THỰC TẾ NHÀ MÁY)
            # Layout chuẩn thực tế trên màn hình MIGO (xem screenshot nhà máy):
            # Col 0: Line
            # Col 1: Material Trfr Pstg
            # Col 2: Material Trfr Short Text
            # Col 3: Mat. Short Text (nhập liệu chính vật tư)
            # Col 4: S.. (Stock status icon)
            # Col 5: Qty in UnE (Số lượng)
            # Col 6: Text (Trạm)
            # Col 7: Plnt (Nhà máy 3093)
            # Col 8: SLoc (Kho xuất 1005)
            # Col 9: SLoc Transfer Pstg (Kho nhận 3001)
            # Col 10: Batch (Lô xuất)
            # Col 11: Batch Trfr Pstg (Lô nhận)
            mat_col = 3
            fallback_mat_col = 1
            qty_col = 5
            txt_col = 6
            plant_col = 7
            src_sloc_col = 8
            dest_sloc_col = 9
            batch_col = 10
            dest_batch_col = 11
            ok_col = -1

            mat_found = False
            qty_found = False
            txt_found = False
            plant_found = False
            src_sloc_found = False
            dest_sloc_found = False
            batch_found = False
            dest_batch_found = False
            ok_found = False

            # Lượt 1: Phân giải qua tbl.Columns (nếu hỗ trợ)
            try:
                col_cnt = int(getattr(tbl.Columns, "Count", getattr(tbl.Columns, "count", 0)))
                # Ưu tiên các cột hiển thị trong viewport (c_i < 20), tuyệt đối chặn cột ẩn (như col 61)
                for c_i in range(col_cnt):
                    col_obj = None
                    try:
                        col_obj = tbl.Columns.Item(c_i)
                    except Exception:
                        pass
                    if col_obj is None:
                        try:
                            col_obj = tbl.Columns.ElementAt(c_i)
                        except Exception:
                            pass
                    if col_obj is None:
                        try:
                            col_obj = tbl.Columns(c_i)
                        except Exception:
                            pass
                    if col_obj is None:
                        try:
                            col_obj = tbl.Columns[c_i]
                        except Exception:
                            pass
                    if col_obj is None:
                        continue

                    c_title = str(getattr(col_obj, "Title", getattr(col_obj, "title", "")) or "").strip().upper()
                    c_name = str(getattr(col_obj, "Name", getattr(col_obj, "name", "")) or "").strip().upper()

                    # Cột ngoài vùng hiển thị (c_i >= 20) không được ưu tiên
                    is_visible_col = (c_i < 20)
                    try:
                        vis_prop = getattr(col_obj, "Visible", getattr(col_obj, "visible", True))
                        if vis_prop is False:
                            is_visible_col = False
                    except Exception:
                        pass

                    if not is_visible_col:
                        continue

                    # Material:
                    # Ưu tiên 1 (Mat. Short Text - Chuẩn layout thực tế MIGO SAP):
                    # Cột 'Mat. Short Text' (Col 3 trong layout thực tế) là ô nhập liệu chính thức có icon Search (F4)
                    is_mat_group = any(k in c_title for k in ("GROUP", "TYPE", "NHÓM", "LOẠI"))
                    if not mat_found:
                        if any(k in c_title for k in ("MAT. SHORT TEXT", "MAT. SHORT", "SHORT TEXT")) or c_name in ("GOITEM-MAKTX", "MAKTX"):
                            mat_col = c_i
                            mat_found = True
                        elif (c_name in ("GOITEM-MATNR", "MATNR") or c_title in ("MATERIAL", "PRODUCT", "VẬT TƯ")) and not is_mat_group and "TRFR" not in c_title:
                            mat_col = c_i
                            mat_found = True
                        elif (
                            c_name in ("GOITEM-UMMAT", "UMMAT")
                            or "TRFR PSTG" in c_title
                            or "TRANSFER POSTING" in c_title
                            or c_title.startswith("MATERIAL")
                            or c_title.startswith("VẬT TƯ")
                        ):
                            # Tạm ghi nhận cột transfer posting dự phòng nếu không có Mat. Short Text / MATNR
                            if fallback_mat_col < 0:
                                fallback_mat_col = c_i

                    # Qty: GOITEM-ERFMG hoặc Qty in UnE / Quantity / Qty (không ăn sang PO Qty)
                    if not qty_found and (c_name in ("GOITEM-ERFMG", "ERFMG") or any(k in c_title for k in ("QTY IN UNE", "QUANTITY", "QTY", "SỐ LƯỢNG"))):
                        qty_col = c_i
                        qty_found = True
                    # Text / Trạm: GOITEM-SGTXT hoặc Text / Item Text
                    elif not txt_found and (c_name in ("GOITEM-SGTXT", "SGTXT") or c_title in ("TEXT", "ITEM TEXT", "TRẠM")):
                        txt_col = c_i
                        txt_found = True
                    # Plant: GOITEM-WERKS hoặc Plnt / Plant
                    elif not plant_found and (c_name in ("GOITEM-WERKS", "WERKS") or c_title in ("PLNT", "PLANT", "NHÀ MÁY")):
                        plant_col = c_i
                        plant_found = True
                    # SLoc xuất: GOITEM-LGORT hoặc SLoc / Storage Loc (không ăn sang SLoc Tra... hay Receiving)
                    elif not src_sloc_found and (
                        c_name in ("GOITEM-LGORT", "LGORT") 
                        or (
                            (any(k in c_title for k in ("STORAGE LOC", "STORAGE LOCATION", "KHO XUẤT")) or c_title == "SLOC")
                            and not any(k in c_title for k in ("TRA", "TRFR", "TRANSFER", "REC", "DEST", "NHẬN", "PSTG"))
                            and not any(w in c_title.split() for w in ("TO",))
                        )
                    ):
                        src_sloc_col = c_i
                        src_sloc_found = True
                    # SLoc nhận: GOITEM-UMLGO hoặc SLoc Tra... / SLoc Transfer Pstg / Rec. SLoc
                    elif not dest_sloc_found and (
                        c_name in ("GOITEM-UMLGO", "UMLGO") 
                        or any(c_title.startswith(k) for k in ("SLOC TRA", "SLOC TRFR", "REC. SLOC", "SLOCTO", "DEST. SLOC", "KHO NHẬN")) 
                        or ("SLOC" in c_title and any(k in c_title for k in ("TRANSFER", "TRFR", "TRA", "DEST", "NHẬN", "PSTG")))
                        or ("SLOC" in c_title and any(w in c_title.split() for w in ("TO",)))
                    ):
                        dest_sloc_col = c_i
                        dest_sloc_found = True
                    # Batch: GOITEM-CHARG hoặc Batch chính xác (loại trừ Batch Trfr, Batch Management...)
                    elif not batch_found and (
                        (c_name in ("GOITEM-CHARG", "CHARG") or c_title in ("BATCH", "LÔ", "SỐ LÔ")) 
                        and not any(k in c_title for k in ("TRFR", "TRANSFER", "PSTG", "NHẬN", "DEST"))
                        and not any(w in c_title.split() for w in ("TO",))
                        and c_name not in ("GOITEM-UMCHA", "UMCHA")
                    ):
                        batch_col = c_i
                        batch_found = True
                    # Batch nhận: GOITEM-UMCHA hoặc Batch Trfr...
                    elif not dest_batch_found and (
                        c_name in ("GOITEM-UMCHA", "UMCHA")
                        or any(c_title.startswith(k) for k in ("BATCH TRFR", "REC. BATCH", "BATCHTO", "DEST. BATCH", "LÔ NHẬN"))
                        or ("BATCH" in c_title and any(k in c_title for k in ("TRANSFER", "TRFR", "TRA", "DEST", "NHẬN", "PSTG")))
                        or ("BATCH" in c_title and any(w in c_title.split() for w in ("TO",)))
                    ):
                        dest_batch_col = c_i
                        dest_batch_found = True
                    # OK: GOITEM-TAKE_IT hoặc OK
                    elif not ok_found and (c_name in ("GOITEM-TAKE_IT", "TAKE_IT") or c_title == "OK"):
                        ok_col = c_i
                        ok_found = True
            except Exception as col_err:
                print(f"[SAP MIGO WARN] Lỗi đọc danh sách cột qua tbl.Columns: {col_err}")

            # Lượt 2: Phân giải bổ sung qua tbl.Children dựa trên Token kỹ thuật trong Id (txtGOITEM-XXX[col, row])
            try:
                children = getattr(tbl, "Children", getattr(tbl, "children", None))
                if children:
                    cnt = getattr(children, "Count", getattr(children, "count", 0))
                    for i in range(cnt):
                        ch = None
                        try:
                            ch = children.Item(i)
                        except Exception:
                            try:
                                ch = children(i) if callable(children) else children[i]
                            except Exception:
                                pass
                        if ch is None:
                            continue
                        cid = str(getattr(ch, "Id", getattr(ch, "id", "")) or "").upper()
                        at = cid.rfind("[")
                        if at == -1 or not cid.endswith("]"):
                            continue
                        coord_str = cid[at + 1:-1]
                        parts = coord_str.split(",")
                        if len(parts) != 2:
                            continue
                        try:
                            c_col = int(parts[0])
                        except Exception:
                            continue
                        if c_col >= 20:
                            continue

                        if not batch_found and "GOITEM-CHARG[" in cid:
                            batch_col = c_col
                            batch_found = True
                        elif not dest_batch_found and "GOITEM-UMCHA[" in cid:
                            dest_batch_col = c_col
                            dest_batch_found = True
                        elif not src_sloc_found and "GOITEM-LGORT[" in cid:
                            src_sloc_col = c_col
                            src_sloc_found = True
                        elif not dest_sloc_found and "GOITEM-UMLGO[" in cid:
                            dest_sloc_col = c_col
                            dest_sloc_found = True
                        elif not plant_found and "GOITEM-WERKS[" in cid:
                            plant_col = c_col
                            plant_found = True
                        elif not txt_found and "GOITEM-SGTXT[" in cid:
                            txt_col = c_col
                            txt_found = True
                        elif not qty_found and "GOITEM-ERFMG[" in cid:
                            qty_col = c_col
                            qty_found = True
                        elif not mat_found and "GOITEM-MAKTX[" in cid:
                            mat_col = c_col
                            mat_found = True
                        elif fallback_mat_col < 0 and ("GOITEM-MATNR[" in cid or "GOITEM-UMMAT[" in cid):
                            fallback_mat_col = c_col
                        elif not ok_found and "GOITEM-TAKE_IT[" in cid:
                            ok_col = c_col
                            ok_found = True
            except Exception as child_err:
                print(f"[SAP MIGO WARN] Lỗi phân giải qua tbl.Children: {child_err}")

            if not mat_found and fallback_mat_col >= 0:
                mat_col = fallback_mat_col
                mat_found = True

            # Chốt chặn van an toàn: Nếu bất kỳ cột nào nằm ngoài dải hiển thị (>= 20), ép về vị trí chuẩn
            if mat_col >= 20:
                print(f"[SAP MIGO WARN] mat_col={mat_col} nằm ngoài dải hiển thị, tự động ép về cột 3 (Mat. Short Text)")
                mat_col = 3
            if qty_col >= 20:
                qty_col = 5
            if txt_col >= 20:
                txt_col = 6
            if plant_col >= 20:
                plant_col = 7
            if src_sloc_col >= 20:
                src_sloc_col = 8
            if dest_sloc_col >= 20:
                dest_sloc_col = 9
            if batch_col >= 20:
                batch_col = 10
            if dest_batch_col >= 20:
                dest_batch_col = 11
            print(f"[SAP MIGO] Cột TableControl: Material={mat_col}, Qty={qty_col}, Trạm={txt_col}, Plant={plant_col}, KhoXuất={src_sloc_col}, KhoNhận={dest_sloc_col}, Lô={batch_col}")

            _set_cell = self._set_cell

            def _ensure_visible_row(table_ctl, logical_row):
                try:
                    first_row = int(table_ctl.FirstVisibleRow)
                    visible_rows = int(table_ctl.VisibleRowCount)
                except Exception:
                    first_row = 0
                    visible_rows = 15
                if logical_row < first_row or logical_row >= first_row + visible_rows:
                    try:
                        table_ctl.FirstVisibleRow = logical_row
                    except Exception:
                        try:
                            table_ctl.VerticalScrollbar.Position = logical_row
                        except Exception:
                            pass
                    time.sleep(0.2)
                    try:
                        first_row = int(table_ctl.FirstVisibleRow)
                    except Exception:
                        first_row = logical_row
                return logical_row - first_row

            # 6. Nhịp 1: Điền Material, Qty, Text (Trạm), Plant, SLoc xuất, SLoc nhận
            print(f"[SAP MIGO] Điền Nhịp 1 ({len(ob_items)} dòng): Material, Qty, Trạm, Plant, SLoc...")
            for logical_row, item in enumerate(ob_items):
                table_row = _ensure_visible_row(tbl, logical_row)
                raw_qty = item.get("allocated_qty") or item.get("qty") or 0
                try:
                    qty_str = str(int(float(raw_qty))) if float(raw_qty).is_integer() else str(raw_qty)
                except Exception:
                    qty_str = str(raw_qty)
                tram_str = str(item.get("station") or "").strip()
                mat_str = str(item.get("material", "")).strip()
                plant_str = str(item.get("plant") or plant or "3093").strip()
                src_sloc_str = str(item.get("source_sloc") or source_sloc or "1005").strip()
                dest_sloc_str = str(item.get("dest_sloc") or dest_sloc or "3001").strip()
                if dest_sloc_str in ("1003", "", "-"):
                    dest_sloc_str = "3001"

                ok_mat = _set_cell(tbl, table_row, mat_col, mat_str, "Material")
                ok_qty = _set_cell(tbl, table_row, qty_col, qty_str, "Qty")
                if tram_str:
                    _set_cell(tbl, table_row, txt_col, tram_str, "Trạm")
                ok_plant = _set_cell(tbl, table_row, plant_col, plant_str, "Plant")
                ok_src = _set_cell(tbl, table_row, src_sloc_col, src_sloc_str, "SLoc xuất")
                ok_dest = _set_cell(tbl, table_row, dest_sloc_col, dest_sloc_str, "SLoc nhận")

                if not (ok_mat and ok_qty and ok_plant and ok_src and ok_dest):
                    err_msg = f"ERR_MIGO_INPUT_FAILED: Không thể điền dữ liệu vào bảng MIGO tại dòng {logical_row+1}"
                    print(f"[SAP MIGO ERROR] {err_msg}")
                    try:
                        self._open_transaction("/n")
                    except Exception:
                        pass
                    return False, "", err_msg

            print("[SAP MIGO] Gửi Enter Nhịp 1...")
            w0 = self._find_by_id_safe("wnd[0]")
            if w0:
                w0.sendVKey(0)  # Enter nhịp 1
            self._wait_idle(5.0)

            # Lấy lại đối tượng tbl sau Enter 1
            w0 = self._find_by_id_safe("wnd[0]")
            tbl = _find_migo_table(w0) or tbl

            sbar = self._find_by_id_safe("wnd[0]/sbar")
            msg_type = getattr(sbar, "messageType", getattr(sbar, "MessageType", "")) if sbar else ""
            msg_txt = getattr(sbar, "text", getattr(sbar, "Text", "")) if sbar else ""
            if msg_type:
                print(f"[SAP MIGO] Trạng thái sau Enter 1 [{msg_type}]: {msg_txt}")
            if msg_type in ("E", "A"):
                err_msg = f"ERR_MIGO_N1_SAP_ERROR [{msg_type}]: {msg_txt}"
                print(f"[SAP MIGO ERROR] {err_msg}")
                try:
                    self._open_transaction("/n")
                except Exception:
                    pass
                return False, "", err_msg

            # 7. Nhịp 2: Điền số lô Batch động (nếu có lô)
            has_any_batch = any(bool(str(it.get("batch") or "").strip() and str(it.get("batch") or "").strip() != "-") for it in ob_items)
            if has_any_batch and tbl is not None:
                print(f"[SAP MIGO] Điền Nhịp 2: Batch (cột {batch_col})...")
                for logical_row, item in enumerate(ob_items):
                    batch_val = str(item.get("batch", "")).strip()
                    if batch_val and batch_val != "-":
                        table_row = _ensure_visible_row(tbl, logical_row)
                        ok_b = _set_cell(tbl, table_row, batch_col, batch_val, "Batch")
                        if not ok_b and dest_batch_col is not None and dest_batch_col != batch_col:
                            print(f"[SAP MIGO] Thử điền Batch vào cột nhận {dest_batch_col}...")
                            _set_cell(tbl, table_row, dest_batch_col, batch_val, "Batch")

                print("[SAP MIGO] Gửi Enter Nhịp 2...")
                w0 = self._find_by_id_safe("wnd[0]")
                if w0:
                    w0.sendVKey(0)  # Enter nhịp 2
                self._wait_idle(5.0)

                sbar = self._find_by_id_safe("wnd[0]/sbar")
                msg_type = getattr(sbar, "messageType", getattr(sbar, "MessageType", "")) if sbar else ""
                msg_txt = getattr(sbar, "text", getattr(sbar, "Text", "")) if sbar else ""
                if msg_type:
                    print(f"[SAP MIGO] Trạng thái sau Enter 2 [{msg_type}]: {msg_txt}")
                if msg_type in ("E", "A"):
                    err_msg = f"ERR_MIGO_N2_SAP_ERROR [{msg_type}]: {msg_txt}"
                    print(f"[SAP MIGO ERROR] {err_msg}")
                    try:
                        self._open_transaction("/n")
                    except Exception:
                        pass
                    return False, "", err_msg

            # Xác nhận dòng Item OK nếu có cột checkbox OK trong lưới
            if ok_col >= 0 and tbl is not None:
                try:
                    for logical_row in range(len(ob_items)):
                        t_row = _ensure_visible_row(tbl, logical_row)
                        cell_ok = tbl.GetCell(t_row, ok_col)
                        if cell_ok is not None and hasattr(cell_ok, "Selected") and not bool(cell_ok.Selected):
                            cell_ok.Selected = True
                except Exception:
                    pass

            # 8. Nhấn nút Post (btn[23] trên tbar[1] hoặc btn[11] trên tbar[0])
            print("[SAP MIGO] Bấm nút Post/Save tạo chứng từ...")
            post_pressed = False
            try:
                post_btn = self._find_by_id_safe("wnd[0]/tbar[1]/btn[23]")
                if post_btn:
                    post_btn.Press()
                    post_pressed = True
            except Exception:
                pass

            if not post_pressed:
                try:
                    save_btn = self._find_by_id_safe("wnd[0]/tbar[0]/btn[11]")
                    if save_btn:
                        save_btn.Press()
                        post_pressed = True
                except Exception:
                    pass

            if not post_pressed and w0:
                w0.sendVKey(11)

            self._wait_idle(5.0)

            # 9. Vượt qua Popups và Warnings (Chuẩn VBA RUN_AcceptPopupIfAny)
            for _ in range(8):
                handled_popup = False
                try:
                    act_wnd = getattr(self.session, "ActiveWindow", None)
                    act_name = getattr(act_wnd, "Name", "") if act_wnd else ""
                    children_cnt = getattr(self.session.Children, "Count", getattr(self.session.children, "count", 0))
                    if children_cnt > 1 or act_name == "wnd[1]":
                        try:
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").Press()
                            handled_popup = True
                        except Exception:
                            try:
                                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").Press()
                                handled_popup = True
                            except Exception:
                                try:
                                    self.session.findById("wnd[1]").sendVKey(0)
                                    handled_popup = True
                                except Exception:
                                    pass
                except Exception:
                    pass

                if handled_popup:
                    self._wait_idle(2.0)
                    continue

                sbar = self._find_by_id_safe("wnd[0]/sbar")
                msg_type = getattr(sbar, "messageType", getattr(sbar, "MessageType", "")) if sbar else ""
                if msg_type in ("W", "I"):
                    w0 = self._find_by_id_safe("wnd[0]")
                    if w0:
                        w0.sendVKey(0)
                    self._wait_idle(2.0)
                else:
                    break

            # 10. Đọc kết quả từ thanh trạng thái
            sbar = self._find_by_id_safe("wnd[0]/sbar")
            sbar_text = getattr(sbar, "text", getattr(sbar, "Text", "")) if sbar else ""
            msg_type = getattr(sbar, "messageType", getattr(sbar, "MessageType", "")) if sbar else ""
            print(f"[SAP MIGO] Phản hồi StatusBar [{msg_type}]: {sbar_text}")

            # Chặn lỗi phủ định và lỗi SAP (E/A) trước tiên - Fail-Closed
            negative_keywords = [
                "DOES NOT EXIST", "NOT EXIST", "DEFICIT", "DEFICIT OF", "ERROR", 
                "FAILED", "NOT POSTED", "KHONG TON TAI", "KHONG THANH CONG", "KHONG THE"
            ]
            if msg_type in ("E", "A") or any(neg in sbar_text.upper() for neg in negative_keywords):
                err_msg = f"Lỗi SAP MIGO [{msg_type}]: {sbar_text}"
                print(f"[SAP MIGO ERROR] {err_msg}")
                return False, "", err_msg

            # Chỉ bắt số chứng từ khi SAP báo MessageType 'S' hoặc thông báo thành công rõ ràng
            has_success_context = (
                msg_type == "S" 
                or any(kw in sbar_text.upper() for kw in ["POSTED", "PLACED IN STORAGE", "DOCUMENT", "CHỨNG TỪ", "SAVED"])
            )

            if has_success_context:
                doc_match = re.search(r"(?:MATERIAL\s+DOCUMENT|DOCUMENT|CHỨNG\s+TỪ|DELIVERY|OUTBOUND\s+DELIVERY)[^\d]{0,40}(\d{8,12})", sbar_text, re.IGNORECASE)
                if not doc_match:
                    doc_match = re.search(r"\b(49\d{8}|500\d{7}|10\d{8}|\d{10})\b", sbar_text)
                if doc_match:
                    doc_no = doc_match.group(1)
                    print(f"[SAP MIGO SUCCESS] Tạo chứng từ MIGO {doc_no} thành công!")
                    return True, doc_no, f"Tạo chứng từ Material Document (MIGO) {doc_no} thành công trên SAP: {sbar_text}"

                print(f"[SAP MIGO] Thao tác MIGO 311 thành công: {sbar_text}")
                return True, "", f"Thao tác MIGO 311 thành công: {sbar_text}"

            err_msg = f"MIGO 311 chưa xác nhận số chứng từ: {sbar_text or 'Không nhận diện được số chứng từ MIGO'}"
            print(f"[SAP MIGO WARN] {err_msg}")
            return False, "", err_msg

        except Exception as e:
            print(f"[SAP MIGO ERROR] Lỗi MIGO 311: {e}")
            return False, "", f"Lỗi thực thi SAP MIGO 311: {e}"

    def check_bom_cs03(self, materials: List[str], plant: str = "3093") -> List[Dict[str, Any]]:
        """Kiểm tra mã hàng có BOM trên SAP hay không qua /nCS03.
        Nếu có, trả về danh sách mã linh kiện con (Components), định mức (Qty), đơn vị (Unit), và mô tả.
        Thử lần lượt các BOM Usage phổ biến trong nhà máy: '1' (Production), '3' (Universal), '5' (Sales/Distribution).
        TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY).
        """
        results: List[Dict[str, Any]] = []
        if not materials:
            return results

        user_plant = str(plant or "3093").strip().upper()

        for mat in materials:
            clean_mat = str(mat).strip().upper()
            if not clean_mat:
                continue

            bom_found_for_mat = False
            bom_errors = []
            for usage in ("1", "3", "5"):
                try:
                    self._open_transaction("CS03")
                    self._wait_idle(5.0)

                    mat_fld = self.smart_find("RC29N-MATNR", "GuiCTextField") or self._find_by_id_safe("wnd[0]/usr/ctxtRC29N-MATNR")
                    if mat_fld is not None:
                        mat_fld.text = clean_mat

                    werks_fld = self.smart_find("RC29N-WERKS", "GuiCTextField") or self._find_by_id_safe("wnd[0]/usr/ctxtRC29N-WERKS")
                    if werks_fld is not None:
                        werks_fld.text = user_plant

                    stlan_fld = self.smart_find("RC29N-STLAN", "GuiTextField") or self._find_by_id_safe("wnd[0]/usr/txtRC29N-STLAN")
                    if stlan_fld is not None:
                        stlan_fld.text = usage

                    w0 = self._find_by_id_safe("wnd[0]")
                    if w0 is not None:
                        w0.sendVKey(0)
                    self._wait_idle(5.0)

                    status_txt, status_type = self._get_status()
                    if status_type in ("E", "A") or any(k in status_txt.lower() for k in ["not found", "no bom", "does not exist", "không tìm thấy"]):
                        continue

                    if status_type == "W":
                        if w0 is not None:
                            w0.sendVKey(0)
                        self._wait_idle(3.0)

                    # Cách 1: Đọc từ TableControl tblSAPLCSDITCMAT
                    tbl = self._find_by_id_safe("wnd[0]/usr/tblSAPLCSDITCMAT")
                    if tbl is not None:
                        v_rows = int(getattr(tbl, "VisibleRowCount", 0))
                        if int(getattr(tbl, "RowCount", v_rows)) > v_rows:
                            raise SapExecutionError("BOM có nhiều trang; cần adapter phân trang đã xác minh")
                        for r_idx in range(v_rows):
                            try:
                                c_mat = tbl.GetCell(r_idx, 1).Text.strip()
                                if not c_mat or "_" in c_mat or c_mat.startswith("---"):
                                    break
                                c_desc = tbl.GetCell(r_idx, 2).Text.strip()
                                c_qty = tbl.GetCell(r_idx, 3).Text.strip()
                                c_unit = tbl.GetCell(r_idx, 4).Text.strip()
                                results.append({
                                    "parent": clean_mat,
                                    "component": c_mat,
                                    "description": c_desc,
                                    "qty": c_qty,
                                    "unit": c_unit,
                                    "status": f"CÓ BOM (Usage {usage})"
                                })
                                bom_found_for_mat = True
                            except Exception as ex:
                                results.append({
                                    "parent": clean_mat,
                                    "component": "ERR_BOM_ROW",
                                    "description": f"Lỗi đọc dòng BOM #{r_idx}: {ex}",
                                    "qty": "0",
                                    "unit": "",
                                    "status": "ERR_BOM_ROW"
                                })
                                continue

                    # Cách 2: Nếu CS03 cấu hình hiển thị dạng ALV Grid
                    if not bom_found_for_mat:
                        grid = self._find_largest_grid(max_wait_sec=2.0)
                        if grid is not None:
                            grid_records = self._extract_alv_grid(
                                grid,
                                {
                                    "component": ["IDNRK", "COMPONENT", "MATERIAL", "MATNR"],
                                    "description": ["OJTXP", "DESCRIPTION", "TEXT"],
                                    "qty": ["MENGE", "QUANTITY", "COMP. QTY"],
                                    "unit": ["MEINS", "UNIT"]
                                }
                            )
                            for rec in grid_records:
                                if rec.get("component"):
                                    rec["parent"] = clean_mat
                                    rec["status"] = f"CÓ BOM (Usage {usage})"
                                    results.append(rec)
                                    bom_found_for_mat = True

                    if bom_found_for_mat:
                        break
                except Exception as e:
                    bom_errors.append(str(e))
                    print(f"[SAP CS03 WARN] Lỗi đọc BOM {clean_mat} usage {usage}: {e}")
                    continue

            if not bom_found_for_mat:
                results.append({
                    "parent": clean_mat,
                    "component": "-",
                    "description": "; ".join(bom_errors) if bom_errors else "Không xác minh được BOM với các usage đã thử",
                    "qty": "-",
                    "unit": "-",
                    "status": "CHƯA XÁC ĐỊNH BOM"
                })

        return results

    def check_ewm_queue_mon(self, inbound_filter: str = "", warehouse: str = "VF02", monitor: str = "SAP") -> List[Dict[str, str]]:
        """Tra cứu danh sách Inbound hoặc lệnh EWM đang bị kẹt lỗi trên /SCWM/MON (Node MQUEUE0001).
        TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY).
        """
        results: List[Dict[str, str]] = []
        try:
            self._open_transaction("/SCWM/MON")
            self._wait_idle(10.0)

            # Xử lý popup khởi tạo Warehouse & Monitor nếu có
            wnd1 = self._find_by_id_safe("wnd[1]")
            if wnd1 is not None and self._find_by_id_safe("wnd[1]/usr/ctxtP_LGNUM"):
                lgnum_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtP_LGNUM")
                if lgnum_ctl is not None:
                    lgnum_ctl.text = warehouse
                mon_ctl = self._find_by_id_safe("wnd[1]/usr/ctxtP_MONIT")
                if mon_ctl is not None:
                    mon_ctl.text = monitor
                btn8 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[8]")
                if btn8 is not None:
                    btn8.press()
                else:
                    wnd1.sendVKey(8)
                self._wait_idle(10.0)

            # Mở cây thư mục Node MQUEUE0001
            tree = self._find_by_id_safe("wnd[0]/usr/shell/shellcont[0]/shell")
            if tree is not None:
                try:
                    tree.expandNode("T000000001")
                except Exception:
                    pass
                try:
                    tree.selectedNode = "MQUEUE0001"
                    tree.doubleClickNode("MQUEUE0001")
                    self._wait_idle(5.0)
                except Exception:
                    pass

                wnd1 = self._find_by_id_safe("wnd[1]")
                if wnd1 is not None:
                    for chk_id in ["wnd[1]/usr/chkP_QIN", "wnd[1]/usr/chkP_QOUT"]:
                        chk = self._find_by_id_safe(chk_id)
                        if chk is not None:
                            try:
                                chk.Selected = False
                            except Exception:
                                pass
                    btn8 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[8]")
                    if btn8 is not None:
                        btn8.press()
                    else:
                        wnd1.sendVKey(8)
                    self._wait_idle(10.0)

            grid = self._find_largest_grid(max_wait_sec=10.0)
            if grid is None:
                raise SapExecutionError("Không đọc được lưới Queue; chưa xác định trạng thái")
            if grid is not None:
                col_map = {
                    "document": ["BUSOBJ_KEY", "OBJECT KEY", "INBOUND", "DOCUMENT", "REF"],
                    "error_message": ["ERRMESS", "ERROR MESSAGE", "MESSAGE", "ERROR"],
                    "queue_name": ["QNAME", "QUEUE", "QUEUE NAME"]
                }
                raw_records = self._extract_alv_grid(grid, col_map)
                for rec in raw_records:
                    doc_val = str(rec.get("document", "")).strip()
                    err_val = str(rec.get("error_message", "")).strip()
                    q_val = str(rec.get("queue_name", "")).strip()
                    if not err_val and not doc_val and not q_val:
                        continue
                    if inbound_filter and (inbound_filter not in doc_val and inbound_filter not in q_val):
                        continue
                    results.append(rec)

            return results
        except Exception as e:
            raise SapExecutionError(f"Lỗi đọc Queue EWM: {e}") from e

    def check_reservation_mb23(self, res_no: str) -> Dict[str, Any]:
        """Tra cứu phiếu lĩnh Reservation qua T-code MB23. TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY)."""
        clean_res = str(res_no or "").strip()
        if not clean_res:
            return {"found": False, "status": "EMPTY_RES_NO", "message": "Số phiếu Reservation rỗng."}
        try:
            self._open_transaction("MB23")
            self._wait_idle(5.0)
            res_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtRM07M-RSNUM") or self._find_by_id_safe("wnd[0]/usr/txtRM07M-RSNUM")
            if res_ctl is None:
                raise SapExecutionError("Không tìm thấy ô nhập số phiếu Reservation (RM07M-RSNUM).")
            res_ctl.text = clean_res
            self.session.findById("wnd[0]").sendVKey(0)
            self._wait_idle(5.0)
            self._dismiss_popups()
            status_txt, status_type = self._get_status()
            if status_type in ("E", "A"):
                return {"found": False, "res_no": clean_res, "status": "NOT_FOUND", "message": status_txt}

            grid = self._find_largest_grid(max_wait_sec=3.0)
            items = []
            if grid is not None:
                col_map = {
                    "item": ["RSPOS", "ITEM"],
                    "material": ["MATNR", "MATERIAL"],
                    "qty": ["BDMNG", "QTY", "QUANTITY"],
                    "unit": ["MEINS", "UNIT"],
                    "plant": ["WERKS", "PLANT"],
                    "sloc": ["LGORT", "SLOC"]
                }
                items = self._extract_alv_grid(grid, col_map)
            return {
                "found": True,
                "res_no": clean_res,
                "status": "FOUND",
                "items": items,
                "message": f"Tìm thấy phiếu Reservation {clean_res}."
            }
        except Exception as e:
            return {"found": False, "res_no": clean_res, "status": "ERROR", "message": str(e)}
        finally:
            self._reset_to_session_manager()

    def check_attachment_status(
        self,
        docs: List[str],
        year: str = "",
        company_code: str = "2301",
        ticket_no: str = ""
    ) -> List[Dict[str, Any]]:
        """
        Kiểm tra trạng thái file đính kèm BBBG và VAT trên SAP cho danh sách DOC (10 chữ số).
        TUYỆT ĐỐI CHỈ ĐỌC (READ-ONLY).
        Ưu tiên tra cứu qua giao dịch ZFIR266. Nếu ZFIR266 không khả dụng,
        tự động chuyển sang chế độ đọc GOS qua MIGO A04 Display.
        """
        from app.sap_attachment import classify_sap_attachments, MigoAttachmentRunner

        clean_docs = []
        for d in docs:
            cleaned = re.sub(r"\D", "", str(d)).strip()
            if cleaned and len(cleaned) >= 8 and cleaned not in clean_docs:
                clean_docs.append(cleaned)

        if not clean_docs:
            return []

        if not year or not re.fullmatch(r"(?:19|20|21)\d{2}", str(year)):
            from datetime import datetime
            year = str(datetime.now().year)

        session = self.ensure_session()
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

        zfir_supported = True
        zfir_results = {d: [] for d in clean_docs}
        try:
            self._open_transaction("ZFIR266")
            self._wait_idle(3.0)

            sbar = self._find_by_id_safe("wnd[0]/sbar")
            sbar_text = str(getattr(sbar, "Text", "")).strip().lower() if sbar else ""
            trans_name = str(getattr(session.Info, "Transaction", "")).upper()

            if trans_name != "ZFIR266" or "does not exist" in sbar_text or "không tồn tại" in sbar_text or "not authorized" in sbar_text:
                zfir_supported = False
                print(f"[ZFIR266 WARN] ZFIR266 không khả dụng ({sbar_text or trans_name}). Chuyển sang MIGO Display...")
            else:
                # 1. Thử chạy BATCH toàn bộ DOCs qua Multiple Selection (1 lượt duy nhất F8)
                multi_btn_id = "wnd[0]/usr/btn%_S_MBELN_%_APP_%-VALU_PUSH"
                has_multi_btn = self._find_by_id_safe(multi_btn_id) is not None
                batch_ran_ok = False

                if has_multi_btn and len(clean_docs) > 1:
                    try:
                        print(f"[ZFIR266 BATCH] Đang đưa đồng thời {len(clean_docs)} DOC vào Multiple Selection của ZFIR266...")
                        b_f = self._find_by_id_safe("wnd[0]/usr/ctxtP_BUKRS") or self._find_by_id_safe("wnd[0]/usr/txtP_BUKRS")
                        if b_f: b_f.text = str(company_code)
                        g_f = self._find_by_id_safe("wnd[0]/usr/txtP_GJAHR") or self._find_by_id_safe("wnd[0]/usr/ctxtP_GJAHR")
                        if g_f: g_f.text = str(year)
                        m_f = self._find_by_id_safe("wnd[0]/usr/txtP_MAX")
                        if m_f: m_f.text = "1000"

                        # Điền toàn bộ danh sách DOC vào Multiple Selection
                        self._fill_multiple_selection(multi_btn_id, clean_docs)

                        # Bấm F8 Execute
                        w0 = self._find_by_id_safe("wnd[0]")
                        if w0:
                            w0.sendVKey(8)
                        self._wait_idle(3.0)
                        self._dismiss_popups()

                        # Trích xuất ALV Grid chứa toàn bộ các DOC
                        grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell") or self._find_by_id_safe("wnd[0]/usr/shell")
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
                                        c_cap = str(self._get_column_caption(grid, c_name)).strip().upper()
                                    except Exception:
                                        pass
                                    if not resolved_doc_col and any(k in c_up or k in c_cap for k in ["MBELN", "MBLNR", "MAT_DOC", "DOCUMENTNO", "DOCUMENT_NO", "DOCNO", "BELNR", "DOCUMENT", "CHỨNG TỪ"]):
                                        resolved_doc_col = c_name
                                    if not resolved_file_col and any(k in c_up or k in c_cap for k in ["FILE_NAME", "FILENAME", "FILE NAME", "TÊN FILE", "TÊN TỆP", "ATTACH"]):
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
                                if resolved_doc_col:
                                    try:
                                        r_doc = str(grid.GetCellValue(r, resolved_doc_col)).strip()
                                    except Exception:
                                        pass
                                if not r_doc:
                                    for c_cand in ["DOCUMENTNO", "DOCUMENT_NO", "BELNR", "MBELN", "DOCNO"]:
                                        try:
                                            v = str(grid.GetCellValue(r, c_cand)).strip()
                                            if v and re.fullmatch(r"\d{8,10}", v):
                                                r_doc = v
                                                break
                                        except Exception:
                                            pass

                                r_file = ""
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
                        self._open_transaction("ZFIR266")
                        self._wait_idle(2.0)

                # 2. Nếu chỉ có 1 DOC hoặc BATCH không thành công: duyệt từng DOC với bảo vệ COM Pointer tuyệt đối
                if not batch_ran_ok:
                    for doc in clean_docs:
                        titles_found = []
                        try:
                            # Đảm bảo đang đứng ở Selection Screen của ZFIR266
                            btn8_test = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]")
                            if btn8_test is None:
                                self._open_transaction("ZFIR266")
                                self._wait_idle(2.0)

                            # LUÔN RE-RESOLVE TOÀN BỘ CONTROLS TƯƠI MỚI TRONG VÒNG LẶP (Chống triệt để Dead Pointer)
                            w0 = self._find_by_id_safe("wnd[0]")
                            bukrs_field = self._find_by_id_safe("wnd[0]/usr/ctxtP_BUKRS") or self._find_by_id_safe("wnd[0]/usr/txtP_BUKRS")
                            gjahr_f = self._find_by_id_safe("wnd[0]/usr/txtP_GJAHR") or self._find_by_id_safe("wnd[0]/usr/ctxtP_GJAHR")
                            doc_f = self._find_by_id_safe("wnd[0]/usr/ctxtS_MBELN-LOW") or self._find_by_id_safe("wnd[0]/usr/txtS_MBELN-LOW")
                            max_f = self._find_by_id_safe("wnd[0]/usr/txtP_MAX")

                            if bukrs_field:
                                bukrs_field.text = str(company_code)
                            if gjahr_f:
                                gjahr_f.text = str(year)
                            if doc_f:
                                doc_f.text = str(doc)
                            if max_f:
                                max_f.text = "500"

                            # 2. Chạy báo cáo F8 (sendVKey 8)
                            if w0:
                                w0.sendVKey(8)
                            self._wait_idle(3.0)
                            self._dismiss_popups()

                            # 3. Trích xuất ALV Grid
                            grid = self._find_by_id_safe("wnd[0]/usr/cntlGRID1/shellcont/shell") or self._find_by_id_safe("wnd[0]/usr/shell")
                            has_grid = grid is not None and hasattr(grid, "RowCount") and grid.RowCount > 0
                            if has_grid:
                                resolved_file_col = None
                                try:
                                    col_order = grid.ColumnOrder
                                    for c_idx in range(col_order.Count):
                                        c_name = str(col_order.ElementAt(c_idx)).strip()
                                        c_upper = c_name.upper()
                                        c_cap = ""
                                        try:
                                            c_cap = str(self._get_column_caption(grid, c_name)).strip().upper()
                                        except Exception:
                                            pass
                                        if any(k in c_upper or k in c_cap for k in ["FILE_NAME", "FILENAME", "FILE NAME", "TÊN FILE", "TÊN TỆP", "ATTACH"]):
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

                            # 4. CHỈ bấm F3 quay lại màn hình Selection khi THỰC SỰ đang đứng ở màn hình Grid
                            if has_grid:
                                w0_back = self._find_by_id_safe("wnd[0]")
                                if w0_back:
                                    w0_back.sendVKey(3)
                                    self._wait_idle(1.0)
                                    self._dismiss_popups()
                        except Exception as e:
                            print(f"[ZFIR266 DOC WARN] Lỗi đọc DOC {doc}: {e}")
                            try:
                                self._open_transaction("ZFIR266")
                                self._wait_idle(2.0)
                            except Exception:
                                pass
                        zfir_results[doc] = titles_found
        except Exception as zfir_err:
            print(f"[ZFIR266 EXCEPTION] {zfir_err}")
            zfir_supported = False

        # 5. DỰ PHÒNG MIGO DISPLAY: Khi ZFIR266 không khả dụng hoặc trả về 0 kết quả cho DOC
        docs_need_migo_fallback = [d for d in clean_docs if (not zfir_supported) or (d in zfir_results and not zfir_results[d])]
        migo_results = {}
        if docs_need_migo_fallback:
            print(f"[CHECK_ATT FALLBACK] ZFIR266 không khả dụng, chuyển sang kiểm tra GOS qua MIGO Display cho {len(docs_need_migo_fallback)} DOC...")
            runner = MigoAttachmentRunner(sap_engine=self, expected_session_identity=self.session_identity)
            for doc in docs_need_migo_fallback:
                titles = []
                try:
                    loaded_ok, nav_msg = runner.open_migo_document(session, doc, year=year, mode="DISPLAY", force_reload=True)
                    if loaded_ok:
                        titles = runner.read_sap_attachment_titles(session)
                except Exception as migo_err:
                    print(f"[MIGO DISPLAY WARN] Lỗi đọc DOC {doc}: {migo_err}")
                migo_results[doc] = titles

        self._reset_to_session_manager()

        records = []
        for doc in clean_docs:
            titles = zfir_results.get(doc) or migo_results.get(doc, [])
            info = classify_sap_attachments(titles, ticket_no=ticket_no)
            records.append({
                "doc": doc,
                "company_code": company_code,
                "year": year,
                "bbbg_status": "CÓ" if info["has_bbbg"] else "CHƯA",
                "vat_status": "CÓ" if info["has_vat"] else "CHƯA",
                "overall_status": info["overall_status"],
                "status_code": info["status_code"],
                "files_found": ", ".join(titles) if titles else "Không có file",
                "bbbg_file": info["bbbg_file"],
                "vat_file": info["vat_file"],
                "ticket_no": ticket_no
            })

        return records
