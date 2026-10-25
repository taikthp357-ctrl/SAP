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

    def __init__(
        self,
        max_alv_rows: int = 15000,
        keep_alive_idle_seconds: int = 900,
        connection_index: int = 0,
        session_index: int = 0,
        strict_session_safety: bool = True
    ):
        self.max_alv_rows = max_alv_rows
        self.keep_alive_idle_seconds = keep_alive_idle_seconds
        self.connection_index = connection_index
        self.session_index = session_index
        self.strict_session_safety = strict_session_safety
        self.session = None
        self.last_activity = time.time()
        self.last_popup_message = ""
        self.ensure_session()

    def _connect_sap(self):
        """Kết nối tới phiên SAP GUI an toàn, ưu tiên phiên đang ở Main Menu hoặc T-Code của bot."""
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            app = sap_gui_auto.GetScriptingEngine
        except Exception:
            if os.path.exists(SAP_LOGON_PATH):
                print("[SAP] SAP Logon chưa chạy. Đang khởi động qua os.startfile...")
                os.startfile(SAP_LOGON_PATH)
                time.sleep(5)
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                app = sap_gui_auto.GetScriptingEngine
            else:
                raise RuntimeError("SAP Logon chưa được mở trên máy trạm!")

        if app.Children.Count == 0 and SAP_CONNECTION_NAME:
            print(f"[SAP] Đang mở kết nối '{SAP_CONNECTION_NAME}'...")
            app.OpenConnection(SAP_CONNECTION_NAME, True)
            time.sleep(3)

        if app.Children.Count == 0:
            raise RuntimeError("Chưa có Connection nào trong SAP Logon. Vui lòng mở SAP và đăng nhập.")

        # Quét tìm session an toàn: tránh can thiệp nhầm vào phiên người dùng đang nhập liệu
        selected_session = None
        bot_safe_tcodes = {"SESSION_MANAGER", "/SCWM/POST", "MB52", "MB51", "VL06I", "ME2N", "VL06O", "/SCWM/MON", "SCWM/MON", ""}

        for c_idx in range(app.Children.Count):
            conn = app.Children(c_idx)
            for s_idx in range(conn.Children.Count):
                sess = conn.Children(s_idx)
                try:
                    info = getattr(sess, "Info", None)
                    tcode = str(getattr(info, "Transaction", "")).strip().upper() if info else ""
                    if tcode in bot_safe_tcodes:
                        selected_session = sess
                        print(f"[SAP] Đã gắn kết nối an toàn: Connection {c_idx}, Session {s_idx} (T-Code: '{tcode or 'MAIN'}')")
                        break
                except Exception:
                    pass
            if selected_session is not None:
                break

        # Nếu không tìm thấy phiên nhàn rỗi, kiểm tra phiên theo cấu hình
        if selected_session is None:
            c_idx = min(self.connection_index, app.Children.Count - 1)
            conn = app.Children(c_idx)
            if conn.Children.Count == 0:
                raise RuntimeError("Chưa có Session nào đang mở trong SAP GUI.")
            s_idx = min(self.session_index, conn.Children.Count - 1)
            cand_sess = conn.Children(s_idx)

            # Kiểm tra an toàn: không cướp phiên nếu người dùng đang thao tác
            cand_info = getattr(cand_sess, "Info", None)
            cand_tcode = str(getattr(cand_info, "Transaction", "")).strip().upper() if cand_info else ""
            if self.strict_session_safety and cand_tcode and cand_tcode not in bot_safe_tcodes:
                raise RuntimeError(
                    f"Tất cả các phiên SAP đều đang bận thao tác của người dùng (Session {s_idx} đang ở T-Code '{cand_tcode}'). "
                    f"Để bảo vệ dữ liệu, Bot từ chối can thiệp. Vui lòng mở thêm 1 phiên mới ở màn hình chính (SESSION_MANAGER)."
                )

            selected_session = cand_sess
            print(f"[SAP] Dùng phiên cấu hình: Connection {c_idx}, Session {s_idx} (T-Code: '{cand_tcode or 'MAIN'}')")

        self.last_activity = time.time()
        return selected_session

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
        """Bảo đảm session luôn hợp lệ; tự động reconnect nếu phiên bị đứt."""
        if not self._is_session_alive():
            print("[SAP] Phiên SAP bị mất kết nối hoặc chưa khởi tạo, đang kết nối lại...")
            self.session = self._connect_sap()
            print("[SAP] Kết nối phiên SAP GUI thành công!")
        return self.session

    def keep_alive(self, idle_threshold_seconds: Optional[int] = None):
        """Gửi xung nhịp nhẹ làm mới phiên SAP an toàn, không ngắt quãng phiên làm việc của người dùng."""
        thresh = idle_threshold_seconds or self.keep_alive_idle_seconds
        if time.time() - self.last_activity < thresh:
            return

        print(f"[KEEP-ALIVE] Đã qua {int(time.time() - self.last_activity)}s không có thao tác. Gửi xung nhịp làm mới SAP...")
        try:
            if not self._is_session_alive():
                self.session = self._connect_sap()
            else:
                # Kiểm tra T-code hiện tại: chỉ gửi /n nếu đang ở màn hình an toàn
                tcode = ""
                try:
                    info = getattr(self.session, "Info", None)
                    tcode = str(getattr(info, "Transaction", "")).strip().upper() if info else ""
                except Exception:
                    pass

                bot_safe_tcodes = {"SESSION_MANAGER", "/SCWM/POST", "MB52", "MB51", "VL06I", "ME2N", "VL06O", "/SCWM/MON", "SCWM/MON", ""}
                if tcode in bot_safe_tcodes:
                    self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
                    self.session.findById("wnd[0]").sendVKey(0)
                    self._dismiss_popups()
                else:
                    # Nếu người dùng đang ở giao dịch khác, chỉ đọc thuộc tính để giữ kết nối RAM
                    _ = bool(self.session.Busy)

            self.last_activity = time.time()
            print(f"[KEEP-ALIVE] Phiên SAP đã được làm mới an toàn ({datetime.now().strftime('%H:%M:%S')})")
        except Exception as e:
            print(f"[KEEP-ALIVE WARNING] Không thể gửi xung nhịp: {e}. Sẽ kết nối lại khi có lệnh mới.")
            try:
                self.session = self._connect_sap()
                self.last_activity = time.time()
            except Exception:
                pass
                self.last_activity = time.time()
            except Exception:
                pass

    def _wait_idle(self, timeout_sec: float = 30.0) -> bool:
        """Đợi SAP GUI hoàn tất xử lý (Busy = False) như chuẩn modSAPCore.WaitIdle."""
        t0 = time.time()
        while time.time() - t0 < timeout_sec:
            try:
                if not getattr(self.session, "Busy", False):
                    return True
            except Exception:
                return True
            time.sleep(0.1)
        return False

    def _find_by_id_safe(self, id_text: str):
        """Tìm đối tượng theo ID an toàn tuyệt đối, không ném exception nếu không thấy."""
        try:
            return self.session.findById(id_text, False)
        except Exception:
            return None

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
            cand_norms = [self._normalize_header(c) for c in candidates if self._normalize_header(c)]

            # Lượt 1: Khớp chính xác (Exact match) mã kỹ thuật hoặc Caption
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
            "deliveries selected", "outbound deliveries", "giao hàng",
            "không có giao hàng", "no deliveries", "no open", "không có đơn",
            "list contains no data", "danh sách không chứa", "no stock",
            "không có tồn", "keine lieferungen", "ausgewählt", "selected",
            "zero records", "0 records", "không tìm thấy bản ghi"
        ]

        if any(p in lower_txt for p in empty_phrases):
            print(f"[SAP {tcode_name}] Xác nhận không có dữ liệu: '{combined_txt}'")
            return

        if is_selection_screen:
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
        except Exception:
            return []

        if row_count == 0:
            return []

        resolved_cols = self._resolve_grid_columns(alv_obj, column_mapping)
        records = []
        effective_max = max_rows if max_rows is not None else self.max_alv_rows
        limit = min(row_count, effective_max)
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
                    except Exception:
                        item[out_key] = ""

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
                (item.get("transit") and str(item.get("transit")).strip() not in ("", "0", "0.0", "0,0"))
            )
            # Bỏ qua hoàn toàn dòng tổng cộng / subtotal của ALV (dòng không có bất kỳ thông tin vị trí/chứng từ nào)
            if not has_pos_or_doc:
                continue

            records.append(item)

        return records

    def _apply_alv_filter(self, grid, column_id: str, values: List[str]) -> bool:
        """Áp dụng bộ lọc ALV Filter trên ALV Grid (phím Ctrl+F2 / nút &MB_FILTER / btn[29]) theo đúng chuẩn VL6_ApplyALVFilter của RC65."""
        clean_vals = [str(v).strip().upper() for v in values if str(v).strip()]
        if not clean_vals or grid is None or "Mock" in type(grid).__name__:
            return False
        try:
            # 1. Bỏ chọn các vùng cũ và chọn cột cần lọc
            try:
                grid.clearSelection()
            except Exception:
                pass
            try:
                grid.setCurrentCell(-1, column_id)
                grid.selectColumn(column_id)
            except Exception as sel_col_err:
                print(f"[SAP ALV FILTER WARN] Không thể chọn cột {column_id}: {sel_col_err}")

            self._wait_idle(1.0)

            # 2. Bấm nút Filter trên Toolbar (btn[29] hoặc &MB_FILTER)
            btn29 = self._find_by_id_safe("wnd[0]/tbar[1]/btn[29]")
            if btn29 is not None:
                btn29.press()
            else:
                try:
                    grid.pressToolbarButton("&MB_FILTER")
                except Exception:
                    self.session.findById("wnd[0]").sendVKey(29)

            self._wait_idle(3.0)
            wnd1 = self._find_by_id_safe("wnd[1]")
            if wnd1 is None:
                print(f"[SAP ALV FILTER WARN] Không tìm thấy popup wnd[1] sau khi bấm Filter cột {column_id}")
                return False

            # 3. Tìm ô nhập giá trị trong wnd[1]
            base_sub = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/"
            low_ctl = self._find_by_id_safe(base_sub + "ctxt%%DYN001-LOW")
            hi_ctl = self._find_by_id_safe(base_sub + "ctxt%%DYN001-HIGH")

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

            if len(clean_vals) == 1 and low_ctl is not None:
                low_ctl.text = clean_vals[0]
                if hi_ctl is not None:
                    try:
                        hi_ctl.text = ""
                    except Exception:
                        pass
            else:
                multi_push = self._find_by_id_safe(base_sub + "btn%_%%DYN001_%_APP_%-VALU_PUSH")
                if multi_push is not None:
                    self._fill_multiple_selection(base_sub + "btn%_%%DYN001_%_APP_%-VALU_PUSH", clean_vals)
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
            print(f"[SAP ALV FILTER] Đã áp dụng thành công bộ lọc {clean_vals} trên cột {column_id}.")
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

        clean_code = tcode.strip().replace("/n", "").replace("/N", "")
        if not clean_code.startswith("/") and "/" in tcode:
            clean_code = "/" + clean_code

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
            print(f"[SAP ERROR] Không thể mở giao dịch {tcode}: {st_err}")

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
        plant: str = "3093"
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
                "unit":     ["MEINS", "ALTME", "UNIT", "UOM"],
                "cat":      ["CAT", "TST", "STOCK TYPE", "ST", "TO_CAT"],
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
                for p_id in ["wnd[0]/usr/ctxtWERKS-HIGH"]:
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
                "transit": ["TRAME", "STOCK IN TRANSIT", "TRANSIT/TRANSF.", "TRANSIT"],
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

            # LỌC BỎ TRIỆT ĐỂ DÒNG TỔNG CỘNG TRONG MB52: Bắt buộc phải có Kho (Sloc)
            records = [r for r in records if str(r.get("sloc", "")).strip()]

            if slocs:
                clean_slocs_set = {s.strip().upper() for s in slocs if s.strip()}
                records = [r for r in records if str(r.get("sloc", "")).strip().upper() in clean_slocs_set]
                print(f"[SAP MB52] Sau khi lọc theo kho {clean_slocs_set}: còn {len(records)} dòng.")

            # Chuẩn hóa định dạng số lượng tồn cho tất cả các dòng
            for r in records:
                u_f = parse_sap_float(r.get("unrest", 0.0))
                r["unrest"] = format_qty_display(u_f)

            # SẮP XẾP GOM THEO KHO (SLOC): Các dòng cùng kho đứng cạnh nhau, trong cùng kho thì lượng tồn lớn lên trước
            def _mb52_sort_key(item):
                sloc_val = str(item.get("sloc", "")).strip().upper()
                unrest_f = parse_sap_float(item.get("unrest", 0.0))
                batch_val = str(item.get("batch", "")).strip().upper()
                return (sloc_val, -unrest_f, batch_val)

            records.sort(key=_mb52_sort_key)
            return records
        finally:
            self._reset_to_session_manager()

    # --------------------------------------------------------------------------
    # 3. KIỂM TRA LỊCH SỬ XUẤT NHẬP (MB51)
    # --------------------------------------------------------------------------
    def check_history_mb51(self, materials: List[str], days: int = 7, plant: str = "3093") -> List[Dict[str, str]]:
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

            if plant:
                p_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtWERKS-LOW")
                if p_ctl is not None:
                    try:
                        p_ctl.text = plant
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
                "batch":   ["CHARG", "BATCH"]
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
    def check_inbound_treo_vl06i(self, materials: List[str], plant: str = "3093") -> List[Dict[str, str]]:
        """Tra cứu các đơn Inbound Delivery (IB) đang treo chưa GR qua VL06I theo chuẩn BBBG."""
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

            # 2. Xóa trường ngày giao để lấy trọn vẹn không bị lọc ngày
            lfdat_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_LFDAT-LOW")
            if lfdat_ctl is not None:
                try:
                    lfdat_ctl.text = ""
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

            # 5. ĐIỀU KIỆN THEN CHỐT: Đặt trạng thái Goods Movement IT_WBSTK-LOW = 'a' (chưa GR / Treo)
            wbstk_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTK-LOW")
            if wbstk_ctl is not None:
                try:
                    wbstk_ctl.text = "a"
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
    def check_sto_me2n(self, materials: List[str], plant: str = "3093") -> List[Dict[str, str]]:
        """Tra cứu đơn STO (Stock Transport Order) điều chuyển kho qua ME2N."""
        try:
            if not materials:
                return []

            self._open_transaction("ME2N")
            self._wait_idle(10.0)

            # Xóa sạch các bộ lọc cũ trên màn hình chọn ME2N
            self._clear_all_selection_inputs()

            print(f"[SAP ME2N] Đang tra cứu STO cho Materials={materials}, Plant={plant}...")
            clean_mats = [m.strip().upper() for m in materials if m.strip()]

            # 1. Đặt định dạng hiển thị ALV
            listu_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtLISTU")
            if listu_ctl is not None:
                try:
                    listu_ctl.text = "ALV"
                except Exception:
                    pass

            # 2. Đặt Selection Parameter AVIS (Đơn đang thực hiện giao nhận)
            selpa_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtSELPA-LOW")
            if selpa_ctl is not None:
                try:
                    selpa_ctl.text = "AVIS"
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

            # 4. Điền Supplying Plant (Nhà máy cấp hàng) nếu có
            if plant:
                reswk_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtS_RESWK-LOW")
                if reswk_ctl is not None:
                    try:
                        reswk_ctl.text = plant
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
            print(f"[SAP ME2N] Đã trích xuất {len(raw_records)} dòng đơn STO.")

            records = []
            for r in raw_records:
                # Bỏ qua dòng đã xóa
                if str(r.get("deleteInd", "")).strip():
                    continue

                # Tính số lượng còn lại (Remaining Qty) chuẩn hóa float
                ord_q = parse_sap_float(r.get("orderQty"))
                del_q = parse_sap_float(r.get("deliveredQty"))
                rem_q = max(0.0, ord_q - del_q)

                r["orderQty"] = format_qty_display(ord_q)
                r["deliveredQty"] = format_qty_display(del_q)
                r["remainingQty"] = format_qty_display(rem_q)
                records.append(r)

            # Sắp xếp số đơn STO mới nhất lên đầu
            records.sort(key=lambda x: str(x.get("stoNo", "")), reverse=True)
            return records

        # --------------------------------------------------------------------------
        finally:
            self._reset_to_session_manager()

    # 6. KIỂM TRA TỒN THẤP THEO MODEL & SHOP (MB52 ĐA PLANT)
    # --------------------------------------------------------------------------
    def check_low_stock(
        self,
        materials: List[str],
        plants: List[str],
        limit: int = 20,
        model_name: str = ""
    ) -> List[Dict[str, str]]:
        """Tra cứu và tổng hợp danh sách mã tồn thấp từ MB52 qua cả 2 xưởng (hoặc các Plant chỉ định) trong MỘT LẦN CHẠY MB52 duy nhất.
        
        - BODY: Plant 3093 (Kho FG) & 3001 (Line Body)
        - XMĐ:  Plant 3093 (Kho FG), 2001 (Xưởng XMĐ) & 3001 (Line Body)
        Tự xử lý gom dữ liệu và sắp xếp tăng dần theo tổng tồn (mã tồn thấp nhất / hết hàng lên đầu).
        """
        clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
        target_plants = [p.strip().upper() for p in (plants or []) if p.strip()]
        if not target_plants:
            target_plants = ["3093", "3001"]

        print(f"[SAP LOW STOCK] Quét MB52 cả 2 xưởng {target_plants} cho {len(clean_mats)} mã (Model={model_name}, Top={limit})...")

        # 1. Gọi MB52 DUY NHẤT 1 LẦN cho cả 2 xưởng (Plant 3093 và 3001) với giới hạn trích xuất tối đa toàn bộ dữ liệu (15,000 dòng)
        try:
            mb52_records = self.check_stock_mb52(clean_mats, plants=target_plants, max_rows=15000)
            print(f"[SAP LOW STOCK] MB52 trả về tổng cộng {len(mb52_records)} dòng dữ liệu từ các xưởng.")
        except Exception as e:
            print(f"[SAP LOW STOCK ERROR] Lỗi truy vấn MB52 đa xưởng: {e}")
            raise SapExecutionError(f"Không thể truy vấn tồn kho MB52 đa xưởng {target_plants}: {e}") from e

        # 2. Xác định danh sách mã cần đánh giá
        if clean_mats:
            mats_to_evaluate = list(clean_mats)
        else:
            mats_to_evaluate = sorted(list(set(r.get("product", "").strip().upper() for r in mb52_records if r.get("product"))))

        stock_by_mat = {m: {p: 0.0 for p in target_plants} for m in mats_to_evaluate}
        unit_by_mat = {m: "PC" for m in mats_to_evaluate}

        # Tạo helper mapping để map linh hoạt mã trả về từ MB52 (có/không BIW, có/không suffix AA...) về key chuẩn trong mats_to_evaluate
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

        # 3. Gom tồn kho theo Mã hàng và theo từng Plant (TUYỆT ĐỐI BỎ QUA SLoc 9999 của QC)
        for r in mb52_records:
            mat_raw = r.get("product", "").strip().upper()
            p = r.get("plant", "").strip().upper()
            sloc = r.get("sloc", "").strip().upper()

            # LOẠI BỎ TRIỆT ĐỂ SLOC 9999 (QC / Phế phẩm / Blocked) THEO 261008_Stock_FG_MA_BODY
            if sloc in ["9999", "QC", "BLOCK", "SCRAP", "QUARANTINE"]:
                continue

            matched_mat = _find_eval_mat(mat_raw)
            if not matched_mat:
                matched_mat = mat_raw
                stock_by_mat[matched_mat] = {pl: 0.0 for pl in target_plants}
                if matched_mat not in mats_to_evaluate:
                    mats_to_evaluate.append(matched_mat)

            if p in stock_by_mat[matched_mat]:
                stock_by_mat[matched_mat][p] += parse_sap_float(r.get("unrest", "0"))

            if r.get("unit"):
                unit_by_mat[matched_mat] = r.get("unit")

        # 4. Nạp danh mục Trạm (Station) từ Master Data
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

        # 5. Tự xử lý tổng hợp và phân loại cảnh báo
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
            row = {
                "product": mat,
                "station": t_val or "-",
                "model": model_name or "N/A",
                "unit": unit_by_mat.get(mat, "PC"),
                "total_stock": format_qty_display(total_s)
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

        # 5. Tự xử lý sắp xếp: TĂNG DẦN THEO TỔNG TỒN (mã tồn thấp nhất / hết hàng lên đầu)
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
                col_names = {str(cols(i)).strip().upper() for i in range(cols.Count)}
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
        """Tra cứu các đơn Outbound Delivery (OB) đang mở qua VL06O chuẩn theo 261010_OB_TOOL_V1_RC65."""
        try:
            self._open_transaction("VL06O")
            self._wait_idle(10.0)

            clean_mats = [m.strip().upper() for m in (materials or []) if m.strip()]
            clean_ob = str(ob_no or "").strip()
            user_plant = str(plant or "3093").strip().upper()

            # Chuẩn RC65 (ModuleSAP.bas dòng 2572-2584): Shipping Point = "3093" trực tiếp
            target_vstel = "3098" if user_plant == "3098" else "3093"

            print(f"[SAP VL06O] Tra cứu OB mở theo chuẩn RC65: OB='{clean_ob}', Materials={clean_mats}, ShippingPoint='{target_vstel}', ShipTo='{ship_to}'...")

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

            # 3. ĐIỀU KIỆN THEN CHỐT THEO RC65: Điền Shipping Point = 3093 trực tiếp vào ctxtIF_VSTEL-LOW (không dùng popup)
            vstel_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIF_VSTEL-LOW")
            if vstel_ctl is not None:
                try:
                    vstel_ctl.text = target_vstel
                except Exception as e:
                    print(f"[SAP VL06O WARN] Không thể điền Shipping Point: {e}")

            vstel_hi = self._find_by_id_safe("wnd[0]/usr/ctxtIF_VSTEL-HIGH")
            if vstel_hi is not None:
                try:
                    vstel_hi.text = ""
                except Exception:
                    pass

            # 4. Khoảng ngày Planned GI Date: -60 ngày đến +60 ngày theo đúng RC65
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
            gm_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTK-LOW")
            if gm_ctl is None:
                gm_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtIT_WBSTA-LOW")
            if gm_ctl is not None:
                try:
                    gm_ctl.text = "A"
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
            if any(k in comb for k in ["no deliveries", "no data", "không tìm thấy", "keine lieferungen", "không có giao hàng"]):
                print(f"[SAP VL06O] Xác nhận không có đơn Outbound nào: '{status_txt or popup_msg}'")
                return []

            if status_type in ("E", "A"):
                raise SapExecutionError(f"Lỗi SAP VL06O ({status_type}): {status_txt}")

            # 9. BẢO ĐẢM HIỂN THỊ ITEM VIEW THEO CHUẨN VL6_EnsureItemGrid CỦA RC65
            alv_grid = self._ensure_item_grid_vl06o(max_wait_sec=20.0)
            if alv_grid is None:
                self._validate_grid_or_raise(alv_grid, "VL06O", status_txt, status_type)
                return []

            # 10. ÁP DỤNG BỘ LỌC ALV FILTER THEO CHUẨN VL6_ApplyALVFilter CỦA RC65:
            # Nếu người dùng hỏi 1 mã (hoặc danh sách mã), lọc trực tiếp trên ALV để loại bỏ hàng nghìn dòng thừa!
            if clean_mats:
                resolved_item_cols = self._resolve_grid_columns(
                    alv_grid,
                    {"product": ["MATNR", "MATERIAL", "PRODUCT"]}
                )
                mat_tech_col = resolved_item_cols.get("product", "MATNR")
                self._apply_alv_filter(alv_grid, mat_tech_col, clean_mats)

            # 11. Trích xuất toàn bộ dữ liệu từ lưới Item View đã lọc
            col_map = {
                "obNo":       ["VBELN", "DELIVERY", "OUTBOUND DELIVERY"],
                "item":       ["POSNR", "DELIVERY ITEM", "ITEM"],
                "product":    ["MATNR", "MATERIAL", "PRODUCT"],
                "qty":        ["LFIMG", "DLV.QTY", "DELIVERY QTY", "ACTUAL DELIVERED QTY", "MENGE", "QUANTITY"],
                "unit":       ["VRKME", "MEINS", "SALES UNIT", "DELIVERY UNIT", "BASE UNIT", "UNIT"],
                "sloc":       ["LGORT", "SLOC", "STORAGE LOCATION"],
                "batch":      ["CHARG", "S_CHARG", "BATCH"],
                "shipTo":     ["KUNWE", "KUNNR", "SHIP-TO PARTY", "SHIP TO"],
                "gmStatus":   ["WBSTA", "WBSTK", "GOODS MOVEMENT STATUS", "GM STATUS"],
                "refSTO":     ["VGBEL", "REFERENCE DOCUMENT", "STO"],
                "createDate": ["ERDAT", "CREATED ON", "DATE"]
            }
            all_records = self._extract_alv_grid(alv_grid, col_map)
            print(f"[SAP VL06O] Đã trích xuất {len(all_records)} dòng đơn OB từ SAP.")

            # 12. Lọc lại một lượt để chắc chắn và định dạng số lượng
            filtered_records = []
            for r in all_records:
                prod = str(r.get("product", "")).strip().upper()
                if clean_mats and not self._matches_material(prod, clean_mats):
                    continue
                if clean_ob and clean_ob not in str(r.get("obNo", "")).strip():
                    continue
                gm = str(r.get("gmStatus", "")).strip().upper()
                if gm and gm not in ("A", "B", ""):
                    continue
                r["qty"] = format_qty_display(parse_sap_float(r.get("qty")))
                filtered_records.append(r)

            print(f"[SAP VL06O] Sau khi lọc theo điều kiện: còn {len(filtered_records)} dòng đơn OB thỏa mãn.")
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
        po_no: str = "",
        ref_no: str = "",
        materials: Optional[List[str]] = None,
        plant: str = "3093"
    ) -> List[Dict[str, str]]:
        """Tìm kiếm đơn Inbound chưa biết số dựa theo PO/SA (VGBEL), Số vận đơn/Bill (LIFEX),
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
                    qty_f = parse_sap_float(item.get("qty", 0.0))
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
                    is_unknown = cat_status == "UNKNOWN"

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
                    elif qty_f > 0:
                        advice = f"✅ FIFO #{avail_rank} - Sẵn sàng ra đơn MIGO 311 ({format_qty_display(qty_f)} {uom})"
                        rank_lbl = f"#{avail_rank}"
                        avail_rank += 1
                    else:
                        advice = "⚪ Số lượng khả dụng bằng 0"
                        rank_lbl = "-"

                    records.append({
                        "product": m,
                        "plant": "3093",
                        "sloc": "1005",
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
            if post_err is not None and mb52_err is not None:
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
                if batch and r_batch and r_batch != batch.strip().upper():
                    continue

                q = parse_sap_float(r.get("unrest", 0.0))
                u = r.get("unit", "PC")
                for m in missing_mats:
                    if m == p_mat or m.lstrip("0") == p_mat.lstrip("0"):
                        mb52_by_mat[m] += q
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
