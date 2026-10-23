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
import re
import time
import unicodedata
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
import win32com.client

from app.number_utils import parse_sap_float, format_qty_display

SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_CONNECTION_NAME = ""


class SapExecutionError(RuntimeError):
    """Lỗi phát sinh trong quá trình thực thi nghiệp vụ trên SAP GUI."""
    pass


class SapEngine:
    """Quản lý kết nối và điều khiển tự động SAP GUI theo chuẩn BBBG."""

    def __init__(
        self,
        max_alv_rows: int = 300,
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
        bot_safe_tcodes = {"SESSION_MANAGER", "/SCWM/POST", "MB52", "MB51", "VL06I", "ME2N", ""}

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

                bot_safe_tcodes = {"SESSION_MANAGER", "/SCWM/POST", "MB52", "MB51", "VL06I", "ME2N", ""}
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
        msg_text = ""
        try:
            count = 0
            while count < 10:
                wnd1 = self._find_by_id_safe("wnd[1]")
                if wnd1 is None:
                    break
                msg_text = getattr(wnd1, "Text", "")
                wnd1.sendVKey(0)
                count += 1
                self._wait_idle(3.0)
        except Exception:
            pass
        return msg_text

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

    def _fill_multiple_selection(self, push_button_id: str, values: List[str]):
        """Điền danh sách nhiều giá trị vào Multiple Selection mà KHÔNG đụng Clipboard."""
        clean_vals = [v.strip().upper() for v in values if v.strip()]
        if not clean_vals:
            return

        push_btn = self._find_by_id_safe(push_button_id)
        if push_btn is None:
            return
        push_btn.press()
        self._wait_idle(5.0)

        # Xóa các lựa chọn cũ nếu có nút btn[16] (Delete whole selection)
        btn16 = self._find_by_id_safe("wnd[1]/tbar[0]/btn[16]")
        if btn16 is not None:
            try:
                btn16.press()
                self._wait_idle(2.0)
            except Exception:
                pass

        tbl_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
        tbl = self._find_by_id_safe(tbl_id)

        for idx, val in enumerate(clean_vals):
            page_row = idx % 7
            if idx > 0 and page_row == 0 and tbl is not None:
                try:
                    tbl.verticalScrollbar.position = idx
                    page_row = 0
                except Exception:
                    pass

            cell_id = f"{tbl_id}/ctxtRSCSEL_255-SLOW_I[1,{page_row}]"
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

        self._wait_idle(5.0)

    def _validate_grid_or_raise(self, alv_grid, tcode_name: str, status_txt: str, status_type: str):
        """Phân biệt chính xác giữa việc không có dữ liệu thực sự và lỗi không mở được bảng kết quả SAP."""
        if alv_grid is not None:
            return

        is_selection_screen = self._find_by_id_safe("wnd[0]/tbar[1]/btn[8]") is not None
        lower_txt = (status_txt or "").lower()

        empty_phrases = [
            "no data", "not found", "không tìm thấy", "keine daten",
            "chưa có", "không có dữ liệu", "no document", "no item",
            "no suitable", "no records", "không có chứng từ"
        ]

        if is_selection_screen:
            if any(p in lower_txt for p in empty_phrases):
                print(f"[SAP {tcode_name}] Xác nhận không có dữ liệu: '{status_txt}'")
                return
            raise SapExecutionError(
                f"Không thể mở bảng kết quả SAP {tcode_name} (màn hình nhập liệu không phản hồi hoặc bị chặn). "
                f"Trạng thái: '{status_txt or 'Chưa rõ'}'"
            )

        if any(p in lower_txt for p in empty_phrases):
            return

        raise SapExecutionError(
            f"Không tìm thấy bảng kết quả ALV Grid trong SAP {tcode_name}. "
            f"Trạng thái: '{status_txt or 'Lỗi hiển thị giao diện SAP'}'"
        )

    def _extract_alv_grid(self, alv_obj, column_mapping: Dict[str, List[str]]) -> List[Dict[str, str]]:
        """Trích xuất ALV Grid kèm cơ chế cuộn dòng, forward-fill mã hàng và lọc bỏ dòng tổng cộng (Summary rows)."""
        try:
            row_count = int(alv_obj.RowCount)
        except Exception:
            return []

        if row_count == 0:
            return []

        resolved_cols = self._resolve_grid_columns(alv_obj, column_mapping)
        records = []
        limit = min(row_count, self.max_alv_rows)
        last_product = ""

        for r in range(limit):
            if r % 25 == 0:
                try:
                    alv_obj.firstVisibleRow = r
                except Exception:
                    pass

            item = {}
            for out_key, sap_col in resolved_cols.items():
                try:
                    val = alv_obj.GetCellValue(r, sap_col)
                    item[out_key] = str(val).strip()
                except Exception:
                    item[out_key] = ""

            # Forward-fill Product nếu ô bị gộp hiển thị trống
            curr_prod = item.get("product", "")
            if curr_prod:
                last_product = curr_prod
            elif last_product and "product" in item:
                item["product"] = last_product

            # LỌC BỎ DÒNG TỔNG CỘNG / DÒNG RỖNG:
            # Dòng dữ liệu hợp lệ trong SAP ALV phải có ít nhất một trường định danh cụ thể:
            # (Batch, Bin, SLoc, DocNo, STO, IB, Item No, Ref No, MvtType)
            has_identifying_data = bool(
                item.get("batch") or item.get("bin") or item.get("sloc") or
                item.get("docNo") or item.get("stoNo") or item.get("ibNo") or
                item.get("item") or item.get("refNo") or item.get("mvtType")
            )
            if not has_identifying_data:
                continue

            records.append(item)

        return records

    def _open_transaction(self, tcode: str):
        """Mở Transaction trên SAP GUI an toàn tuyệt đối qua ô lệnh okcd hoặc StartTransaction."""
        self.ensure_session()
        self.last_activity = time.time()

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

        # 5. Điền số lô (Batch) nếu có
        if batch:
            batch_ctl = self._find_by_id_safe(ID_BATCH_LOW)
            if batch_ctl is not None:
                try:
                    batch_ctl.text = batch.strip()
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
            "product": ["MATNR", "PRODUCT", "MATERIAL", "MATNR_EXT"],
            "batch":   ["CHARG", "BATCH", "DEST.BATCH", "DEST BATCH"],
            "bin":     ["LGPLA", "VLPLA", "NLPLA", "SOURCE STORAGE BIN", "SOURCE BIN", "STORAGE BIN", "BIN"],
            "qty":     ["AQUAN", "AQTY", "AVL_ALT_QTY", "A.QTY AUOM", "A.QTY", "AVAILABLE QTY", "AVAILABLE QUANTITY", "AVLQTY", "QUAN", "TOTAL STK", "TOTAL STOCK", "GESME"],
            "unit":    ["MEINS", "ALTME", "UNIT", "UOM"],
            "cat":     ["CAT", "TST", "STOCK TYPE", "ST", "TO_CAT"],
            "grDate":  ["WDATU", "GR_DATE", "WEDAT", "GR DATE", "GOODS RECEIPT DATE"],
            "grTime":  ["WTIME", "GR_TIME", "UZEIT", "GR TIME", "GOODS RECEIPT TIME"]
        }
        records = self._extract_alv_grid(alv_grid, col_map)
        print(f"[SAP POST] Đã trích xuất {len(records)} dòng từ /SCWM/POST.")

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
    # 2. KIỂM TRA TỒN KHO TỔNG THỂ (MB52 - KÈM VỊ TRÍ BIN LGPBE THEO CHUẨN BBBG)
    # --------------------------------------------------------------------------
    def check_stock_mb52(
        self,
        materials: List[str],
        plant: str = "3093",
        slocs: Optional[List[str]] = None
    ) -> List[Dict[str, str]]:
        """Tra cứu tồn kho tổng hợp MB52 kèm Vị trí Bin (LGPBE) theo modMB52.bas của BBBG.
        Hỗ trợ lọc theo danh sách kho (slocs, ví dụ ['1005', '3093'])."""
        if not materials:
            return []

        self._open_transaction("MB52")
        self._wait_idle(10.0)

        print(f"[SAP MB52] Đang nhập điều kiện tìm kiếm: Materials={materials}, Plant={plant}...")
        clean_mats = [m.strip().upper() for m in materials if m.strip()]

        mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtMATNR-LOW")
        if len(clean_mats) == 1 and mat_single is not None:
            mat_single.text = clean_mats[0]
        else:
            multi_btn_id = "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
            if self._find_by_id_safe(multi_btn_id) is not None:
                self._fill_multiple_selection(multi_btn_id, clean_mats)
            elif mat_single is not None and clean_mats:
                mat_single.text = clean_mats[0]

        if plant:
            for plant_field_id in ["wnd[0]/usr/ctxtWERKS-LOW", "wnd[0]/usr/ctxtRM07M-WERKS"]:
                p_ctl = self._find_by_id_safe(plant_field_id)
                if p_ctl is not None:
                    try:
                        p_ctl.text = plant
                    except Exception:
                        pass
                    break

        # Tùy chọn bỏ dòng tồn bằng 0 và định dạng dạng phẳng nếu có
        nozero_chk = self._find_by_id_safe("wnd[0]/usr/chkNOZERO")
        if nozero_chk is not None:
            try:
                nozero_chk.selected = True
            except Exception:
                pass

        flat_rad = self._find_by_id_safe("wnd[0]/usr/radPA_FLT")
        if flat_rad is not None:
            try:
                flat_rad.select()
            except Exception:
                pass

        # Điền Layout nếu có ô Layout /KHOMA như modMB52.bas
        layout_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtRM07M-VARIANT")
        if layout_ctl is not None:
            try:
                layout_ctl.text = "/KHOMA"
            except Exception:
                pass

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
            "product": ["MATNR", "MATERIAL"],
            "plant":   ["WERKS", "PLANT"],
            "sloc":    ["LGORT", "SLOC", "STORAGE LOCATION"],
            "batch":   ["CHARG", "BATCH"],
            "bin":     ["LGPBE", "BIN", "STORAGE BIN"],
            "unrest":  ["LABST", "UNRESTRICTED", "UNRESTRICTED USE"],
            "qual":    ["INSME", "QUALITY INSPECTION", "IN QUALITY INSP."],
            "blocked": ["SPEME", "BLOCKED"],
            "unit":    ["MEINS", "BASE UNIT OF MEASURE", "UNIT"]
        }
        records = self._extract_alv_grid(alv_grid, col_map)
        print(f"[SAP MB52] Đã trích xuất {len(records)} dòng dữ liệu tồn kho.")

        if slocs:
            clean_slocs = {s.strip().upper() for s in slocs if s.strip()}
            records = [r for r in records if str(r.get("sloc", "")).strip().upper() in clean_slocs]
            print(f"[SAP MB52] Sau khi lọc theo kho {clean_slocs}: còn {len(records)} dòng.")

        # SẮP XẾP: Ưu tiên dòng có Tồn khả dụng lớn nhất và Batch mới nhất lên đầu
        def _mb52_sort_key(item):
            qty_raw = item.get("unrest", "0").replace(".", "").replace(",", ".")
            try:
                qty = float(qty_raw)
            except Exception:
                qty = 0.0
            return (qty, item.get("batch", ""))

        records.sort(key=_mb52_sort_key, reverse=True)
        return records

    # --------------------------------------------------------------------------
    # 3. KIỂM TRA LỊCH SỬ XUẤT NHẬP (MB51)
    # --------------------------------------------------------------------------
    def check_history_mb51(self, materials: List[str], days: int = 7, plant: str = "3093") -> List[Dict[str, str]]:
        """Tra cứu lịch sử xuất nhập kho trong N ngày qua MB51."""
        if not materials:
            return []

        self._open_transaction("MB51")
        self._wait_idle(10.0)

        print(f"[SAP MB51] Đang nhập điều kiện: Materials={materials}, Days={days}, Plant={plant}...")
        clean_mats = [m.strip().upper() for m in materials if m.strip()]

        mat_single = self._find_by_id_safe("wnd[0]/usr/ctxtMATNR-LOW")
        if len(clean_mats) == 1 and mat_single is not None:
            mat_single.text = clean_mats[0]
        else:
            multi_btn_id = "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
            if self._find_by_id_safe(multi_btn_id) is not None:
                self._fill_multiple_selection(multi_btn_id, clean_mats)
            elif mat_single is not None and clean_mats:
                mat_single.text = clean_mats[0]

        if plant:
            p_ctl = self._find_by_id_safe("wnd[0]/usr/ctxtWERKS-LOW")
            if p_ctl is not None:
                try:
                    p_ctl.text = plant
                except Exception:
                    pass

        end_date = datetime.now()
        start_date = end_date - timedelta(days=int(days))
        b_low = self._find_by_id_safe("wnd[0]/usr/ctxtBUDAT-LOW")
        if b_low is not None:
            b_low.text = start_date.strftime("%d.%m.%Y")
        b_high = self._find_by_id_safe("wnd[0]/usr/ctxtBUDAT-HIGH")
        if b_high is not None:
            b_high.text = end_date.strftime("%d.%m.%Y")

        print("[SAP MB51] Đang bấm thực thi (F8)...")
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
    # 4. KIỂM TRA ĐƠN INBOUND (IB) TREO CHƯA GR (VL06I)
    # --------------------------------------------------------------------------
    def check_inbound_treo_vl06i(self, materials: List[str], plant: str = "3093") -> List[Dict[str, str]]:
        """Tra cứu các đơn Inbound Delivery (IB) đang treo chưa GR qua VL06I theo chuẩn BBBG."""
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
    # 5. KIỂM TRA ĐƠN ĐIỀU CHUYỂN KHO (STO - ME2N)
    # --------------------------------------------------------------------------
    def check_sto_me2n(self, materials: List[str], plant: str = "3093") -> List[Dict[str, str]]:
        """Tra cứu đơn STO (Stock Transport Order) điều chuyển kho qua ME2N."""
        if not materials:
            return []

        self._open_transaction("ME2N")
        self._wait_idle(10.0)

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
