# -*- coding: utf-8 -*-
"""Điều khiển SAP GUI Scripting qua COM Object trong RAM (0 PowerShell / 0 Clipboard)."""

import os
import re
import time
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import win32com.client

SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_CONNECTION_NAME = ""


class SapEngine:
    """Quản lý kết nối và điều khiển tự động SAP GUI."""

    def __init__(self, max_alv_rows: int = 300, keep_alive_idle_seconds: int = 900):
        self.max_alv_rows = max_alv_rows
        self.keep_alive_idle_seconds = keep_alive_idle_seconds
        self.session = None
        self.last_activity = time.time()
        self.ensure_session()

    def _connect_sap(self):
        """Kết nối tới phiên SAP GUI đang mở; nếu chưa chạy thì bật qua os.startfile."""
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
        conn = app.Children(0)
        if conn.Children.Count == 0:
            raise RuntimeError("Chưa có Session nào đang mở trong SAP GUI.")
        session = conn.Children(0)
        self.last_activity = time.time()
        return session

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
        """Gửi xung nhịp nhẹ làm mới phiên SAP chống timeout tự động của SAP Basis."""
        thresh = idle_threshold_seconds or self.keep_alive_idle_seconds
        if time.time() - self.last_activity < thresh:
            return

        print(f"[KEEP-ALIVE] Đã qua {int(time.time() - self.last_activity)}s không có thao tác. Gửi xung nhịp làm mới SAP...")
        try:
            if not self._is_session_alive():
                self.session = self._connect_sap()
            else:
                self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
                self.session.findById("wnd[0]").sendVKey(0)
                self._dismiss_popups()
            self.last_activity = time.time()
            print(f"[KEEP-ALIVE] Phiên SAP đã được làm mới an toàn ({datetime.now().strftime('%H:%M:%S')})")
        except Exception as e:
            print(f"[KEEP-ALIVE WARNING] Không thể gửi xung nhịp: {e}. Sẽ kết nối lại khi có lệnh mới.")
            try:
                self.session = self._connect_sap()
                self.last_activity = time.time()
            except Exception:
                pass

    def _dismiss_popups(self) -> str:
        """Kiểm tra và đóng cửa sổ popup wnd[1] nếu có."""
        msg_text = ""
        try:
            wnd1 = self.session.findById("wnd[1]", False)
            if wnd1 is not None:
                msg_text = wnd1.Text
                wnd1.sendVKey(0)
        except Exception:
            pass
        return msg_text

    def _fill_multiple_selection(self, push_button_id: str, values: List[str]):
        """Điền danh sách nhiều giá trị vào Multiple Selection mà KHÔNG đụng Clipboard."""
        clean_vals = [v.strip().upper() for v in values if v.strip()]
        if not clean_vals:
            return

        self.session.findById(push_button_id).press()
        time.sleep(0.3)

        try:
            self.session.findById("wnd[1]/tbar[0]/btn[16]").press()
        except Exception:
            pass

        tbl_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
        for idx, val in enumerate(clean_vals):
            page_row = idx % 7
            if idx > 0 and page_row == 0:
                self.session.findById(tbl_id).verticalScrollbar.position = idx
                page_row = 0

            cell_id = f"{tbl_id}/ctxtRSCSEL_255-SLOW_I[1,{page_row}]"
            try:
                self.session.findById(cell_id).text = val
            except Exception:
                alt_cell_id = cell_id.replace("ctxtRSCSEL_255", "txtRSCSEL_255")
                self.session.findById(alt_cell_id).text = val

        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()

    def _extract_alv_grid(self, alv_obj, column_mapping: Dict[str, List[str]]) -> List[Dict[str, str]]:
        """Trích xuất dữ liệu ALV Grid kèm cơ chế cuộn dòng chống mất dữ liệu."""
        row_count = alv_obj.RowCount
        if row_count == 0:
            return []

        col_collection = alv_obj.ColumnOrder
        available_cols = set()
        for c_idx in range(col_collection.Count):
            available_cols.add(str(col_collection.ElementAt(c_idx)).upper())

        resolved_cols = {}
        for out_key, candidate_cols in column_mapping.items():
            for cand in candidate_cols:
                if cand.upper() in available_cols:
                    resolved_cols[out_key] = cand.upper()
                    break

        records = []
        limit = min(row_count, self.max_alv_rows)

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
            records.append(item)

        return records

    # --------------------------------------------------------------------------
    # 1. KIỂM TRA VỊ TRÍ BIN & NGÀY GIỜ GR QUA POSTING CHANGE (/SCWM/POST)
    # --------------------------------------------------------------------------
    def check_position_posting_change(self, materials: List[str], batch: Optional[str] = None) -> List[Dict[str, str]]:
        self.ensure_session()
        self.last_activity = time.time()
        self.session.StartTransaction("/n/SCWM/POST")
        time.sleep(0.5)
        self._dismiss_popups()

        adv_start_btn_id = (
            "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
            "subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/btnCMD_START_ADVANCED"
        )
        toggle_btn_id = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/btnGV_BUTTON_TEXT"

        if self.session.findById(adv_start_btn_id, False) is None:
            toggle_btn = self.session.findById(toggle_btn_id, False)
            if toggle_btn is not None:
                toggle_btn.press()
                time.sleep(0.3)

        mat_push_id = (
            "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
            "subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/subSUB_SEARCH_CRITERIA:/SCWM/SAPLUI_POST:2100/"
            "btn%_P_MATNR_%_APP_%-VALU_PUSH"
        )
        mat_single_id = (
            "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
            "subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/subSUB_SEARCH_CRITERIA:/SCWM/SAPLUI_POST:2100/"
            "ctxtP_MATNR-LOW"
        )

        if len(materials) == 1 and self.session.findById(mat_single_id, False) is not None:
            self.session.findById(mat_single_id).text = materials[0]
        elif self.session.findById(mat_push_id, False) is not None:
            self._fill_multiple_selection(mat_push_id, materials)

        if batch:
            for batch_field_id in [
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/subSUB_SEARCH_CRITERIA:/SCWM/SAPLUI_POST:2100/ctxtP_CHARG-LOW",
                "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/subSUB_SEARCH_CRITERIA:/SCWM/SAPLUI_POST:2100/ctxtS_CHARG-LOW"
            ]:
                b_field = self.session.findById(batch_field_id, False)
                if b_field is not None:
                    b_field.text = batch
                    break

        start_btn = self.session.findById(adv_start_btn_id, False)
        if start_btn is not None:
            start_btn.press()
        else:
            self.session.findById("wnd[0]").sendVKey(8)

        time.sleep(1.0)
        self._dismiss_popups()

        alv_id = (
            "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
            "subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2200/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
        )
        alv_grid = self.session.findById(alv_id, False)
        if alv_grid is None:
            return []

        col_map = {
            "product": ["MATNR", "PRODUCT", "MATNR_EXT"],
            "batch":   ["CHARG", "BATCH"],
            "bin":     ["LGPLA", "VLPLA", "NLPLA"],
            "grDate":  ["WDATU", "GR_DATE"],
            "grTime":  ["WTIME", "GR_TIME"],
            "qty":     ["AQUAN", "AQTY", "QUAN", "GESME"],
            "unit":    ["MEINS", "ALTME", "UNIT"],
            "cat":     ["CAT", "LGTYP"]
        }
        records = self._extract_alv_grid(alv_grid, col_map)

        # SẮP XẾP ƯU TIÊN DỮ LIỆU GẦN NHẤT (Ngày GR & Giờ GR mới nhất lên đầu)
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
    # 2. KIỂM TRA TỒN KHO TỔNG THỂ (MB52)
    # --------------------------------------------------------------------------
    def check_stock_mb52(self, materials: List[str], plant: str = "3093") -> List[Dict[str, str]]:
        self.ensure_session()
        self.last_activity = time.time()
        self.session.StartTransaction("/nMB52")
        time.sleep(0.5)

        if len(materials) == 1:
            self.session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = materials[0]
        else:
            self._fill_multiple_selection("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH", materials)

        if plant:
            self.session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = plant

        nozero_chk = self.session.findById("wnd[0]/usr/chkNOZERO", False)
        if nozero_chk is not None:
            nozero_chk.selected = True
        flat_rad = self.session.findById("wnd[0]/usr/radPA_FLT", False)
        if flat_rad is not None:
            flat_rad.select()

        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(0.8)
        self._dismiss_popups()

        alv_grid = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
        if alv_grid is None:
            return []

        col_map = {
            "product": ["MATNR"],
            "plant":   ["WERKS"],
            "sloc":    ["LGORT"],
            "batch":   ["CHARG"],
            "unrest":  ["LABST"],
            "qual":    ["INSME"],
            "unit":    ["MEINS"]
        }
        records = self._extract_alv_grid(alv_grid, col_map)

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
        self.ensure_session()
        self.last_activity = time.time()
        self.session.StartTransaction("/nMB51")
        time.sleep(0.5)

        if len(materials) == 1:
            self.session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = materials[0]
        else:
            self._fill_multiple_selection("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH", materials)

        if plant:
            self.session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = plant

        end_date = datetime.now()
        start_date = end_date - timedelta(days=int(days))
        self.session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = start_date.strftime("%d.%m.%Y")
        self.session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = end_date.strftime("%d.%m.%Y")

        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(1.0)
        self._dismiss_popups()

        alv_grid = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
        if alv_grid is None:
            return []

        col_map = {
            "product": ["MATNR"],
            "docNo":   ["MBLNR"],
            "mvtType": ["BWART"],
            "postDate":["BUDAT"],
            "qty":     ["MENGE", "ERFMG"],
            "unit":    ["MEINS", "ERFME"],
            "sloc":    ["LGORT"],
            "batch":   ["CHARG"]
        }
        records = self._extract_alv_grid(alv_grid, col_map)

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
