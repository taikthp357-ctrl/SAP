# -*- coding: utf-8 -*-
"""Dịch vụ tự động phân luồng và in phiếu kho SAP EWM (Warehouse Order & Outbound Delivery).

Quy tắc phân luồng thông minh:
1. Số Warehouse Order (Mã bắt đầu bằng 6*, vd: 6028447808):
   -> Đi thẳng vào màn hình in F00112 (bỏ qua /SCWM/MON).
2. Số Outbound Delivery / ERP Document (Mã bắt đầu bằng 101*, vd: 1013144096):
   -> Vào /SCWM/MON -> Bấm nút 'Warehouse Order' để lấy danh sách mã 6* -> Sang F00112 in.
3. Số bản in: Chuẩn hóa 2 bản (TDCOPIES = 2) trên SAP.
4. Win32 Watchdog: Tự động đóng popup Windows Print (#32770) ngầm bằng PostMessage BM_CLICK/IDOK,
   không lo mất focus khi người dùng đang thao tác cửa sổ khác.
"""

import re
import time
import threading
from typing import List, Dict, Any, Optional, Tuple

try:
    import pythoncom
    import win32com.client
    import win32gui
    import win32con
    import win32print
    import win32clipboard
except ImportError:
    pythoncom = None
    win32com = None
    win32gui = None
    win32con = None
    win32print = None
    win32clipboard = None


class SapWhoPrinterService:
    """Điều phối quy trình tra cứu và in phiếu Warehouse Order chuẩn SAP EWM."""

    DEFAULT_WHSE = "VF02"
    DEFAULT_PRINTER = r"\\10.128.10.67\VF-MA-103-118"

    def __init__(self, warehouse_num: str = DEFAULT_WHSE, default_printer: Optional[str] = DEFAULT_PRINTER):
        self.warehouse_num = warehouse_num
        self.default_printer = default_printer

    @staticmethod
    def classify_document(doc_str: str) -> str:
        """Phân loại chứng từ theo đầu số:
        - Bắt đầu bằng 6 (độ dài 8-10 số): Warehouse Order (WHO)
        - Bắt đầu bằng 101 (độ dài 10 số) hoặc có tiền tố OB: Outbound Delivery (OB)
        """
        clean = re.sub(r"[^A-Za-z0-9]", "", str(doc_str)).strip().upper()
        if re.fullmatch(r"6\d{7,9}", clean):
            return "WHO"
        if re.fullmatch(r"101\d{7}", clean) or clean.startswith("OB") or re.fullmatch(r"1\d{9}", clean):
            return "OB"
        return "UNKNOWN"

    @staticmethod
    def _win32_print_watchdog(win_copies: Optional[int] = None, timeout_sec: float = 15.0) -> bool:
        """Luồng phụ đứng canh cửa sổ Print (#32770) của Windows và xác nhận OK ngầm."""
        if win32gui is None:
            return False
        end_time = time.time() + timeout_sec
        while time.time() < end_time:
            time.sleep(0.2)
            hwnd = win32gui.FindWindow("#32770", "Print")
            if not hwnd:
                # Thử tìm theo tiêu đề tiếng Việt 'In'
                hwnd = win32gui.FindWindow("#32770", "In")
            if hwnd and win32gui.IsWindowVisible(hwnd):
                try:
                    # Nếu có chỉ định số bản in ở Windows Dialog
                    if win_copies is not None:
                        edit_copies = win32gui.GetDlgItem(hwnd, 1154)
                        if edit_copies:
                            win32gui.SendMessage(edit_copies, win32con.WM_SETTEXT, 0, str(win_copies))
                            time.sleep(0.05)

                    # Bấm nút OK (IDOK = 1 trong Win32 Dialog)
                    btn_ok = win32gui.GetDlgItem(hwnd, win32con.IDOK)
                    if btn_ok:
                        # Gửi BM_CLICK và PostMessage WM_COMMAND để đảm bảo đóng dialog
                        win32gui.SendMessage(btn_ok, win32con.BM_CLICK, 0, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, btn_ok)
                        print("[PRINT WATCHDOG] Đã tự động bấm OK trên hộp thoại Print của Windows thành công.")
                        return True
                except Exception as ex:
                    print(f"[PRINT WATCHDOG WARN] Lỗi khi bấm nút OK trên dialog: {ex}")
        print("[PRINT WATCHDOG WARN] Hết thời gian chờ popup in của Windows.")
        return False

    @staticmethod
    def _set_clipboard_text(text: str) -> bool:
        """Ghi chuỗi vào Windows Clipboard để phục vụ nạp qua nút Upload btn[24]."""
        if win32clipboard is None:
            return False
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
            win32clipboard.CloseClipboard()
            return True
        except Exception as e:
            print(f"[CLIPBOARD WARN] Không thể ghi clipboard: {e}")
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            return False

    @staticmethod
    def _click_who_toolbar_button(top_grid) -> bool:
        """Tìm và bấm nút 'Warehouse Order' trên thanh công cụ của bảng Outbound Delivery Order."""
        try:
            btn_count = getattr(top_grid, "ButtonCount", 0)
            for i in range(btn_count):
                btn_id = top_grid.GetButtonId(i)
                btn_text = (top_grid.GetButtonText(i) or "").strip().lower()
                btn_tip = (top_grid.GetButtonTooltip(i) or "").strip().lower()
                if "warehouse order" in btn_text or "warehouse order" in btn_tip or "whse order" in btn_text:
                    top_grid.pressToolbarButton(btn_id)
                    return True
        except Exception:
            pass

        # Fallback về các ID nút quen thuộc nếu không duyệt được thuộc tính
        for candidate_id in ("N000000096", "N000000108", "&WHO"):
            try:
                top_grid.pressToolbarButton(candidate_id)
                return True
            except Exception:
                pass
        return False

    @staticmethod
    def _extract_who_from_bottom_grid(bottom_grid) -> List[str]:
        """Đọc danh sách mã Warehouse Order từ bảng kết quả bên dưới (ALV Grid)."""
        row_count = getattr(bottom_grid, "RowCount", 0)
        if row_count <= 0:
            return []

        target_col = "WHO"
        try:
            col_order = bottom_grid.ColumnOrder
            for i in range(col_order.Count):
                col_name = col_order.ElementAt(i)
                try:
                    col_title = (bottom_grid.GetColumnTitles(col_name).ElementAt(0) or "").strip().lower()
                except Exception:
                    col_title = ""
                if col_name in ("WHO", "WHOID", "WHSE_ORDER", "DOCNO") or "whse order" in col_title or "warehouse order" in col_title:
                    target_col = col_name
                    break
        except Exception:
            pass

        who_list: List[str] = []
        for r in range(row_count):
            try:
                val = str(bottom_grid.GetCellValue(r, target_col)).strip()
                # Chỉ nhận các mã bắt đầu bằng 6 (chuẩn số WHO)
                if val and val.startswith("6") and val not in who_list:
                    who_list.append(val)
            except Exception:
                pass
        return who_list

    def search_who_from_ob(self, session, ob_doc: str) -> List[str]:
        """Truy vấn /SCWM/MON để lấy danh sách mã WHO thuộc đơn Outbound Delivery."""
        clean_ob = re.sub(r"[^0-9]", "", str(ob_doc)).strip()
        if not clean_ob:
            return []

        # 1. Quay về màn hình gốc
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        # 2. Mở node F00077 (/SCWM/MON)
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00077")

        # 3. Chọn nhánh Outbound Delivery Order (N000000010)
        mon_tree = session.findById("wnd[0]/usr/shell/shellcont[0]/shell")
        mon_tree.selectedNode = "N000000010"
        mon_tree.doubleClickNode("N000000010")

        # 4. Nhập mã ERP Document và F8
        session.findById("wnd[1]/usr/txtS_ERP_H-LOW").text = clean_ob
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # 5. Bấm nút [Warehouse Order] trên thanh công cụ trên
        top_grid = session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
        self._click_who_toolbar_button(top_grid)

        # 6. Đọc kết quả từ bảng dưới
        bottom_grid = session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
        return self._extract_who_from_bottom_grid(bottom_grid)

    def print_who_direct(
        self,
        session,
        who_list: List[str],
        copies: int = 2
    ) -> bool:
        """Nhảy thẳng vào màn hình in F00112 để in danh sách mã WHO (không qua Monitor)."""
        if not who_list:
            return False

        # Đặt máy in mặc định trên Windows nếu có cấu hình
        if self.default_printer and win32print is not None:
            try:
                win32print.SetDefaultPrinter(self.default_printer)
            except Exception as pe:
                print(f"[PRINTER WARN] Không thể set máy in mặc định {self.default_printer}: {pe}")

        # 1. Quay về màn hình gốc Easy Access
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        # 2. Điều hướng cây menu F00010 -> F00112
        main_tree = session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
        main_tree.selectedNode = "F00010"
        main_tree.doubleClickNode("F00010")
        main_tree.selectedNode = "F00112"
        main_tree.topNode = "0000000003"
        main_tree.doubleClickNode("F00112")

        # 3. Điền Warehouse Number (VF02)
        session.findById("wnd[0]/usr/ctxtP_LGNUM").text = self.warehouse_num

        # 4. Mở bảng Multiple Selection để nạp danh sách WHO
        session.findById("wnd[0]/usr/btn%_S_WHO_%_APP_%-VALU_PUSH").press()

        # Nạp dữ liệu qua Clipboard (btn[24] Upload) để không bị giới hạn 7 dòng
        clip_ok = self._set_clipboard_text("\r\n".join(who_list))
        uploaded_via_btn24 = False
        if clip_ok:
            try:
                session.findById("wnd[1]/tbar[0]/btn[24]").press()
                uploaded_via_btn24 = True
            except Exception:
                uploaded_via_btn24 = False

        if not uploaded_via_btn24:
            # Fallback nếu nút btn[24] không khả dụng: điền từng dòng
            for idx, who_code in enumerate(who_list):
                if idx < 7:
                    try:
                        session.findById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,{idx}]").text = who_code
                    except Exception:
                        pass

        # 5. Xác nhận bảng Multiple Selection (F8)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # 6. Chạy chương trình in (F8)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # 7. Nhập số bản in trên SAP (chuẩn hóa 2 bản)
        try:
            session.findById("wnd[1]/usr/ctxtSSFPP-TDCOPIES").text = str(copies)
        except Exception:
            pass

        # 8. Khởi chạy luồng Watchdog canh popup Print của Windows TRƯỚC KHI bấm btn[86]
        watchdog = threading.Thread(
            target=self._win32_print_watchdog,
            kwargs={"win_copies": None, "timeout_sec": 15.0},
            daemon=True
        )
        watchdog.start()

        # 9. Bấm nút Print trên SAP (btn[86])
        session.findById("wnd[1]/tbar[0]/btn[86]").press()

        # Đợi Watchdog hoàn tất
        watchdog.join(timeout=15.0)
        return True

    def execute_smart_print(
        self,
        session,
        doc_inputs: List[str],
        copies: int = 2
    ) -> Dict[str, Any]:
        """Bộ định tuyến thông minh:
        - Gom các mã 6* -> Nhánh in thẳng F00112
        - Gom các mã 101* -> Nhánh tra cứu /SCWM/MON lấy mã 6*
        - Hợp nhất và in 1 lượt trọn vẹn 2 tờ.
        """
        direct_who: List[str] = []
        ob_docs: List[str] = []

        for d in doc_inputs:
            c = str(d).strip()
            cat = self.classify_document(c)
            if cat == "WHO":
                if c not in direct_who:
                    direct_who.append(c)
            elif cat == "OB":
                if c not in ob_docs:
                    ob_docs.append(c)

        discovered_who: List[str] = []
        for ob in ob_docs:
            try:
                found = self.search_who_from_ob(session, ob)
                for f in found:
                    if f not in discovered_who and f not in direct_who:
                        discovered_who.append(f)
            except Exception as e:
                print(f"[SMART PRINT ERROR] Không thể tra cứu WHO từ OB {ob}: {e}")

        final_who_list = list(dict.fromkeys(direct_who + discovered_who))

        if not final_who_list:
            return {
                "status": "FAILED",
                "message": "Không tìm thấy mã Warehouse Order (6*) nào để in.",
                "direct_who": direct_who,
                "ob_docs": ob_docs,
                "who_printed": []
            }

        print(f"[SMART PRINT] Bắt đầu in {len(final_who_list)} phiếu WHO (Trực tiếp: {direct_who}, Từ OB: {discovered_who})...")
        success = self.print_who_direct(session, final_who_list, copies=copies)

        return {
            "status": "SUCCEEDED" if success else "FAILED",
            "message": f"Đã gửi lệnh in {len(final_who_list)} phiếu Warehouse Order ({copies} tờ/phiếu) thành công!" if success else "Lỗi trong quá trình gửi lệnh in SAP.",
            "direct_who": direct_who,
            "ob_docs": ob_docs,
            "who_printed": final_who_list,
            "copies": copies,
            "printer": self.default_printer
        }
