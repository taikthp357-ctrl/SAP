# -*- coding: utf-8 -*-
"""
SAP ZONE Core Engine (v1.6)
Tự động hóa toàn diện quy trình Ra đơn NVL, Cấp phát FIFO/FEFO theo Zone,
Tạo đơn MIGO 311, Tra cứu MON lấy số WHO và In phiếu SAP 2 bản lưu PDF.
"""

import os
import sys
import re
import time
import json
import threading
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple, Set

# COM & Win32
try:
    import win32com.client
    import win32gui
    import win32con
    import win32print
    import win32clipboard
except ImportError:
    win32com = None
    win32gui = None
    win32con = None
    win32print = None
    win32clipboard = None

# OpenPyXL
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    openpyxl = None


# ==============================================================================
# 1. QUẢN LÝ CẤU HÌNH & ZONE MAPPING
# ==============================================================================

DEFAULT_CONFIG = {
    "sap_system": "PRD",
    "sap_client": "300",
    "sap_session_idx": 0,
    "ewm_warehouse": "VF02",
    "ewm_storage_type": "APFL",
    "ewm_owner": "3093",
    "plant": "3093",
    "sloc": "3093",
    "max_materials_per_group": 5,
    "max_lines_per_group": 20,
    "allocation_strategy": "FIFO",
    "allow_partial_allocation": True,
    "sap_print_copies": 2,
    "auto_print_pdf": True,
    "pdf_output_folder": os.path.abspath("./OUTPUT_PDF"),
    "printer_name": "Microsoft Print to PDF"
}


def get_base_dir() -> str:
    """Trả về thư mục chứa file thực thi (.exe) hoặc mã nguồn."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundle_dir() -> str:
    """Trả về thư mục bundle tài nguyên (hỗ trợ PyInstaller _MEIPASS)."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return get_base_dir()


class ConfigManager:
    """Quản lý đọc/ghi file cấu hình config.json (luôn nằm cạnh file .exe)."""
    CONFIG_FILE = "config.json"

    @classmethod
    def get_path(cls) -> str:
        return os.path.join(get_base_dir(), cls.CONFIG_FILE)

    @classmethod
    def load_config(cls) -> Dict[str, Any]:
        cfg = DEFAULT_CONFIG.copy()
        path = cls.get_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    user_cfg = json.load(f)
                    cfg.update(user_cfg)
            except Exception as e:
                print(f"[CONFIG WARN] Không thể đọc {path}: {e}")
        return cfg

    @classmethod
    def save_config(cls, cfg: Dict[str, Any]) -> bool:
        path = cls.get_path()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"[CONFIG ERROR] Không thể lưu {path}: {e}")
            return False


class ZoneManager:
    """Quản lý ánh xạ Trạm -> Zone -> Team Leader (TL)."""
    MAPPING_FILE = "zone_map.json"

    def __init__(self):
        self.mappings: List[Dict[str, str]] = []
        self.load_mappings()

    def get_path(self) -> str:
        # 1. Ưu tiên file bên ngoài nằm cạnh file .exe (nếu người dùng tùy biến)
        outer_path = os.path.join(get_base_dir(), self.MAPPING_FILE)
        if os.path.exists(outer_path):
            return outer_path

        # 2. File đóng gói sẵn trong bundle PyInstaller
        bundle_path = os.path.join(get_bundle_dir(), self.MAPPING_FILE)
        if os.path.exists(bundle_path):
            return bundle_path

        # 3. Thư mục làm việc hiện tại
        if os.path.exists(self.MAPPING_FILE):
            return os.path.abspath(self.MAPPING_FILE)

        return outer_path

    def load_mappings(self):
        path = self.get_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    raw_data = json.load(f)
                    # raw_data dạng: [["Limo", "BIW32011852HA", "P", "3.1", "KV-16/Hiếu/Cường Zone 3.1"], ...]
                    self.mappings.clear()
                    for item in raw_data:
                        if isinstance(item, list) and len(item) >= 5:
                            self.mappings.append({
                                "car_model": item[0],
                                "material": item[1],
                                "p_line": item[2],
                                "zone": str(item[3]).strip(),
                                "tl": str(item[4]).strip()
                            })
            except Exception as e:
                print(f"[ZONE WARN] Lỗi đọc {path}: {e}")

    def resolve_zone_tl(self, material: str, station: str) -> Tuple[str, str]:
        """Xác định Zone và TL dựa vào Material hoặc Station."""
        mat_clean = material.strip().upper()
        sta_clean = station.strip().upper()

        # 1. Tìm theo Material chính xác
        for m in self.mappings:
            if m["material"].upper() == mat_clean:
                return m["zone"], m["tl"]

        # 2. Phân tích chuỗi Station tìm Zone (vd: Zone 3.1, Z1.2...)
        m_zone = re.search(r"(?:ZONE|Z)\s*([0-9]+(?:\.[0-9]+)?)", sta_clean, re.IGNORECASE)
        if m_zone:
            z_val = m_zone.group(1)
            # Tìm TL tương ứng
            for m in self.mappings:
                if m["zone"] == z_val:
                    return z_val, m["tl"]
            return z_val, f"TL Zone {z_val}"

        # 3. Fallback mặc định
        return "1.1", "Mặc định Zone 1.1"


# ==============================================================================
# 2. KẾT NỐI SAP GUI SCRIPTING
# ==============================================================================

class SapConnector:
    """Quản lý kết nối tới phiên SAP GUI đang mở."""

    @staticmethod
    def get_session(session_idx: int = 0):
        if win32com is None:
            raise RuntimeError("pywin32 chưa được cài đặt!")
        
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
        except Exception:
            raise RuntimeError("Không tìm thấy tiến trình SAP GUI! Vui lòng đăng nhập SAP trước.")

        application = sap_gui_auto.GetScriptingEngine
        if not application:
            raise RuntimeError("Scripting Engine bị tắt hoặc không được phép trên SAP GUI!")

        connection_count = application.Children.Count
        if connection_count == 0:
            raise RuntimeError("Chưa có kết nối (Connection) nào được mở trong SAP!")

        connection = application.Children(0)
        session_count = connection.Children.Count
        if session_count == 0:
            raise RuntimeError("Chưa có phiên làm việc (Session) nào được mở trong SAP!")

        if session_idx >= session_count:
            session_idx = 0

        session = connection.Children(session_idx)
        return session

    @staticmethod
    def test_connection(session_idx: int = 0) -> Dict[str, Any]:
        try:
            session = SapConnector.get_session(session_idx)
            info = session.Info
            return {
                "ok": True,
                "system": info.SystemName,
                "client": info.Client,
                "user": info.User,
                "tcode": info.Transaction,
                "program": info.Program
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}


# ==============================================================================
# 3. ĐỌC TỒN MON (/SCWM/MON) - BẢN 1.6
# ==============================================================================

class MonStockService:
    """Tải và trích xuất tồn kho trực tiếp từ /SCWM/MON theo ton.vbs & Bản 1.6."""

    @staticmethod
    def copy_to_clipboard(text: str) -> bool:
        if win32clipboard is None:
            return False
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
            win32clipboard.CloseClipboard()
            return True
        except Exception as e:
            print(f"[CLIPBOARD WARN] Lỗi ghi clipboard: {e}")
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            return False

    @staticmethod
    def parse_sap_number(val_str: str) -> float:
        """Chuyển đổi chuỗi số lượng SAP có dấu phân cách hàng nghìn sang float."""
        if not val_str:
            return 0.0
        clean = str(val_str).strip().replace(" ", "")
        # Định dạng SAP châu Âu / VN: dấu chấm là hàng nghìn, phẩy là thập phân
        if "." in clean and "," in clean:
            clean = clean.replace(".", "").replace(",", ".")
        elif "." in clean and "," not in clean:
            # Nếu có dấu chấm nhưng không có phẩy: ví dụ 1.404 hoặc 28.000
            # Kiểm tra vị trí: nếu sau dấu chấm có 3 chữ số thì là hàng nghìn
            parts = clean.split(".")
            if len(parts) == 2 and len(parts[1]) == 3:
                clean = clean.replace(".", "")
            else:
                clean = clean.replace(".", "")
        elif "," in clean:
            clean = clean.replace(",", ".")
        
        try:
            return float(clean)
        except ValueError:
            return 0.0

    def fetch_stock_from_mon(
        self,
        session,
        materials: List[str],
        storage_type: str = "APFL",
        owner: str = "3093"
    ) -> List[Dict[str, Any]]:
        """Truy vấn tồn kho từ /SCWM/MON với bộ lọc APFL + Owner 3093."""
        if not materials:
            return []

        # 1. Quay về màn hình gốc
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        # 2. Mở node F00077 (/SCWM/MON)
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00077")

        # 3. Chọn node tồn kho N000000139 (Stock Overview / Available Stock)
        mon_tree = session.findById("wnd[0]/usr/shell/shellcont[0]/shell")
        mon_tree.selectedNode = "N000000139"
        mon_tree.doubleClickNode("N000000139")

        # 4. Điền Storage Type (APFL) và Owner (3093)
        try:
            session.findById("wnd[1]/usr/txtS_LGTYP-LOW").text = storage_type
        except Exception:
            pass
        try:
            session.findById("wnd[1]/usr/txtS_OWNER-LOW").text = owner
        except Exception:
            pass

        # 5. Nạp danh sách Material qua Multiple Selection
        unique_mats = list(dict.fromkeys([m.strip() for m in materials if m.strip()]))
        self.copy_to_clipboard("\r\n".join(unique_mats))

        session.findById("wnd[1]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        try:
            session.findById("wnd[2]/tbar[0]/btn[24]").press()  # Nút Upload từ Clipboard
        except Exception:
            # Fallback nếu không có btn[24]
            pass
        session.findById("wnd[2]/tbar[0]/btn[8]").press()  # Xác nhận F8

        # 6. Execute (F8)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # 7. Đọc kết quả từ bảng ALV Grid
        grid = session.findById("wnd[0]/usr/shell/shellcont[1]/shell")
        row_count = getattr(grid, "RowCount", 0)
        stock_records: List[Dict[str, Any]] = []

        if row_count <= 0:
            return stock_records

        # Nhận diện cột kỹ thuật
        col_order = grid.ColumnOrder
        col_map = {}
        for i in range(col_order.Count):
            c_name = col_order.ElementAt(i)
            try:
                c_title = (grid.GetColumnTitles(c_name).ElementAt(0) or "").strip().lower()
            except Exception:
                c_title = ""
            
            if "material" in c_title or c_name == "MATNR":
                col_map["material"] = c_name
            elif "bin" in c_title or "storage bin" in c_title or c_name == "LGPLA":
                col_map["bin"] = c_name
            elif "batch" in c_title or c_name in ("CHARG", "BATCH"):
                col_map["batch"] = c_name
            elif "available" in c_title or "avail" in c_title or c_name in ("ABSQUAN", "AVAILABLE_QTY"):
                col_map["available_qty"] = c_name
            elif "quantity" in c_title and "available" not in c_title and "out" not in c_title and "in" not in c_title:
                col_map["total_qty"] = c_name
            elif "outgoing" in c_title or c_name == "OUT_QUAN":
                col_map["outgoing_qty"] = c_name
            elif "date" in c_title or "goods receipt" in c_title or c_name == "UI_WDATU":
                col_map["gr_date"] = c_name
            elif "time" in c_title or c_name == "UI_WDATT":
                col_map["gr_time"] = c_name

        # Quét từng dòng
        for r in range(row_count):
            try:
                mat = str(grid.GetCellValue(r, col_map.get("material", "MATNR"))).strip()
                bin_loc = str(grid.GetCellValue(r, col_map.get("bin", "LGPLA"))).strip()
                batch = str(grid.GetCellValue(r, col_map.get("batch", "CHARG"))).strip()
                avail_raw = str(grid.GetCellValue(r, col_map.get("available_qty", "ABSQUAN"))).strip()
                gr_date = str(grid.GetCellValue(r, col_map.get("gr_date", "UI_WDATU"))).strip()
                gr_time = str(grid.GetCellValue(r, col_map.get("gr_time", "UI_WDATT"))).strip()

                avail_qty = self.parse_sap_number(avail_raw)

                # Quy tắc Bản 1.6: Chỉ lấy dòng có lượng khả dụng > 0
                if mat and avail_qty > 0:
                    stock_records.append({
                        "material": mat,
                        "bin": bin_loc,
                        "batch": batch,
                        "available_qty": avail_qty,
                        "allocated_qty": 0.0,  # Sẽ trừ dần khi phân bổ FIFO
                        "gr_date": gr_date,
                        "gr_time": gr_time
                    })
            except Exception as ex:
                print(f"[STOCK PARSE WARN] Lỗi dòng {r}: {ex}")

        return stock_records


# ==============================================================================
# 4. THUẬT TOÁN PHÂN BỔ FIFO & CHIA NHÓM THEO ZONE
# ==============================================================================

class AllocationEngine:
    """Phân bổ tồn kho theo FIFO/FEFO và gom nhóm theo Zone (Max 5 mã, Max 20 dòng)."""

    def __init__(self, zone_mgr: ZoneManager, max_materials: int = 5, max_lines: int = 20):
        self.zone_mgr = zone_mgr
        self.max_materials = max_materials
        self.max_lines = max_lines

    def allocate(
        self,
        requests: List[Dict[str, Any]],
        stock_records: List[Dict[str, Any]]
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        Phân bổ tồn cho danh sách yêu cầu.
        Trả về: (Danh sách nhóm OB cần tạo MIGO, Danh sách dòng thiếu hàng)
        """
        # 1. Phân loại yêu cầu theo Zone
        requests_by_zone: Dict[str, List[Dict[str, Any]]] = {}
        for req in requests:
            mat = req["material"].strip().upper()
            qty = float(req["qty"])
            sta = req.get("station", "").strip()

            zone, tl = self.zone_mgr.resolve_zone_tl(mat, sta)
            req["zone"] = zone
            req["tl"] = tl

            if zone not in requests_by_zone:
                requests_by_zone[zone] = []
            requests_by_zone[zone].append(req)

        # Tạo bản sao tồn để trừ dần
        stock_pool = [s.copy() for s in stock_records]

        # Sắp xếp tồn theo FIFO: Ngày nhập tăng dần, giờ nhập tăng dần
        stock_pool.sort(key=lambda x: (x.get("gr_date", "99999999"), x.get("gr_time", "999999")))

        allocated_groups: List[Dict[str, Any]] = []
        shortage_items: List[Dict[str, Any]] = []
        group_counter = 1

        for zone, req_list in requests_by_zone.items():
            current_group_lines: List[Dict[str, Any]] = []
            current_group_mats: Set[str] = set()

            for req in req_list:
                mat = req["material"]
                needed_qty = req["qty"]
                allocated_for_req = 0.0

                # Tìm các dòng tồn khớp Material
                for st in stock_pool:
                    if st["material"].upper() == mat.upper():
                        rem_stock = st["available_qty"] - st["allocated_qty"]
                        if rem_stock > 0:
                            take_qty = min(needed_qty - allocated_for_req, rem_stock)
                            st["allocated_qty"] += take_qty
                            allocated_for_req += take_qty

                            # Thêm vào dòng nhóm
                            line_info = {
                                "material": mat,
                                "batch": st["batch"],
                                "bin": st["bin"],
                                "qty": take_qty,
                                "station": req.get("station", ""),
                                "zone": zone,
                                "tl": req.get("tl", "")
                            }

                            # Kiểm tra giới hạn 5 mã / 20 dòng
                            temp_mats = current_group_mats.union({mat})
                            if len(temp_mats) > self.max_materials or len(current_group_lines) >= self.max_lines:
                                # Đóng nhóm hiện tại
                                allocated_groups.append({
                                    "group_id": f"OB_{group_counter:04d}",
                                    "zone": zone,
                                    "tl": req.get("tl", ""),
                                    "lines": current_group_lines
                                })
                                group_counter += 1
                                current_group_lines = [line_info]
                                current_group_mats = {mat}
                            else:
                                current_group_lines.append(line_info)
                                current_group_mats.add(mat)

                            if allocated_for_req >= needed_qty:
                                break

                # Nếu không đủ tồn
                if allocated_for_req < needed_qty:
                    shortage_items.append({
                        "material": mat,
                        "station": req.get("station", ""),
                        "needed": needed_qty,
                        "allocated": allocated_for_req,
                        "missing": needed_qty - allocated_for_req,
                        "zone": zone
                    })

            # Đóng nhóm cuối của Zone nếu còn
            if current_group_lines:
                allocated_groups.append({
                    "group_id": f"OB_{group_counter:04d}",
                    "zone": zone,
                    "tl": req_list[0].get("tl", ""),
                    "lines": current_group_lines
                })
                group_counter += 1

        return allocated_groups, shortage_items


# ==============================================================================
# 5. TẠO ĐƠN MIGO 311 (TRANSFER POSTING)
# ==============================================================================

class MigoPostingService:
    """Thực thi giao dịch MIGO 311 để tạo số chứng từ xuất kho OB."""

    @staticmethod
    def _find_migo_table(session):
        try:
            tbl = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM", False)
            if tbl is not None:
                return tbl
        except Exception:
            pass
        try:
            tbl = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM", False)
            if tbl is not None:
                return tbl
        except Exception:
            pass
        return None

    @staticmethod
    def _set_cell(tbl, row: int, col: int, val: str):
        try:
            cell = tbl.GetCell(row, col)
            if cell:
                cell.Text = str(val)
                return True
        except Exception:
            pass
        return False

    @staticmethod
    def post_group(session, group_data: Dict[str, Any], plant: str = "3093", sloc: str = "3093") -> Tuple[bool, str]:
        """Tạo 1 chứng từ MIGO 311 cho 1 nhóm đơn."""
        try:
            # 1. Quay về màn hình gốc
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)

            # 2. Thiết lập Transfer Posting / Other / 311
            root = session.findById("wnd[0]")
            try:
                action_combo = root.findById("usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_HEADER:SAPLMIGO:0100/tabsTAB_HEADER/tabpHEADER_GEN/ssubSUB_HEADER_GEN:SAPLMIGO:0101/cmbGODYNPRO-ACTION", False)
                if action_combo:
                    action_combo.key = "A08"  # Transfer Posting
            except Exception:
                pass

            try:
                ref_combo = root.findById("usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_HEADER:SAPLMIGO:0100/tabsTAB_HEADER/tabpHEADER_GEN/ssubSUB_HEADER_GEN:SAPLMIGO:0101/cmbGODYNPRO-REFDOC", False)
                if ref_combo:
                    ref_combo.key = "R10"  # Other
            except Exception:
                pass

            try:
                bwart_fld = root.findById("usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_HEADER:SAPLMIGO:0100/tabsTAB_HEADER/tabpHEADER_GEN/ssubSUB_HEADER_GEN:SAPLMIGO:0101/txtGODEFAULT_TV-BWART", False)
                if bwart_fld:
                    bwart_fld.text = "311"
            except Exception:
                pass

            # Header text (OB Group / Zone info)
            try:
                session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/subSUB_HEADER:SAPLMIGO:0100/tabsTAB_HEADER/tabpHEADER_GEN/ssubSUB_HEADER_GEN:SAPLMIGO:0101/txtGOHEAD-BKTXT").text = str(group_data.get("group_id", "ZONE"))
            except Exception:
                pass

            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)

            # 3. Điền từng dòng vào bảng MIGO
            lines = group_data.get("lines", [])
            tbl = MigoPostingService._find_migo_table(session)
            if tbl is not None:
                for idx, line in enumerate(lines):
                    mat = str(line.get("material", "")).strip()
                    qty_val = int(round(float(line.get("qty", 0))))
                    tram = str(line.get("station", "")).strip()
                    batch = str(line.get("batch", "")).strip()

                    # Cột: 3: Material, 4: Qty, 5: Text, 6: Plant, 7: SLoc, 8: SLocTo
                    MigoPostingService._set_cell(tbl, idx, 3, mat)
                    MigoPostingService._set_cell(tbl, idx, 4, str(qty_val))
                    if tram:
                        MigoPostingService._set_cell(tbl, idx, 5, tram)
                    MigoPostingService._set_cell(tbl, idx, 6, plant)
                    MigoPostingService._set_cell(tbl, idx, 7, sloc)
                    MigoPostingService._set_cell(tbl, idx, 8, sloc)

                # Enter để cập nhật dòng
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)

                # Điền Batch nếu có
                for idx, line in enumerate(lines):
                    batch = str(line.get("batch", "")).strip()
                    if batch:
                        # Thường cột Batch là cột 9 hoặc tìm tên
                        try:
                            MigoPostingService._set_cell(tbl, idx, 9, batch)
                        except Exception:
                            pass
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)

            # 4. Check và Post
            try:
                session.findById("wnd[0]/tbar[1]/btn[7]").press()  # Check
                time.sleep(0.5)
            except Exception:
                pass

            try:
                session.findById("wnd[0]/tbar[1]/btn[23]").press() # Post
                time.sleep(0.8)
            except Exception:
                try:
                    session.findById("wnd[0]/tbar[0]/btn[11]").press()
                except Exception:
                    pass

            # 5. Đọc thông báo phản hồi từ status bar
            msg = ""
            try:
                msg = session.findById("wnd[0]/sbar").text
            except Exception:
                pass

            m_doc = re.search(r"\b(101[0-9]{7}|50[0-9]{8}|[0-9]{10})\b", msg)
            if m_doc:
                doc_no = m_doc.group(1)
                return True, doc_no
            else:
                if "Material document" in msg or "chứng từ" in msg.lower() or "posted" in msg.lower():
                    # Thử lấy lại số bất kỳ 10 chữ số
                    digits = re.findall(r"\d{10}", msg)
                    if digits:
                        return True, digits[0]
                return False, f"SAP Phản hồi: {msg}"

        except Exception as ex:
            return False, f"Lỗi thực thi MIGO: {str(ex)}"


# ==============================================================================
# 6. TRA CỨU MON NỐI OB - WHO - MÃ HÀNG (BẢN 1.4 & 1.6)
# ==============================================================================

class MonRelationService:
    """Tra cứu MON để đối soát chính xác 3 cột: OB - Mã hàng - Số đơn (WHO)."""

    @staticmethod
    def query_who_for_ob(session, ob_number: str) -> Dict[str, List[str]]:
        """
        Tra MON theo OB. Trả về: Dict[Material, List[WHO]]
        Tránh bẫy Cartesian cross-join: Chỉ gắn WHO cho đúng Material chứa nó!
        """
        clean_ob = re.sub(r"[^0-9]", "", str(ob_number)).strip()
        mat_who_map: Dict[str, List[str]] = {}

        if not clean_ob:
            return mat_who_map

        try:
            # 1. Mở MON
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00077")

            # 2. Chọn nhánh Outbound Delivery Order
            mon_tree = session.findById("wnd[0]/usr/shell/shellcont[0]/shell")
            mon_tree.selectedNode = "N000000010"
            mon_tree.doubleClickNode("N000000010")

            # 3. Lọc theo OB và F8
            session.findById("wnd[1]/usr/txtS_ERP_H-LOW").text = clean_ob
            session.findById("wnd[1]/tbar[0]/btn[8]").press()

            # 4. Bấm nút [Warehouse Order]
            top_grid = session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
            top_grid.pressToolbarButton("N000000108")

            # 5. Đọc lưới dưới lấy WHO và MATNR
            bottom_grid = session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
            row_count = getattr(bottom_grid, "RowCount", 0)

            for r in range(row_count):
                try:
                    who_val = str(bottom_grid.GetCellValue(r, "WHO")).strip()
                    mat_val = str(bottom_grid.GetCellValue(r, "MATNR")).strip()

                    if who_val.startswith("6") and mat_val:
                        if mat_val not in mat_who_map:
                            mat_who_map[mat_val] = []
                        if who_val not in mat_who_map[mat_val]:
                            mat_who_map[mat_val].append(who_val)
                except Exception:
                    pass

        except Exception as e:
            print(f"[MON REL WARN] Lỗi tra cứu liên kết cho OB {ob_number}: {e}")

        return mat_who_map


# ==============================================================================
# 7. IN PHIẾU SAP & WIN32 WATCHDOG (BẢN 1.1 - 1.6)
# ==============================================================================

class WhoPrintService:
    """In phiếu Warehouse Order qua F00112 (2 bản) và tự động bắt Dialog lưu PDF."""

    @staticmethod
    def _save_pdf_watchdog(pdf_save_path: str, timeout_sec: float = 20.0) -> bool:
        """Luồng phụ Win32 tự động bắt Dialog 'Save Print Output As' và lưu PDF."""
        if win32gui is None:
            return False

        end_time = time.time() + timeout_sec
        while time.time() < end_time:
            time.sleep(0.3)
            # Tìm cửa sổ Save Print Output As (hoặc Lưu Kết quả In Dưới dạng...)
            hwnd = win32gui.FindWindow("#32770", "Save Print Output As")
            if not hwnd:
                hwnd = win32gui.FindWindow("#32770", "Save As")

            if hwnd and win32gui.IsWindowVisible(hwnd):
                try:
                    # Ô nhập tên file trong dialog Win32 chuẩn là Edit 1152 hoặc ComboBox
                    edit_file = win32gui.GetDlgItem(hwnd, 1152)
                    if edit_file:
                        win32gui.SendMessage(edit_file, win32con.WM_SETTEXT, 0, pdf_save_path)
                        time.sleep(0.1)

                    btn_save = win32gui.GetDlgItem(hwnd, win32con.IDOK)
                    if btn_save:
                        win32gui.SendMessage(btn_save, win32con.BM_CLICK, 0, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, btn_save)
                        print(f"[PRINT WATCHDOG] Đã tự lưu PDF: {pdf_save_path}")
                        return True
                except Exception as ex:
                    print(f"[PRINT WATCHDOG ERR]: {ex}")

        return False

    @staticmethod
    def print_whos(
        session,
        who_list: List[str],
        warehouse_num: str = "VF02",
        copies: int = 2,
        pdf_output_path: Optional[str] = None
    ) -> bool:
        """In tập mã WHO qua F00112 với 2 bản in."""
        if not who_list:
            return False

        # 1. Về gốc Easy Access
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        # 2. Điều hướng mở F00112
        main_tree = session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
        main_tree.selectedNode = "F00010"
        main_tree.doubleClickNode("F00010")
        main_tree.selectedNode = "F00112"
        main_tree.doubleClickNode("F00112")

        # 3. Điền Warehouse Number
        session.findById("wnd[0]/usr/ctxtP_LGNUM").text = warehouse_num

        # 4. Nạp WHO qua Clipboard
        session.findById("wnd[0]/usr/btn%_S_WHO_%_APP_%-VALU_PUSH").press()
        MonStockService.copy_to_clipboard("\r\n".join(who_list))
        try:
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
        except Exception:
            pass
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # 5. F8 Chạy chương trình in
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # 6. Đặt số bản in trên SAP = 2
        try:
            session.findById("wnd[1]/usr/ctxtSSFPP-TDCOPIES").text = str(copies)
        except Exception:
            pass

        # 7. Kích hoạt Watchdog nếu có đường dẫn lưu PDF
        if pdf_output_path:
            os.makedirs(os.path.dirname(pdf_output_path), exist_ok=True)
            watchdog = threading.Thread(
                target=WhoPrintService._save_pdf_watchdog,
                args=(pdf_output_path, 20.0),
                daemon=True
            )
            watchdog.start()

        # 8. Bấm Print btn[86]
        session.findById("wnd[1]/tbar[0]/btn[86]").press()
        return True


# ==============================================================================
# 8. XUẤT BẢNG KẾT QUẢ EXCEL 3 CỘT (BẢN 1.4)
# ==============================================================================

class ExcelResultExporter:
    """Xuất file Excel KET_QUA chuẩn 3 cột: OB - Mã hàng - Số đơn (WHO)."""

    @staticmethod
    def export(records: List[Dict[str, str]], output_path: str) -> bool:
        if openpyxl is None:
            return False

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "KET_QUA"

            # Header
            headers = ["Số OB", "Mã hàng", "Số đơn (WHO)"]
            ws.append(headers)

            header_fill = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
            header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            for col in range(1, 4):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            # Dữ liệu
            for r_idx, r in enumerate(records, start=2):
                ws.cell(row=r_idx, column=1, value=str(r.get("ob", "")).strip())
                ws.cell(row=r_idx, column=2, value=str(r.get("material", "")).strip())
                ws.cell(row=r_idx, column=3, value=str(r.get("who", "")).strip())

                for c in range(1, 4):
                    cell = ws.cell(row=r_idx, column=c)
                    cell.number_format = "@"  # Text format
                    cell.border = border
                    cell.font = Font(name="Segoe UI", size=10)
                    if c == 1 or c == 3:
                        cell.alignment = Alignment(horizontal="center")

            # Độ rộng cột
            ws.column_dimensions["A"].width = 18
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 30

            ws.freeze_panes = "A2"
            wb.save(output_path)
            return True

        except Exception as e:
            print(f"[EXCEL EXPORT ERROR] {e}")
            return False


# ==============================================================================
# 9. PHÂN TÍCH DỮ LIỆU DÁN & ĐIỀU PHỐI QUY TRÌNH (PIPELINE)
# ==============================================================================

def parse_paste_data(raw_text: str) -> List[Dict[str, Any]]:
    """
    Phân tích dữ liệu copy từ Excel hoặc văn bản dán vào ô nhập liệu.
    Hỗ trợ định dạng:
      Material <TAB> Qty <TAB> Station
      Material <TAB> Qty
      Station  <TAB> Material <TAB> Qty
    """
    items: List[Dict[str, Any]] = []
    if not raw_text or not raw_text.strip():
        return items

    lines = raw_text.strip().splitlines()
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue

        # Tách cột ưu tiên TAB, rồi phẩy hoặc chấm phẩy
        if "\t" in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split("\t")]
        elif ";" in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split(";")]
        elif "," in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split(",")]
        else:
            parts = [p.strip() for p in cleaned_line.split()]

        # Lọc bỏ phần tử rỗng
        parts = [p for p in parts if p]
        if not parts:
            continue

        # Kiểm tra tiêu đề header
        header_keywords = ["mã", "material", "vật tư", "số lượng", "qty", "trạm", "station", "stt", "item"]
        first_lower = parts[0].lower()
        if any(hk in first_lower for hk in header_keywords) and len(parts) >= 2:
            second_lower = parts[1].lower()
            if any(hk in second_lower for hk in header_keywords):
                continue

        mat = ""
        qty = 0.0
        station = ""

        def try_parse_num(s: str) -> Optional[float]:
            try:
                # Thay dấu phẩy bằng chấm nếu là số thập phân
                s_clean = s.replace(" ", "").replace(",", ".")
                val = float(s_clean)
                return val if val > 0 else None
            except Exception:
                return None

        if len(parts) == 1:
            continue
        elif len(parts) == 2:
            n1 = try_parse_num(parts[1])
            n0 = try_parse_num(parts[0])
            if n1 is not None:
                mat = parts[0]
                qty = n1
            elif n0 is not None:
                qty = n0
                mat = parts[1]
        else:
            # 3 cột trở lên
            n1 = try_parse_num(parts[1])
            n2 = try_parse_num(parts[2])
            n0 = try_parse_num(parts[0])

            if n1 is not None:
                # Col 0: Mat, Col 1: Qty, Col 2: Station
                mat = parts[0]
                qty = n1
                station = parts[2]
            elif n2 is not None:
                # Col 0: Station, Col 1: Mat, Col 2: Qty
                station = parts[0]
                mat = parts[1]
                qty = n2
            elif n0 is not None:
                qty = n0
                mat = parts[1]
                station = parts[2]
            else:
                mat = parts[0]
                station = parts[2]

        if mat and qty > 0:
            items.append({
                "material": mat.strip().upper(),
                "qty": qty,
                "station": station.strip()
            })

    return items


class SapZonePipeline:
    """Điều phối toàn diện luồng: Đọc Tồn -> Phân Bổ FIFO -> MIGO 311 -> Tra cứu WHO -> In Phiếu SAP."""

    def __init__(self, config: Dict[str, Any], zone_mgr: ZoneManager):
        self.config = config
        self.zone_mgr = zone_mgr
        self.stop_requested = False

    def request_stop(self):
        self.stop_requested = True

    def run(
        self,
        requests: List[Dict[str, Any]],
        progress_cb=None,
        status_cb=None,
        is_mock: bool = False
    ) -> Dict[str, Any]:
        """Thực thi toàn bộ luồng nghiệp vụ."""
        self.stop_requested = False
        summary: Dict[str, Any] = {
            "success": False,
            "records": [],
            "groups": [],
            "shortages": [],
            "pdf_files": [],
            "message": ""
        }

        def update_progress(pct: int, msg: str):
            if progress_cb:
                progress_cb(pct)
            if status_cb:
                status_cb(msg)
            print(f"[{pct}%] {msg}")

        if not requests:
            summary["message"] = "Không có danh sách yêu cầu hợp lệ!"
            return summary

        # -------------------------------------------------------------
        # BƯỚC 1: KẾT NỐI SAP
        # -------------------------------------------------------------
        update_progress(5, "Đang kết nối SAP GUI Scripting...")
        session = None
        if not is_mock:
            try:
                session = SapConnector.get_session(self.config.get("sap_session_idx", 0))
            except Exception as e:
                summary["message"] = f"Lỗi kết nối SAP: {e}"
                update_progress(0, f"Lỗi kết nối SAP: {e}")
                return summary
        else:
            update_progress(10, "[CHẾ ĐỘ MÔ PHỎNG] Đã khởi tạo phiên làm việc giả lập.")

        if self.stop_requested:
            summary["message"] = "Đã dừng bởi người dùng."
            return summary

        # -------------------------------------------------------------
        # BƯỚC 2: ĐỌC TỒN TỪ /SCWM/MON (APFL + 3093)
        # -------------------------------------------------------------
        update_progress(20, "Đang đọc tồn kho khả dụng từ /SCWM/MON (APFL, Owner 3093)...")
        unique_materials = list(dict.fromkeys([r["material"] for r in requests]))
        stock_records: List[Dict[str, Any]] = []

        if not is_mock and session:
            try:
                mon_svc = MonStockService()
                stock_records = mon_svc.fetch_stock_from_mon(
                    session=session,
                    materials=unique_materials,
                    storage_type=self.config.get("ewm_storage_type", "APFL"),
                    owner=self.config.get("ewm_owner", "3093")
                )
            except Exception as e:
                summary["message"] = f"Lỗi đọc tồn MON: {e}"
                update_progress(20, f"Lỗi đọc tồn MON: {e}")
                return summary
        else:
            # Mô phỏng tồn dồi dào cho mock mode
            now_str = datetime.now().strftime("%Y%m%d")
            for m in unique_materials:
                stock_records.append({
                    "material": m,
                    "bin": "APFL-01-A",
                    "batch": "B20261101",
                    "available_qty": 50000.0,
                    "allocated_qty": 0.0,
                    "gr_date": now_str,
                    "gr_time": "080000"
                })

        if self.stop_requested:
            summary["message"] = "Đã dừng bởi người dùng."
            return summary

        # -------------------------------------------------------------
        # BƯỚC 3: PHÂN BỔ TỒN FIFO & GOM NHÓM THEO ZONE
        # -------------------------------------------------------------
        update_progress(40, "Đang phân bổ tồn FIFO/FEFO và gom nhóm theo Zone (Max 5 mã, Max 20 dòng)...")
        alloc_engine = AllocationEngine(
            zone_mgr=self.zone_mgr,
            max_materials=int(self.config.get("max_materials_per_group", 5)),
            max_lines=int(self.config.get("max_lines_per_group", 20))
        )
        allocated_groups, shortage_items = alloc_engine.allocate(requests, stock_records)
        summary["groups"] = allocated_groups
        summary["shortages"] = shortage_items

        if not allocated_groups:
            summary["message"] = "Không thể phân bổ được nhóm đơn nào (hết hàng hoặc không đủ tồn)!"
            update_progress(40, summary["message"])
            return summary

        if self.stop_requested:
            summary["message"] = "Đã dừng bởi người dùng."
            return summary

        # -------------------------------------------------------------
        # BƯỚC 4: TẠO ĐƠN MIGO 311 CHO TỪNG NHÓM
        # -------------------------------------------------------------
        update_progress(55, f"Đang tạo MIGO 311 cho {len(allocated_groups)} nhóm đơn...")
        posted_obs: List[Dict[str, Any]] = []  # [{group, ob_number}]

        for idx, grp in enumerate(allocated_groups):
            if self.stop_requested:
                break
            grp_id = grp["group_id"]
            update_progress(55 + int(15 * idx / len(allocated_groups)), f"Đang post MIGO 311 nhóm {grp_id}...")

            if not is_mock and session:
                ok, doc_or_err = MigoPostingService.post_group(
                    session=session,
                    group_data=grp,
                    plant=self.config.get("plant", "3093"),
                    sloc=self.config.get("sloc", "3093")
                )
                if ok:
                    grp["ob_number"] = doc_or_err
                    posted_obs.append({"group": grp, "ob_number": doc_or_err})
                else:
                    print(f"[MIGO POST FAIL] {grp_id}: {doc_or_err}")
            else:
                mock_ob = f"50{datetime.now().strftime('%y%m%d')}{idx+1:02d}"
                grp["ob_number"] = mock_ob
                posted_obs.append({"group": grp, "ob_number": mock_ob})
                time.sleep(0.3)

        if not posted_obs:
            summary["message"] = "Không tạo được chứng từ MIGO nào thành công!"
            update_progress(70, summary["message"])
            return summary

        # -------------------------------------------------------------
        # BƯỚC 5: TRA CỨU MON NỐI OB - WHO - MÃ HÀNG (TRÁNH CROSS-JOIN)
        # -------------------------------------------------------------
        update_progress(75, "Đang tra cứu MON lấy liên kết OB - Số đơn WHO - Mã hàng...")
        final_3col_records: List[Dict[str, str]] = []
        all_whos_to_print: List[str] = []

        for p_item in posted_obs:
            ob_num = p_item["ob_number"]
            grp = p_item["group"]
            lines = grp.get("lines", [])

            mat_who_map: Dict[str, List[str]] = {}
            if not is_mock and session:
                mat_who_map = MonRelationService.query_who_for_ob(session, ob_num)
            else:
                # Mock WHO
                for l_idx, l in enumerate(lines):
                    m_code = l["material"]
                    w1 = f"602871{l_idx*2+1:04d}"
                    w2 = f"602871{l_idx*2+2:04d}"
                    mat_who_map[m_code] = [w1, w2]

            # Xây dựng 3 cột chuẩn
            mats_in_grp = list(dict.fromkeys([l["material"] for l in lines]))
            for m_code in mats_in_grp:
                whos = mat_who_map.get(m_code, [])
                who_str = "/".join(whos) if whos else "(Chưa có WHO)"
                final_3col_records.append({
                    "ob": ob_num,
                    "material": m_code,
                    "who": who_str
                })
                for w in whos:
                    if w not in all_whos_to_print:
                        all_whos_to_print.append(w)

        summary["records"] = final_3col_records

        if self.stop_requested:
            summary["message"] = "Đã dừng bởi người dùng."
            return summary

        # -------------------------------------------------------------
        # BƯỚC 6: IN PHIẾU SAP F00112 & TỰ LƯU PDF
        # -------------------------------------------------------------
        if self.config.get("auto_print_pdf", True) and all_whos_to_print:
            update_progress(90, f"Đang in phiếu SAP F00112 (2 bản) cho {len(all_whos_to_print)} đơn WHO...")
            pdf_folder = self.config.get("pdf_output_folder", os.path.abspath("./OUTPUT_PDF"))
            os.makedirs(pdf_folder, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_filename = f"PHIEU_WHO_{timestamp}.pdf"
            pdf_path = os.path.join(pdf_folder, pdf_filename)

            if not is_mock and session:
                try:
                    WhoPrintService.print_whos(
                        session=session,
                        who_list=all_whos_to_print,
                        warehouse_num=self.config.get("ewm_warehouse", "VF02"),
                        copies=int(self.config.get("sap_print_copies", 2)),
                        pdf_output_path=pdf_path
                    )
                    summary["pdf_files"].append(pdf_path)
                except Exception as ex:
                    print(f"[PRINT EX] {ex}")
            else:
                # Tạo file PDF giả lập
                try:
                    with open(pdf_path, "w", encoding="utf-8") as f:
                        f.write(f"MOCK PDF INVOICE\nCreated: {datetime.now()}\nWHOs: {', '.join(all_whos_to_print)}\n")
                    summary["pdf_files"].append(pdf_path)
                except Exception:
                    pass

        # -------------------------------------------------------------
        # BƯỚC 7: HOÀN TẤT
        # -------------------------------------------------------------
        update_progress(100, f"Thành công! Đã xử lý {len(posted_obs)} OB, {len(final_3col_records)} dòng kết quả.")
        summary["success"] = True
        summary["message"] = "Quy trình hoàn tất thành công!"
        return summary

