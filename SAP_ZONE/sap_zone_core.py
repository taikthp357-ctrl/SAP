# -*- coding: utf-8 -*-
"""
SAP ZONE Core Engine (v1.6 Pro - Revised)
Tự động hóa toàn diện quy trình Ra đơn NVL:
- Quản lý phiên SAP GUI an toàn theo danh tính System/Client chuẩn SAP_EMAIL_TEAMS_BRIDGE v1.1.12
- Đọc tồn kho khả dụng từ /SCWM/MON (Node N000000139, Storage Type APFL, Owner 3093)
- Phân bổ FIFO/FEFO theo 4 tầng ưu tiên Zone (Ưu tiên Zone nhập vào > zone_map.json > gom nhóm theo Trạm)
- Tạo chứng từ MIGO 311 chuyển kho NVL (1005 -> 3001, Plant 3093) chuẩn VBA V31 2 nhịp
- Tra cứu /SCWM/MON (Node N000000010) lấy số đơn WHO cho từng Material, loại bỏ bẫy nhân chéo
- In phiếu SAP F00112 (2 bản) và Win32 Watchdog tự lưu & xác thực file PDF
- Sổ nhật ký trạng thái SQLite (Journal Persistence) chống mất dữ liệu khi ngắt quãng
"""

import os
import sys
import re
import time
import json
import sqlite3
import threading
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple, Set

# COM & Win32
try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32con
    import win32print
    import win32clipboard
except ImportError:
    win32com = None
    pythoncom = None
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


def get_base_dir() -> str:
    """Trả về thư mục chứa file thực thi (.exe) hoặc mã nguồn."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundle_dir() -> str:
    """Trả về thư mục bundle tài nguyên (hỗ trợ PyInstaller _MEIPASS)."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return getattr(sys, "_MEIPASS")
    return get_base_dir()


# ==============================================================================
# 1. QUẢN LÝ CẤU HÌNH & ZONE MAPPING
# ==============================================================================

DEFAULT_CONFIG = {
    "sap_system": "PSV",
    "sap_client": "908",
    "sap_session_idx": -1,
    "ewm_warehouse": "VF02",
    "ewm_storage_type": "APFL",
    "ewm_owner": "3093",
    "plant": "3093",
    "sloc_from": "1005",   # Kho xuất (1005)
    "sloc_to": "3001",     # Kho nhận (3001)
    "max_materials_per_group": 5,
    "max_lines_per_group": 20,
    "allocation_strategy": "FIFO",
    "allow_partial_allocation": True,
    "sap_print_copies": 2,
    "auto_print_pdf": True,
    "pdf_output_folder": os.path.abspath("./OUTPUT_PDF"),
    "printer_name": "Microsoft Print to PDF"
}


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
    """Quản lý ánh xạ Trạm -> Zone -> Team Leader (TL) với 4 tầng ưu tiên thông minh."""
    MAPPING_FILE = "zone_map.json"

    def __init__(self):
        self.mappings: List[Dict[str, str]] = []
        self.load_mappings()

    def get_path(self) -> str:
        # 1. Ưu tiên file bên ngoài nằm cạnh file .exe
        outer_path = os.path.join(get_base_dir(), self.MAPPING_FILE)
        if os.path.exists(outer_path):
            return outer_path
        # 2. File đóng gói sẵn trong bundle PyInstaller
        bundle_path = os.path.join(get_bundle_dir(), self.MAPPING_FILE)
        if os.path.exists(bundle_path):
            return bundle_path
        # 3. Thư mục hiện tại
        if os.path.exists(self.MAPPING_FILE):
            return os.path.abspath(self.MAPPING_FILE)
        return outer_path

    def load_mappings(self):
        path = self.get_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    raw_data = json.load(f)
                    self.mappings.clear()
                    for item in raw_data:
                        if isinstance(item, list) and len(item) >= 5:
                            self.mappings.append({
                                "car_model": str(item[0]).strip(),
                                "material": str(item[1]).strip().upper(),
                                "p_line": str(item[2]).strip(),
                                "zone": str(item[3]).strip(),
                                "tl": str(item[4]).strip()
                            })
            except Exception as e:
                print(f"[ZONE WARN] Lỗi đọc {path}: {e}")

    @staticmethod
    def extract_zone_from_station(station_str: str) -> Tuple[str, str]:
        """
        Bóc tách Zone từ chuỗi Trạm/Người nhận.
        Trả về: (zone_found, remaining_station)
        Ví dụ:
        'Hùng/Long Zone 1.1' -> ('1.1', 'Hùng/Long')
        'Zone 2.1'          -> ('2.1', '')
        'Z3'                -> ('3', '')
        '1.2'               -> ('1.2', '')
        'Hùng/Long'         -> ('', 'Hùng/Long')
        """
        if not station_str:
            return "", ""
        s = station_str.strip()
        # Tìm pattern 'Zone X.Y' hoặc 'Z X.Y' hoặc 'Zone X'
        m = re.search(r"(?:ZONE|Z)\s*([0-9]+(?:\.[0-9]+)?)", s, re.IGNORECASE)
        if m:
            z_val = m.group(1)
            cleaned = re.sub(r"\(?\s*(?:ZONE|Z)\s*[0-9]+(?:\.[0-9]+)?\s*\)?", "", s, flags=re.IGNORECASE).strip(" -/()")
            return z_val, cleaned
        # Nếu chuỗi thuần túy là số zone như '1.1' hoặc '2.2'
        if re.fullmatch(r"[0-9]+(?:\.[0-9]+)?", s):
            return s, ""
        return "", s

    def resolve_zone_tl(self, material: str, station: str = "", input_zone: str = "") -> Tuple[str, str, bool]:
        """
        Xác định Zone và Team Leader (TL) theo 4 tầng ưu tiên:
        1. Ưu tiên 1 (Input Zone): Nếu người dùng chỉ định Zone (input_zone hoặc trích xuất từ station),
           dùng ngay Zone đó làm nhóm chính của đơn hàng. Không bị zone_map.json ghi đè!
        2. Ưu tiên 2 (Catalog Map): Nếu không có input zone, tra cứu theo Material trong zone_map.json.
        3. Ưu tiên 3 (Dynamic Station): Nếu không có trong zone_map.json nhưng có tên Trạm/Người nhận
           (ví dụ: 'Hùng/Long'), tự động gom nhóm động theo Trạm (Zone = station, TL = station).
        4. Ưu tiên 4 (Default): Mặc định nhóm 'Chung' (TL = 'Kho APFL').
        
        Trả về: (zone_name, tl_name, is_mapped)
        """
        mat_clean = material.strip().upper()
        sta_clean = station.strip()
        inp_z = input_zone.strip()

        # Nếu trong chuỗi station có chứa Zone (vd: 'Hùng/Long Zone 1.1' hoặc 'Z2.1' hoặc '1.1')
        if not inp_z and sta_clean:
            extracted_z, cleaned_sta = self.extract_zone_from_station(sta_clean)
            if extracted_z:
                inp_z = extracted_z
                if cleaned_sta:
                    sta_clean = cleaned_sta

        # 1. ƯU TIÊN 1: Zone được chỉ định trực tiếp từ dữ liệu dán
        if inp_z:
            tl = sta_clean
            for m in self.mappings:
                if str(m.get("zone", "")).strip().upper() == inp_z.upper():
                    tl = m.get("tl", tl)
                    break
            return inp_z, (tl if tl else f"TL Zone {inp_z}"), True

        # 2. ƯU TIÊN 2: Tra cứu theo danh mục zone_map.json
        for m in self.mappings:
            if m.get("material", "").strip().upper() == mat_clean:
                return str(m.get("zone", "")).strip(), m.get("tl", ""), True

        # 3. ƯU TIÊN 3: Gom nhóm theo Trạm / Người nhận thực tế (vd: 'Hùng/Long')
        if sta_clean:
            return sta_clean, sta_clean, True

        # 4. ƯU TIÊN 4: Mặc định nhóm Chung
        return "Chung", "Kho APFL", False


# ==============================================================================
# 2. SỔ NHẬT KÝ SQLITE (JOURNAL PERSISTENCE)
# ==============================================================================

class JournalService:
    """Quản lý lưu vết trạng thái ra đơn, chứng từ MIGO và kết quả tra cứu."""

    @classmethod
    def get_db_path(cls) -> str:
        data_dir = os.path.join(get_base_dir(), "data")
        os.makedirs(data_dir, exist_ok=True)
        return os.path.join(data_dir, "sap_zone_journal.db")

    @classmethod
    def init_db(cls):
        path = cls.get_db_path()
        try:
            with sqlite3.connect(path) as conn:
                cur = conn.cursor()
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS runs (
                        run_id TEXT PRIMARY KEY,
                        created_at TEXT,
                        status TEXT,
                        total_requests INTEGER,
                        total_results INTEGER,
                        message TEXT
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS groups (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        run_id TEXT,
                        group_id TEXT,
                        zone TEXT,
                        tl TEXT,
                        ob_number TEXT,
                        status TEXT,
                        message TEXT
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS results (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        run_id TEXT,
                        ob_number TEXT,
                        material TEXT,
                        who TEXT
                    )
                """)
                conn.commit()
        except Exception as e:
            print(f"[SQLITE INIT ERR] {e}")

    @classmethod
    def save_run(cls, run_id: str, status: str, total_req: int, total_res: int, msg: str):
        cls.init_db()
        try:
            with sqlite3.connect(cls.get_db_path()) as conn:
                cur = conn.cursor()
                cur.execute(
                    "INSERT OR REPLACE INTO runs VALUES (?, datetime('now', 'localtime'), ?, ?, ?, ?)",
                    (run_id, status, total_req, total_res, msg)
                )
                conn.commit()
        except Exception as e:
            print(f"[SQLITE SAVE RUN ERR] {e}")

    @classmethod
    def save_group(cls, run_id: str, group_id: str, zone: str, tl: str, ob_number: str, status: str, msg: str):
        cls.init_db()
        try:
            with sqlite3.connect(cls.get_db_path()) as conn:
                cur = conn.cursor()
                cur.execute(
                    "INSERT INTO groups (run_id, group_id, zone, tl, ob_number, status, message) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (run_id, group_id, zone, tl, ob_number, status, msg)
                )
                conn.commit()
        except Exception as e:
            print(f"[SQLITE SAVE GRP ERR] {e}")

    @classmethod
    def save_results(cls, run_id: str, records: List[Dict[str, str]]):
        cls.init_db()
        try:
            with sqlite3.connect(cls.get_db_path()) as conn:
                cur = conn.cursor()
                for r in records:
                    cur.execute(
                        "INSERT INTO results (run_id, ob_number, material, who) VALUES (?, ?, ?, ?)",
                        (run_id, r.get("ob", ""), r.get("material", ""), r.get("who", ""))
                    )
                conn.commit()
        except Exception as e:
            print(f"[SQLITE SAVE RES ERR] {e}")

    @classmethod
    def get_latest_results(cls) -> List[Dict[str, str]]:
        cls.init_db()
        results = []
        try:
            with sqlite3.connect(cls.get_db_path()) as conn:
                cur = conn.cursor()
                cur.execute("SELECT run_id FROM runs ORDER BY created_at DESC LIMIT 1")
                row = cur.fetchone()
                if row:
                    latest_run_id = row[0]
                    cur.execute("SELECT ob_number, material, who FROM results WHERE run_id = ?", (latest_run_id,))
                    for ob, mat, who in cur.fetchall():
                        results.append({"ob": ob, "material": mat, "who": who})
        except Exception as e:
            print(f"[SQLITE GET RES ERR] {e}")
        return results


# ==============================================================================
# 3. KẾT NỐI SAP GUI AN TOÀN (CHUẨN BRIDGE V1.1.12 & COM APARTMENT GUARD)
# ==============================================================================

class SapConnector:
    """Quản lý kết nối tới phiên SAP GUI an toàn, ràng buộc danh tính System/Client."""

    @staticmethod
    def init_thread_com():
        """Khởi tạo căn hộ COM bắt buộc cho mọi luồng Windows."""
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

    @staticmethod
    def uninit_thread_com():
        """Giải phóng căn hộ COM khi hoàn tất luồng."""
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @staticmethod
    def find_by_id_safe(session, id_text: str):
        """Tìm control an toàn, trả về None nếu không tồn tại (không quăng ngoại lệ)."""
        if session is None:
            return None
        try:
            return session.findById(id_text, False)
        except Exception:
            try:
                return session.findById(id_text)
            except Exception:
                return None

    @staticmethod
    def wait_idle(session, timeout_sec: float = 15.0) -> bool:
        """Đợi SAP GUI hoàn tất xử lý (Busy = False) như chuẩn Bridge/VBA."""
        if session is None:
            return True
        t0 = time.time()
        while time.time() - t0 < timeout_sec:
            try:
                if not getattr(session, "Busy", False):
                    return True
            except Exception:
                return True
            time.sleep(0.1)
        return False

    @staticmethod
    def get_status(session) -> Tuple[str, str]:
        """Đọc thanh trạng thái sbar: (message_text, message_type)."""
        if session is None:
            return "", ""
        try:
            sbar = SapConnector.find_by_id_safe(session, "wnd[0]/sbar")
            if sbar is not None:
                txt = str(getattr(sbar, "Text", "")).strip()
                mtype = str(getattr(sbar, "MessageType", "")).strip().upper()
                return txt, mtype
        except Exception:
            pass
        return "", ""

    @staticmethod
    def dismiss_popups(session, max_loops: int = 8) -> str:
        """Đóng mọi cửa sổ popup dialog wnd[1] nếu đang treo."""
        if session is None:
            return ""
        all_msgs = []
        try:
            for _ in range(max_loops):
                wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]")
                if wnd1 is None:
                    break
                t = str(getattr(wnd1, "Text", "")).strip()
                if t:
                    all_msgs.append(t)
                try:
                    wnd1.sendVKey(0)  # Enter
                except Exception:
                    try:
                        btn = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[0]")
                        if btn:
                            btn.press()
                    except Exception:
                        pass
                SapConnector.wait_idle(session, 1.5)
        except Exception:
            pass
        return " | ".join(all_msgs)

    @staticmethod
    def open_transaction(session, tcode: str) -> bool:
        """Chuyển tới T-Code an toàn qua /n<tcode>."""
        if session is None:
            return False
        SapConnector.dismiss_popups(session)
        target = f"/n{tcode.strip().lstrip('/')}"
        try:
            okcd = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
            if okcd is not None:
                okcd.text = target
                session.findById("wnd[0]").sendVKey(0)
                SapConnector.wait_idle(session, 3.0)
                SapConnector.dismiss_popups(session)
                return True
        except Exception:
            pass
        try:
            session.StartTransaction(tcode)
            SapConnector.wait_idle(session, 3.0)
            SapConnector.dismiss_popups(session)
            return True
        except Exception:
            return False

    @staticmethod
    def open_mon_screen(session, warehouse: str = "VF02") -> bool:
        """
        Mở màn hình /SCWM/MON (Warehouse Monitor) an toàn:
        - Kiểm tra xem đã ở trong MON hay chưa (nếu đã ở trong MON thì tái sử dụng)
        - Thử Favorite Node F00077 nếu có
        - Fallback gọi lệnh /n/SCWM/MON qua okcd và xử lý hộp thoại chọn kho wnd[1]
        """
        if session is None:
            return False

        SapConnector.dismiss_popups(session)

        # 1. Kiểm tra nếu đã ở trong màn hình MON
        mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
        if mon_tree is not None:
            return True

        # 2. Thử mở qua favorite F00077 nếu đang ở Easy Access
        fav_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
        if fav_tree is not None:
            try:
                fav_tree.doubleClickNode("F00077")
                SapConnector.wait_idle(session, 2.0)
                if SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell") is not None:
                    return True
            except Exception:
                pass

        # 3. Mở trực tiếp bằng lệnh /n/SCWM/MON
        SapConnector.open_transaction(session, "/SCWM/MON")
        SapConnector.wait_idle(session, 2.0)

        # 4. Nếu có hộp thoại chọn Kho trên wnd[1]
        wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]")
        if wnd1 is not None:
            try:
                txt_whse = SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtP_LGNUM")
                if txt_whse is not None:
                    txt_whse.text = warehouse
                txt_mon = SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtP_MON")
                if txt_mon is not None:
                    txt_mon.text = "SAP"
                wnd1.sendVKey(0)  # Enter hoặc F8
                SapConnector.wait_idle(session, 2.0)
            except Exception:
                pass

        # 5. Xác nhận lại xem đã vào MON chưa
        return SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell") is not None

    @staticmethod
    def get_session(session_idx: int = -1, preferred_sys: str = "", preferred_client: str = ""):
        """
        Kết nối tới phiên SAP GUI an toàn:
        - Bắt buộc gọi CoInitialize cho worker thread
        - Vòng lặp retry 5 lần để đảm bảo SAP GUI đã sẵn sàng
        - Ràng buộc System/Client nếu có cấu hình
        - Dọn dẹp modal popup trước khi bàn giao session
        """
        SapConnector.init_thread_com()
        if win32com is None:
            raise RuntimeError("Thư viện pywin32 chưa được cài đặt!")

        sap_gui_auto = None
        last_ex = None
        for attempt in range(5):
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                if sap_gui_auto is not None:
                    break
            except Exception as ex:
                last_ex = ex
                time.sleep(0.4)

        if sap_gui_auto is None:
            raise RuntimeError(
                f"Không thể kết nối tiến trình SAP GUI qua COM ({last_ex})!\n"
                f"Vui lòng đảm bảo:\n"
                f"1. Ứng dụng SAP Logon đang mở trên máy tính.\n"
                f"2. Bạn đã đăng nhập vào hệ thống (Client 908 / PSV).\n"
                f"3. Tính năng Scripting đã được bật trong SAP GUI Options."
            )

        application = None
        for _ in range(3):
            try:
                application = getattr(sap_gui_auto, "GetScriptingEngine", None)
                if application is not None:
                    break
            except Exception:
                time.sleep(0.3)

        if not application:
            raise RuntimeError(
                "SAP GUI Scripting Engine bị tắt hoặc không được cấp quyền!\n"
                "Hãy vào SAP Logon -> Options -> Accessibility & Scripting -> Scripting -> Tích chọn 'Enable scripting'."
            )

        conn_count = getattr(application.Children, "Count", 0)
        if conn_count == 0:
            raise RuntimeError("Chưa có kết nối nào được mở trong SAP GUI! Hãy đăng nhập vào hệ thống SAP trước khi bấm Chạy.")

        candidates = []
        for c_i in range(conn_count):
            cand_conn = application.Children(c_i)
            s_count = getattr(cand_conn.Children, "Count", 0)
            for s_i in range(s_count):
                sess = cand_conn.Children(s_i)
                try:
                    info = getattr(sess, "Info", None)
                    s_sys = str(getattr(info, "SystemName", "") or "").strip().upper()
                    s_clnt = str(getattr(info, "Client", "") or "").strip()
                    s_user = str(getattr(info, "User", "") or "").strip()
                    s_tcode = str(getattr(info, "Transaction", "") or "").strip().upper()
                    is_busy = bool(getattr(sess, "Busy", False))
                    candidates.append({
                        "session": sess,
                        "conn_idx": c_i,
                        "sess_idx": s_i,
                        "system": s_sys,
                        "client": s_clnt,
                        "user": s_user,
                        "tcode": s_tcode,
                        "busy": is_busy
                    })
                except Exception:
                    pass

        if not candidates:
            raise RuntimeError("Không tìm thấy phiên làm việc nào đang mở trong SAP GUI!")

        pref_sys_clean = preferred_sys.strip().upper()
        pref_clnt_clean = preferred_client.strip()

        chosen_session = None

        # 1. Khớp ưu tiên: System & Client chỉ định
        if pref_sys_clean:
            for c in candidates:
                if c["system"] == pref_sys_clean and (not pref_clnt_clean or c["client"] == pref_clnt_clean) and not c["busy"]:
                    chosen_session = c["session"]
                    break
            if not chosen_session:
                for c in candidates:
                    if c["system"] == pref_sys_clean and (not pref_clnt_clean or c["client"] == pref_clnt_clean):
                        chosen_session = c["session"]
                        break

        # 2. Khớp theo session_idx nếu có chỉ định hợp lệ
        if chosen_session is None and 0 <= session_idx < len(candidates):
            chosen_session = candidates[session_idx]["session"]

        # 3. Ưu tiên tab SESSION_MANAGER (Easy Access) không bận
        if chosen_session is None:
            for c in candidates:
                if not c["busy"] and c["tcode"] in ("SESSION_MANAGER", ""):
                    chosen_session = c["session"]
                    break

        # 4. Phiên bất kỳ không bận
        if chosen_session is None:
            for c in candidates:
                if not c["busy"]:
                    chosen_session = c["session"]
                    break

        # 5. Fallback về session đầu tiên
        if chosen_session is None:
            chosen_session = candidates[0]["session"]

        # Tự động dọn dẹp các popup đang treo trên wnd[1] trước khi trả session
        SapConnector.dismiss_popups(chosen_session)
        return chosen_session

    @staticmethod
    def test_connection(session_idx: int = -1, preferred_sys: str = "", preferred_client: str = "") -> Dict[str, Any]:
        """Kiểm tra và trả về thông tin danh tính của phiên SAP GUI hiện tại."""
        SapConnector.init_thread_com()
        try:
            session = SapConnector.get_session(session_idx, preferred_sys, preferred_client)
            info = getattr(session, "Info", None)
            return {
                "ok": True,
                "system": getattr(info, "SystemName", "Unknown") if info else "Unknown",
                "client": getattr(info, "Client", "Unknown") if info else "Unknown",
                "user": getattr(info, "User", "Unknown") if info else "Unknown",
                "tcode": getattr(info, "Transaction", "") if info else "",
                "program": getattr(info, "Program", "") if info else ""
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}


# ==============================================================================
# 4. ĐỌC TỒN MON (/SCWM/MON) - NODE N000000139 (AVAILABLE STOCK)
# ==============================================================================

class MonStockService:
    """Tải và trích xuất tồn kho khả dụng từ /SCWM/MON (Node N000000139, Storage Type APFL, Owner 3093)."""

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
        if "." in clean and "," in clean:
            clean = clean.replace(".", "").replace(",", ".")
        elif "." in clean and "," not in clean:
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
        """Truy vấn tồn kho từ /SCWM/MON Node N000000139 với bộ lọc APFL + Owner 3093."""
        if not materials or session is None:
            return []

        # 1. Mở màn hình /SCWM/MON an toàn
        SapConnector.open_mon_screen(session)

        # 2. Chọn node tồn kho N000000139 (Stock Overview / Available Stock)
        mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
        if mon_tree is None:
            raise RuntimeError("Không tìm thấy cây điều hướng của /SCWM/MON!")

        try:
            mon_tree.selectedNode = "N000000139"
            mon_tree.doubleClickNode("N000000139")
            SapConnector.wait_idle(session, 2.0)
        except Exception as e:
            raise RuntimeError(f"Lỗi chọn node tồn kho N000000139: {e}")

        # 3. Điền Storage Type (APFL) và Owner (3093) trên wnd[1]
        try:
            txt_lgtyp = SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtS_LGTYP-LOW")
            if txt_lgtyp:
                txt_lgtyp.text = storage_type
        except Exception:
            pass

        try:
            txt_owner = SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtS_OWNER-LOW")
            if txt_owner:
                txt_owner.text = owner
        except Exception:
            pass

        # 4. Nạp danh sách Material qua Multiple Selection
        unique_mats = list(dict.fromkeys([m.strip() for m in materials if m.strip()]))
        self.copy_to_clipboard("\r\n".join(unique_mats))

        btn_mat_push = SapConnector.find_by_id_safe(session, "wnd[1]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH")
        if btn_mat_push:
            btn_mat_push.press()
            SapConnector.wait_idle(session, 1.0)
            try:
                btn_paste = SapConnector.find_by_id_safe(session, "wnd[2]/tbar[0]/btn[24]")
                if btn_paste:
                    btn_paste.press()
            except Exception:
                pass
            try:
                btn_f8_wnd2 = SapConnector.find_by_id_safe(session, "wnd[2]/tbar[0]/btn[8]")
                if btn_f8_wnd2:
                    btn_f8_wnd2.press()
            except Exception:
                pass
            SapConnector.wait_idle(session, 1.0)

        # 5. Execute F8 trên wnd[1]
        btn_f8_wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
        if btn_f8_wnd1:
            btn_f8_wnd1.press()
            SapConnector.wait_idle(session, 4.0)

        # 6. Đọc kết quả từ bảng ALV Grid
        grid = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell")
        if grid is None:
            return []

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
            elif "batch" in c_title or c_name == "CHARG":
                col_map["batch"] = c_name
            elif "storage bin" in c_title or c_name == "LGPLA":
                col_map["bin"] = c_name
            elif "available" in c_title or c_name in ("ABSQUAN", "AVAILABLE_QTY"):
                col_map["available_qty"] = c_name
            elif "gr date" in c_title or c_name == "WDATU":
                col_map["gr_date"] = c_name
            elif "gr time" in c_title or c_name == "WZEIT":
                col_map["gr_time"] = c_name
            elif "unit" in c_title or c_name in ("MEINS", "UNIT"):
                col_map["unit"] = c_name

        mat_col = col_map.get("material", "MATNR")
        batch_col = col_map.get("batch", "CHARG")
        bin_col = col_map.get("bin", "LGPLA")
        avail_col = col_map.get("available_qty", "ABSQUAN")
        gr_date_col = col_map.get("gr_date", "WDATU")
        gr_time_col = col_map.get("gr_time", "WZEIT")
        unit_col = col_map.get("unit", "MEINS")

        for r in range(row_count):
            try:
                m_val = str(grid.GetCellValue(r, mat_col)).strip().upper()
                b_val = str(grid.GetCellValue(r, batch_col)).strip()
                bin_val = str(grid.GetCellValue(r, bin_col)).strip()
                avail_raw = str(grid.GetCellValue(r, avail_col)).strip()
                avail_qty = self.parse_sap_number(avail_raw)

                if not m_val or avail_qty <= 0:
                    continue

                d_val = re.sub(r"[^0-9]", "", str(grid.GetCellValue(r, gr_date_col)).strip())
                t_val = re.sub(r"[^0-9]", "", str(grid.GetCellValue(r, gr_time_col)).strip())
                sort_key = f"{d_val.zfill(8)}_{t_val.zfill(6)}"

                u_val = str(grid.GetCellValue(r, unit_col)).strip() or "PC"

                stock_records.append({
                    "material": m_val,
                    "batch": b_val,
                    "bin": bin_val,
                    "available_qty": avail_qty,
                    "allocated_qty": 0.0,
                    "sort_key": sort_key,
                    "unit": u_val
                })
            except Exception:
                pass

        return stock_records


# ==============================================================================
# 5. THUẬT TOÁN PHÂN BỔ FIFO & GOM NHÓM ZONE (MAX 5 MÃ, MAX 20 DÒNG)
# ==============================================================================

class AllocationEngine:
    """Phân bổ tồn kho theo FIFO và gom nhóm theo Zone (Max 5 mã, Max 20 dòng)."""

    def __init__(self, zone_mgr: ZoneManager, max_materials: int = 5, max_lines: int = 20):
        self.zone_mgr = zone_mgr
        self.max_materials = max_materials
        self.max_lines = max_lines

    def allocate(
        self,
        requests: List[Dict[str, Any]],
        stock_records: List[Dict[str, Any]]
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        Phân bổ tồn cho danh sách yêu cầu theo 4 tầng ưu tiên Zone.
        Trả về: (allocated_groups, shortage_items, unmapped_items)
        """
        requests_by_zone: Dict[str, List[Dict[str, Any]]] = {}
        unmapped_items: List[Dict[str, Any]] = []

        for req in requests:
            mat = req["material"].strip().upper()
            qty = float(req["qty"])
            sta = req.get("station", "").strip()
            inp_z = req.get("input_zone", req.get("zone", "")).strip()

            # Phân giải Zone theo 4 tầng ưu tiên
            zone, tl, is_mapped = self.zone_mgr.resolve_zone_tl(mat, sta, inp_z)
            req["zone"] = zone
            req["tl"] = tl
            req["is_mapped"] = is_mapped

            if not is_mapped and zone == "Chung":
                unmapped_items.append(req)

            if zone not in requests_by_zone:
                requests_by_zone[zone] = []
            requests_by_zone[zone].append(req)

        # Sắp xếp tồn theo FIFO dựa vào sort_key (ngày/giờ nhập)
        stock_pool = [s.copy() for s in stock_records]
        stock_pool.sort(key=lambda x: x.get("sort_key", "99999999_999999"))

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

                for st in stock_pool:
                    if st["material"].upper() == mat.upper():
                        rem_stock = st["available_qty"] - st["allocated_qty"]
                        if rem_stock > 0:
                            take_qty = min(needed_qty - allocated_for_req, rem_stock)
                            st["allocated_qty"] += take_qty
                            allocated_for_req += take_qty

                            line_info = {
                                "material": mat,
                                "batch": st["batch"],
                                "bin": st["bin"],
                                "qty": take_qty,
                                "unit": st.get("unit", "PC"),
                                "station": req.get("station", ""),
                                "zone": zone,
                                "tl": req.get("tl", "")
                            }

                            temp_mats = current_group_mats.union({mat})
                            if len(temp_mats) > self.max_materials or len(current_group_lines) >= self.max_lines:
                                allocated_groups.append({
                                    "group_id": f"GRP_{group_counter:03d}",
                                    "zone": zone,
                                    "tl": req.get("tl", ""),
                                    "lines": current_group_lines
                                })
                                group_counter += 1
                                current_group_lines = [line_info]
                                current_group_mats = {mat}
                            else:
                                current_group_lines.append(line_info)
                                current_group_mats = temp_mats

                            if allocated_for_req >= needed_qty:
                                break

                if allocated_for_req < needed_qty:
                    shortage_items.append({
                        "material": mat,
                        "required": needed_qty,
                        "allocated": allocated_for_req,
                        "shortage": needed_qty - allocated_for_req,
                        "zone": zone,
                        "station": req.get("station", "")
                    })

            if current_group_lines:
                allocated_groups.append({
                    "group_id": f"GRP_{group_counter:03d}",
                    "zone": zone,
                    "tl": req_list[0].get("tl", ""),
                    "lines": current_group_lines
                })
                group_counter += 1

        return allocated_groups, shortage_items, unmapped_items


# ==============================================================================
# 6. TẠO ĐƠN MIGO 311 (KHO XUẤT 1005 -> KHO NHẬN 3001, PLANT 3093)
# ==============================================================================

class MigoPostingService:
    """
    Thực thi giao dịch MIGO 311 chuyển kho NVL (1005 -> 3001, Plant 3093).
    Tuân thủ chuẩn xác quy trình 2 nhịp từ 260810_RA_DON_NVL_V31_ZONE:
    - Nhịp 1: Điền Material, Qty, Text (Trạm), Plant, SLoc xuất (1005), SLoc nhận (3001) -> Enter
    - Nhịp 2: Tìm cột Batch thật, điền Batch -> Enter
    - Post chứng từ và kiểm soát popup / lỗi
    """

    @staticmethod
    def _find_migo_table(session):
        root = session.findById("wnd[0]")
        for sub in ["0004", "0003", "0002", "0001"]:
            for itemlist in ["0200", "0210", "0205"]:
                try:
                    cand = f"usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_ITEMLIST:SAPLMIGO:{itemlist}/tblSAPLMIGOTV_GOITEM"
                    t = root.findById(cand, False)
                    if t is not None:
                        return t
                except Exception:
                    pass
        try:
            session.SendCommand("MIGO_OK_HEADER_CLOSE")
            time.sleep(0.2)
            session.SendCommand("MIGO_OK_DETAIL_CLOSE")
            time.sleep(0.3)
            root = session.findById("wnd[0]")
            for sub in ["0004", "0003"]:
                t = root.findById(f"usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM", False)
                if t is not None:
                    return t
        except Exception:
            pass
        return None

    @staticmethod
    def _set_cell(tbl, row: int, col: int, val: str) -> bool:
        try:
            cell = tbl.GetCell(row, col)
            if cell:
                cell.Text = str(val)
                return True
        except Exception:
            pass
        return False

    @staticmethod
    def post_group(
        session,
        group_data: Dict[str, Any],
        plant: str = "3093",
        source_sloc: str = "1005",
        dest_sloc: str = "3001"
    ) -> Tuple[bool, str]:
        """Tạo 1 chứng từ MIGO 311 cho 1 nhóm đơn."""
        try:
            # 1. Quay về màn hình MIGO
            SapConnector.dismiss_popups(session)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
            session.findById("wnd[0]").sendVKey(0)
            SapConnector.wait_idle(session, 2.0)

            root = session.findById("wnd[0]")

            # 2. Thiết lập Transfer Posting (A08) / Other (R10) / Movement Type (311)
            try:
                for sub in ["0001", "0002", "0003", "0004"]:
                    act = root.findById(f"usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION", False)
                    if act:
                        act.Key = "A08"
                        ref = root.findById(f"usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-REFDOC", False)
                        if ref:
                            ref.Key = "R10"
                        b_txt = root.findById(f"usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGOHEAD-BWART", False)
                        if b_txt:
                            b_txt.Text = "311"
                        root.sendVKey(0)
                        SapConnector.wait_idle(session, 1.5)
                        break
            except Exception:
                pass

            # 3. Điền Header Text = Tên Trạm / Zone
            station_text = group_data.get("tl", "") or group_data.get("zone", "NVL")
            try:
                for sub in ["0001", "0002", "0003", "0004"]:
                    txt_hdr = root.findById(f"usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_HEADER:SAPLMIGO:0101/txtGOHEAD-BKTXT", False)
                    if txt_hdr:
                        txt_hdr.Text = station_text[:25]
                        break
            except Exception:
                pass

            # 4. Tìm bảng chi tiết
            tbl = MigoPostingService._find_migo_table(session)
            if not tbl:
                return False, "Không tìm thấy bảng lưới nhập liệu MIGO (tblSAPLMIGOTV_GOITEM)!"

            # Nhận diện cột
            col_map = {}
            for c_idx in range(tbl.Columns.Count):
                col_name = tbl.Columns.ElementAt(c_idx).Name
                if "MATNR" in col_name:
                    col_map["material"] = c_idx
                elif "ERFMG" in col_name:
                    col_map["qty"] = c_idx
                elif "WERKS" in col_name and "plant" not in col_map:
                    col_map["plant"] = c_idx
                elif "LGORT" in col_map.get("sloc_from", -1) and "sloc_to" not in col_map:
                    col_map["sloc_to"] = c_idx
                elif "LGORT" in col_name and "sloc_from" not in col_map:
                    col_map["sloc_from"] = c_idx
                elif "CHARG" in col_name and "batch" not in col_map:
                    col_map["batch"] = c_idx
                elif "SGTXT" in col_name:
                    col_map["line_text"] = c_idx

            lines = group_data.get("lines", [])

            # Phase 1: Điền dữ liệu chung từng dòng
            for idx, item in enumerate(lines):
                if not MigoPostingService._set_cell(tbl, idx, col_map.get("material", 1), item["material"]):
                    return False, f"Không thể ghi Material vào dòng {idx+1}!"
                MigoPostingService._set_cell(tbl, idx, col_map.get("qty", 4), str(item["qty"]))
                if "plant" in col_map:
                    MigoPostingService._set_cell(tbl, idx, col_map["plant"], plant)
                if "sloc_from" in col_map:
                    MigoPostingService._set_cell(tbl, idx, col_map["sloc_from"], source_sloc)
                if "sloc_to" in col_map:
                    MigoPostingService._set_cell(tbl, idx, col_map["sloc_to"], dest_sloc)
                if "line_text" in col_map:
                    MigoPostingService._set_cell(tbl, idx, col_map["line_text"], item.get("station", station_text)[:50])

            root.sendVKey(0)  # Enter Phase 1
            SapConnector.wait_idle(session, 2.0)
            SapConnector.dismiss_popups(session)

            # Phase 2: Điền Batch
            tbl = MigoPostingService._find_migo_table(session)
            if tbl and "batch" in col_map:
                for idx, item in enumerate(lines):
                    if item.get("batch"):
                        MigoPostingService._set_cell(tbl, idx, col_map["batch"], item["batch"])
                root.sendVKey(0)  # Enter Phase 2
                SapConnector.wait_idle(session, 1.5)
                SapConnector.dismiss_popups(session)

            # 5. Check & Post chứng từ
            try:
                btn_check = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[1]/btn[7]")
                if btn_check:
                    btn_check.press()
                    SapConnector.wait_idle(session, 2.0)
                    SapConnector.dismiss_popups(session)
            except Exception:
                pass

            btn_post = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/btn[11]")
            if btn_post:
                btn_post.press()
                SapConnector.wait_idle(session, 4.0)

            # Đóng popup cảnh báo nếu có
            SapConnector.dismiss_popups(session)

            # 6. Đọc kết quả từ thanh trạng thái sbar
            status_txt, msg_type = SapConnector.get_status(session)
            m_doc = re.search(r"\b(50\d{8}|49\d{8})\b", status_txt)
            if m_doc:
                doc_no = m_doc.group(1)
                return True, doc_no
            elif msg_type == "E":
                return False, f"Lỗi SAP MIGO: {status_txt}"
            elif "posted" in status_txt.lower() or "đã ghi sổ" in status_txt.lower():
                m_any = re.search(r"\b(\d{10})\b", status_txt)
                return True, m_any.group(1) if m_any else "POSTED_OK"
            else:
                return False, f"Phản hồi không xác định: {status_txt or 'Không có thông báo từ SAP'}"

        except Exception as ex:
            return False, f"Lỗi ngoại lệ MIGO: {ex}"


# ==============================================================================
# 7. TRA CỨU MON NỐI OB - WHO - MÃ HÀNG (TRÁNH BẪY CROSS-JOIN)
# ==============================================================================

class MonRelationService:
    """Tra cứu MON để đối soát chính xác 3 cột: OB - Mã hàng - Số đơn (WHO)."""

    @staticmethod
    def query_who_for_ob(session, ob_number: str) -> Dict[str, List[str]]:
        """
        Tra MON theo OB. Trả về: Dict[Material, List[WHO]]
        Chỉ gắn WHO cho đúng Material chứa nó qua khóa dòng của Outbound Order.
        """
        clean_ob = re.sub(r"[^0-9]", "", str(ob_number)).strip()
        mat_who_map: Dict[str, List[str]] = {}

        if not clean_ob or session is None:
            return mat_who_map

        try:
            # 1. Mở màn hình /SCWM/MON an toàn
            SapConnector.open_mon_screen(session)

            # 2. Chọn nhánh Outbound Delivery Order (N000000010)
            mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
            if mon_tree:
                mon_tree.selectedNode = "N000000010"
                mon_tree.doubleClickNode("N000000010")
                SapConnector.wait_idle(session, 2.0)

            # 3. Lọc theo OB và F8 trên wnd[1]
            txt_ob = SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtS_ERP_H-LOW")
            if txt_ob:
                txt_ob.text = clean_ob
            btn_f8 = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
            if btn_f8:
                btn_f8.press()
                SapConnector.wait_idle(session, 3.0)

            # 4. Bấm nút [Warehouse Order] (N000000108)
            top_grid = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
            if top_grid:
                try:
                    top_grid.pressToolbarButton("N000000108")
                    SapConnector.wait_idle(session, 2.0)
                except Exception:
                    pass

            # 5. Đọc lưới dưới lấy WHO và MATNR
            bottom_grid = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
            if bottom_grid:
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
# 8. IN PHIẾU SAP F00112 (2 BẢN) & XÁC THỰC LƯU FILE PDF
# ==============================================================================

class WhoPrintService:
    """In phiếu Warehouse Order qua F00112 (2 bản) và Win32 Watchdog tự lưu & xác thực PDF."""

    @staticmethod
    def _save_pdf_watchdog(pdf_save_path: str, timeout_sec: float = 25.0) -> bool:
        """Luồng phụ Win32 tự động bắt Dialog 'Save Print Output As' và lưu PDF."""
        if win32gui is None:
            return False

        end_time = time.time() + timeout_sec
        while time.time() < end_time:
            time.sleep(0.3)
            hwnd = win32gui.FindWindow("#32770", "Save Print Output As")
            if not hwnd:
                hwnd = win32gui.FindWindow("#32770", "Save As")

            if hwnd and win32gui.IsWindowVisible(hwnd):
                try:
                    edit_file = win32gui.GetDlgItem(hwnd, 1152)
                    if edit_file:
                        win32gui.SendMessage(edit_file, win32con.WM_SETTEXT, 0, pdf_save_path)
                        time.sleep(0.1)

                    btn_save = win32gui.GetDlgItem(hwnd, win32con.IDOK)
                    if btn_save:
                        win32gui.SendMessage(btn_save, win32con.BM_CLICK, 0, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, btn_save)
                        break
                except Exception as ex:
                    print(f"[PRINT WATCHDOG ERR]: {ex}")

        # Xác thực file PDF có tồn tại trên đĩa và dung lượng > 500 byte
        check_end = time.time() + 10.0
        while time.time() < check_end:
            time.sleep(0.5)
            if os.path.exists(pdf_save_path) and os.path.getsize(pdf_save_path) > 500:
                print(f"[PRINT WATCHDOG] Đã xác thực file PDF: {pdf_save_path} ({os.path.getsize(pdf_save_path)} bytes)")
                return True

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
        if not who_list or session is None:
            return False

        # Khởi động thread watchdog bắt file PDF nếu có đường dẫn
        watchdog_thread = None
        if pdf_output_path:
            watchdog_thread = threading.Thread(
                target=WhoPrintService._save_pdf_watchdog,
                args=(pdf_output_path,),
                daemon=True
            )
            watchdog_thread.start()

        try:
            # 1. Về gốc Easy Access
            SapConnector.dismiss_popups(session)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            SapConnector.wait_idle(session, 1.5)

            # 2. Điều hướng mở F00112 (thử favorite F00010 -> F00112)
            main_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
            if main_tree:
                try:
                    main_tree.selectedNode = "F00010"
                    main_tree.doubleClickNode("F00010")
                    SapConnector.wait_idle(session, 1.0)
                    main_tree.selectedNode = "F00112"
                    main_tree.doubleClickNode("F00112")
                    SapConnector.wait_idle(session, 2.0)
                except Exception:
                    pass

            # 3. Điền Warehouse Number và nạp danh sách WHO
            try:
                txt_whse = SapConnector.find_by_id_safe(session, "wnd[0]/usr/ctxtP_LGNUM")
                if txt_whse:
                    txt_whse.text = warehouse_num
            except Exception:
                pass

            # Nạp WHO vào clipboard
            MonStockService.copy_to_clipboard("\r\n".join(who_list))
            btn_who_push = SapConnector.find_by_id_safe(session, "wnd[0]/usr/btn%_S_WHO_%_APP_%-VALU_PUSH")
            if btn_who_push:
                btn_who_push.press()
                SapConnector.wait_idle(session, 1.0)
                try:
                    btn_paste = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[24]")
                    if btn_paste:
                        btn_paste.press()
                except Exception:
                    pass
                btn_f8 = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
                if btn_f8:
                    btn_f8.press()
                SapConnector.wait_idle(session, 1.0)

            # 4. Thực thi (F8)
            session.findById("wnd[0]/tbar[0]/btn[8]").press()
            SapConnector.wait_idle(session, 2.0)

            # 5. Chọn toàn bộ dòng và bấm In
            alv_print = SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlGRID1/shellcont/shell")
            if alv_print:
                alv_print.selectAll()
            btn_print = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[1]/btn[13]")
            if btn_print:
                btn_print.press()
                SapConnector.wait_idle(session, 1.5)

            # 6. Thiết lập 2 bản in trên popup Print Parameters (SSFPP-TDCOPIES)
            wnd1_copies = SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtSSFPP-TDCOPIES")
            if wnd1_copies:
                wnd1_copies.text = str(copies)
            btn_confirm_print = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[86]") or SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[0]")
            if btn_confirm_print:
                btn_confirm_print.press()
                SapConnector.wait_idle(session, 2.0)

            # Chờ watchdog hoàn thành
            if watchdog_thread:
                watchdog_thread.join(timeout=30.0)

            return True
        except Exception as e:
            print(f"[PRINT EX] {e}")
            return False


# ==============================================================================
# 9. XUẤT EXCEL KẾT QUẢ 3 CỘT (SỐ OB | MÃ HÀNG | SỐ ĐƠN WHO)
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

            for r_idx, r in enumerate(records, start=2):
                ws.cell(row=r_idx, column=1, value=str(r.get("ob", "")).strip())
                ws.cell(row=r_idx, column=2, value=str(r.get("material", "")).strip())
                ws.cell(row=r_idx, column=3, value=str(r.get("who", "")).strip())
                for col in range(1, 4):
                    c = ws.cell(row=r_idx, column=col)
                    c.font = Font(name="Segoe UI", size=10)
                    c.alignment = Alignment(horizontal="center" if col != 2 else "left", vertical="center")
                    c.border = border

            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 24
            ws.column_dimensions['C'].width = 30

            wb.save(output_path)
            return True
        except Exception as e:
            print(f"[EXCEL EXPORT ERR] {e}")
            return False


# ==============================================================================
# 10. BỘ PHÂN TÍCH DỮ LIỆU DÁN EXCEL & PIPELINE ĐIỀU PHỐI TOÀN DIỆN
# ==============================================================================

def parse_paste_data(raw_text: str) -> List[Dict[str, Any]]:
    """
    Phân tích dữ liệu copy từ Excel hoặc văn bản dán.
    Hỗ trợ linh hoạt:
    - 2 cột: [Mã hàng]  [Số lượng]
    - 3 cột: [Mã hàng]  [Số lượng]  [Trạm] (hoặc ngược lại [Số lượng] [Mã hàng] [Trạm])
    - 3 cột có Zone: [Mã hàng] [Số lượng] [Zone]
    - 4 cột: [Mã hàng]  [Số lượng]  [Trạm]  [Zone]
    - 4 cột có STT: [STT]  [Mã hàng]  [Số lượng]  [Trạm/Zone]
    - 5 cột có STT: [STT]  [Mã hàng]  [Số lượng]  [Trạm]  [Zone]
    """
    items: List[Dict[str, Any]] = []
    if not raw_text or not raw_text.strip():
        return items

    def try_parse_num(s: str) -> Optional[float]:
        try:
            s_clean = s.replace(" ", "").replace(",", ".")
            val = float(s_clean)
            return val if val > 0 else None
        except Exception:
            return None

    def is_int_index(s: str) -> bool:
        try:
            v = int(s.strip())
            return 1 <= v <= 9999
        except Exception:
            return False

    header_keywords = ["mã", "material", "vật tư", "số lượng", "qty", "trạm", "station", "stt", "item", "zone"]

    lines = raw_text.strip().splitlines()
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue

        if "\t" in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split("\t")]
        elif ";" in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split(";")]
        elif "," in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split(",")]
        else:
            parts = [p.strip() for p in cleaned_line.split()]

        parts = [p for p in parts if p]
        if not parts:
            continue

        # Bỏ qua dòng tiêu đề
        if any(hk in parts[0].lower() for hk in header_keywords) and len(parts) >= 2:
            if any(hk in parts[1].lower() for hk in header_keywords):
                continue

        # Nếu cột đầu tiên là STT số thứ tự (vd: 1, 2, 3...)
        if len(parts) >= 3 and is_int_index(parts[0]) and try_parse_num(parts[1]) is None:
            parts = parts[1:]

        mat = ""
        qty = 0.0
        station = ""
        input_zone = ""

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
        elif len(parts) == 3:
            n1 = try_parse_num(parts[1])
            n2 = try_parse_num(parts[2])
            n0 = try_parse_num(parts[0])
            if n1 is not None:
                mat = parts[0]
                qty = n1
                station = parts[2]
            elif n2 is not None:
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
        else:
            n1 = try_parse_num(parts[1])
            if n1 is not None:
                mat = parts[0]
                qty = n1
                station = parts[2]
                input_zone = parts[3]
            else:
                mat = parts[0]
                station = parts[1]
                input_zone = parts[2]

        # Tự động trích xuất zone từ station nếu có
        if not input_zone and station:
            m_z = re.search(r"(?:ZONE|Z)\s*([0-9]+(?:\.[0-9]+)?)", station, re.IGNORECASE)
            if m_z:
                input_zone = m_z.group(1)
            elif re.fullmatch(r"[0-9]+(?:\.[0-9]+)?", station.strip()):
                input_zone = station.strip()

        if mat and qty > 0:
            items.append({
                "material": mat.strip().upper(),
                "qty": qty,
                "station": station.strip(),
                "input_zone": input_zone.strip()
            })

    return items


class SapZonePipeline:
    """Điều phối toàn diện luồng: Đọc Tồn -> Phân Bổ FIFO -> MIGO 311 -> Tra cứu WHO -> In Phiếu SAP."""

    def __init__(self, config: Dict[str, Any], zone_mgr: ZoneManager):
        self.config = config
        self.zone_mgr = zone_mgr
        self.stop_requested = False
        self.run_id = datetime.now().strftime("RUN_%Y%m%d_%H%M%S")

    def request_stop(self):
        self.stop_requested = True

    def run(
        self,
        requests: List[Dict[str, Any]],
        progress_cb=None,
        status_cb=None,
        is_mock: bool = False
    ) -> Dict[str, Any]:
        """Thực thi toàn bộ luồng nghiệp vụ với kiểm soát COM và Journal SQLite."""
        self.stop_requested = False
        SapConnector.init_thread_com()

        summary: Dict[str, Any] = {
            "success": False,
            "run_id": self.run_id,
            "records": [],
            "groups": [],
            "shortages": [],
            "unmapped": [],
            "pdf_files": [],
            "message": ""
        }

        def update_progress(pct: int, msg: str):
            if progress_cb:
                progress_cb(pct)
            if status_cb:
                status_cb(msg)
            print(f"[{pct}%] {msg}")

        try:
            if not requests:
                summary["message"] = "Không có danh sách yêu cầu hợp lệ!"
                JournalService.save_run(self.run_id, "FAILED", 0, 0, summary["message"])
                return summary

            JournalService.save_run(self.run_id, "STARTED", len(requests), 0, "Bắt đầu quy trình")

            # -------------------------------------------------------------
            # BƯỚC 1: KẾT NỐI SAP GUI AN TOÀN THEO DANH TÍNH
            # -------------------------------------------------------------
            update_progress(5, "Đang kết nối SAP GUI theo danh tính cấu hình...")
            session = None
            if not is_mock:
                try:
                    session = SapConnector.get_session(
                        session_idx=int(self.config.get("sap_session_idx", -1)),
                        preferred_sys=self.config.get("sap_system", "PSV"),
                        preferred_client=self.config.get("sap_client", "908")
                    )
                except Exception as e:
                    summary["message"] = f"Lỗi kết nối SAP: {e}"
                    update_progress(0, summary["message"])
                    JournalService.save_run(self.run_id, "SAP_CONN_ERR", len(requests), 0, summary["message"])
                    return summary
            else:
                update_progress(10, "[CHẾ ĐỘ MÔ PHỎNG] Đã khởi tạo phiên làm việc giả lập.")

            if self.stop_requested:
                summary["message"] = "Đã dừng bởi người dùng."
                return summary

            # -------------------------------------------------------------
            # BƯỚC 2: ĐỌC TỒN KHẢ DỤNG TỪ /SCWM/MON (APFL, OWNER 3093)
            # -------------------------------------------------------------
            update_progress(20, "Đang đọc tồn kho khả dụng từ /SCWM/MON (APFL, Owner 3093)...")
            all_materials = list(dict.fromkeys([r["material"] for r in requests]))
            stock_records: List[Dict[str, Any]] = []

            if not is_mock and session:
                stock_service = MonStockService()
                try:
                    stock_records = stock_service.fetch_stock_from_mon(
                        session=session,
                        materials=all_materials,
                        storage_type=self.config.get("ewm_storage_type", "APFL"),
                        owner=self.config.get("ewm_owner", "3093")
                    )
                except Exception as e:
                    summary["message"] = f"Lỗi đọc tồn kho từ /SCWM/MON: {e}"
                    update_progress(0, summary["message"])
                    JournalService.save_run(self.run_id, "MON_STOCK_ERR", len(requests), 0, summary["message"])
                    return summary
            else:
                # Tạo tồn kho giả lập
                for idx, m in enumerate(all_materials):
                    stock_records.append({
                        "material": m,
                        "batch": f"B{datetime.now().strftime('%y%m%d')}{idx:02d}",
                        "bin": f"BIN-APFL-{idx+1:02d}",
                        "available_qty": 9999.0,
                        "allocated_qty": 0.0,
                        "sort_key": f"20261108_{idx:04d}00",
                        "unit": "PC"
                    })

            if self.stop_requested:
                summary["message"] = "Đã dừng bởi người dùng."
                return summary

            # -------------------------------------------------------------
            # BƯỚC 3: PHÂN BỔ FIFO & GOM NHÓM THEO 4 TẦNG ƯU TIÊN ZONE
            # -------------------------------------------------------------
            update_progress(40, "Đang phân bổ FIFO và gom nhóm theo 4 tầng ưu tiên Zone...")
            engine = AllocationEngine(
                zone_mgr=self.zone_mgr,
                max_materials=int(self.config.get("max_materials_per_group", 5)),
                max_lines=int(self.config.get("max_lines_per_group", 20))
            )
            groups, shortages, unmapped = engine.allocate(requests, stock_records)

            summary["groups"] = groups
            summary["shortages"] = shortages
            summary["unmapped"] = unmapped

            if not groups:
                summary["message"] = "Không thể phân bổ được nhóm nào (do thiếu tồn kho hoặc mã hàng không hợp lệ)!"
                update_progress(40, summary["message"])
                JournalService.save_run(self.run_id, "ALLOC_EMPTY", len(requests), 0, summary["message"])
                return summary

            if self.stop_requested:
                summary["message"] = "Đã dừng bởi người dùng."
                return summary

            # -------------------------------------------------------------
            # BƯỚC 4: TẠO CHỨNG TỪ MIGO 311 (1005 -> 3001) CHO TỪNG NHÓM
            # -------------------------------------------------------------
            total_groups = len(groups)
            posted_obs: List[Dict[str, Any]] = []

            for idx, grp in enumerate(groups):
                if self.stop_requested:
                    break

                pct = 45 + int(25 * (idx + 1) / total_groups)
                grp_id = grp.get("group_id", f"G{idx+1}")
                grp_zone = grp.get("zone", "Chung")
                update_progress(pct, f"Đang tạo MIGO 311 cho nhóm {idx+1}/{total_groups} (Zone: {grp_zone})...")

                if not is_mock and session:
                    ok, doc_or_err = MigoPostingService.post_group(
                        session=session,
                        group_data=grp,
                        plant=self.config.get("plant", "3093"),
                        source_sloc=self.config.get("sloc_from", "1005"),
                        dest_sloc=self.config.get("sloc_to", "3001")
                    )
                    if ok:
                        grp["ob_number"] = doc_or_err
                        posted_obs.append({"group": grp, "ob_number": doc_or_err})
                        JournalService.save_group(self.run_id, grp_id, grp_zone, grp.get("tl", ""), doc_or_err, "POSTED", "Thành công")
                    else:
                        print(f"[MIGO POST FAIL] {grp_id}: {doc_or_err}")
                        JournalService.save_group(self.run_id, grp_id, grp_zone, grp.get("tl", ""), "", "FAILED", doc_or_err)
                else:
                    mock_ob = f"50{datetime.now().strftime('%y%m%d')}{idx+1:02d}"
                    grp["ob_number"] = mock_ob
                    posted_obs.append({"group": grp, "ob_number": mock_ob})
                    JournalService.save_group(self.run_id, grp_id, grp_zone, grp.get("tl", ""), mock_ob, "MOCK_POSTED", "Mô phỏng thành công")
                    time.sleep(0.3)

            if not posted_obs:
                summary["message"] = "Không tạo được chứng từ MIGO nào thành công!"
                update_progress(70, summary["message"])
                JournalService.save_run(self.run_id, "MIGO_FAIL", len(requests), 0, summary["message"])
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
                    for l_idx, l in enumerate(lines):
                        m_code = l["material"]
                        w1 = f"602871{l_idx*2+1:04d}"
                        w2 = f"602871{l_idx*2+2:04d}"
                        mat_who_map[m_code] = [w1, w2]

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
            JournalService.save_results(self.run_id, final_3col_records)

            if self.stop_requested:
                summary["message"] = "Đã dừng bởi người dùng."
                return summary

            # -------------------------------------------------------------
            # BƯỚC 6: IN PHIẾU SAP F00112 & XÁC THỰC LƯU PDF
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
                        ok_print = WhoPrintService.print_whos(
                            session=session,
                            who_list=all_whos_to_print,
                            warehouse_num=self.config.get("ewm_warehouse", "VF02"),
                            copies=int(self.config.get("sap_print_copies", 2)),
                            pdf_output_path=pdf_path
                        )
                        if ok_print:
                            summary["pdf_files"].append(pdf_path)
                    except Exception as ex:
                        print(f"[PRINT EX] {ex}")
                else:
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
            JournalService.save_run(self.run_id, "SUCCESS", len(requests), len(final_3col_records), summary["message"])
            return summary

        finally:
            SapConnector.uninit_thread_com()
