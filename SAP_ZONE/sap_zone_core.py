# -*- coding: utf-8 -*-
"""
SAP ZONE Core Engine (v1.6 Pro - Revised 2)
Tự động hóa toàn diện quy trình Ra đơn NVL:
- Quản lý phiên SAP GUI an toàn theo danh tính System/Client chuẩn SAP_EMAIL_TEAMS_BRIDGE v1.1.12
- Đọc tồn kho khả dụng từ /SCWM/MON (Node N000000139, Storage Type APFL, Owner 3093) chuẩn xác 100% theo VBS
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

# OpenPyXL (Lazy loaded for instant startup)
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
        outer_path = os.path.join(get_base_dir(), self.MAPPING_FILE)
        if os.path.exists(outer_path):
            return outer_path
        bundle_path = os.path.join(get_bundle_dir(), self.MAPPING_FILE)
        if os.path.exists(bundle_path):
            return bundle_path
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
        m = re.search(r"(?:ZONE|Z)\s*([0-9]+(?:\.[0-9]+)?)", s, re.IGNORECASE)
        if m:
            z_val = m.group(1)
            cleaned = re.sub(r"\(?\s*(?:ZONE|Z)\s*[0-9]+(?:\.[0-9]+)?\s*\)?", "", s, flags=re.IGNORECASE).strip(" -/()")
            return z_val, cleaned
        if re.fullmatch(r"[0-9]+(?:\.[0-9]+)?", s):
            return s, ""
        return "", s

    def resolve_zone_tl(self, material: str, station: str = "", input_zone: str = "") -> Tuple[str, str, bool]:
        """
        Xác định Zone và Team Leader (TL) theo thứ tự ưu tiên tối ưu cho người vận hành:
        1. Ưu tiên 1 (Zone trực tiếp): Nếu dữ liệu dán vào có Zone (cột Zone riêng hoặc trong chuỗi Trạm
           có 'Zone X.Y', 'Z X', '1.1'...), dùng ngay Zone đó làm nhóm chính của đơn hàng.
        2. Ưu tiên 2 (Trạm / Người nhận thực tế): Nếu dữ liệu dán có tên Trạm / Người nhận (vd: 'Hùng/Long'),
           tự động gom nhóm theo Trạm đó! Không phụ thuộc hay bắt buộc mã hàng phải có sẵn trong zone_map.
        3. Ưu tiên 3 (Danh mục zone_map.json): Nếu dòng dữ liệu hoàn toàn không có Zone và không có Trạm,
           tra cứu theo mã hàng trong zone_map.json.
        4. Ưu tiên 4 (Mặc định Chung): Gom vào nhóm 'Chung' (TL: 'Kho APFL').
        
        *LƯU Ý: Mọi mã hàng đều được chấp nhận và phân bổ bình thường (is_mapped luôn True),
        không bao giờ chặn hoặc báo lỗi mã hàng thiếu danh mục.
        """
        mat_clean = material.strip().upper()
        sta_clean = station.strip()
        inp_z = input_zone.strip()

        # Bóc tách nếu trong chuỗi station có kèm Zone (vd: 'Hùng/Long Zone 1.1' hoặc 'Z2.1' hoặc '1.1')
        if not inp_z and sta_clean:
            extracted_z, cleaned_sta = self.extract_zone_from_station(sta_clean)
            if extracted_z:
                inp_z = extracted_z
                if cleaned_sta:
                    sta_clean = cleaned_sta

        # 1. ƯU TIÊN 1: Zone được chỉ định trực tiếp từ dữ liệu đưa vào
        if inp_z:
            tl = sta_clean
            for m in self.mappings:
                if str(m.get("zone", "")).strip().upper() == inp_z.upper():
                    tl = m.get("tl", tl)
                    break
            return inp_z, (tl if tl else f"TL Zone {inp_z}"), True

        # 2. ƯU TIÊN 2: Gom nhóm theo Trạm / Người nhận thực tế đưa vào (vd: 'Hùng/Long')
        if sta_clean:
            return sta_clean, sta_clean, True

        # 3. ƯU TIÊN 3: Tra cứu danh mục zone_map.json (nếu không có trạm/zone từ input)
        for m in self.mappings:
            if m.get("material", "").strip().upper() == mat_clean:
                return str(m.get("zone", "")).strip(), m.get("tl", ""), True

        # 4. ƯU TIÊN 4: Mặc định nhóm Chung
        return "Chung", "Kho APFL", True


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
        """Chuyển tới T-Code an toàn qua /n<tcode> (hỗ trợ chính xác namespace như /n/SCWM/MON)."""
        if session is None:
            return False
        SapConnector.dismiss_popups(session)

        tc = tcode.strip()
        if tc.startswith("/n") or tc.startswith("/N"):
            target = tc
        elif tc.startswith("/"):
            target = f"/n/{tc.lstrip('/')}"
        else:
            target = f"/n{tc}"

        try:
            okcd = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
            if okcd is not None:
                okcd.text = target
                session.findById("wnd[0]").sendVKey(0)
                SapConnector.wait_idle(session, 2.5)
                return True
        except Exception:
            pass
        try:
            session.StartTransaction(tcode)
            SapConnector.wait_idle(session, 2.5)
            return True
        except Exception:
            return False

    @staticmethod
    def open_mon_screen(session, warehouse: str = "VF02") -> bool:
        """
        Mở màn hình /SCWM/MON (Warehouse Monitor) an toàn:
        - Tự động maximize cửa sổ wnd[0] ngay từ đầu như script VBS
        - Nếu đã ở trong MON (đã có cây N000000139), tái sử dụng ngay lập tức
        - Mở qua Favorite Node F00077 (chuẩn VBS của người dùng)
        - Đợi đủ thời gian cho cây điều hướng tải xong
        - Fallback gọi lệnh /n/SCWM/MON nếu cần
        """
        if session is None:
            return False

        # Maximize cửa sổ chính như script VBS
        try:
            session.findById("wnd[0]").maximize()
        except Exception:
            pass

        # 1. Kiểm tra nếu đã ở trong màn hình MON
        mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
        if mon_tree is not None:
            return True

        # Dọn dẹp popup cũ trước khi chuyển màn hình
        SapConnector.dismiss_popups(session)

        # 2. Về màn hình gốc Easy Access nếu đang ở màn hình khác
        fav_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
        if fav_tree is None:
            okcd = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
            if okcd:
                okcd.text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                SapConnector.wait_idle(session, 1.5)
            fav_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")

        # 3. Mở Favorite Node F00077 (chuẩn VBS thực tế của người dùng)
        if fav_tree is not None:
            try:
                try:
                    fav_tree.selectedNode = "F00077"
                except Exception:
                    pass
                fav_tree.doubleClickNode("F00077")
                # Đợi cây điều hướng MON xuất hiện (tối đa 10 giây)
                for _ in range(20):
                    SapConnector.wait_idle(session, 0.5)
                    mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
                    if mon_tree is not None:
                        return True
            except Exception as ex:
                print(f"[FAV F00077 EX] {ex}")

        # 4. Fallback: Mở bằng lệnh /n/SCWM/MON nếu không có F00077
        okcd = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
        if okcd:
            okcd.text = "/n/SCWM/MON"
            session.findById("wnd[0]").sendVKey(0)
            SapConnector.wait_idle(session, 2.0)

            # Xử lý popup chọn kho wnd[1] nếu có
            wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]")
            if wnd1 is not None:
                try:
                    txt_whse = (
                        SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtP_LGNUM") or
                        SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtP_LGNUM")
                    )
                    if txt_whse is not None:
                        txt_whse.text = warehouse
                    txt_mon = (
                        SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtP_MON") or
                        SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtP_MON")
                    )
                    if txt_mon is not None:
                        txt_mon.text = "SAP"
                    wnd1.sendVKey(0)  # Enter hoặc F8
                    SapConnector.wait_idle(session, 2.0)
                except Exception:
                    pass

        # 5. Đợi cây điều hướng MON xuất hiện (tối đa 10 giây)
        for _ in range(20):
            SapConnector.wait_idle(session, 0.5)
            mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
            if mon_tree is not None:
                return True

        return False

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
# 4. ĐỌC TỒN MON (/SCWM/MON) - NODE N000000139 (AVAILABLE STOCK CHUẨN VBS)
# ==============================================================================

class MonStockService:
    """Tải và trích xuất tồn kho khả dụng từ /SCWM/MON (Node N000000139, Storage Type APFL, Owner 3093)."""

    @staticmethod
    def copy_to_clipboard(text: str) -> bool:
        """Ghi chuỗi vào Windows Clipboard an toàn với cơ chế thử lại."""
        if not text:
            return False
        for _ in range(5):
            try:
                if win32clipboard is not None:
                    win32clipboard.OpenClipboard()
                    try:
                        win32clipboard.EmptyClipboard()
                        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
                        return True
                    finally:
                        win32clipboard.CloseClipboard()
            except Exception:
                time.sleep(0.08)
        try:
            import tkinter as tk
            r = tk.Tk()
            r.withdraw()
            r.clipboard_clear()
            r.clipboard_append(text)
            r.update()
            r.destroy()
            return True
        except Exception:
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
        """Truy vấn tồn kho từ /SCWM/MON Node N000000139 chuẩn xác theo VBS của người dùng."""
        if not materials or session is None:
            return []

        # 1. Mở màn hình /SCWM/MON an toàn
        ok = SapConnector.open_mon_screen(session)
        if not ok:
            raise RuntimeError("Không thể điều hướng vào màn hình /SCWM/MON hoặc tìm thấy cây điều hướng!")

        # 2. Cây điều hướng MON
        mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
        if mon_tree is None:
            raise RuntimeError("Không tìm thấy cây điều hướng của /SCWM/MON (wnd[0]/usr/shell/shellcont[0]/shell)!")

        # Kiểm tra nếu hộp thoại lọc wnd[1] chưa mở thì chọn và kích hoạt node N000000139
        wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]")
        if wnd1 is None:
            try:
                mon_tree.selectedNode = "N000000139"
                mon_tree.doubleClickNode("N000000139")
                SapConnector.wait_idle(session, 2.5)
            except Exception as e:
                raise RuntimeError(f"Lỗi chọn node N000000139 trên cây MON: {e}")

            # Đợi hộp thoại lọc wnd[1] xuất hiện
            for _ in range(16):
                wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]")
                if wnd1 is not None:
                    break
                time.sleep(0.3)

        if wnd1 is None:
            raise RuntimeError("Hộp thoại bộ lọc tồn kho /SCWM/MON (wnd[1]) không xuất hiện!")

        # Điền Storage Type (APFL) và Owner (3093) trên wnd[1]
        try:
            txt_lgtyp = (
                SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtS_LGTYP-LOW") or
                SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtS_LGTYP-LOW")
            )
            if txt_lgtyp:
                txt_lgtyp.text = storage_type
        except Exception:
            pass

        try:
            txt_owner = (
                SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtS_OWNER-LOW") or
                SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtS_OWNER-LOW")
            )
            if txt_owner:
                txt_owner.text = owner
        except Exception:
            pass

        # 4. Nạp danh sách Material qua Multiple Selection (wnd[2])
        unique_mats = list(dict.fromkeys([m.strip() for m in materials if m.strip()]))
        self.copy_to_clipboard("\r\n".join(unique_mats))

        btn_mat_push = SapConnector.find_by_id_safe(session, "wnd[1]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH")
        if btn_mat_push:
            btn_mat_push.press()
            SapConnector.wait_idle(session, 1.0)

            # Đợi wnd[2] xuất hiện và bấm Upload từ Clipboard (btn[24])
            btn_paste = None
            for _ in range(16):
                btn_paste = SapConnector.find_by_id_safe(session, "wnd[2]/tbar[0]/btn[24]")
                if btn_paste is not None:
                    break
                time.sleep(0.2)

            if btn_paste is not None:
                try:
                    btn_paste.press()
                    SapConnector.wait_idle(session, 1.0)
                except Exception:
                    pass

            # Đợi và bấm F8 trên wnd[2] để chốt danh sách nạp
            btn_f8_wnd2 = None
            for _ in range(16):
                btn_f8_wnd2 = SapConnector.find_by_id_safe(session, "wnd[2]/tbar[0]/btn[8]")
                if btn_f8_wnd2 is not None:
                    break
                time.sleep(0.2)

            if btn_f8_wnd2 is not None:
                try:
                    btn_f8_wnd2.press()
                    SapConnector.wait_idle(session, 1.0)
                except Exception:
                    pass

        # 5. Execute F8 trên wnd[1]
        btn_f8_wnd1 = None
        for _ in range(16):
            btn_f8_wnd1 = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
            if btn_f8_wnd1 is not None:
                break
            time.sleep(0.2)

        if btn_f8_wnd1:
            btn_f8_wnd1.press()
            SapConnector.wait_idle(session, 3.5)

        # 6. Đợi và đọc kết quả từ bảng ALV Grid (chuẩn xác theo VBS: wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell)
        grid = None
        for _ in range(20):
            grid = (
                SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell") or
                SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell") or
                SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlGRID1/shellcont/shell")
            )
            if grid is not None:
                try:
                    if getattr(grid, "RowCount", 0) > 0:
                        break
                except Exception:
                    pass
            time.sleep(0.3)

        if grid is None:
            return []

        row_count = getattr(grid, "RowCount", 0)
        stock_records: List[Dict[str, Any]] = []

        if row_count <= 0:
            return stock_records

        # Tên cột chuẩn từ script VBS của người dùng
        mat_col = "MATNR"
        batch_col = "CHARG"
        bin_col = "LGPLA"
        avail_col = "ABSQUAN"
        gr_date_col = "UI_WDATU"
        gr_time_col = "UI_WDATT"
        unit_col = "MEINS"

        # Nếu grid có ColumnOrder thì phân giải thêm
        try:
            col_order = grid.ColumnOrder
            for i in range(col_order.Count):
                c_name = str(col_order.ElementAt(i)).strip()
                try:
                    c_title = (grid.GetColumnTitles(c_name).ElementAt(0) or "").strip().lower()
                except Exception:
                    c_title = ""

                c_name_up = c_name.upper()
                if c_name_up == "ABSQUAN" or "available" in c_title or "khả dụng" in c_title:
                    avail_col = c_name
                elif c_name_up in ("MATNR", "PRODUCT", "MATERIAL") or "material" in c_title or "mã hàng" in c_title:
                    mat_col = c_name
                elif c_name_up in ("LGPLA", "BIN") or "storage bin" in c_title or "vị trí" in c_title:
                    bin_col = c_name
                elif c_name_up in ("CHARG", "BATCH", "BATCHID", "SO_CHARG", "LICHA") or "batch" in c_title or "lô" in c_title:
                    batch_col = c_name
                elif c_name_up in ("UI_WDATU", "WDATU", "GR_DATE") or "gr date" in c_title or "ngày nhập" in c_title:
                    gr_date_col = c_name
                elif c_name_up in ("UI_WDATT", "WZEIT", "GR_TIME") or "gr time" in c_title or "giờ nhập" in c_title:
                    gr_time_col = c_name
                elif c_name_up in ("MEINS", "UNIT"):
                    unit_col = c_name
        except Exception as ex_col:
            print(f"[MON COL RESOLVE ERR]: {ex_col}")

        for r in range(row_count):
            try:
                m_val = str(grid.GetCellValue(r, mat_col)).strip().upper()
                b_val = str(grid.GetCellValue(r, batch_col)).strip()
                if not b_val:
                    for fb_col in ["CHARG", "BATCH", "SO_CHARG", "LICHA"]:
                        try:
                            fb_val = str(grid.GetCellValue(r, fb_col)).strip()
                            if fb_val:
                                b_val = fb_val
                                break
                        except Exception:
                            pass
                bin_val = str(grid.GetCellValue(r, bin_col)).strip()
                avail_raw = str(grid.GetCellValue(r, avail_col)).strip()
                avail_qty = self.parse_sap_number(avail_raw)

                if not m_val or avail_qty <= 0:
                    continue

                d_val = re.sub(r"[^0-9]", "", str(grid.GetCellValue(r, gr_date_col)).strip())
                t_val = re.sub(r"[^0-9]", "", str(grid.GetCellValue(r, gr_time_col)).strip())
                sort_key = f"{d_val.zfill(8)}_{t_val.zfill(6)}"

                try:
                    u_val = str(grid.GetCellValue(r, unit_col)).strip() or "PC"
                except Exception:
                    u_val = "PC"

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
    Tạo chứng từ chuyển kho MIGO 311 (1005 -> 3001) chuẩn xác 100% theo VBS ghi nhận thực tế của người dùng:
    - Nhận diện các carrier subscreen: 0008 (thực tế máy user), 0004, 0003, 0002, 0001.
    - Điền trực tiếp Movement Type 311 vào ctxtGODEFAULT_TV-BWART.
    - Ghi ô dữ liệu bằng ID trực tiếp chuẩn VBS:
      col 3: ctxtGOITEM-MAKTX (Material), col 4: txtGOITEM-ERFMG (Qty), col 5: txtGOITEM-SGTXT (Text/Trạm),
      col 6: ctxtGOITEM-NAME1 (Plant 3093), col 7: ctxtGOITEM-LGOBE (Kho xuất 1005), col 8: ctxtGOITEM-UMLGOBE (Kho nhận 3001).
    - Focus ctxtGOITEM-MAKTX[3,0] và gửi Enter nhịp 1 (sendVKey 0).
    - Nhịp 2: Ghi Batch vào ctxtGOITEM-CHARG[10,r] -> Enter.
    - Lưu / Post bằng sendVKey 11 (Save/Post chuẩn SAP).
    - Bắt mọi số chứng từ 10 chữ số (\b(\d{10})\b) bao gồm dải số 1013147419 của hệ thống người dùng.
    - Thoát MIGO trở lại Easy Access bằng btn[15] (F3).
    """

    @staticmethod
    def _find_child_by_name(obj, target_name: str):
        """Tìm control đệ quy theo thuộc tính Name an toàn trong Python win32com."""
        if obj is None:
            return None
        try:
            o_name = str(getattr(obj, "Name", "") or "").strip().upper()
            if o_name == target_name.strip().upper():
                return obj
        except Exception:
            pass
        try:
            children = getattr(obj, "Children", None)
            if children:
                for child in children:
                    res = MigoPostingService._find_child_by_name(child, target_name)
                    if res is not None:
                        return res
        except Exception:
            pass
        return None

    @staticmethod
    def _find_child_by_id_contains(obj, needle: str):
        """Tìm control đệ quy theo chuỗi con trong Id an toàn trong Python win32com."""
        if obj is None:
            return None
        try:
            o_id = str(getattr(obj, "Id", "") or "").strip().upper()
            if needle.strip().upper() in o_id:
                return obj
        except Exception:
            pass
        try:
            children = getattr(obj, "Children", None)
            if children:
                for child in children:
                    res = MigoPostingService._find_child_by_id_contains(child, needle)
                    if res is not None:
                        return res
        except Exception:
            pass
        return None

    @staticmethod
    def _get_combo_selected_text(combo) -> str:
        """Đọc text đang được chọn trong combobox an toàn."""
        if combo is None:
            return ""
        try:
            txt = str(getattr(combo, "Text", "") or "").strip()
            if txt:
                return txt
            key = str(getattr(combo, "Key", "") or "").strip()
            entries = getattr(combo, "Entries", None)
            if entries:
                for ent in entries:
                    if str(getattr(ent, "Key", "")).strip() == key:
                        v = str(getattr(ent, "Value", "") or "").strip()
                        return v if v else str(getattr(ent, "Text", "") or "").strip()
        except Exception:
            pass
        return ""

    @staticmethod
    def _set_combo_by_caption(combo, caption1: str, caption2: str = "") -> bool:
        """Chọn mục combobox theo chuỗi hiển thị an toàn."""
        if combo is None:
            return False
        n1 = caption1.strip().upper()
        n2 = caption2.strip().upper() if caption2 else ""
        try:
            entries = getattr(combo, "Entries", None)
            if entries:
                for ent in entries:
                    e_val = str(getattr(ent, "Value", "") or "").strip().upper()
                    e_txt = str(getattr(ent, "Text", "") or "").strip().upper()
                    e_key = str(getattr(ent, "Key", "") or "").strip()
                    norm = e_val if e_val else e_txt
                    if (n1 and n1 in norm) or (n2 and n2 in norm):
                        combo.Key = e_key
                        return True
        except Exception:
            pass
        try:
            if "TRANSFER" in n1:
                combo.Key = "A08"
                return True
            elif "OTHER" in n1 or "KHAC" in n1:
                combo.Key = "R10"
                return True
        except Exception:
            pass
        return False

    @staticmethod
    def _find_migo_table(session) -> Tuple[Any, str]:
        """Tìm bảng dữ liệu MIGO tblSAPLMIGOTV_GOITEM và trả về (tbl_obj, tbl_id)."""
        if session is None:
            return None, ""

        # Ưu tiên subscreen 0008 theo VBS ghi nhận thực tế trên máy người dùng
        candidate_paths = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
        ]
        for p in candidate_paths:
            tbl = SapConnector.find_by_id_safe(session, p)
            if tbl is not None:
                return tbl, p

        # Thu gọn Header / Detail để mở rộng bảng lưới
        try:
            session.SendCommand("MIGO_OK_HEADER_CLOSE")
            time.sleep(0.15)
            session.SendCommand("MIGO_OK_DETAIL_CLOSE")
            time.sleep(0.2)
        except Exception:
            pass

        # Thử lại danh sách candidate paths sau khi thu gọn panel
        for p in candidate_paths:
            tbl = SapConnector.find_by_id_safe(session, p)
            if tbl is not None:
                return tbl, p

        # Quét đệ quy toàn bộ wnd[0]
        root = session.findById("wnd[0]", False) or session.findById("wnd[0]")
        tbl = MigoPostingService._find_child_by_id_contains(root, "tblSAPLMIGOTV_GOITEM")
        if tbl is not None:
            tbl_id = str(getattr(tbl, "Id", "") or "").strip()
            return tbl, tbl_id

        tbl = MigoPostingService._find_child_by_id_contains(root, "GOITEM")
        if tbl is not None:
            tbl_id = str(getattr(tbl, "Id", "") or "").strip()
            return tbl, tbl_id

        return None, ""

    @staticmethod
    def _write_table_cell(
        session,
        tbl_obj,
        tbl_path: str,
        col_names: List[str],
        col_indices: List[int],
        row: int,
        val: str
    ) -> bool:
        """Ghi giá trị vào cell MIGO qua nhiều lớp bảo vệ (Direct ID -> Iterate Children -> GetCell -> ModifyCell)."""
        val_str = str(val).strip()

        # 1. Thử trực tiếp qua session.findById với các định dạng ID đã ghi nhận từ VBS
        if tbl_path:
            for name in col_names:
                for col in col_indices:
                    cell_id = f"{tbl_path}/{name}[{col},{row}]"
                    ctl = SapConnector.find_by_id_safe(session, cell_id)
                    if ctl is not None:
                        try:
                            ctl.text = val_str
                            return True
                        except Exception:
                            try:
                                ctl.Text = val_str
                                return True
                            except Exception:
                                pass

        # 2. Duyệt qua Children của tbl_obj bằng Python iterator
        if tbl_obj is not None:
            try:
                children = getattr(tbl_obj, "Children", None)
                if children:
                    for child in children:
                        c_id = str(getattr(child, "Id", "") or "").strip()
                        for name in col_names:
                            if name in c_id and (f",{row}]" in c_id or f"[{row}]" in c_id):
                                try:
                                    child.text = val_str
                                    return True
                                except Exception:
                                    try:
                                        child.Text = val_str
                                        return True
                                    except Exception:
                                        pass
            except Exception:
                pass

        # 3. Thử qua GetCell
        if tbl_obj is not None:
            for col in col_indices:
                try:
                    cell = tbl_obj.GetCell(row, col)
                    if cell is not None:
                        try:
                            cell.Text = val_str
                            return True
                        except Exception:
                            try:
                                cell.text = val_str
                                return True
                            except Exception:
                                pass
                except Exception:
                    pass

        # 4. Thử qua ModifyCell
        if tbl_obj is not None:
            for col in col_indices:
                try:
                    tbl_obj.ModifyCell(row, col, val_str)
                    return True
                except Exception:
                    pass

        return False

    @staticmethod
    def ensure_migo_311_ready(session, timeout_sec: float = 15.0) -> Tuple[bool, str]:
        """
        Khởi tạo và đảm bảo phiên MIGO sẵn sàng 100%:
        1. Maximize cửa sổ wnd[0].
        2. Nếu chưa ở MIGO, mở /nMIGO và đợi ổn định.
        3. Điền Movement Type 311 vào ctxtGODEFAULT_TV-BWART (ưu tiên subscreen 0008).
        4. Thu gọn Header/Detail để mở rộng bảng lưới GOITEM.
        5. Xác thực bảng tblSAPLMIGOTV_GOITEM.
        """
        if session is None:
            return False, "Phiên SAP GUI chưa được khởi tạo!"

        try:
            session.findById("wnd[0]").maximize()
        except Exception:
            pass

        SapConnector.dismiss_popups(session)

        # 1. Kiểm tra transaction hiện tại
        try:
            info = getattr(session, "Info", None)
            cur_tcode = str(getattr(info, "Transaction", "") or "").strip().upper()
        except Exception:
            cur_tcode = ""

        if cur_tcode != "MIGO":
            okcd = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
            if okcd is not None:
                okcd.text = "/nMIGO"
                session.findById("wnd[0]").sendVKey(0)
            else:
                session.StartTransaction("MIGO")
            SapConnector.wait_idle(session, 3.0)
            SapConnector.dismiss_popups(session)

        # 2. Tìm và điền Movement Type 311
        bwart_candidates = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART",
        ]
        bwart_field = None
        for bp in bwart_candidates:
            bwart_field = SapConnector.find_by_id_safe(session, bp)
            if bwart_field is not None:
                break

        if bwart_field is None:
            root = session.findById("wnd[0]", False) or session.findById("wnd[0]")
            bwart_field = MigoPostingService._find_child_by_name(root, "GODEFAULT_TV-BWART")

        # Nếu không có BWART field, mới cần điều chỉnh combobox Action/RefDoc
        if bwart_field is None:
            root = session.findById("wnd[0]", False) or session.findById("wnd[0]")
            act_combo = MigoPostingService._find_child_by_name(root, "GODYNPRO-ACTION")
            if act_combo is not None:
                cur_act = MigoPostingService._get_combo_selected_text(act_combo).upper()
                if not ("TRANSFER" in cur_act or getattr(act_combo, "Key", "") == "A08"):
                    MigoPostingService._set_combo_by_caption(act_combo, "TRANSFER POSTING", "TRANSFER POST")
                    SapConnector.wait_idle(session, 2.5)

            ref_combo = MigoPostingService._find_child_by_name(root, "GODYNPRO-REFDOC")
            if ref_combo is not None:
                cur_ref = MigoPostingService._get_combo_selected_text(ref_combo).upper()
                if not ("OTHER" in cur_ref or "KHAC" in cur_ref or getattr(ref_combo, "Key", "") == "R10"):
                    MigoPostingService._set_combo_by_caption(ref_combo, "OTHER", "KHAC")
                    SapConnector.wait_idle(session, 2.5)

            for bp in bwart_candidates:
                bwart_field = SapConnector.find_by_id_safe(session, bp)
                if bwart_field is not None:
                    break
            if bwart_field is None:
                root = session.findById("wnd[0]", False) or session.findById("wnd[0]")
                bwart_field = MigoPostingService._find_child_by_name(root, "GODEFAULT_TV-BWART")

        if bwart_field is not None:
            try:
                bwart_field.text = "311"
            except Exception:
                try:
                    bwart_field.Text = "311"
                except Exception:
                    pass

        # 3. Thu gọn Header/Detail để bảng lưới hiển thị đầy đủ
        try:
            session.SendCommand("MIGO_OK_HEADER_CLOSE")
            time.sleep(0.1)
            session.SendCommand("MIGO_OK_DETAIL_CLOSE")
            time.sleep(0.15)
        except Exception:
            pass

        # 4. Kiểm tra bảng MIGO (Retry an toàn tối đa 3 giây)
        tbl = None
        tbl_path = ""
        for _ in range(10):
            tbl, tbl_path = MigoPostingService._find_migo_table(session)
            if tbl is not None or tbl_path:
                break
            time.sleep(0.3)

        if tbl is None and not tbl_path:
            return False, "Đã vào MIGO 311 nhưng không tìm thấy bảng lưới nhập liệu (tblSAPLMIGOTV_GOITEM)!"

        return True, ""

    @staticmethod
    def _press_post(session) -> bool:
        """Nhấn nút Post theo đúng thứ tự ưu tiên."""
        try:
            btn_post1 = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[1]/btn[23]")
            if btn_post1 is not None:
                btn_post1.press()
                return True
        except Exception:
            pass
        try:
            btn_save0 = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/btn[11]")
            if btn_save0 is not None:
                btn_save0.press()
                return True
        except Exception:
            pass
        try:
            session.SendCommand("MIGO_POST")
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
        """
        Tạo chứng từ MIGO 311 cho 1 nhóm đơn theo đúng chuẩn VBS & Layout cố định:
        - Nhịp 1: Điền Material (col 3), Qty (col 4), Trạm/Text (col 5), Plant (col 6), Kho xuất 1005 (col 7), Kho nhận 3001 (col 8).
        - Gửi Enter Nhịp 1 (sendVKey 0) -> SAP chèn cột Stock Determin. (icon ngôi nhà 🏠) vào Cột 4, đẩy cột Batch sang Cột 10.
        - Chờ SAP Server Roundtrip, re-bind lại bảng mới.
        - Nhịp 2: Điền số Batch thật từ /SCWM/MON vào Cột 10 (ctxtGOITEM-CHARG[10,r]).
        - Gửi Enter Nhịp 2 (sendVKey 0) để SAP kiểm tra tính hợp lệ và tick OK dòng.
        - Nhịp 3: Post chứng từ bằng sendVKey 11 (Save/Post) -> Bắt số OB 10 chữ số (ví dụ 1013147419).
        """
        try:
            # 1. Đảm bảo màn hình MIGO 311 sẵn sàng
            ok_ready, err_ready = MigoPostingService.ensure_migo_311_ready(session)
            if not ok_ready:
                return False, f"Lỗi cấu hình màn hình MIGO 311: {err_ready}"

            # 2. Tìm bảng chi tiết
            tbl, tbl_path = MigoPostingService._find_migo_table(session)
            if not tbl_path and tbl is not None:
                tbl_path = str(getattr(tbl, "Id", "") or "").strip()
            if not tbl_path:
                tbl_path = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"

            lines = group_data.get("lines", [])
            station_text = group_data.get("tl", "") or group_data.get("zone", "NVL")

            # 3. GIAI ĐOẠN 1 (PRE-BATCH): Điền dữ liệu chung từng dòng trước khi Enter
            for idx, item in enumerate(lines):
                mat_code = str(item["material"]).strip()
                qty_raw = item["qty"]
                qty_val = int(qty_raw) if float(qty_raw).is_integer() else qty_raw
                line_station = item.get("station", station_text) or station_text
                line_text = str(line_station)[:50]

                # Ghi Material (Cột 3: ctxtGOITEM-MAKTX)
                mat_ok = MigoPostingService._write_table_cell(
                    session, tbl, tbl_path,
                    ["ctxtGOITEM-MAKTX", "ctxtGOITEM-MATNR"],
                    [3, 1, 2],
                    idx,
                    mat_code
                )
                if not mat_ok:
                    return False, f"Không thể ghi Material vào dòng {idx+1} của bảng MIGO!"

                # Ghi Qty (Cột 4: txtGOITEM-ERFMG)
                MigoPostingService._write_table_cell(
                    session, tbl, tbl_path,
                    ["txtGOITEM-ERFMG"],
                    [4, 2, 3],
                    idx,
                    str(qty_val)
                )

                # Ghi Text/Trạm (Cột 5: txtGOITEM-SGTXT)
                MigoPostingService._write_table_cell(
                    session, tbl, tbl_path,
                    ["txtGOITEM-SGTXT"],
                    [5, 4, 6],
                    idx,
                    line_text
                )

                # Ghi Plant (Cột 6: ctxtGOITEM-NAME1)
                MigoPostingService._write_table_cell(
                    session, tbl, tbl_path,
                    ["ctxtGOITEM-NAME1", "ctxtGOITEM-WERKS"],
                    [6, 5, 7],
                    idx,
                    plant
                )

                # Ghi Kho xuất (Cột 7: ctxtGOITEM-LGOBE)
                MigoPostingService._write_table_cell(
                    session, tbl, tbl_path,
                    ["ctxtGOITEM-LGOBE", "ctxtGOITEM-LGORT"],
                    [7, 6, 8],
                    idx,
                    source_sloc
                )

                # Ghi Kho nhận (Cột 8: ctxtGOITEM-UMLGOBE)
                MigoPostingService._write_table_cell(
                    session, tbl, tbl_path,
                    ["ctxtGOITEM-UMLGOBE", "ctxtGOITEM-UMLGO"],
                    [8, 7, 9],
                    idx,
                    dest_sloc
                )

            # Focus vào ctxtGOITEM-MAKTX[3,0] chuẩn VBS dòng 26-28
            try:
                f_mat = SapConnector.find_by_id_safe(session, f"{tbl_path}/ctxtGOITEM-MAKTX[3,0]") or \
                        SapConnector.find_by_id_safe(session, f"{tbl_path}/ctxtGOITEM-MATNR[3,0]")
                if f_mat:
                    f_mat.setFocus()
                    f_mat.caretPosition = 0
            except Exception:
                pass

            # GỬI ENTER NHỊP 1
            print("[MIGO] Gửi Enter Nhịp 1 (để SAP kích hoạt Stock Determin. và mở khóa ô Batch)...")
            session.findById("wnd[0]").sendVKey(0)

            # 4. GIAI ĐOẠN 2: CHỜ SAP HOÀN TẤT ROUNDTRIP, RE-BIND LẠI BẢNG & ĐIỀN BATCH CỘT 10
            # Vòng lặp chờ ô Batch cột 10 xuất hiện (tối đa 10 giây)
            batch_col_idx = 10
            tbl_path_new = tbl_path
            t_wait_batch = time.time()
            batch_cell_found = False

            while time.time() - t_wait_batch < 10.0:
                SapConnector.wait_idle(session, 0.4)
                SapConnector.dismiss_popups(session)

                # Lấy lại đường dẫn bảng mới sau khi SAP re-render
                _, t_path_check = MigoPostingService._find_migo_table(session)
                if t_path_check:
                    tbl_path_new = t_path_check

                # Kiểm tra ô Batch tại Cột 10 dòng 0
                test_cell10 = SapConnector.find_by_id_safe(session, f"{tbl_path_new}/ctxtGOITEM-CHARG[10,0]")
                if test_cell10 is not None:
                    batch_col_idx = 10
                    batch_cell_found = True
                    break

                # Kiểm tra fallback nếu hệ thống có layout khác (Cột 9 hoặc Cột 11)
                test_cell9 = SapConnector.find_by_id_safe(session, f"{tbl_path_new}/ctxtGOITEM-CHARG[9,0]")
                if test_cell9 is not None:
                    batch_col_idx = 9
                    batch_cell_found = True
                    break

                test_cell11 = SapConnector.find_by_id_safe(session, f"{tbl_path_new}/ctxtGOITEM-CHARG[11,0]")
                if test_cell11 is not None:
                    batch_col_idx = 11
                    batch_cell_found = True
                    break

                time.sleep(0.3)

            print(f"[MIGO] Đã định vị cột Batch: Cột {batch_col_idx} (Tìm thấy: {batch_cell_found})")

            # Điền Batch thật đã check từ MON vào từng dòng (Kiểm tra nghiêm ngặt từng dòng)
            for idx, item in enumerate(lines):
                b_str = str(item.get("batch", "")).strip()
                mat_code = str(item.get("material", "")).strip()
                if not b_str:
                    return False, f"Mã hàng {mat_code} (dòng {idx+1}) chưa có số Batch từ /SCWM/MON! Không thể tạo MIGO khi thiếu Batch."

                line_batch_ok = False
                # Thử lần lượt các cột candidate (Cột tìm được, Cột 10, Cột 9, Cột 11)
                for c_cand in [batch_col_idx, 10, 9, 11]:
                    cell_id_target = f"{tbl_path_new}/ctxtGOITEM-CHARG[{c_cand},{idx}]"
                    target_cell = SapConnector.find_by_id_safe(session, cell_id_target)
                    if target_cell is not None:
                        try:
                            target_cell.text = b_str
                            try:
                                target_cell.setFocus()
                                target_cell.caretPosition = len(b_str)
                            except Exception:
                                pass
                            line_batch_ok = True
                            print(f"[MIGO BATCH] Dòng {idx+1}: Đã điền Batch '{b_str}' vào {cell_id_target}")
                            break
                        except Exception as ex_cell:
                            print(f"[MIGO BATCH ERR] Dòng {idx+1} {cell_id_target}: {ex_cell}")

                if not line_batch_ok:
                    ok_w = MigoPostingService._write_table_cell(
                        session, None, tbl_path_new,
                        ["ctxtGOITEM-CHARG"],
                        [batch_col_idx, 10, 9, 11],
                        idx,
                        b_str
                    )
                    if ok_w:
                        line_batch_ok = True
                        print(f"[MIGO BATCH FALLBACK OK] Dòng {idx+1}: Ghi Batch '{b_str}' thành công")

                if not line_batch_ok:
                    return False, f"Không thể điền số Batch '{b_str}' vào dòng {idx+1} (Cột {batch_col_idx}) của bảng MIGO sau khi Enter!"

            # 5. GIAI ĐOẠN 3: GỬI ENTER NHỊP 2 ĐỂ SAP VALIDATE BATCH VÀ TICK OK DÒNG
            print("[MIGO] Gửi Enter Nhịp 2 (để SAP kiểm tra tồn kho theo Lô và tick OK dòng)...")
            session.findById("wnd[0]").sendVKey(0)  # Enter Nhịp 2
            SapConnector.wait_idle(session, 2.5)
            SapConnector.dismiss_popups(session)

            # 6. GIAI ĐOẠN 4: LƯU / POST CHỨNG TỪ (sendVKey 11 chuẩn VBS)
            print("[MIGO] Nhấn Save / Post chứng từ (sendVKey 11)...")
            try:
                session.findById("wnd[0]").sendVKey(11)
            except Exception:
                MigoPostingService._press_post(session)

            # 7. POLLING ĐỌC SỐ CHỨNG TỪ & BẮT CHI TIẾT NẾU CÓ CẢNH BÁO
            last_status = ""
            doc_no = ""
            for loop in range(15):
                SapConnector.wait_idle(session, 1.0)
                pop_txt = SapConnector.dismiss_popups(session)
                sbar_txt, sbar_type = SapConnector.get_status(session)
                combined = f"{sbar_txt} | {pop_txt}".strip(" |")
                if combined:
                    last_status = combined

                # Trích xuất số chứng từ 10 chữ số (chuẩn 1013147419 của hệ thống hoặc dải 50/49/8/5)
                m_doc10 = re.search(r"\b(\d{10})\b", combined)
                if m_doc10:
                    doc_no = m_doc10.group(1)
                    break

                # Trích xuất từ Outbound Delivery / Document
                m_deliv = re.search(r"(?:OUTBOUND\s+DELIVERY|DELIVERY|DELIV\.?|DOCUMENT|CHỨNG TỪ)[^0-9]{0,40}([0-9]{8,12})", combined, re.IGNORECASE)
                if m_deliv:
                    doc_no = m_deliv.group(1)
                    break

                if sbar_type == "S":
                    m_any = re.search(r"\b(\d{10})\b", combined)
                    doc_no = m_any.group(1) if m_any else "POSTED_OK"
                    break

                # Nếu xuất hiện thông báo Display logs, double click status bar để trích xuất chi tiết
                if "DISPLAY LOG" in combined.upper():
                    try:
                        sbar = SapConnector.find_by_id_safe(session, "wnd[0]/sbar")
                        if sbar:
                            sbar.doubleClick()
                            SapConnector.wait_idle(session, 1.0)
                            log_pop = SapConnector.dismiss_popups(session)
                            if log_pop:
                                last_status = f"Chi tiết lỗi SAP: {log_pop}"
                    except Exception:
                        pass

                if sbar_type in ["E", "A"]:
                    return False, f"Lỗi SAP MIGO ({sbar_type}): {sbar_txt}"

                if sbar_type == "W" or pop_txt:
                    try:
                        session.findById("wnd[0]").sendVKey(0)
                        SapConnector.wait_idle(session, 1.0)
                    except Exception:
                        pass
                    continue

                time.sleep(0.4)

            # 8. Sau khi post thành công, thoát MIGO về Easy Access bằng F3 (btn[15])
            if doc_no:
                try:
                    btn15 = SapConnector.find_by_id_safe(session, "wnd[0]/tbar[0]/btn[15]")
                    if btn15:
                        btn15.press()
                        SapConnector.wait_idle(session, 1.0)
                except Exception:
                    pass
                return True, doc_no

            return False, f"MIGO chưa hoàn tất sau khi Post: {last_status or 'Không nhận được số chứng từ từ SAP'}"

        except Exception as ex:
            return False, f"Lỗi ngoại lệ MIGO: {ex}"


class MonRelationService:
    """Tra cứu MON để đối soát chính xác 3 cột: OB - Mã hàng - Số đơn (WHO) và in WHO chuẩn VBS."""

    @staticmethod
    def query_who_for_ob(session, ob_number: str, auto_print: bool = True) -> Dict[str, List[str]]:
        """
        Tra MON theo OB chuẩn xác VBS thực tế:
        1. Mở Favorite F00077 (hoặc /n/SCWM/MON).
        2. Chọn Node N000000010 (Outbound Delivery Order).
        3. Điền OB vào wnd[1]/usr/txtS_ERP_H-LOW và bấm F8 (btn[8]).
        4. Bấm nút N000000011 (Warehouse Tasks / WHO) trên Toolbar lưới trên.
        5. Đọc lưới dưới lấy WHO và MATNR (hoặc PRODUCTNO), DOCNO.
        6. Nếu auto_print=True: selectAll -> Context METHODS -> @M00007 -> Enter popup wnd[1]/tbar[0]/btn[0] -> Refresh.
        """
        clean_ob = re.sub(r"[^0-9]", "", str(ob_number)).strip()
        mat_who_map: Dict[str, List[str]] = {}

        if not clean_ob or session is None:
            return mat_who_map

        try:
            # 1. Mở màn hình /SCWM/MON an toàn qua Favorite F00077 (chuẩn VBS dòng 35)
            fav_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
            if fav_tree is not None:
                try:
                    fav_tree.doubleClickNode("F00077")
                    SapConnector.wait_idle(session, 2.0)
                except Exception:
                    SapConnector.open_mon_screen(session)
            else:
                SapConnector.open_mon_screen(session)

            # 2. Chọn nhánh Outbound Delivery Order (N000000010 - chuẩn VBS dòng 36-37)
            mon_tree = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[0]/shell")
            if mon_tree:
                try:
                    mon_tree.selectedNode = "N000000010"
                    mon_tree.doubleClickNode("N000000010")
                    SapConnector.wait_idle(session, 2.0)
                except Exception:
                    pass

            # 3. Lọc theo OB và F8 trên wnd[1] (chuẩn VBS dòng 38-41)
            txt_ob = (
                SapConnector.find_by_id_safe(session, "wnd[1]/usr/txtS_ERP_H-LOW") or
                SapConnector.find_by_id_safe(session, "wnd[1]/usr/ctxtS_ERP_H-LOW")
            )
            if txt_ob:
                txt_ob.text = clean_ob
                try:
                    txt_ob.setFocus()
                    txt_ob.caretPosition = len(clean_ob)
                except Exception:
                    pass

            btn_f8 = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[8]")
            if btn_f8:
                btn_f8.press()
                SapConnector.wait_idle(session, 3.0)

            # 4. Bấm nút Warehouse Tasks (N000000011 - chuẩn VBS dòng 42)
            top_grid = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
            if top_grid:
                try:
                    top_grid.pressToolbarButton("N000000011")
                    SapConnector.wait_idle(session, 2.5)
                except Exception:
                    try:
                        top_grid.pressToolbarButton("N000000108")
                        SapConnector.wait_idle(session, 2.5)
                    except Exception:
                        pass

            # 5. Đọc lưới dưới lấy WHO và MATNR (chuẩn VBS dòng 45-88)
            bottom_grid = SapConnector.find_by_id_safe(session, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
            if bottom_grid:
                row_count = getattr(bottom_grid, "RowCount", 0)
                for r in range(row_count):
                    try:
                        who_val = str(bottom_grid.GetCellValue(r, "WHO")).strip()
                        mat_val = str(bottom_grid.GetCellValue(r, "MATNR")).strip()
                        if not mat_val:
                            mat_val = str(bottom_grid.GetCellValue(r, "PRODUCTNO")).strip()

                        if who_val and mat_val:
                            if mat_val not in mat_who_map:
                                mat_who_map[mat_val] = []
                            if who_val not in mat_who_map[mat_val]:
                                mat_who_map[mat_val].append(who_val)
                    except Exception:
                        pass

                # 6. Tự động In phiếu WHO trực tiếp trong MON (chuẩn VBS dòng 92-100)
                if auto_print and row_count > 0:
                    try:
                        bottom_grid.selectAll()
                        bottom_grid.pressToolbarContextButton("METHODS")
                        bottom_grid.selectContextMenuItem("@M00007")
                        SapConnector.wait_idle(session, 1.5)

                        # Bấm nút Enter / OK trên popup xác nhận in (wnd[1]/tbar[0]/btn[0])
                        btn_confirm = SapConnector.find_by_id_safe(session, "wnd[1]/tbar[0]/btn[0]")
                        if btn_confirm:
                            btn_confirm.press()
                            SapConnector.wait_idle(session, 1.5)

                        # Bấm nút REFRESH trên toolbar lưới trên (chuẩn VBS dòng 100)
                        if top_grid:
                            try:
                                top_grid.pressToolbarButton("REFRESH")
                                SapConnector.wait_idle(session, 1.0)
                            except Exception:
                                pass
                    except Exception as print_ex:
                        print(f"[MON DIRECT PRINT EX]: {print_ex}")

        except Exception as e:
            print(f"[MON REL WARN] Lỗi tra cứu liên kết cho OB {ob_number}: {e}")

        return mat_who_map


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
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        except ImportError:
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
                input_zone = " ".join(parts[3:])
            else:
                mat = parts[0]
                station = parts[1]
                input_zone = " ".join(parts[2:])

        # Tự động chuẩn hóa và trích xuất zone
        if input_zone:
            m_z = re.search(r"(?:ZONE|Z)?\s*([0-9]+(?:\.[0-9]+)?)", input_zone, re.IGNORECASE)
            if m_z and m_z.group(1):
                input_zone = m_z.group(1)
        elif station:
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
                        JournalService.save_group(self.run_id, grp_id, grp_zone, grp.get("tl", ""), doc_or_err, "FAILED", doc_or_err)
                else:
                    mock_ob = f"50{datetime.now().strftime('%y%m%d')}{idx+1:02d}"
                    grp["ob_number"] = mock_ob
                    posted_obs.append({"group": grp, "ob_number": mock_ob})
                    JournalService.save_group(self.run_id, grp_id, grp_zone, grp.get("tl", ""), mock_ob, "MOCK_POSTED", "Mô phỏng thành công")
                    time.sleep(0.3)

            if not posted_obs:
                last_err = ""
                for g in groups:
                    if g.get("group_id") in [grp.get("group_id") for grp in groups]:
                        # find from journal or doc_or_err
                        pass
                err_detail = f"\nChi tiết lỗi: {doc_or_err}" if 'doc_or_err' in locals() and doc_or_err else ""
                summary["message"] = f"Không tạo được chứng từ MIGO nào thành công!{err_detail}"
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
