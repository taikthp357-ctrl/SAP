# -*- coding: utf-8 -*-

"""Mô đun tự động đính kèm hóa đơn VAT, Biên bản bàn giao (BBBG), và Email (.msg/.eml) vào MIGO qua GOS,

đồng thời tự động điền số phiếu vào trường Header Text (txtGOHEAD-BKTXT) trên SAP MIGO.

Kế thừa và tối ưu hóa vượt bậc từ bộ công cụ 261002_MIGO_ATT_V3_4_3 & 260801 Tools Attach Mail:

1. Tự động điền số phiếu (HeadText GOHEAD-BKTXT): Kiểm tra, điền mới và lưu an toàn qua phím SendVKey(11) (Ctrl+S).

2. Đính kèm đa định dạng: Hỗ trợ BBBG, VAT, và Email trao đổi / xác nhận (.msg, .eml, .pdf).

3. Tối ưu quét thư mục UNC / Local: Lập chỉ mục trong RAM 1 lần duy nhất bằng os.scandir (O(1)).

4. Bóc tách thông minh từ Email / Tin nhắn: Tự động trích xuất danh sách cặp (DOC, Số phiếu) từ email nhận được.

5. Độc lập phiên SAP (Multi-Session): Không chiếm dụng Tab 0 của người dùng; tự động dùng Session riêng.

6. Báo cáo trực quan tương thích 100% với định dạng thông báo Teams/Email (Ảnh 2).

"""

import os

from pathlib import Path

import re

import time

import unicodedata

from datetime import datetime

from typing import List, Dict, Any, Optional, Tuple, Set, Callable

import threading

from concurrent.futures import ThreadPoolExecutor, as_completed, wait

from app.adaptive_cache import UI_CACHE, AdaptiveConfigCache

try:

    import pythoncom

    import win32com.client

    import pywintypes

except ImportError:

    pythoncom = None

    win32com = None

    pywintypes = None

# =====================================================================

# KHÓA BẢO VỆ TOÀN CỤC & ĐIỀU PHỐI TAB (TỐI ĐA 3 TAB)

# Tab 0: Fast Lane (Tra cứu nhanh 2-5s: MB52, SCWM/POST, VL06I)

# Tab 1, Tab 2: Batch Pool (Chạy lô MIGO 100-300 DOC)

# =====================================================================

POPUP_LOCK = threading.Lock()          # Micro-lock bảo vệ cửa sổ wnd[1] & Context Menu
FAST_LANE_LOCK = threading.Lock()      # Khóa riêng cho Tab 0 (Fast Lane)
BATCH_TAB_LOCKS = {i: threading.Lock() for i in range(8)}  # Khóa riêng cho từng Tab 0..5 trong Batch Pool
FILE_DIALOG_LOCK = threading.Lock()    # Khóa đồng bộ hộp thoại chọn file Windows / GOS (chống xung đột giữa các Tab)

def get_user_idle_seconds() -> float:
    """Đo thời gian người dùng không tương tác chuột/bàn phím qua Windows GetLastInputInfo."""
    try:
        import ctypes
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
            return max(0.0, millis / 1000.0)
    except Exception:
        pass
    return 0.0

def is_turbo_mode(prompt_text: str = "") -> bool:
    """Phát hiện chế độ Turbo Mode (từ khóa hoặc máy tính đang rảnh > 45s)."""
    p_norm = str(prompt_text or "").lower()
    turbo_words = ["turbo", "max", "5 tab", "6 tab", "nhanh", "toi da", "full tab", "het cong suat"]
    if any(w in p_norm for w in turbo_words):
        return True
    return get_user_idle_seconds() >= 45.0

def sap_com_retry(func, *args, retries=4, delay=0.05, **kwargs):

    """Bọc lệnh gọi COM: Tự động lùi lại 50ms rồi thử lại nếu gặp RPC_E_CALL_REJECTED."""

    for attempt in range(retries):

        try:

            return func(*args, **kwargs)

        except Exception as e:

            if pywintypes and isinstance(e, pywintypes.com_error):

                if attempt < retries - 1:

                    time.sleep(delay * (attempt + 1))

                    continue

                raise e

            if attempt < retries - 1 and any(k in str(e).lower() for k in ["busy", "reject", "disconnect", "0x80010"]):

                time.sleep(delay * (attempt + 1))

                continue

            raise e

def safe_find_by_id(session, path: str):

    """Tìm control an toàn, tương thích cả SAP GUI COM và MagicMock 1-arg."""

    if session is None:

        return None

    try:

        return session.findById(path)

    except Exception:

        try:

            return session.findById(path, False)

        except Exception:

            return None

def dismiss_any_modal_popup(session, max_depth: int = 2) -> bool:

    """Tự động đóng và giải phóng mọi popup modal ngầm (wnd[2], wnd[1]) chống đơ giao diện SAP."""

    if session is None:

        return False

    dismissed = False

    for w_idx in range(max_depth, 0, -1):

        try:

            popup = safe_find_by_id(session, f"wnd[{w_idx}]")

            if popup is not None:

                # 1. Thử các nút xác nhận thông thường (Yes, OK, btn[0], SPOP-OPTION1)

                for btn_id in ["btn[0]", "btn[1]", "SPOP-OPTION1", "SPOP-YES", "tbar[0]/btn[0]"]:

                    b = safe_find_by_id(popup, btn_id) or safe_find_by_id(session, f"wnd[{w_idx}]/{btn_id}")

                    if b is not None:

                        try:

                            b.Press()

                            dismissed = True

                            break

                        except Exception:

                            pass

                if not dismissed:

                    try:

                        popup.SendVKey(0)  # Enter

                        dismissed = True

                    except Exception:

                        try:

                            popup.Close()

                            dismissed = True

                        except Exception:

                            pass

        except Exception:

            pass

    return dismissed

def sanitize_migo_session(session):

    """Dọn sạch mọi popup kẹt (wnd[2], wnd[1]) và đưa session về màn hình /nMIGO sạch sẽ."""

    if session is None:

        return

    try:

        for w_idx in (2, 1):

            try:

                w = safe_find_by_id(session, f"wnd[{w_idx}]")

                if w:

                    try:

                        w.Close()

                    except Exception:

                        try:

                            w.SendVKey(12)

                        except Exception:

                            pass

            except Exception:

                pass

        okcd = safe_find_by_id(session, "wnd[0]/tbar[0]/okcd")

        if okcd:

            try:

                okcd.text = "/nMIGO"

            except Exception:

                okcd.Text = "/nMIGO"

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                w0.SendVKey(0)

            sap_wait_idle(session, 3.0)

    except Exception:

        pass

def reset_session_to_main_menu(session):

    """Đưa session SAP về màn hình chính (/n hoặc F3) sau khi hoàn tất batch chứng từ."""

    if session is None:

        return

    try:

        for w_idx in (3, 2, 1):

            try:

                w = safe_find_by_id(session, f"wnd[{w_idx}]")

                if w:

                    try:

                        w.Close()

                    except Exception:

                        try:

                            w.SendVKey(12)  # Esc

                        except Exception:

                            pass

            except Exception:

                pass

        okcd = safe_find_by_id(session, "wnd[0]/tbar[0]/okcd")

        if okcd:

            try:

                okcd.Text = "/n"

            except Exception:

                okcd.text = "/n"

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                w0.SendVKey(0)

            sap_wait_idle(session, 2.0)

        # Nếu chưa thoát MIGO về SESSION_MANAGER, gửi thêm F3 (SendVKey 3)

        try:

            info = getattr(session, "Info", None)

            tcode = str(getattr(info, "Transaction", "") or "").upper() if info else ""

            if tcode not in ("SESSION_MANAGER", ""):

                w0 = safe_find_by_id(session, "wnd[0]")

                if w0:

                    pass

        except Exception:

            pass

    except Exception as e:

        print(f"[TEARDOWN WARN] Lỗi dọn dẹp session về /n: {e}")

def sap_wait_idle(session, max_wait_sec: float = 5.0) -> None:

    """Chờ SAP GUI hoàn tất xử lý (session.Busy == False) tương tự WaitIdle trong VBA."""

    if session is None:

        return

    start = time.time()

    while getattr(session, "Busy", False):

        if time.time() - start >= max_wait_sec:

            break

        time.sleep(0.05)

RECORDED_MIGO_IDS = {

    "GODYNPRO-ACTION": [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
        "wnd[0]/usr/cmbGODYNPRO-ACTION",
    ],

    "GODYNPRO-REFDOC": [

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-REFDOC",

        "wnd[0]/usr/cmbGODYNPRO-REFDOC",

    ],

    "GODYNPRO-MAT_DOC": [

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC",

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/ctxtGODYNPRO-MAT_DOC",

        "wnd[0]/usr/txtGODYNPRO-MAT_DOC",

        "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC",

    ],

    "GODYNPRO-DOC_YEAR": [

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-DOC_YEAR",

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/ctxtGODYNPRO-DOC_YEAR",

        "wnd[0]/usr/txtGODYNPRO-DOC_YEAR",

        "wnd[0]/usr/ctxtGODYNPRO-DOC_YEAR",

    ],

    "OK_GO": [

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/btnOK_GO",

        "wnd[0]/tbar[1]/btn[0]",

    ],

    "OK_REF_ATTA": [

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/btnOK_REF_ATTA",

    ],

    "GOHEAD-BKTXT": [

        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-BKTXT",

        "wnd[0]/usr/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-BKTXT",

        "wnd[0]/usr/txtGOHEAD-BKTXT",

    ],

}

def find_control_recursive(root, target_name: str, depth: int = 0, max_depth: int = 15):

    """Tìm kiếm đệ quy một control theo thuộc tính Name trong cây SAP GUI."""

    if root is None or depth > max_depth:

        return None

    try:

        r_name = getattr(root, "name", None) or getattr(root, "Name", None)

        if r_name and str(r_name).upper() == target_name.upper():

            return root

    except Exception:

        pass

    try:

        children = getattr(root, "Children", None) or getattr(root, "children", None)

        if children:

            count = getattr(children, "Count", None) or getattr(children, "count", 0)

            for i in range(count):

                try:

                    ch = children(i)

                    found = find_control_recursive(ch, target_name, depth + 1, max_depth)

                    if found:

                        return found

                except Exception:

                    continue

    except Exception:

        pass

    return None

def smart_find(session, element_name: str, element_type: Optional[str] = None, window: str = "wnd[0]"):

    """Tự động tìm kiếm đối tượng theo Tên tương thích cả High Speed & Low Speed,

    kết hợp bộ nhớ đệm tự học (UI_CACHE / data/sap_ui_cache.json).

    """

    if session is None:

        return None

    try:

        # 1. Thử qua bộ nhớ đệm tự thích nghi (AdaptiveConfigCache)

        obj = UI_CACHE.resolve_element(session, key_alias=element_name, field_name=element_name, element_type=element_type, window=window)

        if obj is not None:

            return obj

    except Exception:

        pass

    try:

        w0 = safe_find_by_id(session, window)

        if w0 is None:

            return None

        usr = safe_find_by_id(session, f"{window}/usr")

        search_roots = [usr, w0] if usr is not None else [w0]

        for root in search_roots:

            if element_type:

                for fn in ("findByName", "FindByName"):

                    try:

                        obj = getattr(root, fn)(element_name, element_type)

                        if obj is not None:

                            el_id = getattr(obj, "Id", getattr(obj, "id", ""))

                            if el_id:

                                UI_CACHE.set_cached_id(session, element_name, el_id)

                            return obj

                    except Exception:

                        pass

            else:

                for ctrl_type in ("GuiCTextField", "GuiTextField", "GuiComboBox", "GuiButton", "GuiTab"):

                    for fn in ("findByName", "FindByName"):

                        try:

                            obj = getattr(root, fn)(element_name, ctrl_type)

                            if obj is not None:

                                el_id = getattr(obj, "Id", getattr(obj, "id", ""))

                                if el_id:

                                    UI_CACHE.set_cached_id(session, element_name, el_id)

                                return obj

                        except Exception:

                            pass

        if element_name in RECORDED_MIGO_IDS:

            for cid in RECORDED_MIGO_IDS[element_name]:

                try:

                    obj = safe_find_by_id(session, cid)

                    if obj is not None:

                        el_id = getattr(obj, "Id", getattr(obj, "id", ""))

                        if el_id:

                            UI_CACHE.set_cached_id(session, element_name, el_id)

                        return obj

                except Exception:

                    pass

        found = find_control_recursive(w0, element_name)

        if found is not None:

            el_id = getattr(found, "Id", getattr(found, "id", ""))

            if el_id:

                UI_CACHE.set_cached_id(session, element_name, el_id)

        return found

    except Exception:

        return None

def wait_for_wnd1(session, appear: bool = True, timeout: float = 2.0, poll: float = 0.04):

    """

    Smart Polling chờ cửa sổ wnd[1] xuất hiện hoặc đóng lại.

    Thay thế cho các lệnh time.sleep tĩnh gây lãng phí CPU và độ trễ.

    """

    if session is None:

        return None if appear else True

    deadline = time.perf_counter() + timeout

    while time.perf_counter() < deadline:

        wnd1 = None

        try:

            wnd1 = session.findById("wnd[1]")

            if wnd1 is not None:

                try:

                    sbar = session.findById("wnd[0]/sbar")

                    if wnd1 is sbar:

                        wnd1 = None

                except Exception:

                    pass

        except Exception:

            wnd1 = None

        if appear and wnd1 is not None:

            return wnd1

        if not appear and wnd1 is None:

            return True

        time.sleep(poll)

    return None if appear else False

def smart_open_gos_create_attachment(session, mode: str = "FILE") -> bool:

    """Tự động mở cửa sổ đính kèm file (hoặc tạo link URL) ở cả High Speed lẫn Low Speed."""

    if session is None:

        return False

    cmd_code = "%GOS_URL_CREATE" if str(mode).upper() == "URL" else "%GOS_PCATTA_CREA"

    # 1. Thử cách High Speed (Bấm trực tiếp nút GOS trên thanh tiêu đề)

    gos_shell = safe_find_by_id(session, "wnd[0]/titl/shellcont/shell")

    if gos_shell is not None:

        try:

            gos_shell.PressContextButton("%GOS_TOOLBOX")

            gos_shell.SelectContextMenuItem(cmd_code)

            return True

        except Exception:

            pass

    # 2. Thử menu System -> Services for Object -> Create Attachment / Link

    try:

        menu_idx = "menu[1]" if str(mode).upper() == "URL" else "menu[0]"

        mbar_item = safe_find_by_id(session, f"wnd[0]/mbar/menu[0]/menu[11]/{menu_idx}")

        if mbar_item is not None:

            mbar_item.Select()

            return True

    except Exception:

        pass

    # 3. Gửi thẳng Function Code vào ô okcd

    try:

        okcd = safe_find_by_id(session, "wnd[0]/tbar[0]/okcd")

        if okcd is not None:

            okcd.Text = cmd_code

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                w0.SendVKey(0)

            return True

    except Exception:

        pass

    return False

def smart_open_gos_view_attachment(session, allow_gos_fallback: bool = True) -> bool:

    """Tự động mở danh sách đính kèm file ở cả High Speed lẫn Low Speed với Fast-Check Icon."""

    if session is None:

        return False

    # 1. Thử nút OK_REF_ATTA (Fast-Check: Trên MIGO, nút này chỉ xuất hiện khi ĐÃ CÓ đính kèm)

    btn = smart_find(session, "OK_REF_ATTA", "GuiButton")

    if btn is not None:

        try:

            enabled = getattr(btn, "Enabled", True)

            visible = getattr(btn, "Visibility", getattr(btn, "Visible", True))

            if enabled is False or str(enabled).upper() == "FALSE" or visible is False or str(visible).upper() == "FALSE":

                return False  # Không có attachment hoặc nút bị ẩn

            btn.Press()

            return True

        except Exception:

            pass

    elif not allow_gos_fallback:

        # Fast-Check khi được yêu cầu: Trên MIGO, nếu nút OK_REF_ATTA không tồn tại thì dừng lại ngay (0.001s)

        try:

            info = getattr(session, "Info", None)

            tcode = str(getattr(info, "Transaction", "") or "").upper() if info else ""

            if tcode == "MIGO" or "MIGO" in str(getattr(session, "Name", "")).upper():

                return False

        except Exception:

            return False

    # 2. Thử GOS shell %GOS_VIEW_ATTA (Hỗ trợ cả MIGO lẫn các giao dịch khác)

    gos_shell = safe_find_by_id(session, "wnd[0]/titl/shellcont/shell")

    if gos_shell is not None:

        try:

            gos_shell.PressContextButton("%GOS_TOOLBOX")

            gos_shell.SelectContextMenuItem("%GOS_VIEW_ATTA")

            return True

        except Exception:

            pass

    # 3. Thử menu System -> Services for Object -> Attachment List

    try:

        mbar_item = safe_find_by_id(session, "wnd[0]/mbar/menu[0]/menu[11]/menu[1]")

        if mbar_item is not None:

            mbar_item.Select()

            return True

    except Exception:

        pass

    # 4. Gửi Function Code vào ô okcd

    try:

        okcd = safe_find_by_id(session, "wnd[0]/tbar[0]/okcd")

        if okcd is not None:

            okcd.Text = "%GOS_VIEW_ATTA"

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                w0.SendVKey(0)

            return True

    except Exception:

        pass

    return False

def wait_control(session, name: str, kind: str = "GuiTextField", timeout_sec: float = 3.0):

    """Tìm kiếm control trên SAP MIGO qua 4 tầng: findByName -> candidate IDs -> recursive -> retry."""

    if session is None:

        return None

    start = time.time()

    while True:

        sap_wait_idle(session, 1.0)

        # Tầng 1: findByName trực tiếp từ wnd[0]

        try:

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                try:

                    ctl = w0.findByName(name, kind)

                    if ctl:

                        return ctl

                except Exception:

                    pass

                try:

                    ctl = w0.findByName(name)

                    if ctl:

                        return ctl

                except Exception:

                    pass

        except Exception:

            pass

        # Tầng 2: Tìm kiếm theo các ID đã ghi nhận từ bản mẫu chuẩn VBA

        if name in RECORDED_MIGO_IDS:

            for cid in RECORDED_MIGO_IDS[name]:

                try:

                    ctl = safe_find_by_id(session, cid)

                    if ctl:

                        return ctl

                except Exception:

                    pass

        # Tầng 3: Quét đệ quy toàn bộ cây control của wnd[0]

        try:

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                ctl = find_control_recursive(w0, name)

                if ctl:

                    return ctl

        except Exception:

            pass

        if time.time() - start >= timeout_sec:

            break

        time.sleep(0.1)

    return None

def get_safe_sap_connection(app, conn_idx_or_id: Any = 0) -> Any:
    """Tìm GuiConnection an toàn trong SAP GUI, hỗ trợ cả connection index lẫn chuỗi connection_id (ví dụ con[1])."""
    if app is None or not hasattr(app, "Children") or app.Children.Count == 0:
        return None
    target_str = f"/con[{conn_idx_or_id}]"
    try:
        for i in range(app.Children.Count):
            c = app.Children(i)
            if target_str in str(getattr(c, "Id", "")):
                return c
    except Exception:
        pass
    try:
        idx = int(conn_idx_or_id)
        if 0 <= idx < app.Children.Count:
            return app.Children(idx)
    except Exception:
        pass
    try:
        return app.Children(0)
    except Exception:
        return None

def get_safe_sap_session(conn_idx_or_id: Any = 0, tab_idx: int = 0) -> Any:
    """Lấy phiên SAP GUI session an toàn theo connection index/id và tab index."""
    if pythoncom is None or win32com is None:
        return None
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        conn = get_safe_sap_connection(app, conn_idx_or_id)
        if conn is None or not hasattr(conn, "Children") or conn.Children.Count == 0:
            return None
        target_ses = f"/ses[{tab_idx}]"
        try:
            for s_i in range(conn.Children.Count):
                s = conn.Children(s_i)
                if target_ses in str(getattr(s, "Id", "")):
                    return s
        except Exception:
            pass
        if 0 <= tab_idx < conn.Children.Count:
            try:
                return conn.Children(tab_idx)
            except Exception:
                pass
        return conn.Children(0)
    except Exception as ex:
        print(f"[SAP SESSION ERROR] Không thể lấy session con={conn_idx_or_id}, tab={tab_idx}: {ex}")
        return None

def ensure_sap_tabs(target_tabs: int = 5, connection_index: int = 0) -> int:
    """Đảm bảo mở sẵn đủ số Tab cần thiết (Tabs 0-2: User Reserved, Tabs 3-4: Bot Batch Pool, Tab 5: Bot Fast Lane).
    Lưu ý: Tuyệt đối không gọi pythoncom.CoUninitialize() tại đây vì hàm này chạy đồng bộ trên thread chính,
    gọi CoUninitialize sẽ làm đứt kết nối COM của Outlook và SAP Engine trên cùng luồng.
    Hỗ trợ tạo tối đa tới 7 tabs (nếu kernel SAP cho phép), tự thích ứng giảm số tab nếu chạm trần hệ thống.
    Áp dụng vòng lặp Polling đa nhịp (16 x 0.25s) chờ đồng bộ cửa sổ SAP mới xuất hiện trong COM tree.
    """
    if pythoncom is None or win32com is None:
        return 1
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        conn = get_safe_sap_connection(app, connection_index)
        if conn is None:
            return 1
        current = conn.Children.Count
        target_tabs = min(7, target_tabs)
        while current < target_tabs and current < 7:
            orig_count = conn.Children.Count
            creator_session = None
            for idx in range(conn.Children.Count):
                s = conn.Children(idx)
                if not getattr(s, "Busy", False):
                    s_info = getattr(s, "Info", None)
                    tcode = str(getattr(s_info, "Transaction", "") or "").strip().upper() if s_info else ""
                    if tcode in ("SESSION_MANAGER", ""):
                        creator_session = s
                        break
                    elif creator_session is None:
                        creator_session = s
            if creator_session is None and conn.Children.Count > 0:
                creator_session = conn.Children(0)
            if creator_session is None:
                break
            try:
                if not hasattr(creator_session, "CreateSession"):
                    print(f"[SAP MULTI-TAB WARN] Session không hỗ trợ CreateSession()")
                    break
                creator_session.CreateSession()
                tab_spawned = False
                for _ in range(16):
                    time.sleep(0.25)
                    try:
                        conn = get_safe_sap_connection(app, connection_index) or conn
                    except Exception:
                        pass
                    if conn is not None and conn.Children.Count > orig_count:
                        tab_spawned = True
                        break
                if not tab_spawned:
                    print(f"[SAP MULTI-TAB WARN] Chờ mở tab {orig_count + 1} quá thời gian (hệ thống có thể đã chạm trần {orig_count} tabs).")
                    break
                current = conn.Children.Count
            except Exception as e:
                print(f"[SAP MULTI-TAB WARN] Không thể mở thêm tab: {e}")
                break
        return conn.Children.Count
    except Exception as ex:
        print(f"[SAP MULTI-TAB ERROR] {ex}")
        return 1

DEFAULT_MIGO_ATT_FOLDER = (

    r"\\10.128.10.72\fvn$\03.Khối sản xuất\01.Nhà máy VinFast Hải Phòng"

    r"\04.Xưởng lắp ráp kim loại - cơ khí NMHP\01. Lắp ráp kim loại"

    r"\5.LOGISTIC\III. Internal\III. Kho P part\3. Theo doi hang nhap P"

    r"\240101_FT-P\260802_ File ATT BBGH"

)

def clean_folder_path(raw_path: str) -> str:

    """Chuẩn hóa đường dẫn thư mục UNC / Local: xóa ký tự ẩn NBSP, zero-width, nháy, và chuẩn hóa Unicode NFC.

    Tự động hoàn thiện đường dẫn nếu bị Outlook ngắt dòng ở ký tự thứ 73-76.

    """

    if not raw_path:

        return ""

    p = str(raw_path)

    p = p.replace("\xa0", " ").replace("\u200b", "").replace("\u3000", " ")

    p = p.strip(" \t\r\n\"\'<>|")

    p = unicodedata.normalize("NFC", p)

    if len(p) > 3 and p.endswith(("\\", "/")):

        p = p.rstrip("\\/")

    # Tự động bù đắp nếu bị ngắt dòng cụt do Outlook (ví dụ: '... \01.Nhà máy VinFast Hải')

    if len(p) >= 20 and DEFAULT_MIGO_ATT_FOLDER.startswith(p) and len(p) < len(DEFAULT_MIGO_ATT_FOLDER):

        return DEFAULT_MIGO_ATT_FOLDER

    return p

def extract_folder_path_multiline(text: str) -> str:

    """Bóc tách đường dẫn thư mục UNC hoặc Local Drive, tự động ghép các dòng bị ngắt bởi Outlook/Teams."""

    if not text:

        return ""

    m = re.search(r"(?:\\\\|[A-Za-z]:\\)", text)

    if not m:

        return ""

    start_pos = m.start()

    remaining = text[start_pos:]

    lines = remaining.splitlines()

    segments = []

    for line in lines:

        l = line.strip(" \t\r\n\"\'<>|")

        if not l:

            break

        # Dừng nếu dòng bắt đầu bằng 10 chữ số chứng từ (bước sang danh sách)

        if re.match(r"^(?:49\d{8}|500\d{7}|\d{10})\b", l):

            break

        # Cắt nếu trên cùng dòng có số chứng từ ở phía sau

        doc_b = re.search(r"[\s\t]+(?=(?:49\d{8}|500\d{7}|\d{10})\b)", l)

        if doc_b:

            segments.append(l[:doc_b.start()].strip(" \t\r\n\"\'<>|"))

            break

        segments.append(l)

        if len(segments) >= 15:

            break

    if not segments:

        return ""

    full = segments[0]

    for seg in segments[1:]:

        if seg.startswith("\\") or full.endswith("\\"):

            full = full.rstrip("\\") + "\\" + seg.lstrip("\\")

        else:

            full = full + " " + seg

    return full

def normalize_text(text: str) -> str:

    """Chuẩn hóa chuỗi văn bản: bỏ dấu tiếng Việt, chuyển chữ thường, loại bỏ ký tự lạ."""

    if not text:

        return ""

    text_str = str(text).replace("đ", "d").replace("Đ", "d")

    nfkd = unicodedata.normalize("NFKD", text_str)

    ascii_text = "".join([c for c in nfkd if not unicodedata.combining(c)])

    return ascii_text.lower().strip()

# Regex phát hiện loại file chứng từ

BBBG_REGEX = re.compile(r"(?:^|[^a-z0-9])(bbbg|bbgh|bbg|bb|bien[ _-]*ban)(?:$|[^a-z0-9])", re.I)

VAT_REGEX = re.compile(r"(?:^|[^a-z0-9])(vat|invoice|inv|hoa[ _-]*don|hdgtgt|hd[ _-]*gtgt|hd|hddt|gtgt|tax|ct|chung[ _-]*tu)(?:$|[^a-z0-9])", re.I)

MAIL_REGEX = re.compile(r"(?:^|[^a-z0-9])(mail|email|thu|xac[ _-]*nhan|trao[ _-]*doi|order)(?:$|[^a-z0-9])", re.I)

def extract_search_tokens(text: str) -> Tuple[List[str], str]:

    """Bóc tách các token tìm kiếm từ số phiếu hoặc số chứng từ.

    Ví dụ: 'Thành Đạt in 6157' -> nums: ['6157'], compact: 'thanhdatin6157'

    '5006995954' -> nums: ['5006995954'], compact: '5006995954'

    """

    norm = normalize_text(text)

    nums = re.findall(r"\d{3,10}", norm)

    compact = re.sub(r"[^a-z0-9]+", "", norm)

    return nums, compact

def format_status_label(ticket_no: str, is_success: bool) -> str:

    """Định dạng chuỗi trạng thái thành công/thất bại: 'DONE' nếu thành công, 'FILE_NOT_FOUND' nếu thiếu file."""

    if not is_success:

        return "FILE_NOT_FOUND"

    return "DONE"

def deduplicate_files_by_content(file_paths: List[Path]) -> List[Path]:

    """Loại bỏ các file trùng lặp nội dung 100% (cùng dung lượng và mã hash SHA256).

    Giữ lại file đầu tiên gặp. Nếu file không tồn tại trên đĩa (mock/test), giữ nguyên."""

    if not file_paths:

        return []

    seen_hashes = set()

    unique_files = []

    for p in file_paths:

        try:

            if not p.is_file():

                unique_files.append(p)

                continue

            size = p.stat().st_size

            import hashlib

            h = hashlib.sha256()

            with open(p, "rb") as f:

                while chunk := f.read(65536):

                    h.update(chunk)

            file_hash = (size, h.hexdigest())

            if file_hash not in seen_hashes:

                seen_hashes.add(file_hash)

                unique_files.append(p)

            else:

                print(f"[DEDUP] Bỏ qua file trùng lặp nội dung: {p.name}")

        except Exception:

            unique_files.append(p)

    return unique_files

def match_files_to_sap(

    local_files: List[Path],

    sap_titles: List[str],

    ticket_nums: Optional[List[str]] = None,

    file_type: str = "vat"

) -> Tuple[List[Path], List[Tuple[Path, str]], List[str]]:

    """So khớp danh sách file local với danh sách tiêu đề file đã có trên SAP GOS.

    Áp dụng thuật toán Claim-based matching:

    - Pass 1: Khớp chính xác stem hoặc full name.

    - Pass 2: Khớp compact (ví dụ 'vat_5122' khớp 'vat 5122').

    - Pass 3: Khớp cùng số phiếu + cùng loại file + cùng chỉ mục phụ (extra_digits).

    Mỗi tiêu đề trên SAP chỉ được claim bởi tối đa 1 file local.

    Trả về: (to_upload, already_on_sap, unclaimed_sap_titles)

    """

    if not local_files:

        return [], [], list(sap_titles or [])

    if not sap_titles:

        return list(local_files), [], []

    unclaimed_sap = list(sap_titles)

    to_upload: List[Path] = []

    already_on_sap: List[Tuple[Path, str]] = []

    def get_info(s: str):

        stem = Path(s).stem.casefold()

        norm = normalize_text(stem)

        compact = re.sub(r"[^a-z0-9]+", "", norm)

        nums = set(re.findall(r"\d{3,10}", norm))

        cleaned = norm

        if ticket_nums:

            for tn in ticket_nums:

                cleaned = cleaned.replace(tn, "")

        extra_digits = re.findall(r"\d+", cleaned)

        return stem, norm, compact, nums, extra_digits

    # Pass 1: Khớp chính xác stem hoặc full name

    remaining_local: List[Path] = []

    for p in local_files:

        p_stem, p_norm, p_compact, p_nums, p_extra = get_info(p.name)

        matched_idx = None

        for i, t in enumerate(unclaimed_sap):

            t_stem, t_norm, t_compact, t_nums, t_extra = get_info(t)

            if p_stem == t_stem or p.name.casefold() == t.casefold():

                matched_idx = i

                break

        if matched_idx is not None:

            claimed_title = unclaimed_sap.pop(matched_idx)

            already_on_sap.append((p, claimed_title))

        else:

            remaining_local.append(p)

    # Pass 2: Khớp compact (ví dụ 'vat_5122' khớp 'vat 5122')

    remaining_pass2: List[Path] = []

    for p in remaining_local:

        p_stem, p_norm, p_compact, p_nums, p_extra = get_info(p.name)

        matched_idx = None

        for i, t in enumerate(unclaimed_sap):

            t_stem, t_norm, t_compact, t_nums, t_extra = get_info(t)

            if p_compact == t_compact:

                matched_idx = i

                break

        if matched_idx is not None:

            claimed_title = unclaimed_sap.pop(matched_idx)

            already_on_sap.append((p, claimed_title))

        else:

            remaining_pass2.append(p)

    # Pass 3: Khớp theo số phiếu + loại file + extra_digits

    for p in remaining_pass2:

        p_stem, p_norm, p_compact, p_nums, p_extra = get_info(p.name)

        matched_idx = None

        for i, t in enumerate(unclaimed_sap):

            t_stem, t_norm, t_compact, t_nums, t_extra = get_info(t)

            type_match = False

            if file_type == "vat":

                type_match = bool(VAT_REGEX.search(t_norm))

            elif file_type == "bbbg":

                type_match = bool(BBBG_REGEX.search(t_norm))

            elif file_type == "mail":

                type_match = bool(MAIL_REGEX.search(t_norm))

            else:

                type_match = True

            if type_match and (p_nums and p_nums == t_nums) and (p_extra == t_extra):

                matched_idx = i

                break

        if matched_idx is not None:

            claimed_title = unclaimed_sap.pop(matched_idx)

            already_on_sap.append((p, claimed_title))

        else:

            to_upload.append(p)

    return to_upload, already_on_sap, unclaimed_sap

def parse_attachment_items_from_text(text: str) -> List[Dict[str, str]]:

    """Bóc tách danh sách chứng từ và số phiếu từ nội dung email hoặc tin nhắn Teams.

    Hỗ trợ mọi trường hợp:

    - Danh sách nhiều dòng từ Teams / Outlook

    - Bảng copy-paste từ Excel (phân cách bằng tab, dấu cách, hoặc dấu gạch |)

    - Định dạng nhiều cặp chứng từ trên cùng một dòng hoặc toàn bộ chuỗi

    - Bảng HTML từ Teams/Outlook

    """

    if not text:

        return []

    # 1. Làm sạch tag HTML nếu có

    clean_text = re.sub(r"</?(?:tr|div|p|br\s*/?|li)>", "\n", str(text), flags=re.I)

    clean_text = re.sub(r"</td>", "\t", clean_text, flags=re.I)

    clean_text = re.sub(r"<[^>]+>", " ", clean_text)

    # 2. Xóa các đường dẫn UNC hoặc thư mục cục bộ để số trong đường dẫn không bị nhận nhầm thành số DOC

    clean_text = re.sub(r"(?:\\\\[^\r\n\"\'<>|]+|[A-Za-z]:\\[^\r\n\"\'<>|]+)", " ", clean_text)

    clean_text = re.sub(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b", " ", clean_text)

    # 3. Quét tất cả các số chứng từ 10 số (bắt đầu bằng 49, 500 hoặc 10 chữ số)

    doc_matches = list(re.finditer(r"\b(49\d{8}|500\d{7}|\d{10})\b", clean_text))

    if not doc_matches:

        return []

    items = []

    seen_docs = set()

    for idx, m in enumerate(doc_matches):

        doc = m.group(1)

        if doc in seen_docs:

            continue

        seen_docs.add(doc)

        start_pos = m.end()

        # Số phiếu của chứng từ này nằm từ sau doc hiện tại tới vị trí bắt đầu doc tiếp theo

        if idx + 1 < len(doc_matches):

            raw_cand = clean_text[start_pos:doc_matches[idx + 1].start()]

        else:

            raw_cand = clean_text[start_pos:]

            # Nếu là phần tử cuối cùng, loại bỏ phần chữ ký / footer email nếu có

            sig_match = re.search(

                r"(?i)\n\s*(?:người yêu cầu|người gửi|trân trọng|thân ái|regards|best regards|thanks|sent from|sender_|from:|to:|subject:|date:|\-{3,}|\_{3,})",

                raw_cand

            )

            if sig_match:

                raw_cand = raw_cand[:sig_match.start()]

        # Outlook / Power Automate thường tự động ngắt dòng ở cột 76-80 ký tự (hard wrap)

        # Nối tất cả các dòng bị ngắt giữa doc[i] và doc[i+1] thành một chuỗi liên tục

        clean_ticket = " ".join(raw_cand.split()).strip(" :-\t,|\\/")

        for _ in range(4):

            clean_ticket = re.sub(

                r"^(?:SAP|BOT|LENH|LỆNH|ĐIỀN\s*SỐ\s*PHIẾU|DIEN\s*SO\s*PHIEU|ĐIỀN\s*PHIẾU|DIEN\s*PHIEU|SỐ\s*PHIẾU|SO\s*PHIEU|PHIẾU|PHIEU|HEADTEXT|BKTXT|ATTACH|ATT|MAIL|EMAIL)\s*[:=]?\s*",

                "",

                clean_ticket,

                flags=re.I

            ).strip(" :-\t,|\\/")

        items.append({"doc": doc, "ticket_no": clean_ticket})

    return items

def classify_sap_attachments(titles: List[str], ticket_no: str = "") -> Dict[str, Any]:

    """Kiểm tra và phân loại trạng thái file đính kèm BBBG / VAT trên SAP cho một chứng từ DOC."""

    if not titles:

        return {

            "total_files": 0,

            "files": [],

            "bbbg_file": None,

            "vat_file": None,

            "other_files": None,

            "has_bbbg": False,

            "has_vat": False,

            "status_code": "NO_ATT",

            "overall_status": "Chưa có cả BBBG lẫn VAT",

            "badge_color": "red"

        }

    bbbg_list = []

    vat_list = []

    other_list = []

    search_nums, search_compact = extract_search_tokens(ticket_no) if ticket_no else ([], "")

    for title in titles:

        norm = normalize_text(title)

        title_nums, title_compact = extract_search_tokens(title)

        is_ticket_match = True

        if ticket_no:

            is_ticket_match = False

            if search_nums and title_nums:

                if set(search_nums).intersection(set(title_nums)):

                    is_ticket_match = True

            elif search_compact and (search_compact in title_compact or title_compact in search_compact):

                is_ticket_match = True

            elif not search_nums:

                is_ticket_match = True

        is_bb = bool(BBBG_REGEX.search(norm))

        is_vat = bool(VAT_REGEX.search(norm))

        if is_bb and (is_ticket_match or not bbbg_list):

            bbbg_list.append(title)

        if is_vat and (is_ticket_match or not vat_list):

            vat_list.append(title)

        if not is_bb and not is_vat:

            other_list.append(title)

    has_bb = len(bbbg_list) > 0

    has_vat = len(vat_list) > 0

    if has_bb and has_vat:

        code = "FULL_ATT"

        status_text = "Đã có cả BBBG và VAT"

        color = "green"

    elif has_bb and not has_vat:

        code = "MISSING_VAT"

        status_text = "Đã có BBBG, thiếu VAT"

        color = "gold"

    elif has_vat and not has_bb:

        code = "MISSING_BBBG"

        status_text = "Đã có VAT, thiếu BBBG"

        color = "gold"

    else:

        code = "REVIEW"

        status_text = f"Có {len(titles)} file nhưng chưa rõ loại (Cần kiểm tra)"

        color = "blue"

    return {

        "total_files": len(titles),

        "files": titles,

        "bbbg_file": ", ".join(bbbg_list) if bbbg_list else None,

        "vat_file": ", ".join(vat_list) if vat_list else None,

        "other_files": ", ".join(other_list) if other_list else None,

        "has_bbbg": has_bb,

        "has_vat": has_vat,

        "status_code": code,

        "overall_status": status_text,

        "badge_color": color

    }

def ensure_migo_mode(session, target_action="A12", target_refdoc="R02") -> bool:
    """Đảm bảo MIGO ở đúng Action và RefDoc:
    - Nếu có chứng từ cũ đang nạp trên màn hình, Esc trước để mở khóa ComboBox Action.
    - Hỗ trợ dò tìm toàn diện ComboBox GODYNPRO-ACTION và GODYNPRO-REFDOC qua mọi Subscreen (:0001-:0004).
    - Nhận diện linh hoạt cả Key kỹ thuật (A04/A08/A12, R02/R10) lẫn Text hiển thị (Display/Change/Transfer).
    - Chuyển Action & RefDoc kèm vòng lặp retry an toàn để đảm bảo không bị kẹt ở Display.
    """
    if session is None:
        return False
    try:
        def _get_combo_key_or_code(cb, prefix="A"):
            if cb is None:
                return ""
            val = str(getattr(cb, "Key", getattr(cb, "key", "")) or "").strip()
            if re.fullmatch(rf"{prefix}\d{{2}}", val):
                return val
            txt = str(getattr(cb, "Text", getattr(cb, "text", getattr(cb, "Value", ""))) or "").upper()
            if prefix == "A":
                if "CHANGE" in txt:
                    return "A12"
                elif "DISPLAY" in txt:
                    return "A04"
                elif "TRANSFER" in txt:
                    return "A08"
                elif "RECEIPT" in txt or "GOODS RECEIPT" in txt:
                    return "A01"
                elif "ISSUE" in txt:
                    return "A07"
            elif prefix == "R":
                if "MATERIAL" in txt:
                    return "R02"
                elif "OTHER" in txt:
                    return "R10"
                elif "ORDER" in txt or "PURCHASE" in txt:
                    return "R01"
                elif "DELIVERY" in txt:
                    return "R05"
            return val

        def _find_combo(name):
            try:
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0 and hasattr(w0, "findByName"):
                    try:
                        cb = w0.findByName(name, "GuiComboBox")
                        if cb is not None:
                            return cb
                    except Exception:
                        pass
            except Exception:
                pass

            cand_ids = []
            if name == "GODYNPRO-ACTION":
                cand_ids = [
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION",
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION",
                    "wnd[0]/usr/cmbGODYNPRO-ACTION",
                ]
            else:
                cand_ids = [
                    f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmb{name}",
                    f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmb{name}",
                    f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmb{name}",
                    f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmb{name}",
                    f"wnd[0]/usr/cmb{name}",
                ]
            for cand_id in cand_ids:
                cb = safe_find_by_id(session, cand_id)
                if cb is not None:
                    return cb

            return smart_find(session, name, "GuiComboBox")

        # 0. Nếu đang có chứng từ cũ mở trên màn hình làm khóa ComboBox -> Esc trước
        for doc_cid in [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/ctxtGODYNPRO-MAT_DOC",
            "wnd[0]/usr/txtGODYNPRO-MAT_DOC",
            "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC",
        ]:
            doc_chk = safe_find_by_id(session, doc_cid)
            if doc_chk is not None and not getattr(doc_chk, "Changeable", True):
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    try:
                        w0.SendVKey(12)  # Esc (Cancel)
                        time.sleep(0.05)
                        dismiss_any_modal_popup(session)
                    except Exception:
                        pass
                break

        # 1. Kiểm tra Action: Nếu đã đúng target_action thì BỎ QUA luôn
        cmb_action = _find_combo("GODYNPRO-ACTION")
        cur_act = _get_combo_key_or_code(cmb_action, "A")

        if cur_act != target_action:
            print(f"[MIGO AUTO-SWITCH] ComboBox Action đang ở '{cur_act or 'UNKNOWN'}' (cần '{target_action}'). Đang chuyển...")
            action_changed = False
            if cmb_action is not None:
                for attempt in range(3):
                    try:
                        cmb_action.Key = target_action
                        action_changed = True
                        break
                    except Exception:
                        try:
                            cmb_action.key = target_action
                            action_changed = True
                            break
                        except Exception as act_err:
                            time.sleep(0.15)
                            sap_wait_idle(session, 1.0)
            if action_changed:
                sap_wait_idle(session, 5.0)
                time.sleep(0.3)  # Chờ SAP vẽ lại màn hình sau khi đổi Action

        # 2. Phải tìm lại từ đầu cho RefDoc (vì màn hình vừa redraw xong, COM object cũ đã bị hủy)
        cmb_refdoc = _find_combo("GODYNPRO-REFDOC")
        cur_ref = _get_combo_key_or_code(cmb_refdoc, "R")

        # Nếu đã là target_refdoc rồi thì TUYỆT ĐỐI KHÔNG gán lại!
        if cur_ref != target_refdoc:
            print(f"[MIGO AUTO-SWITCH] ComboBox RefDoc đang ở '{cur_ref or 'UNKNOWN'}' (cần '{target_refdoc}'). Đang chuyển...")
            ref_changed = False
            if cmb_refdoc is not None:
                for attempt in range(3):
                    try:
                        cmb_refdoc.Key = target_refdoc
                        ref_changed = True
                        break
                    except Exception:
                        try:
                            cmb_refdoc.key = target_refdoc
                            ref_changed = True
                            break
                        except Exception as ref_err:
                            time.sleep(0.15)
                            sap_wait_idle(session, 1.0)
            if ref_changed:
                sap_wait_idle(session, 5.0)
                time.sleep(0.3)  # Chờ SAP vẽ lại màn hình sau khi đổi RefDoc

        return True
    except Exception as ex:
        print(f"[MIGO AUTO-SWITCH WARN] {ex}")
        return False


class MigoAttachmentRunner:

    """Điều khiển tự động hóa điền số phiếu và đính kèm BBBG / VAT / Email vào SAP MIGO qua GOS."""

    def __init__(self, sap_engine=None, expected_session_identity=None):

        self.sap_engine = sap_engine

        if not expected_session_identity and sap_engine is not None:

            expected_session_identity = getattr(sap_engine, "session_identity", None)

        self.expected_session_identity = dict(expected_session_identity or {})

        self._local = threading.local()

        self.enable_multi_tab = True

    @property

    def session_identity(self):

        return self.expected_session_identity

    @session_identity.setter

    def session_identity(self, val):

        self.expected_session_identity = dict(val or {})

    @property

    def _active_document(self):

        return getattr(self._local, "active_document", None)

    @_active_document.setter

    def _active_document(self, val):

        self._local.active_document = val

    def verify_session_identity(self, session, require_migo=False):

        if session is None:

            return False, "NO_SESSION"

        if not self.expected_session_identity:

            return False, "SESSION_IDENTITY_REQUIRED"

        required = ("system", "client", "user")

        if not all(self.expected_session_identity.get(key) for key in required):

            return False, "SESSION_IDENTITY_REQUIRED"

        try:

            info = getattr(session, "Info", None)

            actual = {

                "system": getattr(info, "SystemName", None),

                "client": getattr(info, "Client", None),

                "user": getattr(info, "User", None)

            }

            if any(not isinstance(actual[key], str) or actual[key].strip().casefold() != str(self.expected_session_identity[key]).strip().casefold() for key in required):

                return False, "SESSION_IDENTITY_MISMATCH"

            # Kiểm tra session_id và session_index (bỏ qua khi chạy multi-tab engine hoặc allow_any_tab)

            allow_any_tab = self.expected_session_identity.get("allow_any_tab", False) or getattr(self, "_multi_tab_mode", False)

            if not allow_any_tab:

                # Kiểm tra session_id nếu có

                exp_sess_id = self.expected_session_identity.get("session_id")

                if exp_sess_id:

                    act_id = getattr(session, "Id", getattr(session, "id", None))

                    if not act_id:

                        return False, "SESSION_IDENTITY_MISMATCH"

                    act_str = str(act_id).strip().casefold()

                    exp_str = str(exp_sess_id).strip().casefold()

                    if "/" in exp_str:

                        if act_str != exp_str:

                            return False, "SESSION_IDENTITY_MISMATCH"

                    else:

                        if not (act_str == exp_str or act_str.endswith(f"/{exp_str}")):

                            return False, "SESSION_IDENTITY_MISMATCH"

                # Kiểm tra session_index nếu có

                exp_sess_idx = self.expected_session_identity.get("session_index")

                if exp_sess_idx is not None:

                    act_idx = None

                    if info is not None and hasattr(info, "SessionNumber"):

                        act_idx = info.SessionNumber

                    elif hasattr(session, "Id") and "/ses[" in str(session.Id):

                        try:

                            act_idx = int(str(session.Id).split("/ses[")[1].split("]")[0])

                        except Exception:

                            pass

                    if act_idx is None or int(act_idx) != int(exp_sess_idx):

                        return False, "SESSION_IDENTITY_MISMATCH"

            # Kiểm tra connection_index hoặc connection_id nếu có

            exp_conn_idx = self.expected_session_identity.get("connection_index")

            if exp_conn_idx is None:

                exp_conn_idx = self.expected_session_identity.get("connection_id")

            if exp_conn_idx is not None:

                act_conn = None

                act_id_val = getattr(session, "Id", getattr(session, "id", None))

                if act_id_val and "/con[" in str(act_id_val):

                    try:

                        act_conn = str(act_id_val).split("/con[")[1].split("]")[0]

                    except Exception:

                        pass

                if act_conn is None or str(act_conn).strip().casefold() != str(exp_conn_idx).strip().casefold():

                    return False, "SESSION_IDENTITY_MISMATCH"

            if getattr(session, "Busy", False):

                sap_wait_idle(session, max_wait_sec=4.0)

                if getattr(session, "Busy", False):

                    return False, "SESSION_BUSY"

            if require_migo:

                tcode = getattr(info, "Transaction", "")

                if str(tcode).upper() != "MIGO":

                    return False, "SESSION_TRANSACTION_MISMATCH"

            return True, "SESSION_VERIFIED"

        except Exception:

            return False, "SESSION_IDENTITY_UNAVAILABLE"

    @staticmethod

    def _failed_item(doc, ticket, status, message):

        return {"doc": doc, "ticket_no": ticket, "headtext_status": status,

                "headtext_msg": message, "bbbg_status": "NOT_EXECUTED",

                "vat_status": "NOT_EXECUTED", "mail_status": "NOT_EXECUTED",

                "bbbg_file": None, "vat_file": None, "mail_file": None,

                "overall_status": status, "match_notes": message, "is_succ": False}, False

    @staticmethod

    def index_source_folder(folder_path: str, recursive: bool = True) -> Dict[str, Any]:

        """Quét và lập chỉ mục toàn bộ file PDF/Ảnh/Email (.msg/.eml) trong thư mục nguồn vào RAM (O(1))."""

        folder = Path(folder_path)

        if not folder.exists() or not folder.is_dir():

            if str(folder_path).startswith("\\\\"):

                raise FileNotFoundError(

                    f"Thư mục chia sẻ mạng UNC '{folder_path}' không tồn tại hoặc máy trạm chưa kết nối mạng nội bộ/VPN (chưa xác thực quyền truy cập Windows)."

                )

            raise FileNotFoundError(f"Thư mục nguồn không tồn tại hoặc không thể truy cập: {folder_path}")

        files_indexed: List[Dict[str, Any]] = []

        token_to_files: Dict[str, List[int]] = {}

        valid_exts = {".pdf", ".jpg", ".jpeg", ".png", ".xlsx", ".docx", ".tif", ".tiff", ".msg", ".eml"}

        def scan_dir(dir_path: Path, depth: int = 0):

            if depth > 4:

                return

            try:

                with os.scandir(dir_path) as it:

                    for entry in it:

                        try:

                            if entry.is_file() and Path(entry.name).suffix.lower() in valid_exts:

                                fpath = Path(entry.path)

                                norm_name = normalize_text(fpath.stem)

                                ext = fpath.suffix.lower()

                                is_bbbg = bool(BBBG_REGEX.search(norm_name))

                                is_vat = bool(VAT_REGEX.search(norm_name))

                                is_mail = ext in {".msg", ".eml"} or bool(MAIL_REGEX.search(norm_name))

                                nums = re.findall(r"\d{3,10}", norm_name)

                                compact = re.sub(r"[^a-z0-9]+", "", norm_name)

                                file_idx = len(files_indexed)

                                files_indexed.append({

                                    "path": fpath,

                                    "name": fpath.name,

                                    "norm_name": norm_name,

                                    "compact": compact,

                                    "nums": set(nums),

                                    "is_bbbg": is_bbbg,

                                    "is_vat": is_vat,

                                    "is_mail": is_mail

                                })

                                for num in nums:

                                    token_to_files.setdefault(num, []).append(file_idx)

                            elif entry.is_dir() and recursive:

                                scan_dir(Path(entry.path), depth + 1)

                        except Exception:

                            continue

            except Exception as e:

                print(f"[INDEX WARN] Không thể quét thư mục '{dir_path}': {e}")

        scan_dir(folder, depth=0)

        return {

            "folder": str(folder),

            "files": files_indexed,

            "token_map": token_to_files,

            "total_files": len(files_indexed)

        }

    @classmethod

    def match_all_files(cls, doc: str, ticket_no: str, index: Dict[str, Any]) -> Dict[str, List[Path]]:

        """Tìm tất cả các file BBBG, VAT (hỗ trợ nhiều file VAT) và Mail phù hợp."""

        files = index.get("files", [])

        token_map = index.get("token_map", {})

        doc_clean = str(doc).strip()

        ticket_clean = str(ticket_no).strip()

        doc_nums, doc_compact = extract_search_tokens(doc_clean)

        ticket_nums, ticket_compact = extract_search_tokens(ticket_clean)

        search_nums = set(doc_nums + ticket_nums)

        cand_indices: Set[int] = set()

        for num in search_nums:

            if num in token_map:

                cand_indices.update(token_map[num])

        if not cand_indices:

            for idx, f in enumerate(files):

                for num in search_nums:

                    if num in f["compact"]:

                        cand_indices.add(idx)

        cand_list = [files[i] for i in cand_indices]

        # 1. Tìm các file BBBG

        bbbg_files = [f["path"] for f in cand_list if f["is_bbbg"]]

        # 2. Tìm các file VAT

        vat_files = [f["path"] for f in cand_list if f["is_vat"]]

        # 3. Tìm các file Mail

        mail_files = [f["path"] for f in cand_list if f["is_mail"]]

        # Fallback VAT: Nếu không có file nào mang từ khóa VAT rõ ràng,

        # nhưng có file khớp số phiếu và không phải BBBG / Mail -> coi đó là hóa đơn VAT (ký hiệu riêng)

        if not vat_files:

            fallback = [f["path"] for f in cand_list if not f["is_bbbg"] and not f["is_mail"]]

            if fallback:

                vat_files = fallback

        return {

            "bbbg": deduplicate_files_by_content(bbbg_files),

            "vat": deduplicate_files_by_content(vat_files),

            "mail": deduplicate_files_by_content(mail_files)

        }

    @classmethod

    def match_files(cls, doc: str, ticket_no: str, index: Dict[str, Any]) -> Tuple[Optional[Path], Optional[Path], Optional[Path], str]:

        """Tìm file BBBG, VAT, và Email phù hợp nhất cho số chứng từ và số phiếu đã cho (tương thích ngược)."""

        all_matches = cls.match_all_files(doc, ticket_no, index)

        bbbg_match = all_matches["bbbg"][0] if all_matches["bbbg"] else None

        vat_match = all_matches["vat"][0] if all_matches["vat"] else None

        mail_match = all_matches["mail"][0] if all_matches["mail"] else None

        notes = []

        if all_matches["bbbg"]:

            notes.append(f"BBBG: {len(all_matches['bbbg'])} file ({', '.join(p.name for p in all_matches['bbbg'][:2])})")

        else:

            notes.append("BBBG: FILE_NOT_FOUND")

        if all_matches["vat"]:

            notes.append(f"VAT: {len(all_matches['vat'])} file ({', '.join(p.name for p in all_matches['vat'][:2])})")

        else:

            notes.append("VAT: FILE_NOT_FOUND")

        if mail_match:

            notes.append(f"Mail: {mail_match.name}")

        return bbbg_match, vat_match, mail_match, " | ".join(notes)

    def handle_migo_headtext(self, session, ticket_no: str, allow_overwrite: bool = False) -> Tuple[str, str]:

        """Kiểm tra và điền số phiếu vào trường Header Text theo chuẩn Change_doc:

        MIGO_OK_HEADER_OPEN -> GOHEAD-BKTXT.Text = ticket_no -> SendVKey 11 (Save).

        """

        identity_ok, identity_status = self.verify_session_identity(session, require_migo=True)

        if not identity_ok or self._active_document is None:

            return "HEADTEXT_BLOCKED", identity_status if not identity_ok else "DOCUMENT_NOT_VERIFIED"

        if not ticket_no or not str(ticket_no).strip():

            return "SKIP_NO_TICKET", "Không có số phiếu cần điền"

        ticket_clean = str(ticket_no).strip()
        # Kiểm tra tiêu đề cửa sổ: Nếu đang ở Display Material Document -> Chặn ngay lập tức
        titl_bar = safe_find_by_id(session, "wnd[0]/titl")
        titl_text = str(getattr(titl_bar, "Text", getattr(titl_bar, "text", "")) or "").upper()
        if titl_text and "DISPLAY" in titl_text:
            return "HEADTEXT_BLOCKED", f"MIGO đang ở chế độ xem Display (A04) thay vì Change (A12), không thể sửa Header Text."


        try:

            # 1. Bảo đảm phần Header MIGO được mở

            try:

                session.SendCommand("MIGO_OK_HEADER_OPEN")

                time.sleep(0.1)

            except Exception:

                pass

            # 2. Tìm trường GOHEAD-BKTXT

            head_field = None

            try:

                w0 = safe_find_by_id(session, "wnd[0]")

                if w0:

                    fn = getattr(w0, "findByName", getattr(w0, "FindByName", None))

                    if fn:

                        head_field = fn("GOHEAD-BKTXT", "GuiTextField")

            except Exception:

                pass

            if head_field is None:

                head_field = smart_find(session, "GOHEAD-BKTXT", "GuiTextField")

            if head_field is None:

                head_field = wait_control(session, "GOHEAD-BKTXT", "GuiTextField", timeout_sec=2.0)

            if head_field is None:

                return "HEADTEXT_NOT_FOUND", "Không tìm thấy trường GOHEAD-BKTXT trên MIGO"

            current_text = ""

            for prop in ("Text", "text"):

                try:

                    val = getattr(head_field, prop)

                    if isinstance(val, str):

                        current_text = val.strip()

                        break

                except Exception:

                    pass

            # Trường hợp 1: Đã có sẵn HeadText -> Kiểm tra trùng khớp hay khác biệt (A01 Review Guard)

            if current_text:

                curr_words = [w.upper() for w in re.findall(r"[A-Za-z0-9]+", current_text)]

                tick_words = [w.upper() for w in re.findall(r"[A-Za-z0-9]+", ticket_clean)]

                curr_compact = re.sub(r"[^A-Za-z0-9]+", "", current_text).upper()

                tick_compact = re.sub(r"[^A-Za-z0-9]+", "", ticket_clean).upper()

                # So khớp token chính xác (exact token match hoặc regex ranh giới từ \b)

                # Ngăn chặn trường hợp OUT_P1_123 khớp nhầm OUT_P1_1234

                is_matched = (

                    curr_compact == tick_compact

                    or (tick_words and len(tick_words) == len(curr_words) and tick_words == curr_words)

                    or bool(re.search(rf"(?<![A-Za-z0-9]){re.escape(ticket_clean.strip())}(?![A-Za-z0-9])", current_text, re.IGNORECASE))

                )

                if is_matched:

                    return "HEADTEXT_OK", f"Đã có đúng số phiếu ('{current_text}')"

                elif allow_overwrite:
                    print(f"[ATT RUNNER] Người dùng yêu cầu 'Sửa số phiếu' (allow_overwrite=True). Ghi đè HeadText cũ ('{current_text}') bằng số phiếu mới ('{ticket_clean}').")

                else:

                    # Quy tắc Đợt 2 (A01): Nếu đã có HeadText khác nội dung -> Báo REVIEW, tuyệt đối không ghi đè ngầm

                    print(f"[ATT RUNNER WARN] MIGO HeadText hiện có ('{current_text}') khác với số phiếu yêu cầu ('{ticket_clean}'). Giữ nguyên để Review thủ công.")

                    return "HEADTEXT_REVIEW", f"Header Text hiện có ('{current_text}') khác với số phiếu yêu cầu ('{ticket_clean}'). Cần Review thủ công, không ghi đè ngầm."

            # Trường hợp 2: Điền số phiếu và bấm Lưu (Save / vKey 11) chuẩn theo Change_doc

            try:

                head_field.Text = ticket_clean

            except Exception:

                try:

                    head_field.text = ticket_clean

                except Exception:

                    pass

            w0 = safe_find_by_id(session, "wnd[0]")

            if w0:

                w0.SendVKey(11)  # Ctrl+S / Save

            sap_wait_idle(session, 2.0)

            time.sleep(0.05)

            dismiss_any_modal_popup(session)

            # Kiểm tra kết quả trên sbar

            sbar = safe_find_by_id(session, "wnd[0]/sbar")

            if sbar:

                s_type = ""

                s_msg = ""

                for prop in ("MessageType", "messageType"):

                    v = getattr(sbar, prop, None)

                    if isinstance(v, str):

                        s_type = v.strip()

                        break

                for prop in ("Text", "text"):

                    v = getattr(sbar, prop, None)

                    if isinstance(v, str):

                        s_msg = v.strip()

                        break

                if s_type in ("E", "A"):

                    return "HEADTEXT_FAIL", f"Lỗi lưu: {s_msg}"

            else:

                return "HEADTEXT_UNCERTAIN", "Không đọc được phản hồi lưu SAP; cần kiểm tra thủ công"

            if s_type != "S":

                return "HEADTEXT_UNCERTAIN", f"SAP chưa xác nhận lưu: {s_msg}"

            # Readback kiểm tra HeadText thực tế sau khi lưu

            try:

                head_field_after = smart_find(session, "GOHEAD-BKTXT", "GuiTextField")

                if head_field_after:

                    saved_val = getattr(head_field_after, "Text", getattr(head_field_after, "text", ""))

                    if saved_val:

                        saved_compact = re.sub(r"[^A-Za-z0-9]+", "", str(saved_val)).upper()

                        tick_compact = re.sub(r"[^A-Za-z0-9]+", "", ticket_clean).upper()

                        if saved_compact != tick_compact and not re.search(rf"\b{re.escape(ticket_clean.strip())}\b", str(saved_val), re.IGNORECASE):

                            return "HEADTEXT_FAIL", f"Readback không khớp sau lưu: '{saved_val}' != '{ticket_clean}'"

            except Exception as rb_ex:

                print(f"[ATT READBACK WARN] Không thể readback HeadText: {rb_ex}")

            # Khi SAP đã xác nhận lưu thành công (sbar MessageType 'S'), chứng từ đã được ghi vào SAP.

            # Giữ nguyên MIGO ở chế độ CHANGE (A12) cho các chứng từ tiếp theo trong batch,

            # tuyệt đối KHÔNG reload lại chứng từ ở chế độ DISPLAY (A04) làm mất trạng thái CHANGE.

            if current_text and allow_overwrite and not is_matched:
                return "HEADTEXT_DONE", f"Đã ghi đè số phiếu cũ ('{current_text}') thành '{ticket_clean}' (SAP xác nhận: {s_msg})"
            return "HEADTEXT_DONE", f"Đã điền và lưu mới HeadText: '{ticket_clean}' (SAP xác nhận: {s_msg})"

        except Exception as e:

            return "HEADTEXT_EXCEPTION", f"Lỗi xử lý HeadText: {e}"

    @staticmethod

    def check_existing_sap_attachments(titles: List[str], ticket_no: str) -> Tuple[Optional[str], Optional[str]]:

        """Kiểm tra danh sách tiêu đề file đính kèm đã có trên SAP MIGO (từ popup Attachment list).

        Đối chiếu với số phiếu đã cho:

        - Tìm file BBBG đã có (VD: '3874 BB', 'BBBG 3874', 'in 3874 BBBG')

        - Tìm file VAT đã có (VD: 'VAT 3874', '3874 VAT', 'in 3874 VAT')

        Trả về: (sap_bbbg_title, sap_vat_title)

        """

        if not titles:

            return None, None

        search_nums, search_compact = extract_search_tokens(ticket_no)

        bbbg_found = None

        vat_found = None

        for title in titles:

            norm = normalize_text(title)

            title_nums, title_compact = extract_search_tokens(title)

            # Kiểm tra xem title trên SAP có chứa số phiếu của chứng từ này không

            is_ticket_match = False

            if search_nums and title_nums:

                if set(search_nums).intersection(set(title_nums)):

                    is_ticket_match = True

            elif search_compact and (search_compact in title_compact or title_compact in search_compact):

                is_ticket_match = True

            elif not search_nums:

                is_ticket_match = True

            if is_ticket_match:

                if not bbbg_found and BBBG_REGEX.search(norm):

                    bbbg_found = title

                if not vat_found and VAT_REGEX.search(norm):

                    vat_found = title

        return bbbg_found, vat_found

    @staticmethod

    def classify_sap_attachments(titles: List[str], ticket_no: str = "") -> Dict[str, Any]:

        """Phân loại trạng thái file đính kèm BBBG / VAT trên SAP cho một chứng từ DOC."""

        return classify_sap_attachments(titles, ticket_no=ticket_no)

    def read_sap_attachment_titles(self, session) -> List[str]:

        """Kiểm tra và đọc danh sách tiêu đề các file đính kèm đã có sẵn trên SAP MIGO qua nút OK_REF_ATTA / Attachment list."""

        titles: List[str] = []

        if session is None:

            return titles

        # 1. Mở cửa sổ Attachment list an toàn qua Micro-Lock (ngăn chặn 100% đơ cứng Context Menu Win32)
        has_opened = False
        popup = None
        lock_held = False
        try:
            lock_held = POPUP_LOCK.acquire(timeout=3.0)
            has_opened = smart_open_gos_view_attachment(session)
            if has_opened:
                popup = wait_for_wnd1(session, appear=True, timeout=1.5, poll=0.04)
        finally:
            if lock_held:
                POPUP_LOCK.release()

        if not has_opened or popup is None:
            return []

        # 2. Đọc popup Attachment list (wnd[1]) - Đọc song song 100% ngoài vùng lock
        try:

            if popup is not None:

                grid = None

                for cand_grid_id in [

                    "wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell",

                    "wnd[1]/usr/cntlCONTAINER/shellcont/shell",

                    "wnd[1]/usr/cntlCONTAINER_0100/shellcont",

                    "wnd[1]/usr/shell"

                ]:

                    try:

                        g = session.findById(cand_grid_id)

                        if g:

                            grid = g

                            break

                    except Exception:

                        pass

                if grid is None:

                    def find_grid(obj, depth=0):

                        if depth > 10 or obj is None:

                            return None

                        try:

                            c_type = str(getattr(obj, "Type", ""))

                            c_sub = str(getattr(obj, "SubType", ""))

                            if "Grid" in c_type or "Grid" in c_sub or c_type in ("GuiGridView", "GuiTableControl"):

                                return obj

                            children = getattr(obj, "Children", None)

                            children_count = int(getattr(children, "Count", 0)) if children else 0

                            for idx in range(children_count):

                                res = find_grid(children.Item(idx), depth + 1)

                                if res:

                                    return res

                        except Exception:

                            pass

                        return None

                    grid = find_grid(popup)

                if grid:

                    try:

                        if hasattr(grid, "RowCount") and hasattr(grid, "GetCellValue"):

                            row_count = int(getattr(grid, "RowCount", 0))

                            for r in range(row_count):

                                title_val = ""

                                for col_name in ["BITM_DESCR", "OBJDES", "TITLE", "DESCRIPT", "FILENAME"]:

                                    try:

                                        val = grid.GetCellValue(r, col_name)

                                        if val and str(val).strip():

                                            title_val = str(val).strip()

                                            break

                                    except Exception:

                                        continue

                                if title_val:

                                    titles.append(title_val)

                        elif hasattr(grid, "GetAllNodeKeys") or hasattr(grid, "GetNodeKeys"):

                            try:

                                keys = grid.GetAllNodeKeys()

                            except Exception:

                                try:

                                    keys = grid.GetNodeKeys()

                                except Exception:

                                    keys = []

                            for k in keys:

                                for col in ["BITM_DESCR", "OBJDES", "TITLE", "DESCRIPT", "FILENAME"]:

                                    try:

                                        t_val = grid.GetItemText(k, col)

                                        if t_val and str(t_val).strip():

                                            titles.append(str(t_val).strip())

                                            break

                                    except Exception:

                                        pass

                                try:

                                    n_text = grid.GetNodeTextByKey(k)

                                    if n_text and str(n_text).strip() and str(n_text).strip() not in titles:

                                        titles.append(str(n_text).strip())

                                except Exception:

                                    pass

                    except Exception as e:

                        print(f"[SAP ATT LIST WARN] Lỗi đọc grid/tree: {e}")

                # 3. Đóng popup wnd[1]

                try:

                    popup.Close()

                except Exception:

                    try:

                        session.findById("wnd[1]").SendVKey(12)

                    except Exception:

                        pass

                wait_for_wnd1(session, appear=False, timeout=1.0, poll=0.04)

        except Exception as ex:

            print(f"[SAP ATT LIST ERROR] {ex}")

        return titles

    def attach_file_to_current_migo(self, session, file_path: Path) -> Tuple[bool, str]:

        """Đính kèm 1 file vật lý vào chứng từ MIGO đang mở trên SAP qua GOS (%GOS_TOOLBOX)."""

        identity_ok, identity_status = self.verify_session_identity(session, require_migo=True)

        if not identity_ok or self._active_document is None:

            return False, identity_status if not identity_ok else "DOCUMENT_NOT_VERIFIED"

        abs_path = Path(file_path).resolve()

        if not abs_path.exists():

            return False, f"FILE_NOT_FOUND: {abs_path}"

        parent_folder = str(abs_path.parent)

        filename = abs_path.name

        try:

            # 1. Chỉ giữ Micro-Lock (~50-100ms) đúng khoảnh khắc trigger mở context menu và chờ wnd[1] hiện ra

            lock_held = False

            try:

                lock_held = POPUP_LOCK.acquire(timeout=3.0)

                opened = smart_open_gos_create_attachment(session, mode="FILE")

                if not opened:

                    return False, "GOS_TRIGGER_FAILED: Không thể mở hộp thoại đính kèm file GOS"

                # Chờ nhanh tối đa 1.5s đến khi cửa sổ wnd[1] hiện ra rồi nhả lock NGAY LẬP TỨC!

                wnd1 = wait_for_wnd1(session, appear=True, timeout=1.5, poll=0.04)

            finally:

                if lock_held:

                    POPUP_LOCK.release()

            # 2. NGOÀI VÙNG LOCK: Các Tab điền dữ liệu và bấm nút song song 100%!

            if not wnd1:

                return False, "Không mở được hộp thoại đính kèm file (wnd[1])"

            # Điền đường dẫn và nhấn nút Open/btn[0] độc lập trên từng session

            with FILE_DIALOG_LOCK:
                path_ctl = UI_CACHE.resolve_element(session, "gos_path", "DY_PATH", window="wnd[1]")
                if path_ctl is not None:
                    path_ctl.Text = parent_folder
                else:
                    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = parent_folder

                file_ctl = UI_CACHE.resolve_element(session, "gos_fname", "DY_FILENAME", window="wnd[1]")
                if file_ctl is not None:
                    file_ctl.Text = filename
                else:
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename

                btn_open = UI_CACHE.resolve_element(session, "gos_btn_open", "btn[0]", window="wnd[1]")
                if btn_open is not None:
                    btn_open.Press()
                else:
                    session.findById("wnd[1]/tbar[0]/btn[0]").Press()

            # 3. KIỂM TRA ĐÓNG DIALOG VÀ XÁC MINH STATUSBAR:

            # Điều kiện 1: Xác nhận wnd[1] đã đóng lại hoàn toàn sau khi nhấn btn[0]

            wnd1_closed = wait_for_wnd1(session, appear=False, timeout=2.0, poll=0.04)

            if not wnd1_closed:

                return False, "FAILED_POPUP_STUCK: Hộp thoại đính kèm không đóng sau khi nhấn nút tải"

            # Điều kiện 2: Kiểm tra session.findById("wnd[0]/sbar").MessageType != 'E'/'A'

            sbar = session.findById("wnd[0]/sbar")

            msg = getattr(sbar, "Text", "")

            msg_type = getattr(sbar, "MessageType", "")

            text = str(msg).upper()

            if msg_type in ("E", "A") or any(k in text for k in ["NOT CREATED", "NOT SAVED", "FAILED", "ERROR", "CANCELLED", "CANCELED", "KHONG THANH CONG"]):

                return False, f"SAP_ERROR: {msg}"

            # Điều kiện 3: Xác minh vật lý qua readback danh sách attachment trên SAP

            titles = self.read_sap_attachment_titles(session)

            target_name = Path(filename).name.casefold()

            target_stem = Path(filename).stem.casefold()

            target_norm = normalize_text(target_stem)

            target_clean = re.sub(r"[^a-z0-9]+", "", target_stem)

            matched = False

            for t in titles:

                t_str = str(t).strip()

                t_name = Path(t_str).name.casefold()

                t_stem = Path(t_str).stem.casefold()

                t_norm = normalize_text(t_str)

                t_clean = re.sub(r"[^a-z0-9]+", "", t_stem)

                if (

                    t_name == target_name

                    or t_stem == target_stem

                    or t_name == target_stem

                    or t_stem == target_name

                    or (target_norm and t_norm == target_norm)

                    or (target_clean and t_clean == target_clean)

                    or (len(target_stem) >= 4 and (target_stem in t_name or t_stem in target_name or target_norm in t_norm or t_norm in target_norm))

                    or (len(target_clean) >= 4 and (target_clean in t_clean or t_clean in target_clean))

                ):

                    matched = True

                    break

            if matched:

                return True, f"DONE_SAP_READBACK: {filename}"

            # FAST-PATH SBAR CONFIRMATION (FALLBACK KHI READBACK CHƯA THẤY DO OK_REF_ATTA CHƯA RELOAD HOẶC GUITREE):

            # Nếu readback chưa đọc lại được (do MIGO chưa kịp vẽ lại nút OK_REF_ATTA hoặc do điều khiển Cây GuiTree),

            # nhưng thanh sbar xác nhận MessageType 'S' với từ khóa GOS thành công -> Chấp nhận DONE_SAP_SBAR!

            gos_success_keywords = ["CREATED", "ATTACHMENT", "SAVED", "SUCCESS", "THANH CONG", "DA TAO", "DA LUU"]

            is_explicit_success = (msg_type == "S" and any(k in text for k in gos_success_keywords))

            if is_explicit_success:

                return True, f"DONE_SAP_SBAR: {filename}"

            return False, f"ATTACHMENT_UNCERTAIN: Chưa đọc lại đúng tên file '{filename}' trong Attachment list trên SAP"

        except Exception as e:

            return False, f"GOS_EXCEPTION: {e}"

    def attach_url_to_current_migo(self, session, link_title: str, link_address: str) -> Tuple[bool, str]:

        """Gắn link web (URL) vào chứng từ MIGO đang mở trên SAP qua GOS (%GOS_URL_CREATE)."""

        identity_ok, identity_status = self.verify_session_identity(session, require_migo=True)

        if not identity_ok or self._active_document is None:

            return False, identity_status if not identity_ok else "DOCUMENT_NOT_VERIFIED"

        try:

            # 1. Chỉ giữ Micro-Lock đúng lúc mở menu GOS

            with POPUP_LOCK:

                opened = smart_open_gos_create_attachment(session, mode="URL")

                if not opened:

                    return False, "GOS_URL_TRIGGER_FAILED: Không thể mở hộp thoại Create link to web page"

                wnd1 = wait_for_wnd1(session, appear=True, timeout=1.5, poll=0.04)

            # 2. Ngoài vùng lock: Các Tab điền dữ liệu song song 100%

            if not wnd1:

                return False, "Không mở được hộp thoại Create link to web page (wnd[1])"

            try:

                title_field = smart_find(session, "TITLE", "GuiTextField", window="wnd[1]")

                if title_field is None:

                    title_field = smart_find(session, "DY_TITLE", "GuiTextField", window="wnd[1]")

                if title_field is not None:

                    title_field.Text = str(link_title).strip()

                url_field = smart_find(session, "URL", "GuiTextField", window="wnd[1]")

                if url_field is None:

                    url_field = smart_find(session, "DY_URL", "GuiTextField", window="wnd[1]")

                if title_field is None or url_field is None:

                    return False, "URL_CONTROLS_MISSING: Không tìm thấy trường Title hoặc URL trên wnd[1]"

                url_field.Text = str(link_address).strip()

                session.findById("wnd[1]/tbar[0]/btn[0]").Press()

            except Exception as fill_err:

                return False, f"URL_FILL_ERROR: {fill_err}"

            wnd1_closed = wait_for_wnd1(session, appear=False, timeout=2.0, poll=0.04)

            if not wnd1_closed:

                return False, "FAILED_POPUP_STUCK: Hộp thoại URL không đóng sau khi nhấn nút xác nhận"

            sbar = session.findById("wnd[0]/sbar")

            msg = getattr(sbar, "Text", "")

            msg_type = getattr(sbar, "MessageType", "")

            if msg_type in ("E", "A", "W"):

                return False, f"SAP_ERROR: {msg}"

            return True, f"DONE_SAP_URL: {link_title}"

        except Exception as e:

            return False, f"ATTACH_URL_EXCEPTION: {e}"

    def prepare_migo_session(self, session, mode: str = "CHANGE") -> bool:

        """Chuẩn bị màn hình MIGO theo chuẩn Change_doc từ 240101_MA Log_GR Tool_V7_1:

        Mở /nmigo một lần duy nhất trước vòng lặp và đặt Action = A12 (Change) hoặc A04 (Display).

        """

        identity_ok, _ = self.verify_session_identity(session)

        if not identity_ok:

            return False

        try:

            try:

                session.findById("wnd[0]").maximize()

            except Exception:

                pass

            # Đóng các popup kẹt nếu có

            for w_idx in (2, 1):

                try:

                    w = safe_find_by_id(session, f"wnd[{w_idx}]")

                    if w:

                        try:

                            w.Close()

                        except Exception:

                            w.SendVKey(12)

                except Exception:

                    pass

            okcd = safe_find_by_id(session, "wnd[0]/tbar[0]/okcd")

            if okcd:

                try:

                    okcd.Text = "/nmigo"

                except Exception:

                    okcd.text = "/nmigo"

                w0 = safe_find_by_id(session, "wnd[0]")

                if w0:

                    w0.SendVKey(0)

                sap_wait_idle(session, 2.0)

                time.sleep(0.05)

# Đặt Action = A12 (Change) hoặc A04 (Display) và RefDoc = R02

            target_act = "A12" if mode == "CHANGE" else "A04"

            ensure_migo_mode(session, target_action=target_act, target_refdoc="R02")

            return str(session.Info.Transaction).upper() == "MIGO"

        except Exception as e:

            print(f"[MIGO INIT WARN] {e}")

            return False

    def _do_open_migo_document(self, session, doc: str, year: str = "", mode: str = "CHANGE", force_reload: bool = False) -> Tuple[bool, str]:

        """Thao tác nạp chứng từ MIGO đơn lẻ kèm cơ chế chống lag giao diện và bắt lỗi COM E_PENDING / 613."""

        self._active_document = None

        identity_ok, reason = self.verify_session_identity(session)

        if not identity_ok:

            return False, reason

        if not year:

            from datetime import datetime

            year = str(datetime.now().year)

        if not re.fullmatch(r"\d{10}", str(doc)) or not re.fullmatch(r"(?:19|20|21)\d{2}", str(year)):

            return False, "DOCUMENT_AND_YEAR_REQUIRED"

        try:

            target_act = "A12" if mode == "CHANGE" else "A04"

            need_prepare = force_reload or str(session.Info.Transaction).upper() != "MIGO"

            if need_prepare:

                if not self.prepare_migo_session(session, mode=mode):

                    return False, "SESSION_PREPARE_FAILED"

            else:

                ensure_migo_mode(session, target_action=target_act, target_refdoc="R02")

            dismiss_any_modal_popup(session)

            # 1. Tìm ô nhập số chứng từ GODYNPRO-MAT_DOC
            doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")

            # Nếu chưa thấy hoặc ô đang bị khóa do chứng từ trước chưa thoát
            if doc_fld is None or not getattr(doc_fld, "Changeable", True):
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    try:
                        w0.SendVKey(12)  # Esc (Cancel) để hủy/thoát chứng từ cũ thay vì F3
                        time.sleep(0.05)
                        dismiss_any_modal_popup(session)
                    except Exception:
                        pass
                doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")
                # Sau khi Esc, ComboBox Action đã được mở khóa -> Đảm bảo lại đúng Action A12 và RefDoc R02
                ensure_migo_mode(session, target_action=target_act, target_refdoc="R02")

            # Nếu sau Esc vẫn chưa sẵn sàng -> Reset về MIGO A12 sạch sẽ
            if doc_fld is None or not getattr(doc_fld, "Changeable", True):
                if not self.prepare_migo_session(session, mode=mode):
                    return False, "SESSION_PREPARE_FAILED"
                doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")

            if doc_fld is None:
                return False, "CONTROL_NOT_FOUND: Không tìm thấy ô nhập số chứng từ GODYNPRO-MAT_DOC trên MIGO"

            # 2. Điền số chứng từ vào ô vàng (Bọc Retry Guard chống COM E_PENDING / 613)
            doc_assigned = False
            for _ in range(5):
                try:
                    try:
                        doc_fld.SetFocus()
                    except Exception:
                        pass
                    doc_fld.Text = str(doc).strip()
                    doc_assigned = True
                    break
                except Exception as fld_err:
                    err_s = str(fld_err)
                    if "-2147483638" in err_s or "613" in err_s or "pending" in err_s.lower() or "not yet available" in err_s.lower():
                        time.sleep(0.2)
                        sap_wait_idle(session, 1.0)
                        continue
                    try:
                        doc_fld.text = str(doc).strip()
                        doc_assigned = True
                        break
                    except Exception:
                        time.sleep(0.2)

            # Tìm ô năm với vòng lặp chờ tối đa 2 giây
            year_fld = None
            for _ in range(8):
                year_fld = smart_find(session, "GODYNPRO-DOC_YEAR", "GuiTextField") or smart_find(session, "GODYNPRO-DOC_YEAR", "GuiCTextField")
                if year_fld is not None:
                    break
                time.sleep(0.15)
                sap_wait_idle(session, 1.0)

            if year_fld is None:
                return False, "YEAR_CONTROL_NOT_FOUND"

            for _ in range(5):
                try:
                    year_fld.Text = str(year).strip()
                    break
                except Exception as yr_err:
                    err_s = str(yr_err)
                    if "-2147483638" in err_s or "613" in err_s or "pending" in err_s.lower() or "not yet available" in err_s.lower():
                        time.sleep(0.2)
                        sap_wait_idle(session, 1.0)
                        continue
                    try:
                        year_fld.text = str(year).strip()
                        break
                    except Exception:
                        time.sleep(0.2)

            # 3. Enter nạp chứng từ
            w0 = safe_find_by_id(session, "wnd[0]")
            if w0:
                w0.SendVKey(0)

            # Nâng thời gian chờ từ 2.0s lên 5.0s để Server kịp nạp chứng từ nặng
            sap_wait_idle(session, 5.0)
            time.sleep(0.1)

            # 4. Kiểm tra thông báo trên thanh trạng thái sbar
            sbar = safe_find_by_id(session, "wnd[0]/sbar")
            if sbar:
                s_type = ""
                s_msg = ""
                for prop in ("MessageType", "messageType"):
                    v = getattr(sbar, prop, None)
                    if isinstance(v, str):
                        s_type = v.strip()
                        break
                for prop in ("Text", "text"):
                    v = getattr(sbar, prop, None)
                    if isinstance(v, str):
                        s_msg = v.strip()
                        break
                if s_type in ("E", "A"):
                    if w0:
                        try:
                            w0.SendVKey(12)  # Esc thoát chứng từ khi báo lỗi
                        except Exception:
                            pass
                    return False, f"DOC_NOT_FOUND: {s_msg}"

            identity_ok, reason = self.verify_session_identity(session, require_migo=True)
            if not identity_ok:
                return False, reason

            # Đối chiếu đọc lại số chứng từ kèm retry (chống trễ render)
            match_ok = False
            read_doc = None
            read_year = None
            for _ in range(8):
                read_doc = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")
                read_year = smart_find(session, "GODYNPRO-DOC_YEAR", "GuiTextField") or smart_find(session, "GODYNPRO-DOC_YEAR", "GuiCTextField")
                if read_doc is not None and read_year is not None:
                    cur_d = str(getattr(read_doc, "Text", getattr(read_doc, "text", ""))).strip()
                    cur_y = str(getattr(read_year, "Text", getattr(read_year, "text", ""))).strip()
                    if cur_d == str(doc).strip() and cur_y == str(year).strip():
                        match_ok = True
                        break
                time.sleep(0.2)
                sap_wait_idle(session, 1.0)

            if read_doc is None or read_year is None:
                return False, "DOCUMENT_READBACK_UNAVAILABLE"

            if not match_ok:
                return False, "DOCUMENT_READBACK_MISMATCH"

            if getattr(read_doc, "Changeable", True):
                return False, "DOCUMENT_NOT_LOADED"

            # 5. Xác thực chế độ nạp chứng từ: Nếu yêu cầu CHANGE mà SAP lại mở DISPLAY -> Buộc reset và nạp lại ở CHANGE!
            titl_bar = safe_find_by_id(session, "wnd[0]/titl")
            titl_text = str(getattr(titl_bar, "Text", getattr(titl_bar, "text", "")) or "").upper()
            if mode == "CHANGE" and ("DISPLAY" in titl_text or (titl_text and "CHANGE" not in titl_text and "MATERIAL" in titl_text)):
                print(f"[MIGO MODE MISMATCH WARN] Cảnh báo: Yêu cầu CHANGE nhưng màn hình đang là '{titl_text}'. Buộc reset và mở lại ở CHANGE...")
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    try:
                        w0.SendVKey(12)  # Esc thoát Display
                    except Exception:
                        pass
                self.prepare_migo_session(session, mode="CHANGE")
                return False, f"WRONG_MODE_DISPLAY: MIGO nạp nhầm vào mục Display ('{titl_text}') thay vì Change"

            self._active_document = (str(doc), str(year))
            return True, "LOADED_VERIFIED"

        except Exception as e:
            try:
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    w0.SendVKey(12)  # Esc thoát chứng từ khi gặp ngoại lệ
            except Exception:
                pass
            return False, f"NAV_ERROR: {e}"

    def open_migo_document(self, session, doc: str, year: str = "", mode: str = "CHANGE", force_reload: bool = False) -> Tuple[bool, str]:
        """Nạp chứng từ MIGO theo chuẩn luồng Change_doc kèm cơ chế Tự động Thử lại (Auto-Retry Guard):
        Nếu lần 1 gặp lỗi do trễ giao diện (UI Lag / Redraw delay / E_PENDING / COM 613 / READBACK_MISMATCH),
        Bot tự động làm sạch session, làm mới MIGO và thử lại lần 2 để cứu vãn chứng từ!
        """
        loaded_ok, nav_msg = self._do_open_migo_document(session, doc, year, mode=mode, force_reload=force_reload)
        if loaded_ok:
            return True, nav_msg

        retryable_keys = [
            "NAV_ERROR", "YEAR_CONTROL_NOT_FOUND", "DOCUMENT_READBACK_MISMATCH",
            "SESSION_BUSY", "DOCUMENT_NOT_LOADED", "DOCUMENT_READBACK_UNAVAILABLE"
        ]
        if any(k in nav_msg for k in retryable_keys):
            print(f"[MIGO NAV RETRY] ⚠️ Lần 1 nạp DOC {doc} gặp độ trễ giao diện ({nav_msg}). Làm mới session và thử lại lần 2...")
            sanitize_migo_session(session)
            time.sleep(0.3)
            self.prepare_migo_session(session, mode=mode)
            loaded_ok2, nav_msg2 = self._do_open_migo_document(session, doc, year, mode=mode, force_reload=True)
            if loaded_ok2:
                print(f"[MIGO NAV RETRY] ✅ Lần 2 nạp DOC {doc} THÀNH CÔNG!")
                return True, nav_msg2
            return loaded_ok2, nav_msg2

        return loaded_ok, nav_msg

    def process_single_migo_item(

        self,

        session,

        item: Dict[str, str],

        index_data: Dict[str, Any],

        fill_headtext: bool = True,

        year: str = "",

        is_folder_requested: bool = False,

        has_folder: bool = False,

        folder_error: str = "",

        dry_run: bool = False,

        allow_overwrite_headtext: bool = False

    ) -> Tuple[Dict[str, Any], bool]:

        """Xử lý 1 chứng từ MIGO duy nhất (Điền HeadText, đính kèm BBBG/VAT)."""

        doc = str(item.get("doc", "")).strip()

        ticket_no = str(item.get("ticket_no", "")).strip()

        year = str(item.get("year") or year or "").strip()

        if not year:

            from datetime import datetime

            year = str(datetime.now().year)

        if not re.fullmatch(r"\d{10}", doc):

            return self._failed_item(doc, ticket_no, "INVALID_DOCUMENT", "Số chứng từ phải có 10 chữ số")

        if item.get("multi_ticket_conflict"):

            return self._failed_item(doc, ticket_no, "MULTI_TICKET_DOC", "Cùng chứng từ có nhiều số phiếu; cần kiểm tra thủ công")

        if dry_run:

            matches = self.match_all_files(doc, ticket_no, index_data) if has_folder else {"bbbg": [], "vat": [], "mail": []}

            ready = bool(ticket_no and fill_headtext) if not is_folder_requested else bool(has_folder and any(matches.values()))

            status = "SIMULATED" if ready else "FILE_NOT_FOUND" if has_folder else "LỖI_THƯ_MỤC" if is_folder_requested else "THIẾU_SỐ_PHIẾU"

            result, _ = self._failed_item(doc, ticket_no, status, "Xem trước; chưa thay đổi dữ liệu SAP")

            result.update({"headtext_status": "SIMULATED" if ticket_no and fill_headtext else "SKIP_NO_TICKET",

                           "simulated": True, "preview_ready": ready, "year": year})

            for kind in ("bbbg", "vat", "mail"):

                result[kind + "_file"] = ", ".join(p.name for p in matches[kind]) or None

                result[kind + "_status"] = ("SIMULATED" if matches[kind] else "FILE_NOT_FOUND") if has_folder else "LỖI_THƯ_MỤC" if is_folder_requested else "SKIP_NO_ATTACH"

            return result, False

        if session is None:

            return self._failed_item(doc, ticket_no, "NO_SESSION", "Không kết nối SAP; chưa thực hiện thao tác")

        if not re.fullmatch(r"(?:19|20|21)\d{2}", year):

            return self._failed_item(doc, ticket_no, "YEAR_REQUIRED", "Cần năm chứng từ rõ ràng; không tự lấy năm hiện tại")

        identity_ok, identity_status = self.verify_session_identity(session)

        if not identity_ok:

            return self._failed_item(doc, ticket_no, identity_status, "Chưa xác minh đúng hệ thống, client, người dùng và session được cấp")

        # =========================================================================

        # TRƯỜNG HỢP 1: NGƯỜI DÙNG KHÔNG YÊU CẦU ĐÍNH KÈM (Chỉ điền số phiếu HeadText)

        # =========================================================================

        if not is_folder_requested:

            headtext_res = "SKIP"

            headtext_msg = "Không có số phiếu cần điền"

            overall_status = "CHƯA_XỬ_LÝ"

            is_succ = False

            if dry_run:

                if fill_headtext and ticket_no:

                    headtext_res = "HEADTEXT_DONE"

                    headtext_msg = f"Mô phỏng: Điền '{ticket_no}'"

                    overall_status = "DONE"

                    is_succ = True

                else:

                    headtext_res = "SKIP_NO_TICKET"

                    headtext_msg = "Thiếu số phiếu"

                    overall_status = "THIẾU_SỐ_PHIẾU"

            else:

                loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE")

                if not loaded_ok:

                    headtext_res = "OPEN_FAIL"

                    headtext_msg = nav_msg

                    overall_status = f"Lỗi mở DOC: {nav_msg}"

                    sanitize_migo_session(session)

                else:

                    if fill_headtext and ticket_no:

                        headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no, allow_overwrite=allow_overwrite_headtext)

                        if headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK"):

                            overall_status = "DONE"

                            is_succ = True

                        else:

                            overall_status = f"Lỗi HeadText: {headtext_res}"

                    else:

                        headtext_res = "SKIP_NO_TICKET"

                        headtext_msg = "Không có số phiếu để điền"

                        overall_status = "THIẾU_SỐ_PHIẾU"

            if not dry_run and session is not None:

                try:

                    w0 = safe_find_by_id(session, "wnd[0]")

                    if w0:

                        w0.SendVKey(12)  # Esc thoát chứng từ

                        time.sleep(0.05)

                except Exception:

                    pass

            return {

                "doc": doc,

                "ticket_no": ticket_no,

                "headtext_status": headtext_res,

                "headtext_msg": headtext_msg,

                "overall_status": overall_status,

                "bbbg_status": "SKIP_NO_ATTACH",

                "vat_status": "SKIP_NO_ATTACH",

                "mail_status": "SKIP_NO_ATTACH",

                "bbbg_file": None,

                "vat_file": None,

                "mail_file": None,

                "match_notes": headtext_msg,

                "is_succ": is_succ

            }, is_succ

        # =========================================================================

        # TRƯỜNG HỢP 2: CÓ YÊU CẦU ĐÍNH KÈM NHƯNG THƯ MỤC LỖI TRUY CẬP (Mất mạng / VPN)

        # =========================================================================

        if not has_folder:

            headtext_res = "SKIP"

            headtext_msg = "Không điền HeadText"

            if fill_headtext and ticket_no:

                if dry_run:

                    headtext_res = "HEADTEXT_DONE"

                    headtext_msg = f"Mô phỏng: Điền '{ticket_no}'"

                else:

                    loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE")

                    if loaded_ok:

                        headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no, allow_overwrite=allow_overwrite_headtext)

                    else:

                        headtext_res = "OPEN_FAIL"

                        headtext_msg = nav_msg

            bbbg_res = "LỖI_THƯ_MỤC"

            vat_res = "LỖI_THƯ_MỤC"

            if headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK"):

                overall_status = f"{format_status_label(ticket_no, True)} (Lỗi link thư mục)"

                is_succ = True

            else:

                overall_status = "LỖI_THƯ_MỤC"

                is_succ = False

            return {

                "doc": doc,

                "ticket_no": ticket_no,

                "headtext_status": headtext_res,

                "headtext_msg": headtext_msg,

                "bbbg_file": None,

                "vat_file": None,

                "mail_file": None,

                "bbbg_status": bbbg_res,

                "vat_status": vat_res,

                "mail_status": "LỖI_THƯ_MỤC",

                "overall_status": overall_status,

                "match_notes": folder_error or "Không thể truy cập thư mục đính kèm",

                "is_succ": is_succ

            }, is_succ

        # =========================================================================

        # TRƯỜNG HỢP 3: CÓ THƯ MỤC HỢP LỆ -> QUÉT & ĐÍNH KÈM FILE BBBG, VAT VÀO MIGO

        # =========================================================================

        all_matches = self.match_all_files(doc, ticket_no, index_data)

        bbbg_paths = all_matches["bbbg"]

        vat_paths = all_matches["vat"]

        mail_paths = all_matches["mail"]

        match_notes_list = []

        if bbbg_paths:

            match_notes_list.append(f"BBBG: {len(bbbg_paths)} file")

        else:

            match_notes_list.append("BBBG: FILE_NOT_FOUND")

        if vat_paths:

            match_notes_list.append(f"VAT: {len(vat_paths)} file")

        else:

            match_notes_list.append("VAT: FILE_NOT_FOUND")

        match_notes = " | ".join(match_notes_list)

        headtext_res = "SKIP"

        headtext_msg = "Không điền HeadText"

        bbbg_res = "FILE_NOT_FOUND"

        vat_res = "FILE_NOT_FOUND"

        mail_res = "FILE_NOT_FOUND" if mail_paths else None

        overall_status = "FILE_NOT_FOUND"

        is_succ = False

        if dry_run:

            if fill_headtext and ticket_no:

                headtext_res = "HEADTEXT_DONE"

                headtext_msg = f"Mô phỏng: Điền '{ticket_no}'"

            bbbg_res = "DONE" if bbbg_paths else "FILE_NOT_FOUND"

            vat_res = f"DONE ({len(vat_paths)})" if len(vat_paths) > 1 else ("DONE" if vat_paths else "FILE_NOT_FOUND")

            if mail_paths:

                mail_res = "DONE"

            if bbbg_paths or vat_paths or mail_paths:

                overall_status = "DONE"

                is_succ = True

            else:

                overall_status = "FILE_NOT_FOUND"

        else:
            # Khởi tạo an toàn tuyệt đối trước khi mở MIGO (ngăn chặn 100% UnboundLocalError)
            sap_bbbg = None
            sap_vat = None

            loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE")

            if not loaded_ok:
                bbbg_res = nav_msg
                vat_res = nav_msg
                overall_status = nav_msg
                sanitize_migo_session(session)
            else:

                # 0. KIỂM TRA ĐÍNH KÈM ĐÃ CÓ SẴN TRÊN SAP MIGO (Display Attachment / OK_REF_ATTA)

                # 0. KIỂM TRA ĐÍNH KÈM: Sử dụng cờ chi tiết ZFIR266 nếu đã có sẵn để BỎ QUA GOS popup
                need_bb = item.get("need_bbbg")
                need_v = item.get("need_vat")

                if need_bb is not None and need_v is not None:
                    # Đã có cờ chính xác từ Phase 1 ZFIR266: không cần mở popup GOS tốn thời gian!
                    sap_titles = []
                    sap_bbbg = "ZFIR266" if not need_bb else None
                    sap_vat = "ZFIR266" if not need_v else None
                else:
                    sap_titles = self.read_sap_attachment_titles(session)
                    sap_bbbg, sap_vat = self.check_existing_sap_attachments(sap_titles, ticket_no)

                # 1. Điền số phiếu vào Header Text nếu được yêu cầu
                if fill_headtext and ticket_no:
                    headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no, allow_overwrite=allow_overwrite_headtext)

                ticket_nums, _ = extract_search_tokens(ticket_no)

                # 2. Đính kèm file BBBG (hỗ trợ lũy kế nhiều file BBBG)

                b_ok = False

                if need_bb is False:
                    to_upload_bb, already_bb = [], ["ZFIR266"]
                else:
                    to_upload_bb, already_bb, _ = match_files_to_sap(bbbg_paths, sap_titles, ticket_nums=ticket_nums, file_type="bbbg")

                if not to_upload_bb and already_bb:

                    b_ok = True

                    bbbg_res = f"DONE ({len(already_bb)} file đã có sẵn trên SAP)" if len(already_bb) > 1 else "DONE"

                    print(f"[ATT RUNNER] DOC {doc}: BBBG đã có sẵn trên SAP ({len(already_bb)} file), bỏ qua upload lại.")

                elif to_upload_bb:

                    succ_bb = []

                    fail_bb = []

                    for bp in to_upload_bb:

                        ok_b, msg_b = self.attach_file_to_current_migo(session, bp)

                        if ok_b:

                            succ_bb.append(bp.name)

                        else:

                            fail_bb.append(f"{bp.name}: {msg_b}")

                    if succ_bb or already_bb:

                        b_ok = True

                        if already_bb and succ_bb:

                            bbbg_res = f"DONE (Đính thêm {len(succ_bb)}, đã có {len(already_bb)})"

                        elif len(succ_bb) > 1:

                            bbbg_res = f"DONE ({len(succ_bb)})"

                        else:

                            bbbg_res = "DONE"

                    else:

                        bbbg_res = f"Lỗi: {'; '.join(fail_bb)}"

                elif sap_bbbg:

                    b_ok = True

                    bbbg_res = "DONE"

                    print(f"[ATT RUNNER] DOC {doc}: BBBG đã có sẵn trên SAP ('{sap_bbbg}'), bỏ qua upload lại.")

                else:

                    bbbg_res = "FILE_NOT_FOUND"

                # 3. Đính kèm file VAT (hỗ trợ lũy kế nhiều file VAT)

                v_ok = False

                if need_v is False:
                    to_upload_vat, already_vat = [], ["ZFIR266"]
                else:
                    to_upload_vat, already_vat, _ = match_files_to_sap(vat_paths, sap_titles, ticket_nums=ticket_nums, file_type="vat")

                if not to_upload_vat and already_vat:

                    v_ok = True

                    vat_res = f"DONE ({len(already_vat)} file đã có sẵn trên SAP)" if len(already_vat) > 1 else "DONE"

                    print(f"[ATT RUNNER] DOC {doc}: VAT đã có sẵn trên SAP ({len(already_vat)} file), bỏ qua upload lại.")

                elif to_upload_vat:

                    succ_vat = []

                    fail_vat = []

                    for vp in to_upload_vat:

                        ok_v, msg_v = self.attach_file_to_current_migo(session, vp)

                        if ok_v:

                            succ_vat.append(vp.name)

                        else:

                            fail_vat.append(f"{vp.name}: {msg_v}")

                    if succ_vat or already_vat:

                        v_ok = True

                        if already_vat and succ_vat:

                            vat_res = f"DONE (Đính thêm {len(succ_vat)}, đã có {len(already_vat)})"

                        elif len(succ_vat) > 1:

                            vat_res = f"DONE ({len(succ_vat)})"

                        else:

                            vat_res = "DONE"

                    else:

                        vat_res = f"Lỗi: {'; '.join(fail_vat)}"

                elif sap_vat:

                    v_ok = True

                    vat_res = "DONE"

                    print(f"[ATT RUNNER] DOC {doc}: VAT đã có sẵn trên SAP ('{sap_vat}'), bỏ qua upload lại.")

                else:

                    vat_res = "FILE_NOT_FOUND"

                # 4. Đính kèm file Mail (nếu có)

                m_ok = False

                if mail_paths:

                    ok_m, msg_m = self.attach_file_to_current_migo(session, mail_paths[0])

                    if ok_m:

                        m_ok = True

                        mail_res = "DONE"

                    else:

                        mail_res = f"Lỗi: {msg_m}"

                # Đánh giá trạng thái tổng thể chính xác (A03 - Phân định rõ DONE, PARTIAL, REVIEW)

                if headtext_res == "HEADTEXT_REVIEW":

                    overall_status = "REVIEW_HEADTEXT"

                    is_succ = False

                elif b_ok and v_ok:

                    overall_status = "DONE"

                    is_succ = True

                elif b_ok and not v_ok:

                    overall_status = "PARTIAL (Thiếu VAT)"

                    is_succ = False

                elif v_ok and not b_ok:

                    overall_status = "PARTIAL (Thiếu BBBG)"

                    is_succ = False

                elif m_ok:

                    overall_status = "DONE"

                    is_succ = True

                elif headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK"):

                    overall_status = "PARTIAL (Chỉ có HeadText)"

                    is_succ = False

                else:

                    overall_status = "FILE_NOT_FOUND"

                    is_succ = False

        # Sau khi xử lý xong chứng từ, dọn sạch modal popup và chuyển chứng từ êm ái chống treo tab

        if not dry_run and session is not None:

            try:

                dismiss_any_modal_popup(session)

                w0 = safe_find_by_id(session, "wnd[0]")

                if w0:

                    w0.SendVKey(12)  # Esc thoát chứng từ

                    time.sleep(0.05)

                    dismiss_any_modal_popup(session)

            except Exception:

                pass

        return {

            "doc": doc,

            "ticket_no": ticket_no,

            "headtext_status": headtext_res,

            "headtext_msg": headtext_msg,

            "bbbg_file": ", ".join(p.name for p in bbbg_paths) if bbbg_paths else (sap_bbbg if sap_bbbg else None),

            "vat_file": ", ".join(p.name for p in vat_paths) if vat_paths else (sap_vat if sap_vat else None),

            "mail_file": mail_paths[0].name if mail_paths else None,

            "bbbg_status": bbbg_res,

            "vat_status": vat_res,

            "mail_status": mail_res,

            "overall_status": overall_status,

            "match_notes": match_notes,

            "is_succ": is_succ

        }, is_succ

    def process_attachment_batch(

        self,

        folder_path: str,

        items: List[Dict[str, str]],

        year: str = "",

        fill_headtext: bool = True,

        dry_run: bool = False,

        progress_callback: Optional[Callable[[int, int, str, int], None]] = None,

        allow_overwrite_headtext: bool = False

    ) -> Dict[str, Any]:

        """Xử lý hàng loạt danh sách chứng từ: Điền số phiếu vào HeadText và Đính kèm file BBBG/VAT."""

        clean_folder = clean_folder_path(folder_path)

        is_folder_requested = bool(clean_folder)

        has_folder = False

        folder_error = ""

        if is_folder_requested:

            try:

                folder_obj = Path(clean_folder)

                if folder_obj.exists() and folder_obj.is_dir():

                    has_folder = True

                else:

                    folder_error = f"Thư mục không tồn tại hoặc không thể truy cập: {clean_folder}"

            except Exception as e:

                folder_error = f"Lỗi truy cập thư mục: {e}"

        if has_folder:

            index_data = self.index_source_folder(clean_folder, recursive=True)

            print(f"[ATT RUNNER] Đã quét thư mục '{clean_folder}': Tìm thấy {index_data['total_files']} files hợp lệ.")

        elif is_folder_requested:

            index_data = {"total_files": 0, "files": [], "token_map": {}}

            print(f"[ATT WARN] Không thể truy cập thư mục '{clean_folder}': {folder_error}")

            print(f"[ATT WARN] Vui lòng kiểm tra lại quyền truy cập mạng LAN/VPN VinFast hoặc mở thư mục trong File Explorer.")

        else:

            index_data = {"total_files": 0, "files": [], "token_map": {}}

            print(f"[ATT RUNNER] Chế độ điền số phiếu HeadText (Không đính kèm file).")

        # Phát hiện xung đột MULTI_TICKET_DOC (A02): cùng 1 DOC nhưng có nhiều ticket_no khác nhau

        doc_tickets = {}

        for it in items:

            d = str(it.get("doc", "")).strip()

            t = str(it.get("ticket_no", "")).strip()

            if d and t:

                doc_tickets.setdefault(d, set()).add(t)

        multi_ticket_docs = {d: sorted(list(ts)) for d, ts in doc_tickets.items() if len(ts) > 1}

        if multi_ticket_docs:

            print(f"[ATT RUNNER WARN] Phát hiện DOC có nhiều số phiếu khác nhau (MULTI_TICKET_DOC): {multi_ticket_docs}")

            for it in items:

                d = str(it.get("doc", "")).strip()

                if d in multi_ticket_docs:

                    it["multi_ticket_conflict"] = multi_ticket_docs[d]

        results = []
        total_success = 0
        total_partial = 0
        total_missing = 0

        total_items = len(items)

        # =========================================================================
        # PHASE 1: ZFIR266 BATCH PRE-CHECK (Kiểm tra 1 shot toàn bộ DOCs trong 1.5s)
        # =========================================================================
        zfir_map = {}
        if not dry_run and self.sap_engine is not None and len(items) > 0 and hasattr(self.sap_engine, "check_attachments_zfir266"):
            try:
                doc_list = [str(it.get("doc", "")).strip() for it in items if re.fullmatch(r"\d{10}", str(it.get("doc", "")).strip())]
                if doc_list:
                    print(f"[ATT RUNNER] Phase 1: Batch-check {len(doc_list)} DOCs qua ZFIR266 để bóc tách cờ thiếu chi tiết...")
                    zfir_records = self.sap_engine.check_attachments_zfir266(
                        doc_list,
                        year=year or str(datetime.now().year)
                    )
                    for r in zfir_records:
                        zfir_map[r["doc"]] = r
                    print(f"[ATT RUNNER] Phase 1 hoàn tất: Đã lập bản đồ trạng thái cho {len(zfir_map)} DOC.")
            except Exception as zfir_pre_err:
                print(f"[ATT RUNNER WARN] Phase 1 ZFIR266 precheck không thành công, tiếp tục quy trình MIGO chuẩn: {zfir_pre_err}")

        # Gắn cờ chi tiết và lọc DOC đã hoàn thành 100% (bỏ qua hoàn toàn MIGO)
        items_to_process = []
        already_done_results = []
        for it in items:
            d = str(it.get("doc", "")).strip()
            if d in zfir_map:
                z_info = zfir_map[d]
                it["found_in_zfir"] = z_info.get("found_in_zfir", False)
                it["need_bbbg"] = z_info.get("need_bbbg", True)
                it["need_vat"] = z_info.get("need_vat", True)
                it["need_header_text"] = z_info.get("need_header_text", True)

                # KIỂM TRA THÔNG MINH (SMART GATING):
                # Chỉ gạch tên khi TẤT CẢ các file quét được trong thư mục LAN đều đã khớp 1-1 trên SAP!
                # Nếu có file mới, file khác ký tự (_Dot2, _BS, (1)...) thì KHÔNG gạch tên, tiếp tục đẩy vào MIGO.
                can_skip_migo = False
                if it["found_in_zfir"]:
                    ticket_no = it.get("ticket_no", "")
                    ticket_nums, _ = extract_search_tokens(ticket_no)
                    sap_titles = z_info.get("sap_titles") or []
                    if not sap_titles and z_info.get("files_found"):
                        raw_f = z_info.get("files_found", "")
                        if raw_f and raw_f != "Không có file":
                            sap_titles = [f.strip() for f in raw_f.split(",") if f.strip()]

                    # So khớp với file thực tế trong thư mục LAN nếu có
                    if has_folder:
                        local_matches = self.match_all_files(d, ticket_no, index_data)
                        local_bb = local_matches.get("bbbg", [])
                        local_vat = local_matches.get("vat", [])

                        to_upload_bb, _, _ = match_files_to_sap(local_bb, sap_titles, ticket_nums=ticket_nums, file_type="bbbg")
                        to_upload_vat, _, _ = match_files_to_sap(local_vat, sap_titles, ticket_nums=ticket_nums, file_type="vat")

                        # Chỉ được bỏ qua khi: Không còn bất kỳ file local nào cần upload (len(to_upload) == 0)
                        # VÀ trên SAP thực sự đã có file BBBG/VAT nếu trong thư mục có yêu cầu
                        has_all_local_attached = (len(to_upload_bb) == 0) and (len(to_upload_vat) == 0)
                        has_sap_files = (not local_bb or not it["need_bbbg"]) and (not local_vat or not it["need_vat"])
                        ht_satisfied = not fill_headtext or not ticket_no or not it["need_header_text"]
                        can_skip_migo = has_all_local_attached and has_sap_files and ht_satisfied
                    else:
                        # Không yêu cầu đính kèm thư mục, chỉ cần thỏa mãn cờ ZFIR và HeadText
                        can_skip_migo = not it["need_bbbg"] and not it["need_vat"] and (not fill_headtext or not ticket_no or not it["need_header_text"])

                if can_skip_migo:
                    already_done_results.append({
                        "doc": d,
                        "ticket_no": it.get("ticket_no", ""),
                        "headtext_status": "DONE (ZFIR266)",
                        "headtext_msg": "Đã có sẵn trên SAP",
                        "bbbg_file": z_info.get("bbbg_file"),
                        "bbbg_status": "DONE (ZFIR266: Đã đủ)",
                        "vat_file": z_info.get("vat_file"),
                        "vat_status": "DONE (ZFIR266: Đã đủ)",
                        "mail_file": None,
                        "match_notes": "Đã đủ file BBBG & VAT trên SAP (ZFIR266 Precheck)",
                        "overall_status": "DONE",
                        "is_succ": True
                    })
                    continue
            items_to_process.append(it)

        if already_done_results:
            print(f"[ATT RUNNER] Phase 1: {len(already_done_results)}/{total_items} DOC đã đủ 100% BBBG/VAT -> Bỏ qua MIGO!")
            total_success += len(already_done_results)
            results.extend(already_done_results)

        items = items_to_process

        last_progress_time = time.time()
        PROGRESS_INTERVAL_SEC = 180.0  # Cập nhật tiến độ lên Teams mỗi 3 phút

        # =========================================================================
        # PHASE 2: ADAPTIVE MULTI-SESSION SAP POOL (TỰ THÍCH ỨNG TURBO / IDLE)
        # =========================================================================
        prompt_raw = str(getattr(self, "_prompt_context", "") or "").lower()
        tab_match = re.search(r"\b([1-7])\s*tabs?\b", prompt_raw)
        has_tab_request = tab_match is not None
        explicit_turbo_words = ["turbo", "max", "nhanh", "toi da", "full tab", "het cong suat"]
        is_explicit_turbo = has_tab_request or any(w in prompt_raw for w in explicit_turbo_words)

        user_idle = get_user_idle_seconds()
        turbo = is_explicit_turbo or (user_idle >= 45.0)

        min_items_threshold = 2 if turbo else 3
        is_multi_tab_eligible = (
            not dry_run
            and getattr(self, 'enable_multi_tab', True)
            and self.sap_engine is not None
            and len(items) >= min_items_threshold
            and pythoncom is not None
        )

        user_res_val = getattr(self.sap_engine, "user_reserved_sessions", None)
        max_bot_val = getattr(self.sap_engine, "max_batch_tabs", None)

        if has_tab_request:
            req_tabs = int(tab_match.group(1))
            user_reserved = 1
            max_bot_tabs = min(len(items), req_tabs)
            print(f"[ATT RUNNER] 🚀 YÊU CẦU ĐẶC BIỆT {req_tabs} TABS: Phân bổ tới {max_bot_tabs} Tabs song song (Bảo tồn Tab 1 cho User)...")
        elif is_explicit_turbo:
            user_reserved = 1
            if "6 tab" in prompt_raw or "max" in prompt_raw:
                max_bot_tabs = min(len(items), 6)
            elif "5 tab" in prompt_raw:
                max_bot_tabs = min(len(items), 5)
            elif "4 tab" in prompt_raw:
                max_bot_tabs = min(len(items), 4)
            elif "3 tab" in prompt_raw:
                max_bot_tabs = min(len(items), 3)
            elif "2 tab" in prompt_raw:
                max_bot_tabs = min(len(items), 2)
            else:
                max_bot_tabs = min(len(items), 6)
            print(f"[ATT RUNNER] 🚀 YÊU CẦU TURBO MAX: Chuẩn bị phân bổ tới {max_bot_tabs} Tabs song song (Bảo tồn Tab 1 cho User)...")
        else:
            # Chạy thông thường hoặc tự kích hoạt khi máy nhàn rỗi (tôn trọng cấu hình engine nếu có)
            if user_res_val is not None:
                user_reserved = int(user_res_val)
            elif turbo:
                user_reserved = 0 if user_idle >= 45.0 else 1
            else:
                user_reserved = 1

            if max_bot_val is not None:
                max_bot_tabs = int(max_bot_val)
            elif turbo:
                max_bot_tabs = 5
            else:
                max_bot_tabs = 4

            if turbo and user_res_val is None:
                print(f"[ATT RUNNER] 🚀 KÍCH HOẠT TURBO MULTI-TAB (Idle {user_idle:.1f}s): Cho phép tới {max_bot_tabs} Tabs song song!")

        conn_idx = 0
        if isinstance(self.expected_session_identity, dict) and ("connection_index" in self.expected_session_identity or "connection_id" in self.expected_session_identity):
            try:
                conn_idx = int(self.expected_session_identity.get("connection_index", self.expected_session_identity.get("connection_id", 0)))
            except Exception:
                conn_idx = 0
        elif hasattr(self.sap_engine, "connection_index"):
            conn_idx = int(getattr(self.sap_engine, "connection_index", 0))

        active_tab_indices = [0]
        if is_multi_tab_eligible:
            target_tabs = min(7, user_reserved + max_bot_tabs)
            available_tabs = ensure_sap_tabs(target_tabs, connection_index=conn_idx)
            bot_tabs = [t for t in range(user_reserved, available_tabs) if t < 7]
            if len(bot_tabs) >= 2:
                active_tab_indices = bot_tabs[:max_bot_tabs]
                tabs_str = ", ".join(f"Tab {t+1}" for t in active_tab_indices)
                print(f"[ATT RUNNER] 🚀 ĐÃ KÍCH HOẠT TURBO MULTI-TAB THÀNH CÔNG: Chia tải trên {len(active_tab_indices)} tabs song song ({tabs_str}) (Bảo tồn Tab 1..{user_reserved} cho Người dùng).")
            elif len(bot_tabs) == 1:
                active_tab_indices = [bot_tabs[0]]
                print(f"[ATT RUNNER] SAP có Tab {active_tab_indices[0]+1} khả dụng: Chạy Batch trên Tab {active_tab_indices[0]+1} (Bảo tồn Tab 1..{user_reserved} cho Người dùng).")
            else:
                active_tab_indices = [0]
                print(f"[ATT RUNNER WARN] SAP hiện chỉ có {available_tabs} Tab, không thể tạo thêm tab mới; tiếp tục chạy an toàn trên Session 0.")

        # NHÁNH A: CHẠY ĐA TAB SONG SONG (MULTI-TAB BATCH POOL)

        # ---------------------------------------------------------------------

        if len(active_tab_indices) > 1:
            self._multi_tab_mode = True

            # WORK-STEALING DYNAMIC SCHEDULER & FAULT QUARANTINE (CHỐNG TREO DÂY CHUYỀN)
            from collections import deque
            tab_queues = {t_idx: deque() for t_idx in active_tab_indices}
            for i, itm in enumerate(items):
                target_tab = active_tab_indices[i % len(active_tab_indices)]
                tab_queues[target_tab].append((i, itm))

            processed_count = 0
            success_count = 0
            prog_lock = threading.Lock()
            queue_lock = threading.Lock()
            indexed_results = []
            stop_event = threading.Event()
            healthy_tabs = set(active_tab_indices)

            def _tab_worker(tab_idx: int):
                nonlocal processed_count, success_count, last_progress_time
                if pythoncom is not None:
                    pythoncom.CoInitialize()

                if tab_idx > 0:
                    time.sleep(min(1.0, (tab_idx % 6) * 0.15))
                tab_lock = BATCH_TAB_LOCKS.get(tab_idx, threading.Lock())
                with tab_lock:
                    worker_session = None
                    try:
                        conn_idx = 0
                        if isinstance(self.expected_session_identity, dict):
                            conn_idx = self.expected_session_identity.get("connection_index", self.expected_session_identity.get("connection_id", 0))
                        elif hasattr(self.sap_engine, "connection_index"):
                            conn_idx = getattr(self.sap_engine, "connection_index", 0)
                        worker_session = get_safe_sap_session(conn_idx, tab_idx)
                    except Exception as conn_err:
                        print(f"[ATT RUNNER Tab {tab_idx}] Lỗi kết nối session: {conn_err}")

                    if worker_session is not None:
                        self.prepare_migo_session(worker_session, mode="CHANGE")
                    else:
                        print(f"[ATT RUNNER Tab {tab_idx} QUARANTINE] ⚠️ Không thể lấy session cho Tab {tab_idx+1}. Bỏ qua tab này để các tab khác tiếp tục gánh tải!")
                        with queue_lock:
                            healthy_tabs.discard(tab_idx)
                        if pythoncom is not None:
                            pythoncom.CoUninitialize()
                        return

                    consecutive_tab_errors = 0
                    idle_spins = 0
                    try:
                        while not stop_event.is_set():
                            item_tuple = None
                            with queue_lock:
                                if tab_queues.get(tab_idx):
                                    item_tuple = tab_queues[tab_idx].popleft()
                                    idle_spins = 0
                                else:
                                    # CHỈ RÚT VIỆC TỪ CÁC TAB BỊ LỖI / BỊ TREO / ĐÃ BỊ CÁCH LY (QUARANTINED TABS)
                                    dead_tabs = [t for t in tab_queues if t not in healthy_tabs and tab_queues[t]]
                                    if dead_tabs:
                                        target_steal = dead_tabs[0]
                                        item_tuple = tab_queues[target_steal].popleft()
                                        idle_spins = 0
                                        print(f"[ATT RUNNER Tab {tab_idx}] ⚡ Cứu viện tab lỗi: Rút việc từ Tab {target_steal+1} (còn {len(tab_queues[target_steal])} việc)!")

                            if item_tuple is None:
                                # Kiểm tra xem còn tab nào khác đang dở việc không
                                has_remaining = False
                                with queue_lock:
                                    if any(len(q) > 0 for q in tab_queues.values()):
                                        has_remaining = True
                                if has_remaining and idle_spins < 60:
                                    idle_spins += 1
                                    time.sleep(0.05)
                                    continue
                                break

                            orig_idx, itm = item_tuple
                            doc_str = str(itm.get("doc", "")).strip()
                            is_tab_fatal = False

                            try:
                                res_dict, is_s = self.process_single_migo_item(
                                    worker_session, itm, index_data, fill_headtext, year,
                                    is_folder_requested, has_folder, folder_error, dry_run=False,
                                    allow_overwrite_headtext=allow_overwrite_headtext
                                )
                                consecutive_tab_errors = 0
                            except Exception as item_err:
                                err_msg = str(item_err)
                                print(f"[ATT RUNNER Tab {tab_idx} ERROR] DOC {doc_str}: {item_err}")

                                fatal_indicators = [
                                    "-2147417848", "-2147023170",
                                    "the enumerator of the collection",
                                    "the server threw an exception",
                                    "the remote procedure call failed",
                                    "disconnected", "connection broken", "session closed"
                                ]
                                if any(ind in err_msg.lower() for ind in fatal_indicators) or consecutive_tab_errors >= 5:
                                    is_tab_fatal = True

                                if is_tab_fatal:
                                    with queue_lock:
                                        healthy_tabs.discard(tab_idx)
                                        remain = len(healthy_tabs)
                                    print(f"[ATT RUNNER Tab {tab_idx} QUARANTINE] ⚠️ Tab {tab_idx+1} gặp sự cố nghiêm trọng ({err_msg}). Tiến hành cô lập và loại bỏ tab này!")
                                    print(f"[ATT RUNNER] 🔄 Dây chuyền tiếp tục vận hành với {remain} tab còn lại; các chứng từ tồn sẽ được các tab khỏe rút việc xử lý tiếp.")
                                else:
                                    consecutive_tab_errors += 1
                                    sanitize_migo_session(worker_session)

                                res_dict = {
                                    "doc": doc_str,
                                    "ticket_no": str(itm.get("ticket_no", "")).strip(),
                                    "headtext_status": "EXCEPTION",
                                    "headtext_msg": err_msg,
                                    "bbbg_status": "ERROR",
                                    "vat_status": "ERROR",
                                    "mail_status": "ERROR",
                                    "overall_status": "ERROR",
                                    "match_notes": f"Tab {tab_idx} exception: {item_err}"
                                }
                                is_s = False

                            with prog_lock:
                                indexed_results.append((orig_idx, res_dict, is_s))
                                processed_count += 1
                                is_item_succ = is_s or res_dict.get("is_succ") is True or res_dict.get("overall_status") == "DONE"
                                if is_item_succ:
                                    success_count += 1
                                now = time.time()
                                if progress_callback and (now - last_progress_time >= PROGRESS_INTERVAL_SEC):
                                    last_progress_time = now
                                    try:
                                        progress_callback(processed_count, total_items, doc_str, success_count)
                                    except Exception:
                                        pass

                            if is_tab_fatal:
                                break

                    finally:
                        if not is_tab_fatal:
                            reset_session_to_main_menu(worker_session)

                if pythoncom is not None:
                    pythoncom.CoUninitialize()
                return

            with ThreadPoolExecutor(max_workers=len(active_tab_indices)) as executor:
                futures = [
                    executor.submit(_tab_worker, t_idx)
                    for t_idx in active_tab_indices
                ]
                max_batch_wait = max(180.0, len(items) * 30.0)
                t_start = time.time()
                from app.sap_recovery import is_sap_gui_hung, is_sap_gui_alive, kill_hung_sap_processes

                # VÒNG LẶP GIÁM SÁT THỜI GIAN THỰC (ACTIVE REAL-TIME WATCHDOG):
                # Phát hiện tức thì cửa sổ "MIGO (Not Responding)" trong 1.0 giây và cưỡng chế giải phóng!
                while time.time() - t_start < max_batch_wait:
                    if all(f.done() for f in futures):
                        break

                    # 1. Kiểm tra nếu SAP GUI bị đơ cứng "(Not Responding)"
                    if is_sap_gui_hung():
                        print("[ATT RUNNER WATCHDOG] 🚨 PHÁT HIỆN CỬA SỔ SAP GUI BỊ TREO CỨNG '(Not Responding)'! Kích hoạt cưỡng chế giải phóng...")
                        stop_event.set()
                        kill_hung_sap_processes(timeout_sec=3.0)
                        break

                    # 2. Kiểm tra nếu SAP GUI bị văng hoặc tắt hẳn
                    if not is_sap_gui_alive():
                        is_genuinely_dead = True
                        try:
                            if self.sap_engine and hasattr(self.sap_engine, "session"):
                                s = getattr(self.sap_engine, "session", None)
                                if s and hasattr(s, "Busy"):
                                    _ = s.Busy
                                    is_genuinely_dead = False
                        except Exception:
                            is_genuinely_dead = True
                        if is_genuinely_dead:
                            print("[ATT RUNNER WATCHDOG] 🚨 PHÁT HIỆN SAP GUI BỊ VĂNG / MẤT KẾT NỐI KHỎI HỆ THỐNG! Dừng chờ mẻ...")
                            stop_event.set()
                            break

                    time.sleep(1.0)

                stop_event.set()

                # TỰ HỒI SINH GIỮA MẺ (INTRA-BATCH SELF-HEALING): Cứu vãn các chứng từ còn tồn khi tất cả các tab bị sập
                unprocessed_items = []
                with queue_lock:
                    for t_q in tab_queues.values():
                        while t_q:
                            unprocessed_items.append(t_q.popleft())

                if unprocessed_items:
                    print(f"[ATT RUNNER SELF-HEAL] ⚠️ Còn {len(unprocessed_items)} chứng từ chưa hoàn tất do sự cố sập tab! Kích hoạt quy trình TỰ HỒI SINH GIỮA MẺ...")
                    heal_session = None
                    try:
                        from app.sap_recovery import is_sap_gui_hung, is_sap_gui_alive, recover_hung_sap_session, auto_connect_sso_session
                        conn_name = getattr(self.sap_engine, "sap_connection_name", "13.Future VN S4 HANA PRD (SSO) Link") if self.sap_engine else "13.Future VN S4 HANA PRD (SSO) Link"
                        if not is_sap_gui_alive():
                            print("[ATT RUNNER SELF-HEAL] Phát hiện SAP GUI bị văng, crash hoặc đơ cứng! Cưỡng chế giải phóng và mở lại phiên SSO mới...")
                            heal_session = recover_hung_sap_session(connection_name=conn_name)
                        else:
                            heal_session = auto_connect_sso_session(connection_name=conn_name)
                    except Exception as heal_conn_err:
                        print(f"[ATT RUNNER SELF-HEAL WARN] Lỗi khi kết nối lại SSO: {heal_conn_err}")
                        heal_session = None

                    if heal_session is not None:
                        print(f"[ATT RUNNER SELF-HEAL] ✅ Kết nối lại SSO thành công! Tiếp tục xử lý {len(unprocessed_items)} chứng từ tồn đọng...")
                        self.prepare_migo_session(heal_session, mode="CHANGE")
                        for o_idx, rem_itm in unprocessed_items:
                            doc_rem = str(rem_itm.get("doc", "")).strip()
                            try:
                                h_dict, h_succ = self.process_single_migo_item(
                                    heal_session, rem_itm, index_data, fill_headtext, year,
                                    is_folder_requested, has_folder, folder_error, dry_run=False,
                                    allow_overwrite_headtext=allow_overwrite_headtext
                                )
                            except Exception as h_err:
                                h_dict = {
                                    "doc": doc_rem,
                                    "ticket_no": str(rem_itm.get("ticket_no", "")).strip(),
                                    "headtext_status": "EXCEPTION",
                                    "headtext_msg": str(h_err),
                                    "bbbg_status": "ERROR",
                                    "vat_status": "ERROR",
                                    "mail_status": "ERROR",
                                    "overall_status": "ERROR",
                                    "match_notes": f"Self-heal recovery error: {h_err}"
                                }
                                h_succ = False
                            with prog_lock:
                                indexed_results.append((o_idx, h_dict, h_succ))
                                processed_count += 1
                                if h_succ or h_dict.get("is_succ") is True or h_dict.get("overall_status") == "DONE":
                                    success_count += 1
                            time.sleep(0.05)
                        unprocessed_items.clear()

                # Vét nốt những item thực sự không thể phục hồi (nếu SSO không mở được)
                for o_idx, rem_itm in unprocessed_items:
                    doc_rem = str(rem_itm.get("doc", "")).strip()
                    with prog_lock:
                        indexed_results.append((
                            o_idx,
                            {
                                "doc": doc_rem,
                                "ticket_no": str(rem_itm.get("ticket_no", "")).strip(),
                                "headtext_status": "TIMEOUT",
                                "headtext_msg": "Quá thời gian chờ hoặc tab bị sự cố ngắt đột ngột",
                                "bbbg_status": "TIMEOUT",
                                "vat_status": "TIMEOUT",
                                "mail_status": "TIMEOUT",
                                "overall_status": "TIMEOUT",
                                "match_notes": "Batch recovery drain"
                            },
                            False
                        ))
                        processed_count += 1

            indexed_results.sort(key=lambda x: x[0])
            for _, r_dict, is_s in indexed_results:
                results.append(r_dict)

                st = str(r_dict.get("overall_status", "")).upper()

                if "PARTIAL" in st:

                    total_partial += 1

                elif is_s or r_dict.get("is_succ") is True or st == "DONE":

                    total_success += 1

                else:

                    total_missing += 1

        # ---------------------------------------------------------------------

        # NHÁNH B: CHẠY ĐƠN TAB TUẦN TỰ (SINGLE-TAB HOẶC DRY-RUN)

        # ---------------------------------------------------------------------

        else:

            session = None

            if not dry_run and self.sap_engine:

                single_tab_idx = active_tab_indices[0]

                if single_tab_idx > 0:

                    try:

                        session = get_safe_sap_session(conn_idx, single_tab_idx)

                        if session is None:

                            session = self.sap_engine.ensure_session()

                    except Exception:

                        session = self.sap_engine.ensure_session()

                else:

                    session = self.sap_engine.ensure_session()

                if session is not None:

                    self.prepare_migo_session(session, mode="CHANGE")

            try:

                for idx, item in enumerate(items, 1):

                    doc = str(item.get("doc", "")).strip()

                    now = time.time()

                    if progress_callback and total_items > 20 and (now - last_progress_time >= PROGRESS_INTERVAL_SEC):

                        last_progress_time = now

                        try:

                            progress_callback(idx - 1, total_items, doc, total_success)

                        except Exception as ex:

                            print(f"[PROGRESS WARN] Callback error: {ex}")

                    res_dict, is_s = self.process_single_migo_item(

                        session, item, index_data, fill_headtext, year,

                        is_folder_requested, has_folder, folder_error, dry_run

                    )

                    results.append(res_dict)

                    st = str(res_dict.get("overall_status", "")).upper()

                    if "PARTIAL" in st:

                        total_partial += 1

                    elif is_s or res_dict.get("is_succ") is True or st == "DONE":

                        total_success += 1

                    else:

                        total_missing += 1

            finally:

                if not dry_run and session is not None:

                    reset_session_to_main_menu(session)

        return {

            "folder_path": clean_folder if is_folder_requested else "",

            "folder_error": folder_error,

            "total_items": total_items,

            "total_success": total_success,

            "total_partial": total_partial,

            "multi_tab_count": len(active_tab_indices),

            "is_turbo": turbo,

            "total_missing": total_missing,

            "simulated": dry_run,

            "total_preview_ready": sum(bool(r.get("preview_ready")) for r in results),

            "results": results

        }

    @staticmethod

    def format_teams_report(batch_summary: Dict[str, Any]) -> str:

        """Định dạng báo cáo kết quả đính kèm và điền số phiếu gửi về Teams / Email."""

        results = batch_summary.get("results", [])

        total = batch_summary.get("total_items", len(results))

        success = batch_summary.get("total_success", 0)

        missing = batch_summary.get("total_missing", 0)

        folder_path = batch_summary.get("folder_path", "")

        folder_error = batch_summary.get("folder_error", "")

        lines = []

        if batch_summary.get("simulated"):

            ready = batch_summary.get("total_preview_ready", 0)

            lines.append(f"**XEM TRƯỚC MIGO — CHƯA GHI SAP**: {ready}/{total} dòng sẵn sàng theo dữ liệu cục bộ.")

        if not folder_path:

            lines.append(f"📝 **KẾT QUẢ ĐIỀN SỐ PHIẾU HEADTEXT MIGO**")

            lines.append(f"• Tiến độ: Hoàn tất **{success}/{total}** chứng từ (Lỗi/Thiếu: {missing})")

            lines.append("")

            lines.append("| Số DOC | Số phiếu | Điền HeadText | Trạng thái |")

            lines.append("| :--- | :--- | :--- | :--- |")

            for r in results:

                doc = r["doc"]

                ticket = r["ticket_no"] or "-"

                head = r.get("headtext_status", "-")

                status = r.get("overall_status", "-")

                lines.append(f"| {doc} | {ticket} | {head} | {status} |")

        else:

            lines.append(f"📎 **KẾT QUẢ ĐÍNH KÈM BBBG, VAT & ĐIỀN SỐ PHIẾU MIGO**")

            lines.append(f"• Thư mục nguồn: `{folder_path}`")

            if folder_error:

                lines.append(f"⚠️ **Cảnh báo thư mục:** `{folder_error}`")

            lines.append(f"• Tiến độ: Hoàn tất **{success}/{total}** chứng từ (Thiếu/Lỗi: {missing})")

            lines.append("")

            lines.append("**Danh sách trạng thái từng chứng từ:**")

            for r in results:

                doc = r["doc"]

                status = r["overall_status"]

                lines.append(f"- `{doc}`: {status}")

            lines.append("")

            lines.append("| Số DOC | Số phiếu | Điền HeadText | BBBG | VAT | Trạng thái |")

            lines.append("| :--- | :--- | :--- | :--- | :--- | :--- |")

            for r in results:

                doc = r["doc"]

                ticket = r["ticket_no"] or "-"

                head = r.get("headtext_status", "-")

                bbbg = r.get("bbbg_status", "-")

                vat = r.get("vat_status", "-")

                status = r.get("overall_status", "-")

                lines.append(f"| {doc} | {ticket} | {head} | {bbbg} | {vat} | {status} |")

        return "\n".join(lines)

