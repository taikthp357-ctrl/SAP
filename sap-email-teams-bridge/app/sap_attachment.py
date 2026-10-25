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
from concurrent.futures import ThreadPoolExecutor, as_completed
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
BATCH_TAB_LOCKS = {                    # Khóa riêng cho từng Tab trong Batch Pool
    1: threading.Lock(),
    2: threading.Lock()
}


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
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/cmbGODYNPRO-ACTION",
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
        "wnd[0]/titl/shellcont/shell",
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


def smart_open_gos_create_attachment(session) -> bool:
    """Tự động mở cửa sổ đính kèm file ở cả High Speed lẫn Low Speed."""
    if session is None:
        return False
    # 1. Thử cách High Speed (Bấm trực tiếp nút GOS trên thanh tiêu đề)
    gos_shell = safe_find_by_id(session, "wnd[0]/titl/shellcont/shell")
    if gos_shell is not None:
        try:
            gos_shell.PressContextButton("%GOS_TOOLBOX")
            gos_shell.SelectContextMenuItem("%GOS_PCATTA_CREA")
            return True
        except Exception:
            pass

    # 2. Thử menu System -> Services for Object -> Create Attachment
    try:
        mbar_item = safe_find_by_id(session, "wnd[0]/mbar/menu[0]/menu[11]/menu[0]")
        if mbar_item is not None:
            mbar_item.Select()
            return True
    except Exception:
        pass

    # 3. Gửi thẳng Function Code vào ô okcd
    try:
        okcd = safe_find_by_id(session, "wnd[0]/tbar[0]/okcd")
        if okcd is not None:
            okcd.Text = "%GOS_PCATTA_CREA"
            w0 = safe_find_by_id(session, "wnd[0]")
            if w0:
                w0.SendVKey(0)
            return True
    except Exception:
        pass

    return False


def smart_open_gos_view_attachment(session) -> bool:
    """Tự động mở danh sách đính kèm file ở cả High Speed lẫn Low Speed."""
    if session is None:
        return False
    # 1. Thử nút OK_REF_ATTA
    btn = smart_find(session, "OK_REF_ATTA", "GuiButton")
    if btn is not None:
        enabled = getattr(btn, "Enabled", True)
        if enabled is False or str(enabled).upper() == "FALSE":
            return False  # Không có attachment
        try:
            btn.Press()
            return True
        except Exception:
            pass

    # 2. Thử GOS shell %GOS_VIEW_ATTA
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


def ensure_sap_tabs(target_tabs: int = 3) -> int:
    """Đảm bảo mở sẵn đủ số Tab cần thiết (Tab 0: Fast Lane, Tab 1 & 2: Batch Pool)."""
    if pythoncom is None or win32com is None:
        return 1
    pythoncom.CoInitialize()
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        if app.Children.Count == 0:
            return 1
        conn = app.Children(0)
        current = conn.Children.Count
        while current < target_tabs and current < 6:
            try:
                conn.Children(0).CreateSession()
                time.sleep(0.8)
                current = conn.Children.Count
            except Exception as e:
                print(f"[SAP MULTI-TAB WARN] Không thể mở thêm tab: {e}")
                break
        return conn.Children.Count
    except Exception as ex:
        print(f"[SAP MULTI-TAB ERROR] {ex}")
        return 1
    finally:
        pythoncom.CoUninitialize()


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
VAT_REGEX = re.compile(r"(?:^|[^a-z0-9])(vat|invoice|inv|hoa[ _-]*don|hd|hddt|gtgt|tax|ct|chung[ _-]*tu)(?:$|[^a-z0-9])", re.I)
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


class MigoAttachmentRunner:
    """Điều khiển tự động hóa điền số phiếu và đính kèm BBBG / VAT / Email vào SAP MIGO qua GOS."""

    def __init__(self, sap_engine=None, expected_session_identity=None):
        self.sap_engine = sap_engine
        self.expected_session_identity = dict(expected_session_identity or {})
        self._active_document = None

    def verify_session_identity(self, session, require_migo=False):
        if session is None:
            return False, "NO_SESSION"
        required = ("system", "client", "user", "session_id")
        if not all(self.expected_session_identity.get(key) for key in required):
            return False, "SESSION_IDENTITY_REQUIRED"
        try:
            actual = {"system": session.Info.SystemName, "client": session.Info.Client,
                      "user": session.Info.User, "session_id": session.Id}
            if any(not isinstance(actual[key], str) or actual[key].strip().casefold() != str(self.expected_session_identity[key]).strip().casefold() for key in required):
                return False, "SESSION_IDENTITY_MISMATCH"
            if session.Busy:
                return False, "SESSION_BUSY"
            if require_migo and str(session.Info.Transaction).upper() != "MIGO":
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
            "bbbg": bbbg_files,
            "vat": vat_files,
            "mail": mail_files
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

    def handle_migo_headtext(self, session, ticket_no: str) -> Tuple[str, str]:
        """Kiểm tra và điền số phiếu vào trường Header Text theo chuẩn Change_doc:
        MIGO_OK_HEADER_OPEN -> GOHEAD-BKTXT.Text = ticket_no -> SendVKey 11 (Save).
        """
        identity_ok, identity_status = self.verify_session_identity(session, require_migo=True)
        if not identity_ok or self._active_document is None:
            return "HEADTEXT_BLOCKED", identity_status if not identity_ok else "DOCUMENT_NOT_VERIFIED"
        if not ticket_no or not str(ticket_no).strip():
            return "SKIP_NO_TICKET", "Không có số phiếu cần điền"

        ticket_clean = str(ticket_no).strip()

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
                curr_tokens, curr_compact = extract_search_tokens(current_text)
                tick_tokens, tick_compact = extract_search_tokens(ticket_clean)
                if curr_compact == tick_compact:
                    return "HEADTEXT_OK", f"Đã có đúng số phiếu ('{current_text}')"
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
            time.sleep(0.3)

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
            doc, year = self._active_document
            loaded, reason = self.open_migo_document(session, doc, year, mode="DISPLAY", force_reload=True)
            if not loaded:
                return "HEADTEXT_UNCERTAIN", f"Chưa đọc lại chứng từ sau lưu: {reason}"
            readback = smart_find(session, "GOHEAD-BKTXT", "GuiTextField")
            if readback is None or str(getattr(readback, "Text", "")).strip() != ticket_clean:
                return "HEADTEXT_UNCERTAIN", "Header Text đọc lại không khớp; cần kiểm tra thủ công"
            return "HEADTEXT_DONE", f"Đã điền và lưu mới HeadText, đọc lại khớp: '{ticket_clean}'"
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

    def read_sap_attachment_titles(self, session) -> List[str]:
        """Kiểm tra và đọc danh sách tiêu đề các file đính kèm đã có sẵn trên SAP MIGO qua nút OK_REF_ATTA / Attachment list."""
        titles: List[str] = []
        if session is None:
            return titles

        with POPUP_LOCK:
            try:
                # 1. Mở cửa sổ Attachment list thông minh (OK_REF_ATTA / GOS Shell / Menu / Function Code)
                has_opened = smart_open_gos_view_attachment(session)
                if not has_opened:
                    return []

                time.sleep(0.5)

                # 2. Đọc popup Attachment list (wnd[1])
                popup = None
                try:
                    popup = session.findById("wnd[1]")
                except Exception:
                    pass

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
                                if children:
                                    for idx in range(children.Count):
                                        res = find_grid(children.Item(idx), depth + 1)
                                        if res:
                                            return res
                            except Exception:
                                pass
                            return None
                        grid = find_grid(popup)

                    if grid:
                        try:
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
                        except Exception as e:
                            print(f"[SAP ATT LIST WARN] Lỗi đọc grid: {e}")

                    # 3. Đóng popup wnd[1]
                    try:
                        popup.Close()
                    except Exception:
                        try:
                            session.findById("wnd[1]").SendVKey(12)
                        except Exception:
                            pass

                    time.sleep(0.3)
            except Exception as ex:
                print(f"[SAP ATT LIST ERROR] {ex}")

        return titles

    def attach_file_to_current_migo(self, session, file_path: Path) -> Tuple[bool, str]:
        """Đính kèm 1 file vật lý vào chứng từ MIGO đang mở trên SAP qua GOS (%GOS_TOOLBOX)."""
        identity_ok, identity_status = self.verify_session_identity(session, require_migo=True)
        if not identity_ok or self._active_document is None:
            return False, identity_status if not identity_ok else "DOCUMENT_NOT_VERIFIED"
        abs_path = file_path.resolve()
        if not abs_path.exists():
            return False, f"FILE_NOT_FOUND: {abs_path}"

        parent_folder = str(abs_path.parent)
        filename = abs_path.name

        try:
            # 1. Nhấn nút GOS Toolbox & Chọn Create Attachment bên trong Micro-lock
            with POPUP_LOCK:
                def _trigger_gos_and_fill():
                    opened = smart_open_gos_create_attachment(session)
                    if not opened:
                        raise RuntimeError("GOS_TRIGGER_FAILED: Không thể mở hộp thoại đính kèm file GOS")

                    time.sleep(0.4)

                    # Điền đường dẫn và nhấn nút Open/btn[0] với hỗ trợ AdaptiveConfigCache
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

                # Do not retry a sequence containing upload: a lost response may
                # follow a successful write and retrying could duplicate files.
                _trigger_gos_and_fill()
                # Vừa nhấn btn[0] xong là wnd[1] bắt đầu đóng -> nhả POPUP_LOCK ngay cho Tab kia!
                time.sleep(0.4)

            # 4. Kiểm tra thông báo trạng thái tại thanh sbar
            sbar = session.findById("wnd[0]/sbar")
            msg = getattr(sbar, "Text", "")
            msg_type = getattr(sbar, "MessageType", "")

            text = str(msg).upper()
            if msg_type in ("E", "A") or any(k in text for k in ["NOT CREATED", "NOT SAVED", "FAILED", "ERROR", "CANCELLED", "CANCELED", "KHONG THANH CONG"]):
                return False, f"SAP_ERROR: {msg}"
            if msg_type != "S":
                return False, f"ATTACHMENT_UNCERTAIN: {msg}"
            titles = self.read_sap_attachment_titles(session)
            if not any(Path(str(title)).name.casefold() == filename.casefold() for title in titles):
                return False, "ATTACHMENT_UNCERTAIN: Chưa đọc lại đúng tên file trên SAP"
            return True, f"DONE_SAP_READBACK: {filename}"

        except Exception as e:
            return False, f"GOS_EXCEPTION: {e}"

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
                time.sleep(0.3)

            # Đặt Action = A12 (Change)
            target_act = "A12" if mode == "CHANGE" else "A04"
            act_cb = None
            try:
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    act_cb = w0.findByName("GODYNPRO-ACTION", "GuiComboBox")
            except Exception:
                pass
            if act_cb is None:
                act_cb = wait_control(session, "GODYNPRO-ACTION", "GuiComboBox", timeout_sec=2.0)

            if act_cb is not None:
                try:
                    act_cb.Key = target_act
                except Exception:
                    try:
                        act_cb.key = target_act
                    except Exception:
                        pass
                time.sleep(0.3)
                if str(getattr(act_cb, "Key", "")) != target_act:
                    return False
            else:
                return False
            return str(session.Info.Transaction).upper() == "MIGO"
        except Exception as e:
            print(f"[MIGO INIT WARN] {e}")
            return False

    def open_migo_document(self, session, doc: str, year: str = "", mode: str = "CHANGE", force_reload: bool = False) -> Tuple[bool, str]:
        """Nạp chứng từ MIGO theo chuẩn luồng Change_doc:
        Điền số DOC vào GODYNPRO-MAT_DOC và gửi Enter (SendVKey 0).
        """
        self._active_document = None
        identity_ok, reason = self.verify_session_identity(session)
        if not identity_ok:
            return False, reason
        if not re.fullmatch(r"\d{10}", str(doc)) or not re.fullmatch(r"(?:19|20|21)\d{2}", str(year)):
            return False, "DOCUMENT_AND_YEAR_REQUIRED"
        try:
            if force_reload or str(session.Info.Transaction).upper() != "MIGO":
                if not self.prepare_migo_session(session, mode=mode):
                    return False, "SESSION_PREPARE_FAILED"
            # 1. Tìm ô nhập số chứng từ GODYNPRO-MAT_DOC
            doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")

            # Nếu chưa thấy hoặc ô đang bị khóa do chứng từ trước chưa thoát
            if doc_fld is None or not getattr(doc_fld, "Changeable", True):
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    try:
                        w0.SendVKey(12)  # Esc (Cancel) để hủy/thoát chứng từ cũ thay vì F3
                        time.sleep(0.2)
                    except Exception:
                        pass
                doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")

            # Nếu sau Esc vẫn chưa sẵn sàng -> Reset về MIGO A12 sạch sẽ
            if doc_fld is None or not getattr(doc_fld, "Changeable", True):
                if not self.prepare_migo_session(session, mode=mode):
                    return False, "SESSION_PREPARE_FAILED"
                doc_fld = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")

            if doc_fld is None:
                return False, "CONTROL_NOT_FOUND: Không tìm thấy ô nhập số chứng từ GODYNPRO-MAT_DOC trên MIGO"

            # 2. Điền số chứng từ vào ô vàng
            try:
                doc_fld.SetFocus()
            except Exception:
                pass
            try:
                doc_fld.Text = str(doc).strip()
            except Exception:
                doc_fld.text = str(doc).strip()

            year_fld = smart_find(session, "GODYNPRO-DOC_YEAR", "GuiTextField") or smart_find(session, "GODYNPRO-DOC_YEAR", "GuiCTextField")
            if year_fld is None:
                return False, "YEAR_CONTROL_NOT_FOUND"
            year_fld.Text = str(year)
            # 3. Enter nạp chứng từ
            w0 = safe_find_by_id(session, "wnd[0]")
            if w0:
                w0.SendVKey(0)
            time.sleep(0.3)

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
            read_doc = smart_find(session, "GODYNPRO-MAT_DOC", "GuiTextField") or smart_find(session, "GODYNPRO-MAT_DOC", "GuiCTextField")
            read_year = smart_find(session, "GODYNPRO-DOC_YEAR", "GuiTextField") or smart_find(session, "GODYNPRO-DOC_YEAR", "GuiCTextField")
            if read_doc is None or read_year is None:
                return False, "DOCUMENT_READBACK_UNAVAILABLE"
            if str(getattr(read_doc, "Text", "")).strip() != str(doc) or str(getattr(read_year, "Text", "")).strip() != str(year):
                return False, "DOCUMENT_READBACK_MISMATCH"
            if getattr(read_doc, "Changeable", True):
                return False, "DOCUMENT_NOT_LOADED"
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
        dry_run: bool = False
    ) -> Tuple[Dict[str, Any], bool]:
        """Xử lý 1 chứng từ MIGO duy nhất (Điền HeadText, đính kèm BBBG/VAT)."""
        doc = str(item.get("doc", "")).strip()
        ticket_no = str(item.get("ticket_no", "")).strip()
        year = str(item.get("year") or year or "").strip()
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
                        headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no)
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
                        time.sleep(0.2)
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
                        headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no)
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
            loaded_ok, nav_msg = self.open_migo_document(session, doc, year, mode="CHANGE" if (fill_headtext and ticket_no) else "DISPLAY")
            if not loaded_ok:
                bbbg_res = nav_msg
                vat_res = nav_msg
                overall_status = nav_msg
            else:
                # 0. KIỂM TRA ĐÍNH KÈM ĐÃ CÓ SẴN TRÊN SAP MIGO (Display Attachment / OK_REF_ATTA)
                sap_titles = self.read_sap_attachment_titles(session)
                sap_bbbg, sap_vat = self.check_existing_sap_attachments(sap_titles, ticket_no)

                # 1. Điền số phiếu vào Header Text nếu được yêu cầu
                if fill_headtext and ticket_no:
                    headtext_res, headtext_msg = self.handle_migo_headtext(session, ticket_no)

                # 2. Đính kèm file BBBG
                b_ok = False
                if sap_bbbg:
                    b_ok = True
                    bbbg_res = "DONE"
                    print(f"[ATT RUNNER] DOC {doc}: BBBG đã có sẵn trên SAP ('{sap_bbbg}'), bỏ qua upload lại.")
                elif bbbg_paths:
                    succ_bb = []
                    fail_bb = []
                    for bp in bbbg_paths:
                        ok_b, msg_b = self.attach_file_to_current_migo(session, bp)
                        if ok_b:
                            succ_bb.append(bp.name)
                        else:
                            fail_bb.append(f"{bp.name}: {msg_b}")
                    if succ_bb:
                        b_ok = True
                        bbbg_res = f"DONE ({len(succ_bb)})" if len(succ_bb) > 1 else "DONE"
                    else:
                        bbbg_res = f"Lỗi: {'; '.join(fail_bb)}"
                else:
                    bbbg_res = "FILE_NOT_FOUND"

                # 3. Đính kèm file VAT (hỗ trợ 1 hoặc nhiều file VAT)
                v_ok = False
                if sap_vat:
                    v_ok = True
                    vat_res = "DONE"
                    print(f"[ATT RUNNER] DOC {doc}: VAT đã có sẵn trên SAP ('{sap_vat}'), bỏ qua upload lại.")
                elif vat_paths:
                    succ_vat = []
                    fail_vat = []
                    for vp in vat_paths:
                        ok_v, msg_v = self.attach_file_to_current_migo(session, vp)
                        if ok_v:
                            succ_vat.append(vp.name)
                        else:
                            fail_vat.append(f"{vp.name}: {msg_v}")
                    if succ_vat:
                        v_ok = True
                        vat_res = f"DONE ({len(succ_vat)})" if len(succ_vat) > 1 else "DONE"
                    else:
                        vat_res = f"Lỗi: {'; '.join(fail_vat)}"
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
                    is_succ = True
                elif v_ok and not b_ok:
                    overall_status = "PARTIAL (Thiếu BBBG)"
                    is_succ = True
                elif m_ok:
                    overall_status = "DONE"
                    is_succ = True
                elif headtext_res in ("HEADTEXT_DONE", "HEADTEXT_OK"):
                    overall_status = "PARTIAL (Chỉ có HeadText)"
                    is_succ = True
                else:
                    overall_status = "FILE_NOT_FOUND"
                    is_succ = False

        # Sau khi xử lý xong chứng từ, gửi Esc (SendVKey 12) để MIGO trở về màn hình ô vàng cho chứng từ tiếp theo
        if not dry_run and session is not None:
            try:
                w0 = safe_find_by_id(session, "wnd[0]")
                if w0:
                    w0.SendVKey(12)  # Esc thoát chứng từ
                    time.sleep(0.2)
            except Exception:
                pass

        return {
            "doc": doc,
            "ticket_no": ticket_no,
            "headtext_status": headtext_res,
            "headtext_msg": headtext_msg,
            "bbbg_file": ", ".join(p.name for p in bbbg_paths) if bbbg_paths else None,
            "vat_file": ", ".join(p.name for p in vat_paths) if vat_paths else None,
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
        progress_callback: Optional[Callable[[int, int, str, int], None]] = None
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
        total_missing = 0

        total_items = len(items)
        last_progress_time = time.time()
        PROGRESS_INTERVAL_SEC = 180.0  # Cập nhật tiến độ lên Teams mỗi 3 phút

        # Xác định khả năng chạy Multi-Tab (Danh sách > 15 chứng từ, có kết nối SAP và không dry_run)
        # Index-based tab selection is not proof of ownership. Keep writes on the
        # single explicitly configured and reverified session.
        is_multi_tab_eligible = False

        active_tab_indices = [0]
        if is_multi_tab_eligible:
            available_tabs = ensure_sap_tabs(3)
            if available_tabs >= 3:
                # Tab 0 dành riêng cho Fast Lane! Tab 1 và Tab 2 dùng cho Batch Pool
                active_tab_indices = [1, 2]
                print(f"[ATT RUNNER] Đã kích hoạt Multi-Tab Engine: Chia tải trên Tab 1 & Tab 2 (Chừa Tab 0 cho Fast Lane).")
            elif available_tabs >= 2:
                active_tab_indices = [1]
                print(f"[ATT RUNNER] SAP có 2 tabs: Chạy Batch trên Tab 1 (Chừa Tab 0 cho Fast Lane).")
            else:
                active_tab_indices = [0]

        # ---------------------------------------------------------------------
        # NHÁNH A: CHẠY ĐA TAB SONG SONG (MULTI-TAB BATCH POOL)
        # ---------------------------------------------------------------------
        if len(active_tab_indices) > 1:
            tab_chunks = {t_idx: [] for t_idx in active_tab_indices}
            for i, itm in enumerate(items):
                target_tab = active_tab_indices[i % len(active_tab_indices)]
                tab_chunks[target_tab].append((i, itm))

            processed_count = 0
            success_count = 0
            prog_lock = threading.Lock()
            indexed_results = []

            def _tab_worker(tab_idx: int, chunk: List[Tuple[int, Dict[str, str]]]):
                nonlocal processed_count, success_count, last_progress_time
                if pythoncom is not None:
                    pythoncom.CoInitialize()
                worker_results = []
                tab_lock = BATCH_TAB_LOCKS.get(tab_idx, threading.Lock())

                with tab_lock:
                    worker_session = None
                    try:
                        sap_gui = win32com.client.GetObject("SAPGUI")
                        conn = sap_gui.GetScriptingEngine.Children(0)
                        if tab_idx < conn.Children.Count:
                            worker_session = conn.Children(tab_idx)
                    except Exception as conn_err:
                        print(f"[ATT RUNNER Tab {tab_idx}] Lỗi kết nối session: {conn_err}")

                    if worker_session is not None:
                        self.prepare_migo_session(worker_session, mode="CHANGE")

                    for orig_idx, itm in chunk:
                        doc_str = str(itm.get("doc", "")).strip()
                        try:
                            res_dict, is_s = self.process_single_migo_item(
                                worker_session, itm, index_data, fill_headtext, year,
                                is_folder_requested, has_folder, folder_error, dry_run=False
                            )
                        except Exception as item_err:
                            print(f"[ATT RUNNER Tab {tab_idx} ERROR] DOC {doc_str}: {item_err}")
                            sanitize_migo_session(worker_session)
                            res_dict = {
                                "doc": doc_str,
                                "ticket_no": str(itm.get("ticket_no", "")).strip(),
                                "headtext_status": "EXCEPTION",
                                "headtext_msg": str(item_err),
                                "bbbg_status": "ERROR",
                                "vat_status": "ERROR",
                                "mail_status": "ERROR",
                                "overall_status": "ERROR",
                                "match_notes": f"Tab {tab_idx} exception: {item_err}"
                            }
                            is_s = False

                        worker_results.append((orig_idx, res_dict, is_s))

                        with prog_lock:
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

                if pythoncom is not None:
                    pythoncom.CoUninitialize()
                return worker_results

            with ThreadPoolExecutor(max_workers=len(active_tab_indices)) as executor:
                futures = [
                    executor.submit(_tab_worker, t_idx, tab_chunks[t_idx])
                    for t_idx in active_tab_indices
                ]
                for fut in as_completed(futures):
                    indexed_results.extend(fut.result())

            # Khôi phục đúng thứ tự ban đầu của danh sách
            indexed_results.sort(key=lambda x: x[0])
            for _, r_dict, is_s in indexed_results:
                results.append(r_dict)
                is_actually_succ = is_s or r_dict.get("is_succ") is True or r_dict.get("overall_status") == "DONE" 
                if is_actually_succ:
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
                        conn = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0)
                        session = conn.Children(single_tab_idx)
                    except Exception:
                        session = self.sap_engine.ensure_session()
                else:
                    session = self.sap_engine.ensure_session()

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
                is_actually_succ = is_s or res_dict.get("is_succ") is True or res_dict.get("overall_status") == "DONE" 
                if is_actually_succ:
                    total_success += 1
                else:
                    total_missing += 1

        return {
            "folder_path": clean_folder if is_folder_requested else "",
            "folder_error": folder_error,
            "total_items": len(items),
            "total_success": total_success,
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
