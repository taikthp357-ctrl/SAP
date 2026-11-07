# -*- coding: utf-8 -*-
"""Các thao tác nguyên tử (Atomic Primitives) trên SAP GUI.
TUYỆT ĐỐI KHÔNG CHỨA LOGIC NGHIỆP VỤ (Không phân biệt MB52, MIGO, ZFIR266...).
"""
import time
from contextlib import contextmanager
from typing import Optional, Tuple, Set

try:
    import pythoncom
except Exception:
    pythoncom = None

try:
    import win32com.client
except Exception:
    win32com = None


def copy_to_clipboard_safely(text: str) -> bool:
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


def get_clipboard_safely() -> Optional[str]:
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


def find_by_id_safe(session, id_text: str):
    """Tìm đối tượng theo ID an toàn tuyệt đối, tương thích cả 1 tham số và 2 tham số."""
    if session is None:
        return None
    try:
        return session.findById(id_text)
    except Exception:
        try:
            return session.findById(id_text, False)
        except Exception:
            return None


def wait_idle(session, timeout_sec: float = 30.0) -> bool:
    """Đợi SAP GUI hoàn tất xử lý (Busy = False) như chuẩn modSAPCore.WaitIdle."""
    if session is None:
        return False
    t0 = time.time()
    while time.time() - t0 < timeout_sec:
        try:
            if not getattr(session, "Busy", False):
                return True
        except Exception:
            return False
        time.sleep(0.1)
    return False


def smart_find(session, element_name: str, element_type: Optional[str] = None, window: str = "wnd[0]"):
    """Tự động tìm kiếm đối tượng theo Tên trên window."""
    if session is None:
        return None
    w = find_by_id_safe(session, window)
    if w and hasattr(w, "findByName"):
        try:
            return w.findByName(element_name, element_type) if element_type else w.findByName(element_name)
        except Exception:
            pass
    return None


def get_status(session) -> Tuple[str, str]:
    """Đọc thông điệp và kiểu trạng thái từ StatusBar (wnd[0]/sbar) như modSAPCore."""
    if session is None:
        return "", ""
    try:
        sbar = find_by_id_safe(session, "wnd[0]/sbar")
        if sbar is not None:
            txt = getattr(sbar, "Text", getattr(sbar, "text", ""))
            mtype = getattr(sbar, "MessageType", getattr(sbar, "messageType", ""))
            return str(txt).strip(), str(mtype).strip().upper()
    except Exception:
        pass
    return "", ""


def dismiss_popups(session) -> str:
    """Kiểm tra và đóng cửa sổ popup wnd[1] nếu có như modSAPCore.ClosePopups."""
    if session is None:
        return ""
    all_msgs = []
    try:
        count = 0
        while count < 10:
            wnd1 = find_by_id_safe(session, "wnd[1]")
            if wnd1 is None:
                break
            t = str(getattr(wnd1, "Text", "")).strip()
            b_txt = ""
            try:
                usr1 = find_by_id_safe(session, "wnd[1]/usr")
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
            wait_idle(session, 3.0)
    except Exception:
        pass
    return " | ".join(all_msgs)


def open_transaction(session, tcode: str) -> None:
    """Mở Transaction trên SAP GUI an toàn tuyệt đối qua ô lệnh okcd hoặc StartTransaction."""
    if session is None:
        raise RuntimeError("Phiên SAP chưa khởi tạo (session is None)")

    raw = tcode.strip()
    if raw.upper().startswith("/N"):
        raw = raw[2:].strip()
    elif raw.upper().startswith("/O"):
        raw = raw[2:].strip()

    if raw.startswith("/") and raw.count("/") >= 2:
        clean_code = raw
    else:
        clean_code = raw.lstrip("/")

    okcd_command = f"/n{clean_code}"
    print(f"[SAP] Đang mở giao dịch: {okcd_command}...")

    try:
        okcd = find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
        if okcd is not None:
            okcd.text = okcd_command
            w0 = find_by_id_safe(session, "wnd[0]")
            if w0:
                w0.sendVKey(0)
            wait_idle(session, 5.0)
            dismiss_popups(session)
            return
    except Exception as e:
        print(f"[SAP WARN] Không thể nhập qua okcd: {e}")

    try:
        raw_tcode = clean_code
        print(f"[SAP] Thử mở bằng StartTransaction: {raw_tcode}...")
        session.StartTransaction(raw_tcode)
        wait_idle(session, 5.0)
        dismiss_popups(session)
    except Exception as st_err:
        raise RuntimeError(f"Không thể mở giao dịch {tcode}") from st_err


def reset_to_session_manager(session) -> None:
    """Đưa SAP GUI về màn hình gốc /n (SESSION_MANAGER) để giải phóng toàn bộ ALV grid, bộ nhớ và session."""
    if session is None:
        return
    try:
        dismiss_popups(session)
        okcd = find_by_id_safe(session, "wnd[0]/tbar[0]/okcd")
        if okcd is not None:
            okcd.text = "/n"
            w0 = find_by_id_safe(session, "wnd[0]")
            if w0:
                w0.sendVKey(0)
            wait_idle(session, 2.0)
        else:
            open_transaction(session, "SESSION_MANAGER")
            wait_idle(session, 2.0)
        print("[SAP] Đã đưa phiên làm việc về màn hình gốc /n thành công.")
    except Exception as e:
        print(f"[SAP RESET WARN] Lỗi khi đưa về màn hình gốc /n: {e}")


@contextmanager
def transaction_scope(session, tcode: str):
    """Đảm bảo mọi Reader/Writer luôn bắt đầu sạch và trả lại màn hình /n khi kết thúc."""
    try:
        dismiss_popups(session)
        open_transaction(session, tcode)
        yield session
    finally:
        try:
            dismiss_popups(session)
            reset_to_session_manager(session)
        except Exception:
            pass


def get_safe_sap_connection(app=None, conn_idx_or_id: int = 0):
    """Lấy đối tượng SAP GuiConnection an toàn theo index hoặc id."""
    try:
        if app is None:
            import win32com.client
            sap_gui = win32com.client.GetObject("SAPGUI")
            app = sap_gui.GetScriptingEngine
        if app is None or not hasattr(app, "Children") or app.Children.Count == 0:
            return None
        if isinstance(conn_idx_or_id, int):
            if 0 <= conn_idx_or_id < app.Children.Count:
                return app.Children(conn_idx_or_id)
        return app.Children(0)
    except Exception:
        return None


def get_safe_sap_session(conn_idx: int = 0, tab_idx: int = 0):
    """Lấy phiên GuiSession theo tab_idx (0-5) an toàn cho từng luồng/tiến trình.
    TUYỆT ĐỐI KHÔNG FALLBACK về conn.Children(0) nếu tab_idx > 0 để tránh xung đột phiên (Session Collision).
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        import win32com.client
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        conn = get_safe_sap_connection(app, conn_idx)
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
        # Chỉ fallback về tab 0 nếu người gọi thực sự yêu cầu tab 0
        if tab_idx == 0 and conn.Children.Count > 0:
            return conn.Children(0)
        return None
    except Exception as ex:
        print(f"[SAP SESSION ERROR] conn={conn_idx}, tab={tab_idx}: {ex}")
        return None


def get_all_unique_sap_sessions(conn_idx: int = 0) -> list:
    """Lấy danh sách tất cả các GuiSession độc lập và duy nhất (không trùng lặp Id) trên connection."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    sessions = []
    seen_ids = set()
    try:
        import win32com.client
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        conn = get_safe_sap_connection(app, conn_idx)
        if conn is None or not hasattr(conn, "Children"):
            return []
        for idx in range(conn.Children.Count):
            try:
                s = conn.Children(idx)
                s_id = str(getattr(s, "Id", "") or "").strip()
                if not s_id:
                    s_id = f"conn_{conn_idx}_ses_{idx}"
                if s_id not in seen_ids:
                    seen_ids.add(s_id)
                    sessions.append(s)
            except Exception:
                pass
    except Exception as ex:
        print(f"[SAP SESSIONS ERROR] {ex}")
    return sessions


def call_com_resilient(func, *args, retries: int = 3, delay: float = 0.2, **kwargs):
    """Thực thi hàm COM an toàn tuyệt đối với cơ chế tự động thử lại khi gặp RPC_E_CALL_REJECTED."""
    last_err = None
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_err = e
            err_str = str(e).lower()
            if "rejected by callee" in err_str or "-2147418111" in err_str or "server busy" in err_str:
                time.sleep(delay * (attempt + 1))
                continue
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                raise
    if last_err is not None:
        raise last_err


def ensure_sap_tabs(target_tabs: int = 4, connection_index: int = 0) -> int:
    """Mở sẵn số Tab (Session) trên SAP GUI nếu chưa đủ, tối đa 6 tab."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        import win32com.client
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        conn = get_safe_sap_connection(app, connection_index)
        if conn is None:
            return 1
        current = conn.Children.Count
        target_tabs = min(6, max(1, target_tabs))
        while current < target_tabs and current < 6:
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
                    print("[SAP MULTI-TAB WARN] Session không hỗ trợ CreateSession()")
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
                    print(f"[SAP MULTI-TAB WARN] Chờ mở tab {orig_count + 1} quá thời gian.")
                    break
                current = conn.Children.Count
            except Exception as e:
                print(f"[SAP MULTI-TAB WARN] Không thể mở thêm tab: {e}")
                break
        return conn.Children.Count
    except Exception as ex:
        print(f"[SAP MULTI-TAB ERROR] {ex}")
        return 1


