# -*- coding: utf-8 -*-
"""Các thao tác nguyên tử (Atomic Primitives) trên SAP GUI.
TUYỆT ĐỐI KHÔNG CHỨA LOGIC NGHIỆP VỤ (Không phân biệt MB52, MIGO, ZFIR266...).
"""
import time
from typing import Optional, Tuple, Set


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
    """Tự động tìm kiếm đối tượng theo Tên tương thích cả High Speed & Low Speed,
    kết hợp bộ nhớ đệm tự học (UI_CACHE).
    """
    if session is None:
        return None
    try:
        from app.sap_attachment import UI_CACHE
        return UI_CACHE.resolve_element(session, key_alias=element_name, field_name=element_name, element_type=element_type, window=window)
    except Exception:
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
