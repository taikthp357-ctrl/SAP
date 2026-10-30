# -*- coding: utf-8 -*-
"""
Module sap_recovery.py: Tự động phục hồi, mở lại kết nối SAP qua SSO và vượt Popup Đa phiên.
Hỗ trợ:
1. Tự động khởi chạy saplogon.exe nếu chưa chạy.
2. Tự động mở Connection chỉ định (mặc định: '13.Future VN S4 HANA PRD (SSO) Link') không cần user/password (Windows Domain SSO).
3. Tự động phát hiện và vượt qua popup 'License Information for Multiple Logons':
   - Tích chọn: 'Continue with this logon and end any other logons in the system' (radMULTI_LOGON_OPT2/OPT1)
   - Nhấn nút Tích xanh (VKey 0 / btn[0] / Enter).
4. Tự động phát hiện và xử lý SAP GUI bị treo cứng (Hard Hang) qua Win32 API.
"""

import os
import sys
import time
import subprocess
from typing import Optional, Any, Tuple

try:
    import pythoncom
    import win32com.client
    import win32gui
    import win32process
except ImportError:
    pythoncom = None
    win32com = None
    win32gui = None
    win32process = None

DEFAULT_SAP_CONNECTION = "13.Future VN S4 HANA PRD (SSO) Link"
DEFAULT_SAPLOGON_PATHS = [
    r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
    r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
]


def is_saplogon_running() -> bool:
    """Kiểm tra xem tiến trình saplogon.exe có đang chạy trong Windows không."""
    try:
        output = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq saplogon.exe", "/FO", "CSV", "/NH"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            text=True,
            timeout=3.0,
        )
        return "saplogon.exe" in output.lower()
    except Exception:
        return False


def launch_saplogon_if_needed(custom_path: Optional[str] = None) -> bool:
    """Khởi chạy saplogon.exe nếu chưa chạy và chờ ứng dụng sẵn sàng."""
    if is_saplogon_running():
        return True

    exe_candidates = []
    if custom_path and os.path.isfile(custom_path):
        exe_candidates.append(custom_path)
    exe_candidates.extend(DEFAULT_SAPLOGON_PATHS)

    chosen_path = None
    for p in exe_candidates:
        if os.path.isfile(p):
            chosen_path = p
            break

    if not chosen_path:
        print("[SAP RECOVERY WARN] Không tìm thấy file saplogon.exe tại đường dẫn tiêu chuẩn.")
        return False

    try:
        print(f"[SAP RECOVERY] Đang khởi chạy '{chosen_path}'...")
        subprocess.Popen(
            [chosen_path],
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
            close_fds=True,
        )
        for _ in range(12):
            time.sleep(0.5)
            if is_saplogon_running():
                print("[SAP RECOVERY] saplogon.exe đã khởi động thành công.")
                time.sleep(1.0)
                return True
        return False
    except Exception as ex:
        print(f"[SAP RECOVERY ERROR] Lỗi khi khởi chạy saplogon: {ex}")
        return False


def kill_hung_sap_processes(timeout_sec: float = 3.0) -> bool:
    """Dọn dẹp cưỡng chế saplogon.exe khi bị treo cứng."""
    try:
        print("[SAP RECOVERY] ⚠️ Đang giải phóng tiến trình saplogon.exe bị treo...")
        subprocess.run(
            ["taskkill", "/F", "/IM", "saplogon.exe", "/T"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            timeout=timeout_sec,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(1.5)
        return True
    except Exception as ex:
        print(f"[SAP RECOVERY WARN] Không thể kill saplogon: {ex}")
        return False


def handle_multiple_logons_popup(session, max_wait_sec: float = 2.5) -> bool:
    """
    Xử lý popup 'License Information for Multiple Logons' (Ảnh 2):
    - Tìm radio button: 'Continue with this logon and end any other logons in the system.'
    - Nhấn nút Tích Xanh (Enter / sendVKey(0) / btn[0]).
    """
    if session is None:
        return False

    t0 = time.time()
    handled = False

    while time.time() - t0 < max_wait_sec:
        try:
            # Kiểm tra wnd[1] hoặc wnd[0] có popup License
            for wnd_id in ("wnd[1]", "wnd[0]"):
                wnd = None
                try:
                    wnd = session.findById(wnd_id)
                except Exception:
                    pass

                if wnd is None:
                    continue

                title = str(getattr(wnd, "Text", "")).strip().lower()
                # Kiểm tra tiêu đề popup
                if "license information" in title or "multiple logon" in title or "already logged on" in str(getattr(wnd, "PopupDialogText", "")).lower():
                    print(f"[SAP RECOVERY] Phát hiện popup '{getattr(wnd, 'Text', '')}'. Đang chọn phương án tiếp tục phiên...")

                    # 1. Tích chọn radio button Tiếp tục và chấm dứt phiên khác
                    radio_clicked = False
                    for rad_name in ("radMULTI_LOGON_OPT2", "radMULTI_LOGON_OPT1", "radMULTI_LOGON_OPT3"):
                        for rad_path in (f"{wnd_id}/usr/{rad_name}", f"{wnd_id}/{rad_name}"):
                            try:
                                rad = session.findById(rad_path)
                                if rad is not None:
                                    rad.select()
                                    radio_clicked = True
                                    break
                            except Exception:
                                pass
                        if radio_clicked:
                            break

                    # 2. Bấm nút Tích Xanh (btn[0] hoặc sendVKey 0)
                    try:
                        btn0 = session.findById(f"{wnd_id}/tbar[0]/btn[0]")
                        if btn0 is not None:
                            btn0.Press()
                        else:
                            wnd.sendVKey(0)
                    except Exception:
                        try:
                            wnd.sendVKey(0)
                        except Exception:
                            pass

                    print("[SAP RECOVERY] ✅ Đã nhấn nút Tích Xanh (Enter) để vượt qua popup License!")
                    handled = True
                    time.sleep(1.0)
                    return True

                # Kiểm tra nếu có popup thông báo hệ thống thông thường (System message/Information)
                elif "information" in title or "system message" in title or "copyright" in title:
                    print(f"[SAP RECOVERY] Bỏ qua popup thông báo hệ thống: '{getattr(wnd, 'Text', '')}'")
                    try:
                        wnd.sendVKey(0)
                    except Exception:
                        pass
                    handled = True
                    time.sleep(0.5)

        except Exception:
            pass
        time.sleep(0.3)

    return handled


def auto_connect_sso_session(
    connection_name: str = DEFAULT_SAP_CONNECTION,
    custom_saplogon_path: Optional[str] = None,
    timeout_sec: float = 15.0,
    max_retries: int = 2,
) -> Optional[Any]:
    """
    Tự động kết nối tới SAP qua kết nối SSO (Ảnh 1) và vượt qua popup Đa phiên (Ảnh 2).
    Hoàn toàn không cần tài khoản hay mật khẩu (Xác thực Windows SSO).
    """
    if pythoncom is None or win32com is None:
        print("[SAP RECOVERY ERROR] Thiếu thư viện pywin32.")
        return None

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    for attempt in range(1, max_retries + 1):
        try:
            print(f"[SAP RECOVERY] 🚀 [Lần {attempt}/{max_retries}] Bắt đầu mở phiên SSO: '{connection_name}'...")

            # 1. Đảm bảo saplogon.exe đang chạy
            launch_saplogon_if_needed(custom_saplogon_path)

            # 2. Bắt tay với Scripting Engine
            sap_gui = None
            for _ in range(10):
                try:
                    sap_gui = win32com.client.GetObject("SAPGUI")
                    if sap_gui is not None:
                        break
                except Exception:
                    time.sleep(0.5)

            if sap_gui is None:
                print("[SAP RECOVERY WARN] Chưa thể lấy đối tượng SAPGUI từ COM.")
                continue

            app = getattr(sap_gui, "GetScriptingEngine", None)
            if app is None:
                print("[SAP RECOVERY WARN] GetScriptingEngine trả về None.")
                continue

            # 3. Mở kết nối SSO
            conn = None
            try:
                # OpenConnection(name, Sync=True)
                conn = app.OpenConnection(connection_name, True)
            except Exception as conn_err:
                print(f"[SAP RECOVERY WARN] OpenConnection '{connection_name}' văng lỗi: {conn_err}")
                # Thử tìm connection đã có sẵn
                if app.Children.Count > 0:
                    for i in range(app.Children.Count):
                        c = app.Children(i)
                        c_desc = str(getattr(c, "Description", "") or "")
                        if connection_name.lower() in c_desc.lower():
                            conn = c
                            break
                    if conn is None:
                        conn = app.Children(0)

            if conn is None or not hasattr(conn, "Children") or conn.Children.Count == 0:
                print("[SAP RECOVERY WARN] Connection mở thành công nhưng chưa có Session.")
                time.sleep(1.0)
                continue

            session = conn.Children(0)

            # 4. Bắt và xử lý popup License Information for Multiple Logons nếu có
            handle_multiple_logons_popup(session, max_wait_sec=3.0)

            # 5. Chờ màn hình SAP sẵn sàng
            time.sleep(0.5)
            s_info = getattr(session, "Info", None)
            tcode = str(getattr(s_info, "Transaction", "") or "").strip().upper() if s_info else ""
            user = str(getattr(s_info, "User", "") or "").strip().upper() if s_info else ""
            client = str(getattr(s_info, "Client", "") or "").strip() if s_info else ""

            print(f"[SAP RECOVERY] ✅ KẾT NỐI SSO THÀNH CÔNG! User: '{user}', Client: '{client}', Màn hình: '{tcode or 'MAIN'}'")
            return session

        except Exception as ex:
            print(f"[SAP RECOVERY ERROR] Lỗi trong lượt thử {attempt}: {ex}")
            time.sleep(1.0)

    return None
