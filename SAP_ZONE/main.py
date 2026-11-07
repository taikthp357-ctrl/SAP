# -*- coding: utf-8 -*-
"""
SAP ZONE Standalone Application Entry Point (v1.6 Pro)
Hỗ trợ cả Giao diện Chuẩn Tkinter 2 Tab tối giản lẫn Web UI Hiện đại (Edge Chromium / PyWebView).
"""

import sys
import os

def main():
    # Đảm bảo thư mục hiện hành là thư mục chứa file exe/script
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    # Chế độ kiểm tra nhanh không mở cửa sổ (Smoke test)
    if len(sys.argv) > 1 and sys.argv[1] == "--smoke-test":
        from sap_zone_core import ZoneManager, ConfigManager
        zm = ZoneManager()
        cfg = ConfigManager.load_config()
        print(f"SMOKE_TEST_OK: mappings={len(zm.mappings)}, cfg_whse={cfg.get('ewm_warehouse')}")
        sys.exit(0)

    # Chế độ Web UI nếu người dùng truyền cờ --web
    if len(sys.argv) > 1 and sys.argv[1] == "--web":
        from sap_zone_web import run_app
        run_app()
        return

    # Mặc định: Giao diện Tkinter 2 Tab Chuẩn Công Nghiệp (Khởi động tức thì, 0 lỗi port, tương thích 100% mọi PC SAP)
    import tkinter as tk
    from sap_zone_gui import SapZoneApp
    root = tk.Tk()
    app = SapZoneApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
