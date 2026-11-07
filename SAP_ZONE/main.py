# -*- coding: utf-8 -*-
"""
SAP ZONE Standalone Application Entry Point (v1.6)
Mặc định: Giao diện Web UI Hiện Đại (Edge Chromium / PyWebView)
Hỗ trợ fallback: Tkinter legacy (--tk) & Smoke test (--smoke-test)
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

    # Chế độ GUI Tkinter cũ nếu có cờ --tk
    if len(sys.argv) > 1 and sys.argv[1] == "--tk":
        import tkinter as tk
        from sap_zone_gui import SapZoneApp
        root = tk.Tk()
        app = SapZoneApp(root)
        root.mainloop()
        return

    # Mặc định: Giao diện Web UI Hiện Đại
    from sap_zone_web import run_app
    run_app()

if __name__ == "__main__":
    main()
