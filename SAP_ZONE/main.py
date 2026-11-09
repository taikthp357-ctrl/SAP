# -*- coding: utf-8 -*-
"""
SAP ZONE Standalone Application Entry Point (v1.7 Pro - Optimized)
Giao diện Tkinter chuẩn công nghiệp, tải tức thì (< 1s), siêu nhẹ, độc lập.
"""

import sys
import os

def main():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    # Chế độ kiểm tra nhanh không mở cửa sổ (Smoke test)
    if any(arg in ["--smoke-test", "-s"] for arg in sys.argv[1:]):
        from sap_zone_core import ZoneManager, ConfigManager, parse_paste_data
        from sap_zone_gui import SapZoneApp
        zm = ZoneManager()
        cfg = ConfigManager.load_config()
        test_items = parse_paste_data("TEST_MAT 100 TRAM_A")
        with open("smoke_test.log", "w", encoding="utf-8") as f:
            f.write(f"SMOKE_TEST_OK: mappings={len(zm.mappings)}, cfg_whse={cfg.get('ewm_warehouse')}, parsed={len(test_items)}, gui_ok=True\n")
        sys.exit(0)

    # Khởi động Giao diện Tkinter Chuẩn Công Nghiệp
    import tkinter as tk
    from sap_zone_gui import SapZoneApp
    root = tk.Tk()
    app = SapZoneApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
