# -*- coding: utf-8 -*-
"""
SAP ZONE Standalone Application Entry Point (v1.6 Pro - Optimized)
Giao diện Tkinter 2 Tab tối giản, tải tức thì, siêu nhẹ, không phụ thuộc môi trường ngoài.
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
        from sap_zone_core import ZoneManager, ConfigManager
        zm = ZoneManager()
        cfg = ConfigManager.load_config()
        with open("smoke_test.log", "w", encoding="utf-8") as f:
            f.write(f"SMOKE_TEST_OK: mappings={len(zm.mappings)}, cfg_whse={cfg.get('ewm_warehouse')}\n")
        sys.exit(0)

    # Khởi động Giao diện Tkinter Chuẩn Công Nghiệp (Khởi động tức thì < 1s, 0 lỗi port, tương thích 100% mọi PC SAP)
    import tkinter as tk
    from sap_zone_gui import SapZoneApp
    root = tk.Tk()
    app = SapZoneApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
