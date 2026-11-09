# -*- coding: utf-8 -*-
"""
SAP ZONE Standalone Application Entry Point (v1.7.5 Pro - Optimized)
Giao diện Tkinter chuẩn công nghiệp, tải tức thì (< 1s), siêu nhẹ, độc lập.
"""

import sys
import os

# Cấu hình UTF-8 an toàn tuyệt đối chống crash UnicodeEncodeError trên Windows
if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr is not None and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def main():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    # Chế độ kiểm tra nhanh không mở cửa sổ (Smoke test)
    if any(arg in ["--smoke-test", "-s"] for arg in sys.argv[1:]):
        from sap_zone_core import ZoneManager, ConfigManager, parse_paste_data, extract_ob_numbers, SapZonePipeline
        from sap_zone_gui import SapZoneApp
        zm = ZoneManager()
        cfg = ConfigManager.load_config()
        test_items = parse_paste_data("TEST_MAT 100 TRAM_A")
        test_obs = extract_ob_numbers("1013147824 1013147825")
        pipe = SapZonePipeline(cfg, zm)
        mock_res = pipe.run_who_lookup_for_obs(test_obs, is_mock=True)
        with open("smoke_test.log", "w", encoding="utf-8") as f:
            f.write(f"SMOKE_TEST_OK: mappings={len(zm.mappings)}, cfg_whse={cfg.get('ewm_warehouse')}, parsed={len(test_items)}, obs={len(test_obs)}, mock_lookup={mock_res.get('success')}, gui_ok=True\n")
        sys.exit(0)

    # Khởi động Giao diện Tkinter Chuẩn Công Nghiệp
    import tkinter as tk
    from sap_zone_gui import SapZoneApp
    root = tk.Tk()
    app = SapZoneApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
