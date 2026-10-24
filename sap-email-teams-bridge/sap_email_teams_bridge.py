# -*- coding: utf-8 -*-
"""Entry point alias cho main.py (hỗ trợ gọi bằng cả tên main.py lẫn sap_email_teams_bridge.py)."""
import sys
from main import run_bridge, run_self_test

if __name__ == "__main__":
    if "--test-parser" in sys.argv:
        run_self_test()
    else:
        run_bridge()
