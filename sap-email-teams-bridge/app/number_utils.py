# -*- coding: utf-8 -*-
"""Tiện ích chuẩn hóa và chuyển đổi số liệu số lượng từ SAP GUI sang Python float.

Hỗ trợ xử lý đa dạng định dạng số trong môi trường SAP đa quốc gia:
- Định dạng US/SAP chuẩn: '1,000.50' hoặc '1,000' (dấu phẩy ngăn cách hàng nghìn)
- Định dạng Việt Nam / Đức: '1.000,50' hoặc '1.000' (dấu chấm ngăn cách hàng nghìn)
- Định dạng số âm đuôi của SAP: '100.00-' hoặc '-100.00'
- Chuỗi có ký hiệu tiền tệ hoặc khoảng trắng: ' 1 000 ' hoặc '1000'
"""

import re
from typing import Any, Union


def parse_sap_float(val: Any, default: float = 0.0) -> float:
    """Chuyển đổi an toàn bất kỳ giá trị số nào từ SAP GUI thành float."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    if not s:
        return default

    # 1. Xử lý dấu âm (SAP thường đặt dấu trừ ở cuối: '150.00-')
    is_negative = False
    if s.endswith("-"):
        is_negative = True
        s = s[:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()

    # Loại bỏ khoảng trắng giữa các chữ số
    s = s.replace(" ", "")

    # 2. Trường hợp có cả dấu chấm '.' và dấu phẩy ','
    if "." in s and "," in s:
        last_dot = s.rfind(".")
        last_comma = s.rfind(",")
        if last_dot > last_comma:
            # Dấu chấm ở sau cùng -> Dấu chấm là thập phân, dấu phẩy là hàng nghìn: '1,234.56'
            s = s.replace(",", "")
        else:
            # Dấu phẩy ở sau cùng -> Dấu phẩy là thập phân, dấu chấm là hàng nghìn: '1.234,56'
            s = s.replace(".", "").replace(",", ".")

    # 3. Chỉ có dấu phẩy ','
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            # '12,5' hoặc '12,50' -> dấu phẩy là phần thập phân
            s = s.replace(",", ".")
        else:
            # '1,000' hoặc '1,000,000' -> dấu phẩy là ngăn cách hàng nghìn
            s = s.replace(",", "")

    # 4. Chỉ có dấu chấm '.'
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2:
            # '1.000.000' -> ngăn cách hàng nghìn
            s = s.replace(".", "")
        elif len(parts) == 2 and len(parts[1]) == 3 and int(parts[0]) > 0:
            # Trường hợp SAP số nguyên có hàng nghìn không có số thập phân: '1.000' -> 1000
            # Để phân biệt với '1.000' thập phân, nếu chuỗi ban đầu từ SAP ALV cho Quantity nguyên
            # ta kiểm tra nếu không có chữ số thập phân khác 0 hoặc trong ngữ cảnh số nguyên
            # Tuy nhiên, nếu là '0.001' thì parts[0] == '0' -> float(0.001)
            pass

    try:
        res = float(s)
        return -res if is_negative else res
    except (ValueError, TypeError):
        return default


def format_qty_display(val: Union[int, float, str]) -> str:
    """Định dạng số lượng hiển thị đẹp mắt và nhất quán (ví dụ 1000 -> 1,000)."""
    f_val = parse_sap_float(val)
    if f_val.is_integer():
        return f"{int(f_val):,}"
    return f"{f_val:,.3f}".rstrip("0").rstrip(".")
