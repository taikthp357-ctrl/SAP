# -*- coding: utf-8 -*-
"""Tiện ích chuẩn hóa và chuyển đổi số liệu số lượng từ SAP GUI sang Python float.

Hỗ trợ xử lý đa dạng định dạng số trong môi trường SAP đa quốc gia:
- Cấu hình rõ ràng: 'US' (1,234.56), 'EU' / 'VN' (1.234,56), hoặc 'AUTO'
- Định dạng US/SAP chuẩn: '1,000.50' hoặc '1,000' (dấu phẩy ngăn cách hàng nghìn)
- Định dạng Việt Nam / Đức: '1.000,50' hoặc '1.000' (dấu chấm ngăn cách hàng nghìn)
- Định dạng số âm đuôi của SAP: '100.00-' hoặc '-100.00'
- Chuỗi chứa ký tự chữ hoặc rác: 'abc.123', ' 1 000 ', 'N/A' (không bao giờ gây unhandled exception)
"""

import re
from typing import Any, Union, Optional

_GLOBAL_SAP_NUMBER_FORMAT = "AUTO"


def set_global_sap_number_format(fmt: str):
    """Thiết lập định dạng số SAP mặc định cho toàn hệ thống ('AUTO', 'US', 'EU', 'VN')."""
    global _GLOBAL_SAP_NUMBER_FORMAT
    if fmt:
        _GLOBAL_SAP_NUMBER_FORMAT = fmt.strip().upper()


def parse_sap_float(val: Any, default: float = 0.0, fmt: Optional[str] = None) -> float:
    """Chuyển đổi an toàn bất kỳ giá trị số nào từ SAP GUI thành float.
    
    Không bao giờ văng ngoại lệ dù dữ liệu đầu vào là chuỗi lạ ('abc.123', None, {}).
    """
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

    # Loại bỏ khoảng trắng
    s = s.replace(" ", "")

    # Kiểm tra xem có chữ số nào không
    digits = re.findall(r"\d", s)
    if not digits:
        return default

    # Loại bỏ các ký tự lạ không phải chữ số, dấu chấm hoặc dấu phẩy
    # Ví dụ: "abc.123" -> ".123", "QTY: 1,000 PC" -> "1,000"
    s = re.sub(r"[^\d.,]", "", s)
    if not s:
        return default

    active_fmt = (fmt or _GLOBAL_SAP_NUMBER_FORMAT).upper()

    # 2. Xử lý theo cấu hình định dạng rõ ràng
    if active_fmt == "US":
        # Chuẩn US: Dấu phẩy là hàng nghìn, dấu chấm là thập phân
        # '1,000' -> 1000.0, '1,000.50' -> 1000.5, '1.000' -> 1.0
        s = s.replace(",", "")

    elif active_fmt in ("EU", "VN", "DE"):
        # Chuẩn EU / Việt Nam: Dấu chấm là hàng nghìn, dấu phẩy là thập phân
        # '1.000' -> 1000.0, '1.000,50' -> 1000.5, '1,234' -> 1.234
        s = s.replace(".", "").replace(",", ".")

    else:
        # Cấu hình AUTO (Tự động suy luận thông minh)
        if "." in s and "," in s:
            last_dot = s.rfind(".")
            last_comma = s.rfind(",")
            if last_dot > last_comma:
                # '1,234.56' -> dấu chấm là thập phân
                s = s.replace(",", "")
            else:
                # '1.234,56' -> dấu phẩy là thập phân
                s = s.replace(".", "").replace(",", ".")

        elif "," in s:
            parts = s.split(",")
            if len(parts) == 2 and len(parts[1]) != 3:
                # '12,5' hoặc '12,50' -> dấu phẩy là phần thập phân
                s = s.replace(",", ".")
            else:
                # '1,000' hoặc '1,000,000' -> dấu phẩy là ngăn cách hàng nghìn
                s = s.replace(",", "")

        elif "." in s:
            parts = s.split(".")
            if len(parts) > 2:
                # '1.000.000' -> ngăn cách hàng nghìn
                s = s.replace(".", "")
            elif len(parts) == 2:
                p1, p2 = parts[0], parts[1]
                # Nếu chuỗi dạng '1.000' và số lượng nguyên trong SAP (p2 == '000' và p1 nguyên > 0)
                # trong chế độ AUTO, nếu p1 > 0 và p2 == '000', ta giữ nguyên số chấm float(1.000)
                # trừ khi được chỉ định rõ qua cấu hình US hoặc EU.
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
