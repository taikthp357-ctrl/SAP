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


def parse_sap_datetime_sort_key(d_raw: Any, t_raw: Any) -> tuple[str, str]:
    """Phân giải Ngày và Giờ GR từ SAP thành canonical tuple (YYYYMMDD, HHMMSS) để sắp xếp FIFO chuẩn xác.

    Hỗ trợ định dạng:
    - DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY
    - YYYY-MM-DD, YYYY.MM.DD
    - YYYYMMDD
    - Giờ: HH:MM:SS, HH:MM, HHMMSS
    Đặc biệt (R04): Ngày/giờ bị thiếu, rỗng hoặc N/A sẽ được xếp xuống cuối cùng ('99999999', '999999'),
    tuyệt đối không bao giờ bị xếp lên đầu trong sắp xếp FIFO tăng dần.
    """
    d_clean = str(d_raw or "").strip().upper()
    t_clean = str(t_raw or "").strip().upper()

    if not d_clean or d_clean in ["-", "N/A", "NONE", "NULL"]:
        return ("99999999", "999999")

    # 1. Phân giải ngày
    m_iso = re.match(r"^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$", d_clean)
    if m_iso:
        yyyy, mm, dd = m_iso.groups()
        d_sort = f"{yyyy}{mm.zfill(2)}{dd.zfill(2)}"
    else:
        m_eu = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", d_clean)
        if m_eu:
            dd, mm, yyyy = m_eu.groups()
            d_sort = f"{yyyy}{mm.zfill(2)}{dd.zfill(2)}"
        elif len(d_clean) == 8 and d_clean.isdigit():
            d_sort = d_clean
        else:
            return ("99999999", "999999")

    # 2. Phân giải giờ
    t_clean_no_colon = t_clean.replace(":", "").replace(".", "")
    if len(t_clean_no_colon) == 6 and t_clean_no_colon.isdigit():
        t_sort = t_clean_no_colon
    elif len(t_clean_no_colon) == 4 and t_clean_no_colon.isdigit():
        t_sort = t_clean_no_colon + "00"
    else:
        t_parts = re.split(r"[:.]", t_clean)
        if len(t_parts) == 3 and all(p.isdigit() for p in t_parts):
            t_sort = "".join(p.zfill(2) for p in t_parts)
        elif len(t_parts) == 2 and all(p.isdigit() for p in t_parts):
            t_sort = "".join(p.zfill(2) for p in t_parts) + "00"
        else:
            t_sort = "000000"

    return (d_sort, t_sort)
