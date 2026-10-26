# -*- coding: utf-8 -*-
"""Strict SAP numeric boundaries. Quantities remain numeric after parsing.

AUTO accepts unambiguous notation only. For 1.234 / 1,234 select US or EU
explicitly; guessing a factor of 1000 is unsafe for stock decisions.
"""
import math
import re
from datetime import datetime, time as dt_time
from decimal import Decimal, InvalidOperation
from typing import Any, Optional, Union


class NumericParseError(ValueError):
    """Ném ra khi bóc tách số liệu SAP không hợp lệ hoặc gây mơ hồ (Fail-Closed)."""
    pass

_GLOBAL_SAP_NUMBER_FORMAT = "AUTO"


def set_global_sap_number_format(fmt: str):
    global _GLOBAL_SAP_NUMBER_FORMAT
    value = str(fmt).strip().upper()
    if value not in {"AUTO", "US", "EU", "VN", "DE"}:
        raise ValueError("SAP number format must be AUTO, US, EU, VN or DE")
    _GLOBAL_SAP_NUMBER_FORMAT = value


def parse_sap_decimal(value: Any, fmt: Optional[str] = None) -> Decimal:
    """Parse a whole finite quantity or raise ValueError (never strip junk)."""
    if isinstance(value, bool) or value is None:
        raise ValueError("Missing or invalid quantity")
    if isinstance(value, (int, float, Decimal)):
        try:
            result = Decimal(str(value))
        except InvalidOperation as exc:
            raise ValueError("Invalid quantity") from exc
        if not result.is_finite():
            raise ValueError("Non-finite quantity")
        return result
    if not isinstance(value, str):
        raise ValueError("Invalid quantity type")
    s = value.strip().replace("\u00a0", " ")
    sign = ""
    if s.endswith("-"):
        sign, s = "-", s[:-1].strip()
    elif s.startswith(("-", "+")):
        sign, s = s[0], s[1:].strip()
    if not s or not re.fullmatch(r"[0-9., ]+", s):
        raise ValueError("Invalid quantity text")
    if " " in s:
        if not re.fullmatch(r"\d{1,3}(?: \d{3})+(?:[.,]\d+)?", s):
            raise ValueError("Invalid space grouping")
        s = s.replace(" ", "")
    active = str(fmt or _GLOBAL_SAP_NUMBER_FORMAT).upper()
    # Nếu chứa cả hai dấu phân cách thì format đã rõ ràng 100%
    if "." in s and "," in s:
        active = "US" if s.rfind(".") > s.rfind(",") else "EU"
    elif active == "AUTO":
        if s.count(".") > 1:
            active = "EU"
        elif s.count(",") > 1:
            active = "US"
        elif "." in s or "," in s:
            sep = "." if "." in s else ","
            left, right = s.split(sep)
            if left == "0":
                active = "US" if sep == "." else "EU"
            elif sep == ".":
                if len(right) != 3:
                    active = "US"
                elif 1 <= len(left) <= 3 and int(left or "0") != 0 and right != "000":
                    raise NumericParseError("Ambiguous quantity: configure US or EU format")
                else:
                    active = "US"
            else:
                if len(right) != 3:
                    active = "EU"
                elif 1 <= len(left) <= 3 and right != "000":
                    raise NumericParseError("Ambiguous quantity: configure US or EU format")
                else:
                    active = "US"
        else:
            active = "US"
    if active == "US":
        if not re.fullmatch(r"(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?|\.\d+", s):
            raise ValueError("Invalid US quantity")
        canonical = s.replace(",", "")
    elif active in {"EU", "VN", "DE"}:
        if "," not in s and "." in s:
            left, right = s.split(".")
            if left == "0" or len(right) != 3:
                # Dạng số lẻ canonical (ví dụ: 0.001, 100.00, 50.5) trong môi trường EU
                canonical = s
            else:
                if not re.fullmatch(r"(?:\d+|\d{1,3}(?:\.\d{3})+)", s):
                    raise NumericParseError("Invalid EU quantity")
                canonical = s.replace(".", "")
        else:
            if not re.fullmatch(r"(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?|,\d+", s):
                raise NumericParseError("Invalid EU quantity")
            canonical = s.replace(".", "").replace(",", ".")
    else:
        raise ValueError("Unknown SAP number format")
    try:
        result = Decimal(sign + canonical)
    except InvalidOperation as exc:
        raise ValueError("Invalid quantity") from exc
    if not result.is_finite():
        raise ValueError("Non-finite quantity")
    return result


def parse_sap_float(val: Any, default: float = 0.0, fmt: Optional[str] = None) -> float:
    """Compatibility wrapper; decision code should use parse_sap_decimal."""
    try:
        result = float(parse_sap_decimal(val, fmt))
        return result if math.isfinite(result) else default
    except (ValueError, TypeError, OverflowError, InvalidOperation):
        return default


def format_qty_display(val: Union[int, float, str, Decimal]) -> str:
    """Render canonical quantities without locale-sensitive grouping."""
    if isinstance(val, bool) or val is None:
        return "0"
    if isinstance(val, (int, float, Decimal)):
        d_val = float(val)
        if math.isnan(d_val) or math.isinf(d_val):
            return "0"
        if d_val.is_integer():
            return str(int(d_val))
        return f"{d_val:.4f}".rstrip("0").rstrip(".")
    try:
        try:
            value = parse_sap_decimal(val)
        except Exception:
            value = parse_sap_decimal(val, fmt="EU")
        d_val = float(value)
        if d_val.is_integer():
            return str(int(d_val))
        return f"{d_val:.4f}".rstrip("0").rstrip(".")
    except Exception:
        clean_v = str(val or "0").strip()
        return clean_v if clean_v else "0"


def parse_sap_datetime_sort_key(d_str: Any, t_str: Any = "") -> tuple:
    """Phân giải chuỗi ngày giờ từ SAP thành tuple (YYYYMMDD, HHMMSS) để sắp xếp FIFO chuẩn xác.

    Hỗ trợ các định dạng ngày:
    - DD.MM.YYYY (Chuẩn SAP Việt Nam / Đức, ví dụ: 15.10.2026)
    - DD/MM/YYYY (Chuẩn ngày tháng Việt Nam phổ biến, ví dụ: 15/10/2026)
    - YYYY-MM-DD (Chuẩn ISO, ví dụ: 2026-10-15)
    - YYYYMMDD (Chuẩn thô SAP Date, ví dụ: 20261015)
    """
    date_val = "99999999"
    time_val = "999999"

    d_clean = str(d_str or "").strip()
    if d_clean and d_clean.upper() not in ("N/A", "-", "NONE", "NULL"):
        m_eu = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", d_clean)
        m_slash = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", d_clean)
        m_iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", d_clean)
        m_raw8 = re.match(r"^(\d{4})(\d{2})(\d{2})$", d_clean)

        y_i, m_i, d_i = None, None, None
        if m_eu:
            d_i, m_i, y_i = int(m_eu.group(1)), int(m_eu.group(2)), int(m_eu.group(3))
        elif m_slash:
            d_i, m_i, y_i = int(m_slash.group(1)), int(m_slash.group(2)), int(m_slash.group(3))
        elif m_iso:
            y_i, m_i, d_i = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        elif m_raw8:
            y_i, m_i, d_i = int(m_raw8.group(1)), int(m_raw8.group(2)), int(m_raw8.group(3))

        if y_i is not None:
            try:
                dt_obj = datetime(y_i, m_i, d_i)
                date_val = dt_obj.strftime("%Y%m%d")
            except (ValueError, OverflowError):
                return ("INVALID_DATE", "INVALID_TIME")

    t_clean = str(t_str or "").strip()
    if t_clean and t_clean.upper() not in ("N/A", "-", "NONE", "NULL"):
        m_time = re.match(r"^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$", t_clean)
        m_raw6 = re.match(r"^(\d{2})(\d{2})(\d{2})$", t_clean)

        h_i, min_i, sec_i = None, None, 0
        if m_time:
            h_i, min_i = int(m_time.group(1)), int(m_time.group(2))
            sec_i = int(m_time.group(3)) if m_time.group(3) is not None else 0
        elif m_raw6:
            h_i, min_i, sec_i = int(m_raw6.group(1)), int(m_raw6.group(2)), int(m_raw6.group(3))

        if h_i is not None:
            try:
                tm_obj = dt_time(h_i, min_i, sec_i)
                time_val = tm_obj.strftime("%H%M%S")
            except (ValueError, OverflowError):
                return ("INVALID_DATE", "INVALID_TIME")

    return (date_val, time_val)
