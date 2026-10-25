# -*- coding: utf-8 -*-
"""Strict SAP numeric boundaries. Quantities remain numeric after parsing.

AUTO accepts unambiguous notation only. For 1.234 / 1,234 select US or EU
explicitly; guessing a factor of 1000 is unsafe for stock decisions.
"""
import math
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Optional, Union

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
    if active == "AUTO":
        if "." in s and "," in s:
            active = "US" if s.rfind(".") > s.rfind(",") else "EU"
        elif s.count(".") > 1:
            active = "EU"
        elif s.count(",") > 1:
            active = "US"
        elif "." in s or "," in s:
            sep = "." if "." in s else ","
            left, right = s.split(sep)
            if sep == ".":
                if len(right) == 3 and 1 <= len(left) <= 3 and int(left or "0") != 0 and right != "000":
                    raise ValueError("Ambiguous quantity: configure US or EU format")
                active = "US"
            else:
                active = "US" if len(right) == 3 else "EU"
        else:
            active = "US"
    if active == "US":
        if not re.fullmatch(r"(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?|\.\d+", s):
            raise ValueError("Invalid US quantity")
        canonical = s.replace(",", "")
    elif active in {"EU", "VN", "DE"}:
        if not re.fullmatch(r"(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?|,\d+", s):
            raise ValueError("Invalid EU quantity")
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


def format_qty_display(val: Union[int, float, str]) -> str:
    """Render canonical quantities without locale-sensitive grouping."""
    try:
        value = parse_sap_decimal(val, fmt="US")
    except ValueError:
        return "N/A"
    rendered = format(value, "f")
    return rendered.rstrip("0").rstrip(".") if "." in rendered else rendered


def parse_sap_datetime_sort_key(d_raw: Any, t_raw: Any) -> tuple[str, str]:
    """Validate real calendar dates/times; unknown values sort after known ones."""
    d_text, t_text = str(d_raw or "").strip(), str(t_raw or "").strip()
    parsed_date = None
    for pattern in ("%Y%m%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            parsed_date = datetime.strptime(d_text, pattern)
            break
        except ValueError:
            continue
    if parsed_date is None:
        return "99999999", "999999"
    parsed_time = None
    for pattern in ("%H:%M:%S", "%H:%M", "%H.%M.%S", "%H.%M", "%H%M%S", "%H%M"):
        try:
            if pattern in ("%H%M%S", "%H%M") and len(t_text) != (6 if pattern == "%H%M%S" else 4):
                continue
            parsed_time = datetime.strptime(t_text, pattern)
            break
        except ValueError:
            continue
    return parsed_date.strftime("%Y%m%d"), parsed_time.strftime("%H%M%S") if parsed_time else "999999"
