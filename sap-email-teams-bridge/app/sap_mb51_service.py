# -*- coding: utf-8 -*-
"""Service tra cứu lịch sử chứng từ vật tư SAP MB51 (LICH_SU_MB51).

Khắc phục triệt để lỗi (617, 'The virtual key is not enabled'):
1. Không dùng các phím ảo thừa (sendVKey 2, 12, 4).
2. Điền trực tiếp chuỗi ngày Posting Date (BUDAT-LOW / BUDAT-HIGH) không qua popup lịch F4.
3. Chuẩn hóa wildcard '*' cho mã vật tư chưa có hậu tố (VD: BIW32011434 -> BIW32011434*).
4. Điều hướng an toàn bằng '/n' trước khi mở MB51, triệt tiêu xung đột trạng thái giữa MB52 và MB51.
5. Trích xuất chính xác 10 trường dữ liệu cốt lõi từ ALV Grid (cntlGRID1) và lọc bỏ dòng tổng rỗng.
"""

import re
from datetime import datetime, timedelta, date
from typing import List, Dict, Any, Optional, Tuple

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


def parse_mb51_date_range(text: str, default_days: int = 7) -> Tuple[datetime, datetime]:
    """Trích xuất khoảng ngày từ câu lệnh người dùng cho MB51.
    
    Hỗ trợ:
    - 'hôm nay', 'hom nay' -> (now, now)
    - 'hôm qua', 'hom qua' -> (yesterday, yesterday)
    - 'N ngày', 'N ngay' (VD: 'trong 15 ngày', '30 ngày') -> (now - (N-1), now)
    - 'ngày 27/10', '27/10/2026', 'từ 25/10 đến 28/10'
    - Mặc định: (now - default_days, now)
    """
    now = datetime.now()
    clean_text = (text or "").lower()

    if "hôm nay" in clean_text or "hom nay" in clean_text:
        return now.replace(hour=0, minute=0, second=0, microsecond=0), now
    if "hôm qua" in clean_text or "hom qua" in clean_text:
        yesterday = now - timedelta(days=1)
        return yesterday.replace(hour=0, minute=0, second=0, microsecond=0), yesterday.replace(hour=23, minute=59, second=59, microsecond=0)

    # Tìm mẫu "N ngày" (VD: "trong 15 ngày", "30 ngày")
    m_days = re.search(r"\b(\d{1,3})\s*ng[aà]y\b", clean_text)
    if m_days:
        days = int(m_days.group(1))
        start_d = now - timedelta(days=max(days - 1, 0))
        return start_d.replace(hour=0, minute=0, second=0, microsecond=0), now

    # Tìm các ngày cụ thể dạng dd/mm hoặc dd/mm/yyyy hoặc dd-mm
    date_matches = re.findall(r"\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b", clean_text)
    parsed_dates = []
    for d_str, m_str, y_str in date_matches:
        try:
            day, month = int(d_str), int(m_str)
            if 1 <= day <= 31 and 1 <= month <= 12:
                year = int(y_str) if y_str else now.year
                if year < 100:
                    year += 2000
                parsed_dates.append(datetime(year, month, day))
        except (ValueError, OverflowError):
            pass

    if len(parsed_dates) >= 2:
        return min(parsed_dates[0], parsed_dates[1]), max(parsed_dates[0], parsed_dates[1])
    elif len(parsed_dates) == 1:
        return parsed_dates[0], parsed_dates[0]

    # Mặc định: Lấy default_days ngày gần nhất
    start_d = now - timedelta(days=int(default_days))
    return start_d.replace(hour=0, minute=0, second=0, microsecond=0), now


class SapMb51Service:
    """Service tra cứu lịch sử chứng từ vật tư SAP MB51 an toàn."""

    def __init__(self, default_plant: str = "3093"):
        self.default_plant = default_plant

    @staticmethod
    def normalize_wildcard_mat(raw_mat: str) -> str:
        """Chuẩn hóa wildcard '*' cho mã vật tư.
        
        - 'BIW32011434' (11 ký tự, 3 chữ cái + 8 số, chưa có hậu tố HA/HB/AA/AB) -> 'BIW32011434*'
        - '32011434' (8 số) -> '*32011434*'
        - 'BIW32011434HA' (đã có hậu tố đầy đủ) -> giữ nguyên
        - Đã có '*' -> giữ nguyên
        """
        mat = (raw_mat or "").strip().upper()
        if not mat or mat.endswith("*"):
            return mat
        if len(mat) == 8 and mat.isdigit():
            return f"*{mat}*"
        if len(mat) == 11 and mat[:3].isalpha() and mat[3:].isdigit():
            return f"{mat}*"
        return mat

    @staticmethod
    def format_sap_date(dt: datetime) -> str:
        """Định dạng ngày chuẩn dd.mm.yyyy cho trường ngày SAP."""
        return dt.strftime("%d.%m.%Y")

    @staticmethod
    def _safe_get_cell(grid: Any, row: int, col_name: str, available_cols: set) -> str:
        if col_name not in available_cols:
            return ""
        try:
            return str(grid.GetCellValue(row, col_name) or "").strip()
        except Exception:
            return ""

    def fetch_movement_history(
        self,
        materials: List[str],
        date_from: datetime,
        date_to: datetime,
        plant: Optional[str] = None,
        sloc: str = "",
        session: Any = None,
        max_rows: int = 100
    ) -> List[Dict[str, Any]]:
        """Tra cứu lịch sử xuất nhập vật tư trong khoảng ngày qua MB51."""
        clean_mats = [self.normalize_wildcard_mat(m) for m in materials if m and m.strip()]
        if not clean_mats:
            return []

        active_session = session
        if active_session is None:
            if pythoncom is not None:
                pythoncom.CoInitialize()
            if win32com is not None:
                try:
                    SapGuiAuto = win32com.client.GetObject("SAPGUI")
                    application = SapGuiAuto.GetScriptingEngine
                    connection = application.Children(0)
                    active_session = connection.Children(0)
                except Exception as e:
                    raise RuntimeError(f"Không thể kết nối phiên SAP GUI: {e}")

        if not active_session:
            raise RuntimeError("Phiên SAP GUI không khả dụng.")

        target_plant = plant or self.default_plant

        try:
            try:
                active_session.findById("wnd[0]").maximize()
            except Exception:
                pass

            # 1. ĐIỀU HƯỚNG AN TOÀN: Thoát sạch màn hình cũ (MB52) bằng /n
            okcd = active_session.findById("wnd[0]/tbar[0]/okcd")
            okcd.text = "/n"
            active_session.findById("wnd[0]").sendVKey(0)

            # Mở MB51: Thử click node F00023, nếu không có thì gọi thẳng /nMB51
            try:
                tree = active_session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
                tree.selectedNode = "F00023"
                tree.doubleClickNode("F00023")
            except Exception:
                okcd = active_session.findById("wnd[0]/tbar[0]/okcd")
                okcd.text = "/nMB51"
                active_session.findById("wnd[0]").sendVKey(0)

            # 2. XÓA SẠCH VÀ ĐIỀN ĐIỀU KIỆN LỌC
            # Xóa ô CHARG an toàn (Không dùng VKey 2 hay 12!)
            try:
                charg_ctl = active_session.findById("wnd[0]/usr/ctxtCHARG-LOW")
                charg_ctl.text = ""
            except Exception:
                pass

            # Điền Material
            if len(clean_mats) == 1:
                mat_low = active_session.findById("wnd[0]/usr/ctxtMATNR-LOW")
                mat_low.text = clean_mats[0]
            else:
                try:
                    active_session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
                    for idx, m_code in enumerate(clean_mats[:12]):
                        cell_id = f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,{idx}]"
                        try:
                            active_session.findById(cell_id).text = m_code
                        except Exception:
                            break
                    active_session.findById("wnd[1]/tbar[0]/btn[8]").press()
                except Exception:
                    mat_low = active_session.findById("wnd[0]/usr/ctxtMATNR-LOW")
                    mat_low.text = clean_mats[0]

            # Điền Plant (3093)
            try:
                p_ctl = active_session.findById("wnd[0]/usr/ctxtWERKS-LOW")
                p_ctl.text = target_plant
            except Exception:
                pass

            # Điền Sloc nếu có
            if sloc:
                try:
                    s_ctl = active_session.findById("wnd[0]/usr/ctxtLGORT-LOW")
                    s_ctl.text = sloc.strip()
                except Exception:
                    pass

            # Điền Posting Date trực tiếp (Không dùng VKey 4 mở popup lịch!)
            str_from = self.format_sap_date(date_from)
            str_to = self.format_sap_date(date_to)
            try:
                b_low = active_session.findById("wnd[0]/usr/ctxtBUDAT-LOW")
                b_low.text = str_from
            except Exception:
                pass

            try:
                b_high = active_session.findById("wnd[0]/usr/ctxtBUDAT-HIGH")
                b_high.text = str_to if str_to != str_from else ""
            except Exception:
                pass

            # 3. THỰC THI (F8 / Execute)
            try:
                active_session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                active_session.findById("wnd[0]").sendVKey(8)

            # Kiểm tra Status Bar
            try:
                sbar = active_session.findById("wnd[0]/sbar")
                if sbar.MessageType in ("E", "W") and any(w in sbar.Text.lower() for w in ["no ", "không ", "không có", "not found"]):
                    return []
            except Exception:
                pass

            # 4. ĐỌC DỮ LIỆU TỪ ALV GRID
            try:
                grid = active_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            except Exception:
                return []

            row_count = int(getattr(grid, "RowCount", 0))
            if row_count <= 0:
                return []

            available_cols = set()
            try:
                col_order = grid.ColumnOrder
                for i in range(col_order.Count):
                    available_cols.add(col_order.ElementAt(i))
            except Exception:
                available_cols = {
                    "BUDAT", "MATNR", "MAKTX", "WERKS", "LGORT",
                    "UMLGO", "BWART", "MENGE", "ERFMG", "MEINS",
                    "EBELN", "USNAM", "CHARG", "MBLNR", "SHKZG"
                }

            records = []
            for r in range(row_count):
                bwart = self._safe_get_cell(grid, r, "BWART", available_cols)
                matnr = self._safe_get_cell(grid, r, "MATNR", available_cols)
                budat = self._safe_get_cell(grid, r, "BUDAT", available_cols)

                # Bỏ qua các dòng tổng cộng rỗng do SAP tự sinh ở cuối
                if not bwart and not budat and not matnr:
                    continue

                raw_qty = (
                    self._safe_get_cell(grid, r, "MENGE", available_cols)
                    or self._safe_get_cell(grid, r, "ERFMG", available_cols)
                )
                shkzg = self._safe_get_cell(grid, r, "SHKZG", available_cols).upper()

                # Định dạng dấu số lượng (+/-)
                try:
                    num_qty = float(raw_qty.replace(".", "").replace(",", ".").replace(" ", "") or 0)
                    if shkzg == "H":  # Xuất / Giảm
                        formatted_qty = f"-{abs(num_qty):,.0f}" if num_qty.is_integer() else f"-{abs(num_qty):,.2f}"
                    elif shkzg == "S":  # Nhập / Tăng
                        formatted_qty = f"+{abs(num_qty):,.0f}" if num_qty.is_integer() else f"+{abs(num_qty):,.2f}"
                    else:
                        formatted_qty = f"{num_qty:,.0f}" if num_qty.is_integer() else f"{num_qty:,.2f}"
                except Exception:
                    formatted_qty = raw_qty

                records.append({
                    "budat": budat,
                    "matnr": matnr,
                    "maktx": self._safe_get_cell(grid, r, "MAKTX", available_cols),
                    "werks": self._safe_get_cell(grid, r, "WERKS", available_cols),
                    "lgort": self._safe_get_cell(grid, r, "LGORT", available_cols),
                    "umlgo": self._safe_get_cell(grid, r, "UMLGO", available_cols),
                    "bwart": bwart,
                    "qty": formatted_qty,
                    "raw_qty": raw_qty,
                    "meins": self._safe_get_cell(grid, r, "MEINS", available_cols),
                    "ebeln": self._safe_get_cell(grid, r, "EBELN", available_cols),
                    "usnam": self._safe_get_cell(grid, r, "USNAM", available_cols),
                    "charg": self._safe_get_cell(grid, r, "CHARG", available_cols),
                    "mblnr": self._safe_get_cell(grid, r, "MBLNR", available_cols),
                })

                if len(records) >= max_rows:
                    break

            return records

        finally:
            # Thoát an toàn về màn hình chính
            try:
                okcd = active_session.findById("wnd[0]/tbar[0]/okcd")
                okcd.text = "/n"
                active_session.findById("wnd[0]").sendVKey(0)
            except Exception:
                pass


def generate_mb51_table_html(records: List[Dict[str, Any]], summary_text: str = "") -> str:
    """Tạo bảng HTML trực quan cho Microsoft Teams hiển thị 10 cột lịch sử MB51."""
    if not records:
        return "<div style='font-family:Segoe UI,sans-serif;color:#64748b;padding:8px;'>Không có dữ liệu giao dịch phát sinh.</div>"

    rows_html = []
    for r in records:
        bwart = str(r.get("bwart", "") or r.get("mvtType", "")).strip()
        # Badges cho loại biến động
        if bwart == "311":
            badge_bwart = "<span style='background:#dbeafe;color:#1e40af;padding:2px 6px;border-radius:4px;font-weight:600;font-size:11px;'>311 (Chuyển kho)</span>"
        elif bwart == "101":
            badge_bwart = "<span style='background:#dcfce7;color:#166534;padding:2px 6px;border-radius:4px;font-weight:600;font-size:11px;'>101 (Nhập GR)</span>"
        elif bwart == "261":
            badge_bwart = "<span style='background:#ffedd5;color:#9a3412;padding:2px 6px;border-radius:4px;font-weight:600;font-size:11px;'>261 (Xuất SX)</span>"
        elif bwart in ("102", "262", "312"):
            badge_bwart = f"<span style='background:#fee2e2;color:#991b1b;padding:2px 6px;border-radius:4px;font-weight:600;font-size:11px;'>{bwart} (Hủy)</span>"
        else:
            badge_bwart = f"<span style='background:#f1f5f9;color:#334155;padding:2px 6px;border-radius:4px;font-weight:600;font-size:11px;'>{bwart}</span>" if bwart else "-"

        # Định dạng màu số lượng (+/-)
        qty_str = str(r.get("qty", "")).strip()
        if qty_str.startswith("+"):
            qty_colored = f"<span style='color:#16a34a;font-weight:bold;'>{qty_str}</span>"
        elif qty_str.startswith("-"):
            qty_colored = f"<span style='color:#dc2626;font-weight:bold;'>{qty_str}</span>"
        else:
            qty_colored = f"<span style='color:#0f172a;'>{qty_str}</span>"

        budat = r.get("budat", "") or r.get("postDate", "")
        matnr = r.get("matnr", "") or r.get("product", "")
        werks = r.get("werks", "") or r.get("plant", "")
        lgort = r.get("lgort", "") or r.get("sloc", "")
        umlgo = r.get("umlgo", "")
        meins = r.get("meins", "") or r.get("unit", "")
        ebeln = r.get("ebeln", "") or r.get("docNo", "")
        usnam = r.get("usnam", "")

        rows_html.append(
            f"<tr style='border-bottom:1px solid #e2e8f0;'>"
            f"<td style='padding:8px 10px;white-space:nowrap;'>{budat}</td>"
            f"<td style='padding:8px 10px;font-weight:600;color:#0369a1;'>{matnr}</td>"
            f"<td style='padding:8px 10px;text-align:center;'>{werks}</td>"
            f"<td style='padding:8px 10px;text-align:center;font-weight:600;'>{lgort}</td>"
            f"<td style='padding:8px 10px;text-align:center;color:#64748b;'>{umlgo or '-'}</td>"
            f"<td style='padding:8px 10px;'>{badge_bwart}</td>"
            f"<td style='padding:8px 10px;text-align:right;'>{qty_colored}</td>"
            f"<td style='padding:8px 10px;text-align:center;'>{meins}</td>"
            f"<td style='padding:8px 10px;color:#475569;'>{ebeln or '-'}</td>"
            f"<td style='padding:8px 10px;color:#475569;font-size:11px;'>{usnam or '-'}</td>"
            f"</tr>"
        )

    summary_block = ""
    if summary_text:
        summary_block = (
            f"<div style='background:#f8fafc;border:1px solid #cbd5e1;border-left:4px solid #0284c7;"
            f"padding:10px 14px;border-radius:6px;margin-bottom:10px;'>"
            f"<div style='font-weight:bold;color:#0369a1;font-size:14px;margin-bottom:2px;'>📊 LỊCH SỬ GIAO DỊCH VẬT TƯ (SAP MB51)</div>"
            f"<div style='font-size:12px;color:#334155;'>{summary_text}</div>"
            f"</div>"
        )

    table_html = (
        f"<div style='font-family:Segoe UI,sans-serif;margin-top:6px;overflow-x:auto;'>"
        f"{summary_block}"
        f"<table style='width:100%;border-collapse:collapse;font-size:12px;text-align:left;background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;'>"
        f"<thead>"
        f"<tr style='background:#f1f5f9;color:#334155;border-bottom:2px solid #cbd5e1;font-size:11px;text-transform:uppercase;'>"
        f"<th style='padding:8px 10px;'>Ngày HT</th>"
        f"<th style='padding:8px 10px;'>Mã NVL</th>"
        f"<th style='padding:8px 10px;text-align:center;'>Plant</th>"
        f"<th style='padding:8px 10px;text-align:center;'>Kho</th>"
        f"<th style='padding:8px 10px;text-align:center;'>Kho đích</th>"
        f"<th style='padding:8px 10px;'>Loại GD</th>"
        f"<th style='padding:8px 10px;text-align:right;'>Số lượng</th>"
        f"<th style='padding:8px 10px;text-align:center;'>ĐVT</th>"
        f"<th style='padding:8px 10px;'>Số PO / Đơn</th>"
        f"<th style='padding:8px 10px;'>User SAP</th>"
        f"</tr>"
        f"</thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        f"</table>"
        f"</div>"
    )
    return table_html

