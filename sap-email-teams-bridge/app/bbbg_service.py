# -*- coding: utf-8 -*-
"""Preview-only BBBG quantity reconciliation and spreadsheet export.

Matching preserves inbound/material/batch/unit/scope evidence. Quantity equality
is not proof of a posted goods receipt or an approved handover document.
"""

import os
import re
import math
from collections import defaultdict
from html import escape
from uuid import uuid4
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app.number_utils import parse_sap_float, format_qty_display


class BbbgService:
    """Xử lý đối soát và xuất file BBBG Thành phẩm."""

    DEFAULT_TEMPLATE_PATH = None
    DEFAULT_EXPORTS_DIR = Path(__file__).resolve().parent.parent / "exports" / "bbbg"

    SHOP_RECIPIENTS_CONFIG = {
        "BODY": {
            "name": "Xưởng Hàn (Body)",
            "plant": "3093",
            "sloc": "3000",
            "recipients": []
        },
        "XMD": {
            "name": "Xưởng Xe máy điện (XMĐ)",
            "plant": "2001",
            "sloc": "2001",
            "recipients": []
        },
        "CO_KHI": {
            "name": "Xưởng Cơ khí",
            "plant": "3093",
            "sloc": "3093",
            "recipients": []
        },
        "EDC": {
            "name": "Xưởng Sơn EDC",
            "plant": "4007",
            "sloc": "4004",
            "recipients": []
        }
    }

    @staticmethod
    def normalize_key(val: Any) -> str:
        """Chuẩn hóa khóa chuỗi để so khớp."""
        if val is None:
            return ""
        s = str(val).strip().upper()
        # Bỏ các số 0 vô nghĩa ở đầu nếu là số thuần
        if s.isdigit():
            return s.lstrip("0") or "0"
        return s

    @staticmethod
    def _value(row: Dict[str, Any], *keys: str) -> Any:
        """Keep a real zero; never substitute a different quantity for it."""
        for key in keys:
            if key in row and row[key] is not None and str(row[key]).strip() != "":
                return row[key]
        return ""

    @classmethod
    def reconcile(
        cls,
        mb51_rows: List[Dict[str, Any]],
        vl06i_rows: List[Dict[str, Any]],
        manual_rows: Optional[List[Dict[str, Any]]] = None,
        *, source_complete: bool = False,
    ) -> Dict[str, Any]:
        """Compare quantity evidence, never infer a confirmed GR from equality.

        Match only explicit inbound + material + batch + unit + plant + sloc
        and, when available on both sides, inbound item. Blank scope does not
        match a nonblank scope. RefNo is accepted as an inbound link only when
        its value equals the actual inbound number; arbitrary bill references
        are not joined by material alone. 101/102 rows are netted; other known
        movement types cannot prove receipt. Missing source completeness keeps
        the entire result a preview requiring verification.
        """
        groups = {"MB51": defaultdict(list), "VL06I": defaultdict(list)}
        review = []
        duplicate_count = 0
        seen = {}
        ambiguous = set()

        for source, rows in (("MB51", mb51_rows), ("VL06I", vl06i_rows)):
            for ordinal, raw in enumerate(rows):
                row = dict(raw)
                mat = cls.normalize_key(cls._value(row, "product", "matnr", "material"))
                bat = cls.normalize_key(cls._value(row, "batch", "charg"))
                unit = str(cls._value(row, "unit", "meins", "uom")).strip().upper()
                ib = cls.normalize_key(cls._value(row, "ibNo", "inboundNo", "inbound_no", "deliveryNo", "vbeln", *( ["refNo", "xblnr"] if source == "MB51" else [])))
                plant = cls.normalize_key(cls._value(row, "plant", "werks"))
                sloc = cls.normalize_key(cls._value(row, "sloc", "lgort", "storageLocation"))
                ib_item = cls.normalize_key(cls._value(row, "ibItem", "inboundItem", "deliveryItem", *( ["item", "posnr"] if source == "VL06I" else [])))
                key = (ib, mat, bat, unit, plant, sloc, ib_item)
                qty_raw = cls._value(row, *( ["qty", "menge"] if source == "MB51" else ["actualQty", "qty", "menge"] ))
                qty = parse_sap_float(qty_raw, default=float("nan"))
                mvt = cls.normalize_key(cls._value(row, "mvtType", "movementType", "bwart"))
                reason = ""
                if not ib or not mat or not unit:
                    reason = "Thiếu khóa IB, vật tư hoặc đơn vị; chưa đủ bằng chứng ghép dòng."
                elif not math.isfinite(qty):
                    reason = "Số lượng thiếu hoặc không hợp lệ."
                elif source == "MB51" and mvt and mvt not in {"101", "102"}:
                    reason = f"Movement {mvt} không thuộc phép đối soát nhận hàng 101/102."
                elif source == "VL06I" and qty < 0:
                    reason = "Số lượng Inbound âm cần kiểm tra."
                elif source == "MB51" and mvt == "101" and qty < 0:
                    reason = "Số lượng 101 âm không được tự chuyển thành số dương."
                if source == "MB51" and mvt == "102" and math.isfinite(qty):
                    qty = -abs(qty)
                doc = str(cls._value(row, "docNo", "mblnr")).strip()
                year = str(cls._value(row, "docYear", "year", "mjahr")).strip()
                doc_item = cls.normalize_key(cls._value(row, "docItem", "zeile", "item"))
                identity = (source, doc, year, doc_item) if source == "MB51" and doc and year and doc_item else ((source, ib, ib_item) if source == "VL06I" and ib and ib_item else None)
                signature = (key, qty, mvt)
                if identity is not None:
                    if identity in seen:
                        if seen[identity] == signature:
                            duplicate_count += 1
                            continue
                        ambiguous.add(seen[identity][0])
                        reason = "Một chứng từ/năm/dòng có dữ liệu mâu thuẫn; cần kiểm tra nguồn."
                    else:
                        seen[identity] = signature
                elif any(x["doc"] == doc and x["year"] == year for x in groups[source][key]) and (doc or source == "VL06I"):
                    ambiguous.add(key)
                    reason = "Nhiều dòng cùng khóa nhưng thiếu số dòng chứng từ; không tự cộng trùng."
                record = {"raw": row, "qty": qty if math.isfinite(qty) else None, "doc": doc, "year": year, "mvt": mvt, "key": key}
                if reason:
                    review.append(cls._review_item(record, source, reason))
                else:
                    groups[source][key].append(record)

        items = []
        for key in sorted(set(groups["MB51"]) | set(groups["VL06I"])):
            mb, vl = groups["MB51"].get(key, []), groups["VL06I"].get(key, [])
            if not mb and not vl:
                continue
            mb_qty = sum(r["qty"] for r in mb)
            vl_qty = sum(r["qty"] for r in vl)
            diff = mb_qty - vl_qty
            if key in ambiguous:
                status, note = "REVIEW_REQUIRED", "Khóa chứng từ có dòng trùng/mâu thuẫn; chưa xác nhận số lượng."
            elif mb and vl:
                status = "MATCHED" if math.isclose(diff, 0.0, rel_tol=0.0, abs_tol=1e-9) else "MISMATCH_QTY"
                note = "Số lượng khớp; chưa xác nhận GR trên SAP." if status == "MATCHED" else "Số lượng thuần MB51 và VL06I có chênh lệch."
            elif mb:
                status, note = "MISSING_INBOUND", "Không có dòng Inbound cùng đầy đủ khóa trong kết quả nguồn."
            else:
                status, note = "MISSING_MB51", "Không có dòng MB51 cùng đầy đủ khóa trong kết quả nguồn; không kết luận chưa GR."
            representative = (mb or vl)[0]["raw"]
            items.append({
                "material": cls._value(representative, "product", "matnr", "material"),
                "batch": cls._value(representative, "batch", "charg"),
                "bill_no": cls._value(representative, "refNo", "xblnr"),
                "doc_no": key[0], "inbound_item": key[6], "plant": key[4], "sloc": key[5],
                "unit": key[3], "handover_qty": mb_qty if key not in ambiguous else None,
                "received_qty": vl_qty if key not in ambiguous else None,
                "diff_qty": diff if key not in ambiguous else None, "gr_status": "Chưa xác nhận GR", "gr_confirmed": False,
                "mat_doc": ", ".join(dict.fromkeys(f"{r['doc']}/{r['year']}" if r["year"] else r["doc"] for r in mb if r["doc"])),
                "note": note, "is_manual": False, "status": status,
                "source_row_count": {"MB51": len(mb), "VL06I": len(vl)},
            })
        items.extend(review)
        for row in manual_rows or []:
            qty = parse_sap_float(row.get("qty"), default=float("nan"))
            valid = math.isfinite(qty) and qty >= 0
            items.append({
                "material": cls._value(row, "material", "product"), "batch": row.get("batch", ""),
                "bill_no": cls._value(row, "bill_no", "refNo"), "doc_no": row.get("doc_no", ""),
                "unit": row.get("unit", ""), "handover_qty": qty if valid else None,
                "received_qty": None, "diff_qty": None, "gr_status": "Ghi nhận tay; chưa xác nhận nhận hàng",
                "gr_confirmed": False, "mat_doc": "", "note": row.get("note", "Hàng ngoài hệ thống"),
                "is_manual": True, "status": "MANUAL_ENTRY" if valid else "REVIEW_REQUIRED",
            })
        counts = {name: sum(x["status"] == name for x in items) for name in ("MATCHED", "MISMATCH_QTY", "MISSING_INBOUND", "MISSING_MB51", "MANUAL_ENTRY", "REVIEW_REQUIRED")}
        totals = {}
        for item in items:
            unit = item["unit"] or "UNKNOWN"
            total = totals.setdefault(unit, {"handover_qty": 0.0, "received_qty": 0.0})
            for field in total:
                if item.get(field) is not None:
                    total[field] += item[field]
        single = next(iter(totals.values())) if len(totals) == 1 and "UNKNOWN" not in totals and not counts["REVIEW_REQUIRED"] else None
        return {
            "batch_id": f"BBBG-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:12]}",
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_items": len(items), "matched_count": counts["MATCHED"],
            "mismatch_count": counts["MISMATCH_QTY"], "missing_inbound_count": counts["MISSING_INBOUND"],
            "missing_mb51_count": counts["MISSING_MB51"], "manual_count": counts["MANUAL_ENTRY"],
            "review_required_count": counts["REVIEW_REQUIRED"], "duplicate_count": duplicate_count,
            "total_handover_qty": single["handover_qty"] if single else None,
            "total_received_qty": single["received_qty"] if single else None,
            "totals_by_unit": totals, "source_complete": source_complete,
            "requires_review": not source_complete or any(x["status"] != "MATCHED" for x in items),
            "has_mismatch": any(x["status"] not in {"MATCHED", "MANUAL_ENTRY"} for x in items),
            "status": "PREVIEW", "items": items,
        }

    @classmethod
    def _review_item(cls, record, source, reason):
        row, key, qty = record["raw"], record["key"], record["qty"]
        return {
            "material": cls._value(row, "product", "matnr", "material"),
            "batch": cls._value(row, "batch", "charg"), "bill_no": cls._value(row, "refNo", "xblnr"),
            "doc_no": key[0], "unit": key[3], "plant": key[4], "sloc": key[5],
            "handover_qty": None, "received_qty": None, "source_quantity": qty,
            "source": source, "diff_qty": None,
            "gr_status": "Chưa xác nhận GR", "gr_confirmed": False, "mat_doc": record["doc"],
            "note": reason, "is_manual": False, "status": "REVIEW_REQUIRED",
        }

    @classmethod
    def export_excel(
        cls,
        reconciliation_result: Dict[str, Any],
        template_path: Optional[Path] = None,
        output_dir: Optional[Path] = None,
        meta_header: Optional[Dict[str, str]] = None
    ) -> str:
        """Xuất bản preview Excel BBBG_NEW; không giả định biểu mẫu đã phê duyệt.

        Đảm bảo giữ đúng layout:
        - Dòng 2: Tiêu đề và Doc ID FVN_SX_KHO03.08
        - Dòng 4-9: Thông tin Bên giao, Bên nhận, Ngày, SĐT, Bộ phận
        - Dòng 12: Tiêu đề cột A:K
        - Dòng 13+: Dữ liệu đối soát
        - Dòng cuối: Tổng cộng kèm công thức SUM
        """
        if template_path is not None:
            raise ValueError("Chưa hỗ trợ áp dụng template ngoài; bỏ template_path để xuất mẫu preview chuẩn của bridge.")
        if output_dir is None:
            output_dir = cls.DEFAULT_EXPORTS_DIR
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        batch_id = re.sub(r"[^A-Za-z0-9_-]", "_", str(reconciliation_result.get("batch_id", "BBBG")))[:80]
        file_name = f"{batch_id}-{uuid4().hex[:12]}.xlsx"
        output_path = output_dir / file_name

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BBBG_NEW"

        # Thiết lập font và border
        font_header_title = Font(name="Segoe UI", size=14, bold=True, color="002060")
        font_doc_id = Font(name="Segoe UI", size=9, italic=True)
        font_bold = Font(name="Segoe UI", size=10, bold=True)
        font_regular = Font(name="Segoe UI", size=10)
        font_th = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

        fill_th_blue = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        fill_th_green = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        fill_zebra = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
        fill_warn = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
        fill_manual = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        thin_side = Side(style="thin", color="D0D7DE")
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        border_double_bottom = Border(left=thin_side, right=thin_side, top=thin_side, bottom=Side(style="double", color="000000"))

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        # 1. Header (Dòng 2)
        ws.merge_cells("C2:F2")
        ws["C2"] = "BIÊN BẢN GIAO NHẬN\nTRANSACTION RECORD FORM"
        ws["C2"].font = font_header_title
        ws["C2"].alignment = align_center

        ws.merge_cells("G2:K2")
        ws["G2"] = "BẢN XEM TRƯỚC / DRAFT\nChưa ký nhận hoặc xác nhận GR"
        ws["G2"].font = font_doc_id
        ws["G2"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

        # 2. Metadata (Dòng 4-9)
        meta = meta_header or {}
        now_date_str = meta.get("date") or datetime.now().strftime("%d/%m/%Y")
        giver_name = meta.get("giver_name") or ""
        giver_dept = meta.get("giver_dept") or ""
        giver_code = meta.get("giver_code") or ""
        receiver_name = meta.get("receiver_name") or ""
        receiver_dept = meta.get("receiver_dept") or meta.get("shop") or ""

        ws["A4"] = "Ngày (Date):"
        ws["B4"] = now_date_str
        ws.merge_cells("C4:F4")
        ws["C4"] = "BÊN BÀN GIAO (Hand over party)"
        ws.merge_cells("G4:J4")
        ws["G4"] = "BÊN NHẬN BÀN GIAO (Take over party)"
        ws["K4"] = "Chữ ký người nhận"

        ws["C4"].font = font_bold
        ws["G4"].font = font_bold
        ws["K4"].font = font_bold
        ws["C4"].alignment = align_center
        ws["G4"].alignment = align_center
        ws["K4"].alignment = align_center

        fields = [
            (5, "Họ và Tên (Full Name):", giver_name, receiver_name),
            (6, "Mã NV (Employee code):", giver_code, meta.get("receiver_code", "")),
            (7, "Bộ Phận (Department):", giver_dept, receiver_dept),
            (8, "Chức vụ (Title):", meta.get("giver_title", ""), meta.get("receiver_title", "")),
            (9, "Số điện thoại (Phone):", meta.get("giver_phone", ""), meta.get("receiver_phone", ""))
        ]
        for r_idx, lbl, g_val, rec_val in fields:
            ws.cell(row=r_idx, column=1, value=lbl).font = font_regular
            ws.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=6)
            ws.cell(row=r_idx, column=3, value=g_val).font = font_regular
            ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=10)
            ws.cell(row=r_idx, column=7, value=rec_val).font = font_regular

        # 3. Tiêu đề bảng (Dòng 11-12)
        ws.merge_cells("A11:H11")
        ws["A11"] = "Nhân viên làm shipment điền (Shipment staff fill in)"
        ws["A11"].font = font_bold
        ws.merge_cells("I11:J11")
        ws["I11"] = "Người GR điền (GR person fills in)"
        ws["I11"].font = font_bold
        ws["K11"] = "Ghi chú (Notes)"
        ws["K11"].font = font_bold

        headers = [
            ("A12", "TT\n(No)"),
            ("B12", "Mã Hàng Hóa\n(Part No.)"),
            ("C12", "Tên/ Batch hoặc Serial\n(Name/Batch)"),
            ("D12", "Vận đơn / Cont\n(Bill/Cont No.)"),
            ("E12", "Số Inbound/Doc xuất\n(Inbound/Doc)"),
            ("F12", "Đơn vị\n(Unit)"),
            ("G12", "SL bàn giao\n(Qty handover)"),
            ("H12", "SL thực nhận\n(Qty received)"),
            ("I12", "Tình trạng GR\n(GR status)"),
            ("J12", "Số Material Doc\n(Mat Doc)"),
            ("K12", "Ghi chú chi tiết\n(Notes)")
        ]
        for cell_coord, text in headers:
            cell = ws[cell_coord]
            cell.value = text
            cell.font = font_th
            cell.fill = fill_th_blue if cell_coord[0] <= "H" else fill_th_green
            cell.alignment = align_center
            cell.border = border_all

        # 4. Ghi dữ liệu từng dòng (Dòng 13+)
        items = reconciliation_result.get("items", [])
        start_row = 13
        current_row = start_row

        for idx, it in enumerate(items, 1):
            row = current_row
            ws.cell(row=row, column=1, value=idx).alignment = align_center
            ws.cell(row=row, column=2, value=it.get("material", "")).alignment = align_left
            ws.cell(row=row, column=3, value=it.get("batch", "")).alignment = align_center
            ws.cell(row=row, column=4, value=it.get("bill_no", "")).alignment = align_center
            ws.cell(row=row, column=5, value=it.get("doc_no", "")).alignment = align_center
            ws.cell(row=row, column=6, value=it.get("unit", "PC")).alignment = align_center

            h_qty = it.get("handover_qty")
            r_qty = it.get("received_qty")
            c_g = ws.cell(row=row, column=7, value=h_qty if h_qty is not None else "N/A")
            c_g.alignment = align_right
            if isinstance(h_qty, (int, float)):
                c_g.number_format = "#,##0.00"

            c_h = ws.cell(row=row, column=8, value=r_qty if r_qty is not None else "N/A")
            c_h.alignment = align_right
            if isinstance(r_qty, (int, float)):
                c_h.number_format = "#,##0.00"

            ws.cell(row=row, column=9, value=it.get("gr_status", "")).alignment = align_center
            ws.cell(row=row, column=10, value=it.get("mat_doc", "")).alignment = align_center
            ws.cell(row=row, column=11, value=it.get("note", "")).alignment = align_left

            # Styling màu nền theo trạng thái
            fill_to_apply = None
            if it.get("is_manual"):
                fill_to_apply = fill_manual
            elif it.get("status") in ["MISMATCH_QTY", "MISSING_INBOUND", "MISSING_MB51", "REVIEW_REQUIRED"]:
                fill_to_apply = fill_warn
            elif idx % 2 == 0:
                fill_to_apply = fill_zebra

            for col in range(1, 12):
                c = ws.cell(row=row, column=col)
                c.font = font_regular
                c.border = border_all
                if fill_to_apply:
                    c.fill = fill_to_apply

            current_row += 1

        # 5. Dòng tổng cộng (Total row)
        tot_row = current_row
        ws.cell(row=tot_row, column=1, value="TỔNG CỘNG").alignment = align_center
        ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=6)

        same_unit = (not items or (len({it.get("unit") for it in items}) == 1 and all(it.get("unit") for it in items)))
        totals_known = same_unit and not any(it.get("status") == "REVIEW_REQUIRED" for it in items)
        c_tot_g = ws.cell(row=tot_row, column=7, value=(f"=SUM(G{start_row}:G{tot_row - 1})" if items else 0) if totals_known else "N/A: cần đối soát")
        c_tot_g.alignment = align_right
        c_tot_g.font = font_bold
        c_tot_g.number_format = "#,##0.00"

        c_tot_h = ws.cell(row=tot_row, column=8, value=(f"=SUM(H{start_row}:H{tot_row - 1})" if items else 0) if totals_known else "N/A: cần đối soát")
        c_tot_h.alignment = align_right
        c_tot_h.font = font_bold
        c_tot_h.number_format = "#,##0.00"

        ws.cell(row=tot_row, column=9, value="")
        ws.cell(row=tot_row, column=10, value="")
        ws.cell(row=tot_row, column=11, value="")

        for col in range(1, 12):
            c = ws.cell(row=tot_row, column=col)
            c.font = font_bold
            c.border = border_double_bottom

        # 6. Tự động điều chỉnh độ rộng cột
        col_widths = {
            "A": 6, "B": 18, "C": 16, "D": 18, "E": 18,
            "F": 8, "G": 14, "H": 14, "I": 20, "J": 16, "K": 32
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # 7. Thiết lập vùng in (Print Area)
        ws.print_area = f"A1:K{tot_row + 2}"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4

        # User/SAP text is data, never an Excel formula. Only our total formulas remain formulas.
        for row in ws.iter_rows():
            for cell in row:
                if cell.row != tot_row and isinstance(cell.value, str) and cell.value.startswith(("=", "+", "-", "@")):
                    cell.data_type = "s"
        wb.save(str(output_path))
        wb.close()
        return str(output_path.resolve())

    @classmethod
    def generate_mail_preview(
        cls,
        reconciliation_result: Dict[str, Any],
        shop_key: str = "BODY",
        recipients_override: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """Tạo bản xem trước Email HTML kèm danh sách người nhận và nút xác nhận."""
        shop_info = cls.SHOP_RECIPIENTS_CONFIG.get(shop_key.upper(), cls.SHOP_RECIPIENTS_CONFIG["BODY"])
        recipients = list(recipients_override) if recipients_override is not None else []

        tot = reconciliation_result.get("total_items", 0)
        matched = reconciliation_result.get("matched_count", 0)
        mismatch = reconciliation_result.get("mismatch_count", 0)
        missing_ib = reconciliation_result.get("missing_inbound_count", 0)
        missing_mb = reconciliation_result.get("missing_mb51_count", 0)
        manual = reconciliation_result.get("manual_count", 0)

        badge_style = "display:inline-block;padding:4px 8px;border-radius:4px;font-weight:bold;margin-right:8px;font-size:12px;"
        kpi_badges = [
            f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>Số lượng khớp: {matched}</span>",
        ]
        if mismatch > 0:
            kpi_badges.append(f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>⚠️ Lệch SL: {mismatch}</span>")
        if missing_ib > 0:
            kpi_badges.append(f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>🔍 Thiếu Inbound: {missing_ib}</span>")
        if missing_mb > 0:
            kpi_badges.append(f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>🔍 Thiếu MB51: {missing_mb}</span>")
        if manual > 0:
            kpi_badges.append(f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>📝 Ngoài hệ thống: {manual}</span>")

        html_body = [
            f"<div style='font-family:Segoe UI,sans-serif;font-size:13px;line-height:1.5;color:#24292f;'>",
            f"<h3 style='color:#1f497d;margin-bottom:8px;'>📦 XEM TRƯỚC BIÊN BẢN BÀN GIAO THÀNH PHẨM (BBBG FG) - {shop_info['name']}</h3>",
            f"<p><strong>Batch ID:</strong> {escape(str(reconciliation_result.get('batch_id', '')))} | <strong>Thời điểm:</strong> {escape(str(reconciliation_result.get('created_at', '')))}</p>",
            f"<p><strong>Người nhận dự kiến:</strong> {escape(', '.join(recipients)) or 'Chưa cấu hình'}</p>",
            f"<div style='margin-bottom:12px;'>{''.join(kpi_badges)}</div>",
            f"<table style='border-collapse:collapse;width:100%;font-size:12px;margin-bottom:14px;'>",
            f"<thead><tr style='background:#1f497d;color:#ffffff;'>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>STT</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>Mã hàng</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>Số lô</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>Số Inbound/Doc</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>SL Bàn giao</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>SL Nhận</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>Chênh lệch</th>",
            f"<th style='padding:6px;border:1px solid #d0d7de;'>Trạng thái</th>",
            f"</tr></thead><tbody>"
        ]

        items = reconciliation_result.get("items", [])
        for idx, it in enumerate(items[:15], 1):
            bg = "#fff" if idx % 2 == 1 else "#f6f8fa"
            if it.get("status") in ["MISMATCH_QTY", "MISSING_INBOUND", "MISSING_MB51", "REVIEW_REQUIRED"]:
                bg = "#ffebe9"
            elif it.get("is_manual"):
                bg = "#fff8c5"

            diff = it.get("diff_qty", 0.0)
            diff_str = "N/A" if diff is None else (f"+{diff}" if diff > 0 else str(diff))

            html_body.append(
                f"<tr style='background:{bg};'>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{idx}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;'><strong>{escape(str(it.get('material', '')))}</strong></td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{escape(str(it.get('batch', '')))}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{escape(str(it.get('doc_no', '')))}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:right;'>{format_qty_display(it['handover_qty']) if it.get('handover_qty') is not None else 'N/A'}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:right;'>{format_qty_display(it['received_qty']) if it.get('received_qty') is not None else 'N/A'}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:right;font-weight:bold;'>{diff_str}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{escape(str(it.get('gr_status', '')))}</td>"
                f"</tr>"
            )

        if len(items) > 15:
            html_body.append(
                f"<tr><td colspan='8' style='padding:6px;text-align:center;color:#656d76;background:#f6f8fa;'>"
                f"<em>... Còn {len(items) - 15} dòng nữa trong file Excel đầy đủ.</em></td></tr>"
            )

        html_body.append("</tbody></table>")
        html_body.append(
            f"<div style='background:#f6f8fa;border:1px solid #d0d7de;padding:10px 12px;border-radius:4px;'>"
            f"<strong>Lưu ý:</strong> Đây là bản xem trước; số lượng khớp không xác nhận đã GR. "
            f"Chức năng gửi email/đính kèm chưa được tích hợp. Cần kiểm tra nguồn và người nhận trước khi gửi thủ công.</div>"
            f"</div>"
        )

        return {
            "batch_id": reconciliation_result.get("batch_id"),
            "shop": shop_info["name"],
            "recipients": recipients,
            "html_preview": "".join(html_body),
            "requires_confirmation": True,
            "send_supported": False,
            "status": "PREVIEW_GENERATED"
        }
