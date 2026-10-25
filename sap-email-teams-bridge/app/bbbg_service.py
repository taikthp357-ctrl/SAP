# -*- coding: utf-8 -*-
"""Dịch vụ Đối soát Biên bản bàn giao Thành phẩm (BBBG FG Reconciliation Service).

Kế thừa và chuẩn hóa toàn bộ nghiệp vụ từ 261008_BBBG_FG_V2_2_15.xlsm:
1. Đối soát 2 chiều giữa MB51 (Movement 101/102/311/601) và VL06I (Inbound Delivery).
2. Khóa ghép cặp: (Material, Batch, Delivery/Doc).
3. Bóc tách rõ ràng:
   - MATCHED: Số lượng bàn giao khớp số lượng thực nhận, đã GR.
   - MISMATCH_QTY: Có chênh lệch số lượng (diff_qty != 0).
   - MISSING_INBOUND: Có trên MB51 nhưng không tìm thấy Inbound tương ứng.
   - MISSING_MB51: Có Inbound nhưng chưa phát sinh chứng từ MB51.
   - MANUAL_ENTRY: Hàng ngoài hệ thống, tách riêng luồng nhập tay không làm giả kết quả SAP.
4. Xuất file Excel đúng chuẩn sheet 'BBBG_NEW' bảo toàn layout cột A:K, style, viền, chữ ký.
5. Tạo bản xem trước email (HTML Preview) theo Shop (BODY, XMĐ, CƠ KHÍ, EDC), tách rời gửi mail tự động.
"""

import os
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app.number_utils import parse_sap_float, format_qty_display


class BbbgService:
    """Xử lý đối soát và xuất file BBBG Thành phẩm."""

    DEFAULT_TEMPLATE_PATH = Path("G:/SAP/261008_BBBG_FG_V2_2_15.xlsm")
    DEFAULT_EXPORTS_DIR = Path("G:/SAP/sap-email-teams-bridge/exports/bbbg")

    SHOP_RECIPIENTS_CONFIG = {
        "BODY": {
            "name": "Xưởng Hàn (Body)",
            "plant": "3093",
            "sloc": "3000",
            "recipients": ["kho.body@vinfast.vn", "qlkho.body@vinfast.vn"]
        },
        "XMD": {
            "name": "Xưởng Xe máy điện (XMĐ)",
            "plant": "2001",
            "sloc": "2001",
            "recipients": ["kho.xmd@vinfast.vn"]
        },
        "CO_KHI": {
            "name": "Xưởng Cơ khí",
            "plant": "3093",
            "sloc": "3093",
            "recipients": ["kho.cokhi@vinfast.vn"]
        },
        "EDC": {
            "name": "Xưởng Sơn EDC",
            "plant": "4007",
            "sloc": "4004",
            "recipients": ["kho.edc@vinfast.vn"]
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

    @classmethod
    def reconcile(
        cls,
        mb51_rows: List[Dict[str, Any]],
        vl06i_rows: List[Dict[str, Any]],
        manual_rows: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """Thực hiện đối soát 2 chiều giữa MB51 và VL06I.

        mb51_rows: [{matnr/product, batch, mblnr/docNo, qty, postDate, unit, xblnr/refNo, ...}]
        vl06i_rows: [{ibNo, product, batch, qty/actualQty, unit, refNo, status, ...}]
        manual_rows: [{material, batch, qty, note, ...}]
        """
        batch_id = f"BBBG-{datetime.now().strftime('%Y%m%d-%H%M%S')}"

        # 1. Lập chỉ mục VL06I theo (material_norm, batch_norm) và (ibNo_norm)
        vl_map: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        vl_by_ib: Dict[str, List[Dict[str, Any]]] = {}

        for r in vl06i_rows:
            mat = cls.normalize_key(r.get("product") or r.get("matnr") or "")
            bat = cls.normalize_key(r.get("batch") or r.get("charg") or "")
            ib = cls.normalize_key(r.get("ibNo") or r.get("vbeln") or "")

            key = (mat, bat)
            vl_map.setdefault(key, []).append(r)
            if ib:
                vl_by_ib.setdefault(ib, []).append(r)

        matched_items: List[Dict[str, Any]] = []
        mismatch_items: List[Dict[str, Any]] = []
        missing_ib_items: List[Dict[str, Any]] = []
        visited_vl_indices = set()

        # 2. Quét qua từng dòng MB51 để khớp với VL06I
        for mb in mb51_rows:
            mat = cls.normalize_key(mb.get("product") or mb.get("matnr") or "")
            bat = cls.normalize_key(mb.get("batch") or mb.get("charg") or "")
            doc = str(mb.get("docNo") or mb.get("mblnr") or "").strip()
            ref = str(mb.get("refNo") or mb.get("xblnr") or "").strip()
            unit = str(mb.get("unit") or mb.get("meins") or "PC").strip()
            qty_mb = abs(parse_sap_float(mb.get("qty") or mb.get("menge") or 0.0))

            matched_vl = None
            # Thử khớp theo (mat, bat)
            candidates = vl_map.get((mat, bat), [])
            for c in candidates:
                c_id = id(c)
                if c_id not in visited_vl_indices:
                    matched_vl = c
                    visited_vl_indices.add(c_id)
                    break

            # Nếu chưa thấy, thử tìm theo ref (ibNo)
            if not matched_vl and ref:
                ref_norm = cls.normalize_key(ref)
                ib_candidates = vl_by_ib.get(ref_norm, [])
                for c in ib_candidates:
                    c_id = id(c)
                    if c_id not in visited_vl_indices:
                        matched_vl = c
                        visited_vl_indices.add(c_id)
                        break

            if matched_vl:
                qty_vl = abs(parse_sap_float(matched_vl.get("qty") or matched_vl.get("actualQty") or 0.0))
                ib_no = str(matched_vl.get("ibNo") or matched_vl.get("vbeln") or ref).strip()
                diff = qty_mb - qty_vl

                item_record = {
                    "material": mb.get("product") or mb.get("matnr") or mat,
                    "batch": mb.get("batch") or mb.get("charg") or bat,
                    "bill_no": ref,
                    "doc_no": ib_no or doc,
                    "unit": unit,
                    "handover_qty": qty_mb,
                    "received_qty": qty_vl,
                    "diff_qty": diff,
                    "gr_status": "Đã GR" if diff == 0 else f"Lệch ({format_qty_display(diff)})",
                    "mat_doc": doc,
                    "note": "Khớp 100%" if diff == 0 else f"Chênh lệch số lượng MB51={qty_mb} vs VL06I={qty_vl}",
                    "is_manual": False,
                    "status": "MATCHED" if diff == 0 else "MISMATCH_QTY"
                }

                if diff == 0:
                    matched_items.append(item_record)
                else:
                    mismatch_items.append(item_record)
            else:
                # Không tìm thấy Inbound đối ứng
                missing_ib_items.append({
                    "material": mb.get("product") or mb.get("matnr") or mat,
                    "batch": mb.get("batch") or mb.get("charg") or bat,
                    "bill_no": ref,
                    "doc_no": doc,
                    "unit": unit,
                    "handover_qty": qty_mb,
                    "received_qty": 0.0,
                    "diff_qty": qty_mb,
                    "gr_status": "Thiếu Inbound VL06I",
                    "mat_doc": doc,
                    "note": "Không tìm thấy Inbound Delivery trên VL06I tương ứng",
                    "is_manual": False,
                    "status": "MISSING_INBOUND"
                })

        # 3. Quét các Inbound còn lại trong VL06I chưa được khớp (Missing MB51)
        missing_mb_items: List[Dict[str, Any]] = []
        for r in vl06i_rows:
            if id(r) not in visited_vl_indices:
                qty_vl = abs(parse_sap_float(r.get("qty") or r.get("actualQty") or 0.0))
                missing_mb_items.append({
                    "material": r.get("product") or r.get("matnr") or "",
                    "batch": r.get("batch") or r.get("charg") or "",
                    "bill_no": r.get("refNo") or "",
                    "doc_no": r.get("ibNo") or r.get("vbeln") or "",
                    "unit": r.get("unit") or "PC",
                    "handover_qty": 0.0,
                    "received_qty": qty_vl,
                    "diff_qty": -qty_vl,
                    "gr_status": "Chưa GR (Thiếu MB51)",
                    "mat_doc": "",
                    "note": "Inbound tồn tại nhưng chưa có chứng từ MB51",
                    "is_manual": False,
                    "status": "MISSING_MB51"
                })

        # 4. Hàng xuất thủ công ngoài hệ thống (Manual unlisted entries)
        manual_items: List[Dict[str, Any]] = []
        if manual_rows:
            for mr in manual_rows:
                qty_man = abs(parse_sap_float(mr.get("qty") or 0.0))
                manual_items.append({
                    "material": mr.get("material") or mr.get("product") or "MANUAL_ITEM",
                    "batch": mr.get("batch") or "N/A",
                    "bill_no": mr.get("bill_no") or mr.get("refNo") or "",
                    "doc_no": mr.get("doc_no") or "MANUAL_DOC",
                    "unit": mr.get("unit") or "PC",
                    "handover_qty": qty_man,
                    "received_qty": qty_man,
                    "diff_qty": 0.0,
                    "gr_status": "Hàng ngoài hệ thống (Ghi nhận tay)",
                    "mat_doc": "N/A",
                    "note": mr.get("note") or "Mã ngoài hệ thống - không kiểm tra trên SAP",
                    "is_manual": True,
                    "status": "MANUAL_ENTRY"
                })

        all_items = matched_items + mismatch_items + missing_ib_items + missing_mb_items + manual_items

        tot_handover = sum(it["handover_qty"] for it in all_items)
        tot_received = sum(it["received_qty"] for it in all_items)

        result = {
            "batch_id": batch_id,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_items": len(all_items),
            "matched_count": len(matched_items),
            "mismatch_count": len(mismatch_items),
            "missing_inbound_count": len(missing_ib_items),
            "missing_mb51_count": len(missing_mb_items),
            "manual_count": len(manual_items),
            "total_handover_qty": tot_handover,
            "total_received_qty": tot_received,
            "has_mismatch": (len(mismatch_items) > 0 or len(missing_ib_items) > 0 or len(missing_mb_items) > 0),
            "items": all_items
        }

        return result

    @classmethod
    def export_excel(
        cls,
        reconciliation_result: Dict[str, Any],
        template_path: Optional[Path] = None,
        output_dir: Optional[Path] = None,
        meta_header: Optional[Dict[str, str]] = None
    ) -> str:
        """Xuất file Excel BBBG chuẩn theo template BBBG_NEW của 261008_BBBG_FG_V2_2_15.xlsm.

        Đảm bảo giữ đúng layout:
        - Dòng 2: Tiêu đề và Doc ID FVN_SX_KHO03.08
        - Dòng 4-9: Thông tin Bên giao, Bên nhận, Ngày, SĐT, Bộ phận
        - Dòng 12: Tiêu đề cột A:K
        - Dòng 13+: Dữ liệu đối soát
        - Dòng cuối: Tổng cộng kèm công thức SUM
        """
        if output_dir is None:
            output_dir = cls.DEFAULT_EXPORTS_DIR
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        batch_id = reconciliation_result.get("batch_id", "BBBG")
        file_name = f"{batch_id}.xlsx"
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
        ws["G2"] = "Mã số/Doc ID: FVN_SX_KHO03.08\nPhiên bản/Status: 6.0/Đã phê duyệt\nNgày hiệu lực: 14/07/2026"
        ws["G2"].font = font_doc_id
        ws["G2"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

        # 2. Metadata (Dòng 4-9)
        meta = meta_header or {}
        now_date_str = meta.get("date") or datetime.now().strftime("%d/%m/%Y")
        giver_name = meta.get("giver_name") or "Nguyễn Văn Kho"
        giver_dept = meta.get("giver_dept") or "Bộ phận Quản lý Kho FG 3093"
        giver_code = meta.get("giver_code") or "VFE-3093"
        receiver_name = meta.get("receiver_name") or "Lưu Thị Ngọc Mai"
        receiver_dept = meta.get("receiver_dept") or meta.get("shop") or "Xưởng Sản Xuất"

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
            (6, "Mã NV (Employee code):", giver_code, ""),
            (7, "Bộ Phận (Department):", giver_dept, receiver_dept),
            (8, "Chức vụ (Title):", "Nhân viên Kho", "Điều phối viên"),
            (9, "Số điện thoại (Phone):", "0901234567", "")
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

            c_g = ws.cell(row=row, column=7, value=it.get("handover_qty", 0.0))
            c_g.alignment = align_right
            c_g.number_format = "#,##0.00"

            c_h = ws.cell(row=row, column=8, value=it.get("received_qty", 0.0))
            c_h.alignment = align_right
            c_h.number_format = "#,##0.00"

            ws.cell(row=row, column=9, value=it.get("gr_status", "")).alignment = align_center
            ws.cell(row=row, column=10, value=it.get("mat_doc", "")).alignment = align_center
            ws.cell(row=row, column=11, value=it.get("note", "")).alignment = align_left

            # Styling màu nền theo trạng thái
            fill_to_apply = None
            if it.get("is_manual"):
                fill_to_apply = fill_manual
            elif it.get("status") in ["MISMATCH_QTY", "MISSING_INBOUND", "MISSING_MB51"]:
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

        c_tot_g = ws.cell(row=tot_row, column=7, value=f"=SUM(G{start_row}:G{tot_row - 1})")
        c_tot_g.alignment = align_right
        c_tot_g.font = font_bold
        c_tot_g.number_format = "#,##0.00"

        c_tot_h = ws.cell(row=tot_row, column=8, value=f"=SUM(H{start_row}:H{tot_row - 1})")
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

        wb.save(str(output_path))
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
        recipients = recipients_override or shop_info["recipients"]

        tot = reconciliation_result.get("total_items", 0)
        matched = reconciliation_result.get("matched_count", 0)
        mismatch = reconciliation_result.get("mismatch_count", 0)
        missing_ib = reconciliation_result.get("missing_inbound_count", 0)
        missing_mb = reconciliation_result.get("missing_mb51_count", 0)
        manual = reconciliation_result.get("manual_count", 0)

        badge_style = "display:inline-block;padding:4px 8px;border-radius:4px;font-weight:bold;margin-right:8px;font-size:12px;"
        kpi_badges = [
            f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>✅ Khớp 100%: {matched}</span>",
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
            f"<p><strong>Batch ID:</strong> {reconciliation_result.get('batch_id')} | <strong>Thời điểm:</strong> {reconciliation_result.get('created_at')}</p>",
            f"<p><strong>Người nhận dự kiến:</strong> {', '.join(recipients)}</p>",
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
            if it.get("status") in ["MISMATCH_QTY", "MISSING_INBOUND", "MISSING_MB51"]:
                bg = "#ffebe9"
            elif it.get("is_manual"):
                bg = "#fff8c5"

            diff = it.get("diff_qty", 0.0)
            diff_str = f"+{diff}" if diff > 0 else str(diff)

            html_body.append(
                f"<tr style='background:{bg};'>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{idx}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;'><strong>{it.get('material')}</strong></td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{it.get('batch')}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{it.get('doc_no')}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:right;'>{format_qty_display(it.get('handover_qty', 0))}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:right;'>{format_qty_display(it.get('received_qty', 0))}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:right;font-weight:bold;'>{diff_str}</td>"
                f"<td style='padding:5px;border:1px solid #d0d7de;text-align:center;'>{it.get('gr_status')}</td>"
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
            f"⚠️ <strong>Lưu ý:</strong> Thư này chỉ là bản xem trước (Preview). Hệ thống KHÔNG tự động gửi email cho đến khi bạn gõ lệnh xác nhận: "
            f"<code>XÁC NHẬN GỬI BBBG {reconciliation_result.get('batch_id')}</code>.</div>"
            f"</div>"
        )

        return {
            "batch_id": reconciliation_result.get("batch_id"),
            "shop": shop_info["name"],
            "recipients": recipients,
            "html_preview": "".join(html_body),
            "requires_confirmation": True,
            "status": "PREVIEW_GENERATED"
        }
