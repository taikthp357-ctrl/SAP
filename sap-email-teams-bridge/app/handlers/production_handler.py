# -*- coding: utf-8 -*-
"""Handler chuyên trách quản lý sản xuất: BOM định mức (CS03), Outbound OB (VL06O), RC65 và Queue EWM."""
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler
from app.ob_service import ObService


class ProductionHandler(BaseHandler):
    """Xử lý nhóm lệnh Kế hoạch & Điều độ Sản xuất:
    - CHECK_BOM, BOM, CS03 (Định mức linh kiện)
    - CHECK_OB, VL06O (Đơn Outbound mở)
    - CHECK_FG_OB, DIEN_OB (Điều độ Thành Phẩm & OB)
    - DIEU_PHOI_OB, RC65, XAC_NHAN_OB (Kiểm soát cổng xuất nguyên lô)
    - CHECK_QUEUE, QUEUE (Lỗi kẹt Queue EWM /SCWM/MON)
    - CHECK_RES, MB23, MB25 (Phiếu lĩnh Reservation)
    """

    def handle(self, ctx: RequestContext) -> CommandResult:
        cmd = (ctx.command or "").upper()
        records: List[Dict[str, Any]] = []
        headers: Dict[str, str] = {}
        summary = ""
        custom_html = None
        result_status = "SUCCEEDED"
        audit_target = ""
        audit_message = ""

        raw_mats = ctx.raw_materials or []
        resolved_mats = ctx.resolved_materials or []
        mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""

        # -------------------------------------------------------------
        # 1. TRA CỨU PHIẾU LĨNH RESERVATION (MB23/MB25)
        # -------------------------------------------------------------
        if cmd in ["CHECK_RES", "MB23", "MB25"]:
            req_id_display = ctx.meta.get("requestId") or ctx.request_id
            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#fff8c5;border:1px solid #d4a72c;"
                f"border-left:4px solid #9a6700;padding:12px 14px;border-radius:4px;color:#24292f;'>"
                f"<div style='font-weight:bold;color:#9a6700;font-size:14px;margin-bottom:6px;'>"
                f"📋 Kiểm tra Phiếu Lĩnh / Reservation (MB23/MB25)</div>"
                f"<div style='font-size:12px;line-height:1.6;'>"
                f"• <strong>Mã yêu cầu:</strong> {req_id_display}<br>"
                f"• <strong>Trạng thái:</strong> Hệ thống đã tiếp nhận yêu cầu kiểm tra Reservation.<br>"
                f"• <strong>Hướng dẫn:</strong> Vui lòng cung cấp số phiếu Reservation (RSNUM 10 số) để tra cứu chi tiết."
                f"</div></div>"
            )
            summary = "Tra cứu thông tin phiếu lĩnh Reservation (MB23/MB25)"
            audit_target = req_id_display
            audit_message = summary

        # -------------------------------------------------------------
        # 2. ĐƠN OUTBOUND (OB) ĐANG MỞ (VL06O)
        # -------------------------------------------------------------
        elif cmd in ["CHECK_OB", "VL06O"]:
            ob_no = ctx.meta.get("obNo", "")
            records = ctx.sap.check_open_ob_vl06o(resolved_mats, ob_no=ob_no, plant=ctx.plant)
            headers = {
                "obNo": "Số OB", "product": "Mã hàng", "qty": "SL xuất", "unit": "ĐVT",
                "sloc": "Sloc", "batch": "Batch", "shipTo": "Khách nhận", "createDate": "Ngày tạo"
            }
            target_label = ob_no or ', '.join(resolved_mats) or 'Tất cả'
            summary = f"Danh sách đơn Outbound (OB) đang mở (VL06O) cho: {target_label}"
            audit_target = target_label
            audit_message = summary

        # -------------------------------------------------------------
        # 3. KIỂM TRA ĐỊNH MỨC BOM (CS03)
        # -------------------------------------------------------------
        elif cmd in ["CHECK_BOM", "BOM", "CS03"]:
            records = ctx.sap.check_bom_cs03(resolved_mats, plant=ctx.plant)
            headers = {
                "parent": "Mã cha", "component": "Mã con (BL)",
                "description": "Tên linh kiện", "qty": "Định mức",
                "unit": "ĐVT", "status": "Trạng thái"
            }
            has_bom_cnt = sum(1 for r in records if str(r.get("status", "")).upper().startswith("CÓ BOM"))
            if has_bom_cnt > 0:
                summary = f"Kết quả kiểm tra định mức BOM (CS03) Plant {ctx.plant} cho: {', '.join(resolved_mats)}{mapping_info}"
            else:
                summary = f"Kết quả kiểm tra BOM cho {', '.join(resolved_mats)} tại Plant {ctx.plant}; xem trạng thái nguồn trong bảng.{mapping_info}"
            audit_target = ", ".join(resolved_mats[:5])
            audit_message = summary

        # -------------------------------------------------------------
        # 4. KIỂM TRA KẸT QUEUE EWM (/SCWM/MON)
        # -------------------------------------------------------------
        elif cmd in ["CHECK_QUEUE", "QUEUE"]:
            inbound_filter = ctx.meta.get("inboundNo", "") or (raw_mats[0] if raw_mats else "")
            records = ctx.sap.check_ewm_queue_mon(inbound_filter=inbound_filter)
            headers = {
                "document": "Số IB / Doc", "error_message": "Nội dung lỗi Queue",
                "queue_name": "Tên Queue"
            }
            filter_info = f" (Lọc theo: {inbound_filter})" if inbound_filter else ""
            summary = f"Danh sách lỗi kẹt Queue EWM (/SCWM/MON){filter_info}"
            audit_target = inbound_filter or "ALL"
            audit_message = summary

        # -------------------------------------------------------------
        # 5. ĐIỀU ĐỘ THÀNH PHẨM (FG) & ĐƠN OB
        # -------------------------------------------------------------
        elif cmd in ["CHECK_FG_OB", "DIEN_OB"]:
            records = ctx.sap.check_fg_ob_status(resolved_mats, plant=ctx.plant)
            headers = {
                "product": "Mã hàng", "fgStock": "Tồn FG 3093",
                "openObList": "Đơn OB đang mở", "openObQty": "SL Đang giữ",
                "availableQty": "Khả dụng xuất", "unit": "ĐVT",
                "workAdvice": "Đánh giá công việc"
            }
            summary = f"Điều độ Thành Phẩm (FG) & Đơn OB cho: {', '.join(resolved_mats)}"
            audit_target = ", ".join(resolved_mats[:5])
            audit_message = summary

        # -------------------------------------------------------------
        # 6. ĐIỀU PHỐI OB & CỔNG XUẤT NGUYÊN LÔ (RC65)
        # -------------------------------------------------------------
        elif cmd in ["DIEU_PHOI_OB", "RC65", "XAC_NHAN_OB"]:
            ob_no = ctx.meta.get("obNo", "")
            ob_items = ctx.sap.check_open_ob_vl06o(resolved_mats, ob_no=ob_no, plant=ctx.plant)
            stock_records = ctx.sap.check_position_posting_change(resolved_mats, batch="", plant=ctx.plant or "3093")
            gate_eval = ObService.validate_whole_batch_gate(ob_no or "BATCH", ob_items, stock_records, source_complete=True)
            result_status = "PREVIEW"
            summary = f"Xem trước Outbound {ob_no or 'Lô'}: {gate_eval['action_advice']}. Chưa điều phối hoặc xác nhận trên SAP."
            records = gate_eval.get("items", [])
            headers = {
                "item_no": "Item", "material": "Mã hàng", "batch": "Số lô",
                "required_qty": "SL Yêu cầu", "available_qty": "SL Khả dụng",
                "deficit_qty": "Thâm hụt", "status": "Trạng thái Cổng"
            }
            audit_target = ob_no or "BATCH"
            audit_message = "Chỉ kiểm tra điều kiện; chưa có executor SAP đã xác minh"

        return CommandResult(
            status=result_status,
            summary=summary,
            records=records,
            headers=headers,
            custom_html=custom_html,
            audit_target=audit_target,
            audit_message=audit_message
        )
