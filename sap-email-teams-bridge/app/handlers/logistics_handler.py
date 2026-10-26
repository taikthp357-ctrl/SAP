# -*- coding: utf-8 -*-
"""Handler chuyên trách nghiệp vụ Logistics: Inbound (VL06I), PO (ME2N), STO điều chuyển và ASN tracking."""
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler
from app.extended_query_service import ExtendedQueryService


class LogisticsHandler(BaseHandler):
    """Xử lý nhóm lệnh Logistics & Mua hàng:
    - CHECK_IB, IB_DETAIL, KIEMTRA_IB (Chi tiết đơn Inbound)
    - TIM_IB, FIND_IB (Tìm Inbound theo PO / Bill)
    - IB_TREO, IB, INBOUND, VL06I (Inbound treo chưa GR)
    - CHECK_PO, PO_STATUS, KIEMTRA_PO (Tình trạng đơn PO)
    - STO, ME2N, MB5T, DIEUCHUYEN (Đơn điều chuyển STO)
    - TRA_CUU_ASN, ASN (Tra cứu vận đơn ASN)
    """

    def handle(self, ctx: RequestContext) -> CommandResult:
        cmd = (ctx.command or "").upper()
        records: List[Dict[str, Any]] = []
        headers: Dict[str, str] = {}
        summary = ""
        custom_html = None
        result_status = "SUCCEEDED"

        raw_mats = ctx.raw_materials or []
        resolved_mats = ctx.resolved_materials or []
        mapping_info = f" (Quy đổi: {', '.join(raw_mats)} ➔ {', '.join(resolved_mats)})" if raw_mats != resolved_mats else ""

        # -------------------------------------------------------------
        # 1. DANH SÁCH INBOUND TREO CHƯA GR (VL06I)
        # -------------------------------------------------------------
        if cmd in ["IB_TREO", "IB", "INBOUND", "VL06I"]:
            records = ctx.sap.check_inbound_treo_vl06i(resolved_mats, plant=ctx.plant)
            headers = {
                "ibNo": "Số IB", "refNo": "Số OB / Vận đơn", "product": "Mã hàng",
                "batch": "Số Lô", "qty": "SL giao", "unit": "ĐVT",
                "plant": "Plant", "sloc": "Kho nhận", "status": "Trạng thái"
            }
            summary = f"Danh sách đơn Inbound (IB) treo chưa GR (VL06I) cho mã: {', '.join(resolved_mats)}{mapping_info}"

        # -------------------------------------------------------------
        # 2. ĐƠN ĐIỀU CHUYỂN STO (ME2N)
        # -------------------------------------------------------------
        elif cmd in ["STO", "ME2N", "MB5T", "DIEUCHUYEN"]:
            records = ctx.sap.check_sto_me2n(resolved_mats, plant=ctx.plant)
            headers = {
                "stoNo": "Số STO", "item": "Item", "product": "Mã hàng",
                "orderQty": "SL Đặt", "deliveredQty": "Đã xuất", "remainingQty": "Còn nợ",
                "unit": "ĐVT", "supplyingPlant": "Plant cấp", "receivingPlant": "Plant nhận"
            }
            summary = f"Danh sách đơn điều chuyển STO (ME2N) cho mã: {', '.join(resolved_mats)}{mapping_info}"

        # -------------------------------------------------------------
        # 3. CHI TIẾT ĐƠN INBOUND (VL06I ITEM VIEW)
        # -------------------------------------------------------------
        elif cmd in ["CHECK_IB", "IB_DETAIL", "KIEMTRA_IB"]:
            inbound_no = ctx.meta.get("inboundNo", "") or (raw_mats[0] if raw_mats and raw_mats[0].startswith("18") else "")
            expected_qty = ctx.meta.get("expectedQty", 0.0)
            records = ctx.sap.check_inbound_detail_vl06i(inbound_no=inbound_no, materials=resolved_mats, plant=ctx.plant, expected_qty=expected_qty)
            headers = {
                "ibNo": "Số IB", "item": "Item", "product": "Mã hàng",
                "actualQty": "SL Inbound", "expectedQty": "SL Yêu cầu", "delta": "Chênh lệch", "qtyCheck": "Đối chiếu SL",
                "unit": "ĐVT", "sloc": "Kho", "status": "Trạng thái GR",
                "refPo": "PO tham chiếu", "poItem": "PO Item"
            }
            summary = f"Chi tiết đơn Inbound {inbound_no or ', '.join(resolved_mats)} (VL06I Item View)"

        # -------------------------------------------------------------
        # 4. TÌM KIẾM INBOUND THEO THAM CHIẾU (PO / BILL)
        # -------------------------------------------------------------
        elif cmd in ["TIM_IB", "FIND_IB"]:
            po_no = ctx.meta.get("poNo", "")
            ref_no = ctx.meta.get("refNo", "")
            records = ctx.sap.find_inbound_vl06i(po_no=po_no, ref_no=ref_no, materials=resolved_mats, plant=ctx.plant)
            headers = {
                "ibNo": "Số IB", "item": "Item", "product": "Mã hàng",
                "qty": "SL giao", "unit": "ĐVT", "sloc": "Kho",
                "status": "Trạng thái", "refPo": "Số PO", "refNo": "Vận đơn / Bill"
            }
            target_lbl = po_no or ref_no or (', '.join(resolved_mats) if resolved_mats else 'Tất cả')
            summary = f"Tìm kiếm Inbound theo tham chiếu: {target_lbl}"

        # -------------------------------------------------------------
        # 5. TÌNH TRẠNG ĐƠN ĐẶT HÀNG PO (ME2N)
        # -------------------------------------------------------------
        elif cmd in ["CHECK_PO", "PO_STATUS", "KIEMTRA_PO"]:
            po_no = ctx.meta.get("poNo", "") or (raw_mats[0] if raw_mats and raw_mats[0].startswith(("45", "55")) else "")
            po_item = ctx.meta.get("poItem", "")
            records = ctx.sap.check_po_me2n(po_no=po_no, materials=resolved_mats, plant=ctx.plant, po_item=po_item)
            headers = {
                "poNo": "Số PO", "item": "Item", "product": "Mã hàng",
                "orderQty": "SL Đặt", "deliveredQty": "Đã nhận", "openQty": "Còn nợ",
                "unit": "ĐVT", "plant": "Plant", "deleteInd": "Xóa/Khóa", "releaseStatus": "Duyệt"
            }
            summary = f"Tình trạng đơn đặt hàng PO {po_no or ', '.join(resolved_mats)} (ME2N)"

        # -------------------------------------------------------------
        # 6. TRA CỨU VẬN ĐƠN ASN
        # -------------------------------------------------------------
        elif cmd in ["TRA_CUU_ASN", "ASN"]:
            inbound_no = ctx.meta.get("inboundNo", "") or (raw_mats[0] if raw_mats else "")
            raw_asn = ctx.sap.find_inbound_vl06i(materials=resolved_mats, plant=ctx.plant, inbound_no=inbound_no)
            asn_res = ExtendedQueryService.track_asn_provenance(inbound_no, raw_asn)
            records = asn_res.get("items", [])
            headers = {
                "material": "Mã hàng", "batch": "Số lô", "qty": "SL giao",
                "unit": "ĐVT", "po_no": "Số PO", "status": "Trạng thái"
            }
            summary = asn_res["summary"]

        return CommandResult(
            status=result_status,
            summary=summary,
            records=records,
            headers=headers,
            custom_html=custom_html,
            audit_target=ctx.meta.get("inboundNo", "") or ctx.meta.get("poNo", "") or (", ".join(resolved_mats[:5]) if resolved_mats else ""),
            audit_message=f"{cmd}: {summary}"
        )
