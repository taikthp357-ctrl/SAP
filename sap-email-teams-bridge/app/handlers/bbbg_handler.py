# -*- coding: utf-8 -*-
"""Handler chuyên trách đối soát và tạo Biên bản giao nhận Thành phẩm (BBBG FG)."""
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler
from app.bbbg_service import BbbgService


class BbbgHandler(BaseHandler):
    """Xử lý nhóm lệnh Biên bản giao nhận:
    - TAO_BBBG (Đối soát và xuất file Excel BBBG FG)
    - XEM_BBBG, BBBG_FG (Xem trước bảng đối soát trên Teams)
    """

    def handle(self, ctx: RequestContext) -> CommandResult:
        cmd = (ctx.command or "").upper()
        days_cnt = ctx.meta.get("days", 1)
        resolved_mats = ctx.resolved_materials or []

        mb51_rows = ctx.sap.check_history_mb51(resolved_mats, days=days_cnt, plant=ctx.plant, explicit_plant=True)
        vl06i_rows = ctx.sap.check_inbound_treo_vl06i(resolved_mats, plant=ctx.plant)
        recon = BbbgService.reconcile(mb51_rows, vl06i_rows, source_complete=True)

        custom_html = None
        audit_target = recon.get("batch_id", "")
        audit_message = ""

        if cmd == "TAO_BBBG":
            out_dir = ctx.data_dir / "exports" / "bbbg" if ctx.data_dir else None
            file_path = BbbgService.export_excel(recon, meta_header={"shop": ctx.meta.get("shop", "BODY")}, output_dir=out_dir)
            summary = f"Biên bản giao nhận Thành phẩm (BBBG FG) - {recon['matched_count']}/{recon['total_items']} khớp số lượng (File lưu tại máy worker: {file_path}; chưa gửi tệp lên Teams)"
            audit_message = f"Đã xuất file BBBG: {file_path}"
        else:
            preview = BbbgService.generate_mail_preview(recon, shop_key=ctx.meta.get("shop", "BODY"))
            custom_html = preview.get("html_preview")
            summary = f"Xem trước BBBG FG ({recon['matched_count']}/{recon['total_items']} khớp số lượng)"
            audit_message = summary

        records = recon.get("items", [])
        headers = {
            "material": "Mã hàng", "batch": "Số lô", "doc_no": "Inbound / Doc",
            "handover_qty": "SL Bàn giao", "received_qty": "SL Nhận", "diff_qty": "Chênh lệch",
            "gr_status": "Trạng thái", "note": "Ghi chú"
        }

        return CommandResult(
            status="SUCCEEDED",
            summary=summary,
            records=records,
            headers=headers,
            custom_html=custom_html,
            audit_target=audit_target,
            audit_message=audit_message
        )
