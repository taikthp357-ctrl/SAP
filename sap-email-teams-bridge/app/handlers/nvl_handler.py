# -*- coding: utf-8 -*-
"""Handler chuyên trách nghiệp vụ ra đơn NVL MIGO 311 và đối soát Posting Change Kho 1005."""
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler
from app.nvl_service import NvlService


class NvlHandler(BaseHandler):
    """Xử lý nhóm lệnh: RA_DON_NVL, KE_HOACH_NVL, CHECK_NVL_MIGO, RA_DON_MIGO, MIGO_311, CHECK_NVL, POSTING_CHANGE, POSTING."""

    def handle(self, ctx: RequestContext) -> CommandResult:
        cmd = (ctx.command or "").upper()

        # -------------------------------------------------------------
        # NHÁNH 1: KIỂM TRA POSTING CHANGE KHO 1005 (READ-ONLY)
        # -------------------------------------------------------------
        if cmd in ["CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL", "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING"]:
            records = ctx.sap.check_nvl_migo_balance(ctx.resolved_materials, plant=ctx.plant or "3093", batch=ctx.batch)
            headers = {
                "product": "Mã hàng", "plant": "Plant", "sloc": "Kho",
                "bin": "Vị trí Bin", "batch": "Số lô",
                "qty": "SL Khả dụng", "unit": "ĐVT",
                "grDate": "Ngày GR", "grTime": "Giờ GR",
                "cat": "Trạng thái", "fifoRank": "Thứ tự FIFO",
                "migoAdvice": "Khuyến nghị ra đơn MIGO 311"
            }
            summary = f"Posting Change NVL (Plant {ctx.plant or '3093'} - Kho 1005) cho: {', '.join(ctx.resolved_materials)}"

            return CommandResult(
                status="SUCCEEDED",
                summary=summary,
                records=records,
                headers=headers,
                audit_target=", ".join(ctx.resolved_materials[:5]),
                audit_message=f"Kiểm tra Posting Change NVL ({len(records)} dòng)"
            )

        # -------------------------------------------------------------
        # NHÁNH 2: RA ĐƠN NVL (PHÂN BỔ FIFO & TẠO ĐƠN OB MIGO 311)
        # -------------------------------------------------------------
        elif cmd in ["RA_DON_NVL", "KE_HOACH_NVL"]:
            raw_orders = ctx.meta.get("nvlOrders") or ctx.meta.get("orders") or []
            if raw_orders:
                orders = []
                for o in raw_orders:
                    raw_m = o.get("material", "")
                    resolved_cand, _ = ctx.resolver.resolve(raw_m)
                    resolved_m = resolved_cand[0] if resolved_cand else raw_m
                    orders.append({
                        "material": resolved_m,
                        "qty": o.get("qty", 0.0),
                        "station": o.get("station") or ctx.meta.get("station", "LINE"),
                        "unit": o.get("unit") or ctx.meta.get("unit") or ctx.meta.get("uom") or "PC",
                        "plant": ctx.plant,
                        "sloc": o.get("sloc") or ctx.meta.get("sloc", "1005")
                    })
            else:
                if ctx.meta.get("qty") is None or not ctx.meta.get("station") or not (ctx.meta.get("unit") or ctx.meta.get("uom")):
                    raise ValueError("Kế hoạch NVL cần qty, station và unit rõ ràng; không tự điền mặc định")
                orders = [
                    {
                        "material": m,
                        "qty": ctx.meta["qty"],
                        "station": ctx.meta["station"],
                        "unit": ctx.meta.get("unit") or ctx.meta.get("uom"),
                        "plant": ctx.plant,
                        "sloc": ctx.meta.get("sloc", "1005")
                    }
                    for m in ctx.resolved_materials
                ]

            all_query_mats = list(dict.fromkeys(o["material"] for o in orders if o.get("material")))
            stock_records = ctx.sap.check_position_posting_change(all_query_mats, batch="", plant=ctx.plant)
            plan = NvlService.allocate_orders(
                orders, stock_records, source_plant=ctx.plant,
                db_path=ctx.state_mgr.db_path, request_id=ctx.request_id, source_complete=True
            )
            plan_id = plan["plan_id"]

            is_preview = ctx.meta.get("dryRun", False) is True
            delivery_no = ""
            post_msg = ""
            if not is_preview:
                print(f"[SAP MIGO] Bắt đầu tạo đơn OB MIGO 311 cho {len(plan.get('allocated_items', []))} dòng NVL (Plant {ctx.plant}, 1005 -> 3001)...")
                ok_post, doc_num, post_msg = ctx.sap.post_migo_311_ob_group(
                    plan.get("allocated_items", []),
                    plant=ctx.plant,
                    source_sloc="1005",
                    dest_sloc="3001",
                    dry_run=False
                )
                print(f"[SAP MIGO] Kết quả tạo đơn MIGO 311: ok={ok_post} | Số chứng từ='{doc_num}' | Chi tiết: {post_msg}")
                delivery_no = doc_num if ok_post else ""
                result_status = "SUCCEEDED" if ok_post else "FAILED"
            else:
                delivery_no = ""
                result_status = "PREVIEW"

            if ctx.config.get("allow_nvl_reservations"):
                plan_id = NvlService.save_plan_and_reserve(plan, created_by=ctx.meta.get("senderName") or "BOT", db_path=ctx.state_mgr.db_path)

            records = NvlService.group_into_ob_groups(
                plan.get("allocated_items", []),
                shortages=plan.get("shortages", []),
                delivery_no=delivery_no
            )
            headers = {
                "stt": "OB",
                "delivery_no": "Số OB",
                "material": "Mã NVL",
                "allocated_qty": "SL",
                "batch": "Lô",
                "bin": "Bin",
                "station": "Trạm",
                "status": "TT"
            }
            alloc_unit = plan['allocated_items'][0]['unit'] if plan.get('allocated_items') else 'PC'
            if result_status == "SUCCEEDED" and delivery_no:
                summary = f"Đã chạy FIFO và tạo đơn OB cấp NVL Kho 1005 (Plant {ctx.plant}) - Số OB: {delivery_no} ({plan['total_allocated']}/{plan['total_ordered']} {alloc_unit})"
            elif result_status == "PREVIEW" or is_preview:
                summary = f"Xem trước kế hoạch ra đơn NVL Kho 1005 (Plant {ctx.plant}) ({plan_id}): Có thể phân bổ {plan['total_allocated']}/{plan['total_ordered']} {alloc_unit}"
            else:
                summary = f"Kế hoạch ra đơn NVL Kho 1005 ({post_msg or 'Chưa tạo được đơn OB trên SAP'}) - Đã phân bổ {plan['total_allocated']}/{plan['total_ordered']} {alloc_unit}"

            return CommandResult(
                status=result_status,
                summary=summary,
                records=records,
                headers=headers,
                audit_target=delivery_no or plan_id,
                audit_message=f"Đã xử lý ra đơn NVL ({delivery_no or 'Chưa tạo OB'})"
            )

        raise ValueError(f"NvlHandler không hỗ trợ lệnh: {cmd}")
