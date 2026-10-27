# -*- coding: utf-8 -*-
"""Handler chuyên trách nghiệp vụ ra đơn NVL MIGO 311 và đối soát Posting Change Kho 1005."""
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.core.security import assert_write_permission, WritePermissionError
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
        # NHÁNH 2: KẾ HOẠCH NVL & RA ĐƠN NVL (PHÂN BỔ FIFO & TẠO ĐƠN OB MIGO 311)
        # -------------------------------------------------------------
        elif cmd in ["RA_DON_NVL", "KE_HOACH_NVL"]:
            is_plan_only = (cmd == "KE_HOACH_NVL")
            is_preview = is_plan_only or (ctx.meta.get("dryRun", False) is True)

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
                if ctx.meta.get("qty") is None or not ctx.meta.get("station"):
                    raise ValueError("Kế hoạch NVL cần qty và station rõ ràng")
                req_unit = ctx.meta.get("unit") or ctx.meta.get("uom") or "PC"
                orders = [
                    {
                        "material": m,
                        "qty": ctx.meta["qty"],
                        "station": ctx.meta["station"],
                        "unit": req_unit,
                        "plant": ctx.plant,
                        "sloc": ctx.meta.get("sloc", "1005")
                    }
                    for m in ctx.resolved_materials
                ]

            all_query_mats = list(dict.fromkeys(o["material"] for o in orders if o.get("material")))
            stock_records = ctx.sap.check_position_posting_change(all_query_mats, batch="", plant=ctx.plant)

            # Đánh giá tính toàn vẹn của nguồn dữ liệu tồn: không tự bịa thông tin
            has_invalid_records = any(
                not r.get("plant") or not r.get("sloc") or not r.get("unit")
                for r in (stock_records or [])
            )
            is_source_complete = (stock_records is not None) and (not has_invalid_records)

            plan = NvlService.allocate_orders(
                orders, stock_records, source_plant=ctx.plant,
                db_path=ctx.state_mgr.db_path, request_id=ctx.request_id,
                source_complete=is_source_complete
            )
            plan_id = plan["plan_id"]
            alloc_unit = plan['allocated_items'][0]['unit'] if plan.get('allocated_items') else 'PC'

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

            # ---------------------------------------------------------
            # 2A. CHẾ ĐỘ KẾ HOẠCH / XEM TRƯỚC (PREVIEW - 100% READ-ONLY)
            # ---------------------------------------------------------
            if is_preview:
                records = NvlService.group_into_ob_groups(
                    plan.get("allocated_items", []),
                    shortages=plan.get("shortages", []),
                    delivery_no=""
                )
                summary = (
                    f"Xem trước kế hoạch ra đơn NVL Kho 1005 (Plant {ctx.plant}) ({plan_id}): "
                    f"Có thể phân bổ {plan['total_allocated']}/{plan['total_ordered']} {alloc_unit}"
                )
                return CommandResult(
                    status="PREVIEW",
                    summary=summary,
                    records=records,
                    headers=headers,
                    audit_target=plan_id,
                    audit_message="Xem trước kế hoạch phân bổ NVL (PREVIEW - 0 ghi SAP)"
                )

            # ---------------------------------------------------------
            # 2B. CHẾ ĐỘ RA ĐƠN THỰC TẾ (RA_DON_NVL - WRITE TO SAP MIGO 311)
            # ---------------------------------------------------------
            # Điều kiện 1: Kiểm tra quyền ghi tập trung & người gửi thật
            assert_write_permission(ctx)

            # Điều kiện 2: Trạng thái kế hoạch không bị BLOCKED
            if plan.get("status") == "BLOCKED":
                records = NvlService.group_into_ob_groups([], shortages=plan.get("shortages", []), delivery_no="")
                return CommandResult(
                    status="BLOCKED",
                    summary=f"Kế hoạch phân bổ NVL bị chặn (BLOCKED): {'; '.join(plan.get('errors', ['Lỗi an toàn']))}",
                    records=records,
                    headers=headers,
                    audit_target=plan_id,
                    audit_message="Kế hoạch phân bổ NVL bị chặn trước khi ghi SAP"
                )

            # Điều kiện 3: Bắt buộc có dòng phân bổ hợp lệ
            allocated_lines = plan.get("allocated_items", [])
            if not allocated_lines:
                records = NvlService.group_into_ob_groups([], shortages=plan.get("shortages", []), delivery_no="")
                return CommandResult(
                    status="FAILED",
                    summary="Không có dòng hàng nào được phân bổ để tạo đơn MIGO",
                    records=records,
                    headers=headers,
                    audit_target=plan_id,
                    audit_message="Không có dòng phân bổ để tạo đơn MIGO"
                )

            # Điều kiện 4: Bắt buộc phân bổ đủ 100% (Fail-closed khi thiếu tồn)
            if not plan.get("is_full_allocation") or plan.get("shortages"):
                records = NvlService.group_into_ob_groups(
                    allocated_lines,
                    shortages=plan.get("shortages", []),
                    delivery_no="---"
                )
                return CommandResult(
                    status="BLOCKED_PARTIAL_STOCK",
                    summary=(
                        f"Kế hoạch phân bổ thiếu tồn ({plan['total_allocated']}/{plan['total_ordered']} {alloc_unit}). "
                        f"Chính sách an toàn chặn tự ý xuất tồn lẻ khi chưa đủ số lượng yêu cầu."
                    ),
                    records=records,
                    headers=headers,
                    audit_target=plan_id,
                    audit_message="Chặn tạo MIGO do thiếu tồn (BLOCKED_PARTIAL_STOCK)"
                )

            # Vòng đời Reservation 2 pha (Two-Phase Reservation)
            has_reservations = bool(ctx.config.get("allow_nvl_reservations"))
            if has_reservations:
                try:
                    NvlService.save_plan_and_reserve(
                        plan,
                        created_by=ctx.meta.get("actualSenderEmail") or ctx.meta.get("senderName") or "BOT",
                        db_path=ctx.state_mgr.db_path
                    )
                except Exception as res_err:
                    print(f"[RESERVATION WARN] Không thể lưu giữ chỗ snapshot: {res_err}")

            print(f"[SAP MIGO] Bắt đầu tạo đơn OB MIGO 311 cho {len(allocated_lines)} dòng NVL (Plant {ctx.plant}, 1005 -> 3001)...")
            ok_post, doc_num, post_msg = ctx.sap.post_migo_311_ob_group(
                allocated_lines,
                plant=ctx.plant,
                source_sloc="1005",
                dest_sloc="3001",
                dry_run=False
            )
            print(f"[SAP MIGO] Kết quả tạo đơn MIGO 311: ok={ok_post} | Số chứng từ='{doc_num}' | Chi tiết: {post_msg}")

            delivery_no = doc_num if ok_post else ""
            if ok_post and doc_num:
                result_status = "SUCCEEDED"
                if has_reservations:
                    try:
                        NvlService.confirm_plan_reservation(plan_id, delivery_no, db_path=ctx.state_mgr.db_path)
                    except Exception as db_sync_err:
                        print(f"[RESERVATION WARN] SAP post thành công ({delivery_no}) nhưng lỗi DB sync: {db_sync_err}")
                        result_status = "AMBIGUOUS_RESERVATION"
            else:
                result_status = "FAILED"
                if has_reservations:
                    try:
                        NvlService.release_plan_reservation(plan_id, db_path=ctx.state_mgr.db_path)
                    except Exception:
                        pass

            records = NvlService.group_into_ob_groups(
                allocated_lines,
                shortages=plan.get("shortages", []),
                delivery_no=delivery_no
            )

            if result_status == "SUCCEEDED" and delivery_no:
                summary = f"Đã chạy FIFO và tạo đơn OB cấp NVL Kho 1005 (Plant {ctx.plant}) - Số OB: {delivery_no} ({plan['total_allocated']}/{plan['total_ordered']} {alloc_unit})"
            elif result_status == "AMBIGUOUS_RESERVATION":
                summary = f"Đã tạo chứng từ MIGO thành công trên SAP (Số: {delivery_no}) nhưng gặp sự cố đồng bộ dữ liệu DB nội bộ (AMBIGUOUS_RESERVATION). Vui lòng đối soát kiểm tra trước khi chạy lại."
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
