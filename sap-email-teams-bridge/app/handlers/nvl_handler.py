# -*- coding: utf-8 -*-
"""Handler chuyên trách nghiệp vụ ra đơn NVL MIGO 311 và đối soát Posting Change Kho 1005."""
import os
import json
import re
from pathlib import Path
from decimal import Decimal
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.core.security import assert_write_permission, WritePermissionError
from app.handlers.base_handler import BaseHandler
from app.nvl_service import NvlService
from app.nvl_master import NvlMasterService
from app.nvl_tracking import NvlTrackingService


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
            # Layer 3 Handler Guard: Chống nhầm lẫn số chứng từ DOC (50x/49x) vào luồng ra đơn NVL
            raw_text_check = str(ctx.meta.get("rawText") or "")
            has_att_marker = any(k in raw_text_check.lower() for k in ["att mail", "dinh kem", "bktxt", "headtext", "zfir266"])
            has_explicit_ra_don = any(k in raw_text_check.lower() for k in ["ra don", "don nvl", "ke hoach nvl", "phan bo nvl"])
            if has_att_marker and not has_explicit_ra_don:
                return CommandResult(
                    status="REJECTED_INTENT_CONFLICT",
                    summary="Lệnh chứa từ khóa đính kèm (ATT/HeadText) nhưng bị gửi nhầm sang Handler NVL. Đã chặn an toàn.",
                    records=[],
                    headers={},
                    audit_target="INTENT_GUARD",
                    audit_message="Chặn nhầm luồng ATT sang NVL Handler"
                )

            is_plan_only = (cmd == "KE_HOACH_NVL")
            is_preview = is_plan_only or (ctx.meta.get("dryRun", False) is True)

            req_src_sloc = ctx.meta.get("source_sloc") or ctx.meta.get("sourceSloc") or ctx.meta.get("sloc", "1005")
            req_dest_sloc = ctx.meta.get("dest_sloc") or ctx.meta.get("destSloc") or ctx.meta.get("sloc_line", "3001")
            req_target_bin = ctx.meta.get("target_bin") or ctx.meta.get("bin") or ""

            raw_orders = ctx.meta.get("nvlOrders") or ctx.meta.get("orders") or []
            if raw_orders:
                orders = []
                for o in raw_orders:
                    raw_m = o.get("material", "")
                    resolved_cand, _ = ctx.resolver.resolve(raw_m)
                    resolved_m = resolved_cand[0] if resolved_cand else raw_m
                    mat_src_sloc = o.get("sloc") or ("200C" if NvlMasterService.is_henkel_glue(resolved_m) else req_src_sloc)
                    mat_dest_sloc = o.get("dest_sloc") or req_dest_sloc
                    mat_bin = o.get("target_bin") or o.get("bin") or req_target_bin
                    orders.append({
                        "material": resolved_m,
                        "qty": o.get("qty", 0.0),
                        "station": o.get("station") or ctx.meta.get("station", "LINE"),
                        "unit": o.get("unit") or ctx.meta.get("unit") or ctx.meta.get("uom") or "PC",
                        "plant": ctx.plant,
                        "sloc": mat_src_sloc,
                        "dest_sloc": mat_dest_sloc,
                        "target_bin": mat_bin
                    })
            else:
                if ctx.meta.get("qty") is None or not ctx.meta.get("station"):
                    raise ValueError("Kế hoạch NVL cần qty và station rõ ràng")
                req_unit = ctx.meta.get("unit") or ctx.meta.get("uom") or "PC"
                orders = []
                for m in ctx.resolved_materials:
                    mat_src_sloc = "200C" if NvlMasterService.is_henkel_glue(m) else req_src_sloc
                    orders.append({
                        "material": m,
                        "qty": ctx.meta["qty"],
                        "station": ctx.meta["station"],
                        "unit": req_unit,
                        "plant": ctx.plant,
                        "sloc": mat_src_sloc,
                        "dest_sloc": req_dest_sloc,
                        "target_bin": req_target_bin
                    })

            # Hỗ trợ tự động khống chế theo giới hạn thùng V31 nếu có yêu cầu (cap_boxes / theo thùng)
            enforce_box_limit = (
                bool(ctx.meta.get("capBoxes") or ctx.meta.get("cap_boxes"))
                or bool(ctx.config.get("enforce_v31_box_limit", False))
                or any(k in raw_text_check.lower() for k in ["theo thung", "gioi han thung", "cap thung", "khong che thung", "ep thung"])
            )
            if enforce_box_limit:
                capped_orders = []
                for o in orders:
                    o_copy = dict(o)
                    m = o_copy.get("material", "")
                    q = float(o_copy.get("qty", 0.0))
                    b_split = NvlMasterService.calculate_box_split(m, q)
                    if b_split.get("is_box") and b_split.get("v31_capped_qty", 0) > 0:
                        capped_q = float(b_split["v31_capped_qty"])
                        if q > capped_q:
                            o_copy["qty"] = capped_q
                            o_copy["original_qty"] = q
                    capped_orders.append(o_copy)
                orders = capped_orders

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
                "allocated_qty": "SL Cấp",
                "packing_desc": "Quy cách đóng thùng",
                "box_limit": "Trần thùng V31",
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

            # Gom các dòng cùng vật tư, lô, trạm để điền MIGO 311 (chuẩn RESULT của Excel V31)
            # MIGO 311 không quản lý Bin; gom dòng giúp bảng MIGO chuẩn xác và tránh trùng lặp dòng nhập
            migo_items = []
            migo_consolidated = {}
            for line in allocated_lines:
                m_key = (
                    str(line.get("material") or "").strip().upper(),
                    str(line.get("batch") or "").strip().upper(),
                    str(line.get("station") or "").strip(),
                    str(line.get("source_plant") or ctx.plant).strip(),
                    str(line.get("source_sloc") or "1005").strip(),
                    str(line.get("dest_sloc") or "3001").strip()
                )
                if m_key not in migo_consolidated:
                    row_c = dict(line)
                    migo_consolidated[m_key] = row_c
                    migo_items.append(row_c)
                else:
                    migo_consolidated[m_key]["allocated_qty"] = float(
                        Decimal(str(migo_consolidated[m_key]["allocated_qty"])) + Decimal(str(line.get("allocated_qty", 0)))
                    )

            final_src_sloc = str(migo_items[0].get("source_sloc") or req_src_sloc or "1005") if migo_items else "1005"
            final_dest_sloc = str(migo_items[0].get("dest_sloc") or req_dest_sloc or "3001") if migo_items else "3001"
            print(f"[SAP MIGO] Bắt đầu tạo đơn OB MIGO 311 cho {len(migo_items)} dòng MIGO (từ {len(allocated_lines)} vị trí bin, Plant {ctx.plant}, {final_src_sloc} -> {final_dest_sloc})...")
            ok_post, doc_num, post_msg = ctx.sap.post_migo_311_ob_group(
                migo_items,
                plant=ctx.plant,
                source_sloc=final_src_sloc,
                dest_sloc=final_dest_sloc,
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

            sync_msg = ""
            if result_status == "SUCCEEDED" and delivery_no:
                try:
                    sync_res = NvlTrackingService.sync_dispatch_to_tracking_file(
                        orders_dispatched=migo_items,
                        delivery_no=delivery_no,
                        shift=ctx.meta.get("shift"),
                        directory=ctx.config.get("tracking_files_dir")
                    )
                    if sync_res.get("status") == "SUCCEEDED" and sync_res.get("updated_count", 0) > 0:
                        sync_msg = f" | Đã tự động cập nhật {sync_res['updated_count']} dòng vào file '{sync_res['tracking_file']}' (Trạng thái: 'Đã ra đơn chờ cấp')"
                except Exception as sync_err:
                    print(f"[TRACKING SYNC WARN] {sync_err}")

            if result_status == "SUCCEEDED" and delivery_no:
                summary = f"Đã chạy FIFO và tạo đơn OB cấp NVL Kho 1005 (Plant {ctx.plant}) - Số OB: {delivery_no} ({plan['total_allocated']}/{plan['total_ordered']} {alloc_unit}){sync_msg}"
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

        # -------------------------------------------------------------
        # NHÁNH 3: BÁO CÁO TIẾN ĐỘ & TÌNH TRẠNG CẤP NVL TRONG CA
        # -------------------------------------------------------------
        elif cmd in ["TINH_TRANG_NVL", "BAO_CAO_NVL", "TIEN_DO_NVL", "TINH_TRANG_CAP_NVL", "BAO_CAO_CAP_NVL"]:
            shift_val = ctx.meta.get("shift")
            model_val = ctx.meta.get("model")
            date_val = ctx.meta.get("date") or ctx.meta.get("targetDate")
            dir_val = ctx.config.get("tracking_files_dir")

            status_res = NvlTrackingService.get_shift_supply_status(
                shift=shift_val,
                model=model_val,
                target_date=date_val,
                directory=dir_val,
                sap_engine=ctx.sap
            )

            if status_res.get("status") != "OK":
                return CommandResult(
                    status="NOT_FOUND",
                    summary=status_res.get("summary", "Không tìm thấy dữ liệu theo dõi ca"),
                    records=[],
                    headers={},
                    audit_target=f"{shift_val or ''}_{model_val or ''}",
                    audit_message="Không tìm thấy file theo dõi ca NVL"
                )

            headers = {
                "station": "Trạm cấp",
                "material": "Mã NVL",
                "ncc": "Nhà cung cấp",
                "qty_ordered": "SL Order",
                "qty_supplied": "Đã cấp",
                "gap": "Còn thiếu (Gap)",
                "stock_1005": "Tồn Kho 1005",
                "status_label": "Trạng thái",
                "advice": "Khuyến nghị xử lý"
            }

            display_records = status_res.get("shortages") or status_res.get("records") or []
            kpi = status_res.get("kpi", {})

            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#ffffff;border:1px solid #d0d7de;"
                f"border-radius:6px;padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.05);'>"
                f"<div style='font-size:16px;font-weight:bold;color:#0969da;margin-bottom:12px;display:flex;align-items:center;gap:8px;'>"
                f"📊 BÁO CÁO TIẾN ĐỘ CẤP HÀNG NVL (Ca {shift_val or 'Hiện tại'} - {model_val or 'Tất cả Model'})"
                f"</div>"
                f"<div style='display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px;'>"
                f"<div style='background:#f6f8fa;border:1px solid #d0d7de;padding:10px;border-radius:6px;text-align:center;'>"
                f"<div style='font-size:11px;color:#57606a;'>Tổng số Order</div>"
                f"<div style='font-size:18px;font-weight:bold;color:#24292f;'>{kpi.get('total_orders', 0)}</div>"
                f"</div>"
                f"<div style='background:#dafbe1;border:1px solid #aceebb;padding:10px;border-radius:6px;text-align:center;'>"
                f"<div style='font-size:11px;color:#1a7f37;'>Hoàn thành (100%)</div>"
                f"<div style='font-size:18px;font-weight:bold;color:#1a7f37;'>{kpi.get('completed_count', 0)} ({kpi.get('completion_rate', 0)}%)</div>"
                f"</div>"
                f"<div style='background:#fff8c5;border:1px solid #f1e05a;padding:10px;border-radius:6px;text-align:center;'>"
                f"<div style='font-size:11px;color:#9a6700;'>Cấp 1 phần</div>"
                f"<div style='font-size:18px;font-weight:bold;color:#9a6700;'>{kpi.get('partial_count', 0)}</div>"
                f"</div>"
                f"<div style='background:#ffebe9;border:1px solid #ff8182;padding:10px;border-radius:6px;text-align:center;'>"
                f"<div style='font-size:11px;color:#cf222e;'>Chờ cấp / Pending</div>"
                f"<div style='font-size:18px;font-weight:bold;color:#cf222e;'>{kpi.get('pending_count', 0)}</div>"
                f"</div>"
                f"</div>"
                f"<div style='font-size:12px;color:#57606a;margin-bottom:6px;'>Nguồn dữ liệu: <code>{kpi.get('tracking_file', '')}</code></div>"
                f"</div>"
            )

            return CommandResult(
                status="SUCCEEDED",
                summary=status_res["summary"],
                records=display_records,
                headers=headers,
                custom_html=custom_html,
                audit_target=f"{shift_val or ''}_{model_val or ''}",
                audit_message=f"Báo cáo tiến độ cấp NVL ({len(display_records)} dòng hiển thị)"
            )

        # -------------------------------------------------------------
        # NHÁNH 5: TRUY VẤN ĐỊNH DANH CẤP PHÁT NVL ĐA MODEL (MULTI-MODEL TRACE)
        # -------------------------------------------------------------
        elif cmd in ["TRA_CUU_CAP_NVL", "DINH_DANH_NVL", "TRACE_NVL"]:
            mats = ctx.resolved_materials or ctx.raw_materials or []
            hint_model = ctx.meta.get("model")
            hint_shift = ctx.meta.get("shift")
            tracking_dir = ctx.meta.get("folderPath") or ctx.meta.get("tracking_files_dir") or ctx.config.get("tracking_files_dir")

            if not mats:
                return CommandResult(
                    status="NEED_INPUT",
                    summary="Vui lòng cung cấp ít nhất một mã vật tư (hoặc dán tin nhắn xưởng) để tra cứu cấp phát.",
                    records=[],
                    headers={},
                    audit_target="NO_MAT",
                    audit_message="Thiếu mã vật tư khi tra cứu cấp NVL"
                )

            trace_res = NvlTrackingService.trace_multiple_materials(
                materials=mats,
                hint_model=hint_model,
                hint_shift=hint_shift,
                directory=tracking_dir,
                sap_engine=ctx.sap
            )

            if trace_res.get("status") == "PASSWORD_REQUIRED":
                return CommandResult(
                    status="NEED_INPUT",
                    summary=trace_res.get("summary") or "File theo dõi ca yêu cầu mật khẩu mở file.",
                    records=[],
                    headers={},
                    custom_html=trace_res.get("custom_html") or trace_res.get("summary"),
                    audit_target=", ".join(mats[:5]),
                    audit_message="File theo dõi NVL yêu cầu mật khẩu mở file (PASSWORD_REQUIRED)"
                )

            headers = {
                "material": "Mã NVL",
                "model": "Model xe",
                "shift_chay": "Ca chạy",
                "station": "Trạm cấp",
                "sp_display": "Số phiếu",
                "order_qty": "SL Order",
                "total_cap": "Đã cấp",
                "gap_display": "Cấp thiếu (Gap)",
                "status_pick": "TT Nhặt hàng"
            }

            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;'>"
                f"{trace_res.get('html_table', '')}"
                f"<div style='background:#f6f8fa;border-left:4px solid #0969da;padding:12px;border-radius:6px;font-size:13px;line-height:1.6;margin-top:10px;'>"
                f"{trace_res['summary']}"
                f"</div>"
                f"</div>"
            )

            summary_text = trace_res.get("markdown_table") or trace_res["summary"]

            return CommandResult(
                status="SUCCEEDED",
                summary=summary_text,
                records=trace_res["records"],
                headers=headers,
                custom_html=custom_html,
                audit_target=", ".join(mats[:5]),
                audit_message=f"Định danh cấp NVL đa Model ({len(trace_res['records'])} dòng khớp)"
            )

        # -------------------------------------------------------------
        # NHÁNH 4: TRA CỨU & HƯỚNG DẪN IN PHIẾU WHT / WT TRONG SAP
        # -------------------------------------------------------------
        elif cmd in ["IN_WHT", "PRINT_WHT", "HUONG_DAN_IN_WHT", "IN_WHO", "IN_PHIEU"]:
            docs_to_print = []
            if ctx.meta.get("obNo"):
                docs_to_print.append(str(ctx.meta["obNo"]).strip())
            if ctx.meta.get("materialDocument"):
                docs_to_print.append(str(ctx.meta["materialDocument"]).strip())
            if ctx.raw_materials:
                for m in ctx.raw_materials:
                    c = str(m).strip()
                    if c not in docs_to_print and (c.startswith("6") or c.startswith("101") or (c.isdigit() and len(c) >= 8)):
                        docs_to_print.append(c)

            # Nếu có SAP session kết nối và không phải chế độ xem trước -> Thực thi in tự động 2 tờ
            if ctx.sap and not is_preview and docs_to_print:
                from app.sap_who_printer import SapWhoPrinterService
                printer_svc = SapWhoPrinterService()
                try:
                    active_sess = getattr(ctx.sap, "session", None) or getattr(ctx.sap, "active_session", None)
                    if active_sess:
                        res = printer_svc.execute_smart_print(session=active_sess, doc_inputs=docs_to_print, copies=2)
                        who_list = res.get("who_printed", [])
                        custom_html = (
                            f"<div style='font-family:Segoe UI,sans-serif;background:#ecfdf5;border:1px solid #a7f3d0;"
                            f"border-left:4px solid #10b981;padding:14px;border-radius:4px;color:#065f46;'>"
                            f"<div style='font-weight:bold;font-size:14px;margin-bottom:8px;'>"
                            f"🖨️ KẾT QUẢ IN PHIẾU KHO TỰ ĐỘNG (2 BẢN / PHIẾU)</div>"
                            f"<div style='font-size:12px;line-height:1.6;'>"
                            f"<p><b>Trạng thái:</b> {res.get('message', 'Đã in thành công')}</p>"
                            f"<p><b>Danh sách phiếu WHO đã in ({len(who_list)}):</b> {', '.join(who_list) if who_list else 'Không có'}</p>"
                            f"<p><b>Máy in cấu hình:</b> {printer_svc.default_printer}</p>"
                            f"</div></div>"
                        )
                        return CommandResult(
                            status=res["status"],
                            summary=res["message"],
                            records=[{"who": w, "copies": 2, "printer": printer_svc.default_printer, "status": "PRINTED"} for w in who_list],
                            headers={"who": "Số Warehouse Order", "copies": "Số bản", "printer": "Máy in", "status": "Trạng thái"},
                            custom_html=custom_html,
                            audit_target=", ".join(docs_to_print),
                            audit_message=res["message"]
                        )
                except Exception as ex:
                    print(f"[SMART PRINT ERROR] Không thể thực thi in trực tiếp: {ex}")

            doc_no = docs_to_print[0] if docs_to_print else ""
            if not doc_no and ctx.resolved_materials:
                doc_no = ctx.resolved_materials[0]

            guidance = NvlTrackingService.get_wht_print_guidance(doc_no=doc_no)
            steps_html = "".join([f"<li style='margin-bottom:6px;'>{s}</li>" for s in guidance["steps"]])

            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#f0f9ff;border:1px solid #b9e6fe;"
                f"border-left:4px solid #0284c7;padding:14px;border-radius:4px;color:#0c4a6e;'>"
                f"<div style='font-weight:bold;font-size:14px;margin-bottom:8px;'>"
                f"🖨️ HƯỚNG DẪN IN PHIẾU WHT (WAREHOUSE TASK) - CHỨNG TỪ: {doc_no or 'Chưa chỉ định'}</div>"
                f"<div style='font-size:12px;line-height:1.6;'>"
                f"<ul style='padding-left:18px;margin:0 0 10px 0;'>{steps_html}</ul>"
                f"<div style='background:#e0f2fe;padding:8px;border-radius:4px;font-size:11px;color:#0369a1;'>"
                f"<b>Mẹo:</b> Khi SAP mở sẵn màn hình /SCWM/MON hoặc /SCWM/PRDO, Bot có thể trực tiếp gửi lệnh in đến máy in kho cấu hình."
                f"</div></div></div>"
            )

            return CommandResult(
                status="SUCCEEDED",
                summary=guidance["summary"],
                records=[],
                headers={},
                custom_html=custom_html,
                audit_target=doc_no or "WHT_GUIDANCE",
                audit_message=f"Hướng dẫn in WHT cho chứng từ {doc_no}"
            )

        # -------------------------------------------------------------
        # NHÁNH 6: CẤU HÌNH & THAY ĐỔI THƯ MỤC THEO DÕI NVL THEO THÁNG
        # -------------------------------------------------------------
        elif cmd in ["SET_FOLDER", "CHON_FOLDER", "THU_MUC_NVL"]:
            folder_path = str(ctx.meta.get("folderPath") or "").strip()
            if not folder_path:
                raw_text = str(ctx.meta.get("rawText") or "").strip()
                m = re.search(r'(?:[A-Za-z]:\\[^<>":|?*\r\n\t]+|\\[^<>":|?*\r\n\t]+)', raw_text)
                if m:
                    folder_path = m.group(0).strip()

            if not folder_path:
                current_dir = ctx.config.get("tracking_files_dir") or NvlTrackingService.DEFAULT_TRACKING_DIR
                return CommandResult(
                    status="FAILED",
                    summary=(
                        f"Vui lòng cung cấp đường dẫn thư mục cần cấu hình.\n"
                        f"📁 Thư mục hiện tại đang dùng: `{current_dir}`\n"
                        f"👉 Cú pháp mẫu: `chọn folder: G:\\SAP\\2026_10` hoặc `thư mục nvl: D:\\THEO_DOI\\T11`"
                    ),
                    records=[],
                    headers={},
                    audit_target="SET_FOLDER_NO_PATH",
                    audit_message="Thiếu đường dẫn khi cấu hình thư mục theo dõi NVL"
                )

            clean_folder = folder_path.strip().strip('"').strip("'")
            if not os.path.isdir(clean_folder):
                return CommandResult(
                    status="FAILED",
                    summary=f"❌ Thư mục `{clean_folder}` không tồn tại trên hệ thống Windows. Vui lòng kiểm tra lại đường dẫn.",
                    records=[],
                    headers={},
                    audit_target=clean_folder,
                    audit_message=f"Thư mục không tồn tại: {clean_folder}"
                )

            found_files = NvlTrackingService.find_tracking_files(directory=clean_folder)
            file_names = [f.name for f in found_files[:5]]
            file_list_str = ", ".join(file_names) if file_names else "Chưa có file nào"

            # Lưu cấu hình vào config.json trên đĩa để bền vững qua các phiên khởi động
            config_file = Path(__file__).resolve().parent.parent.parent / "config.json"
            if config_file.exists():
                try:
                    with open(config_file, "r", encoding="utf-8") as f:
                        cfg_data = json.load(f)
                    cfg_data["tracking_files_dir"] = clean_folder
                    with open(config_file, "w", encoding="utf-8") as f:
                        json.dump(cfg_data, f, indent=2, ensure_ascii=False)
                except Exception as save_err:
                    print(f"[SET_FOLDER WARN] Không thể ghi config.json: {save_err}")

            # Cập nhật cache runtime
            ctx.config["tracking_files_dir"] = clean_folder
            NvlTrackingService.DEFAULT_TRACKING_DIR = clean_folder

            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#ffffff;border:1px solid #d0d7de;border-radius:6px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.05);'>"
                f"<div style='font-size:16px;font-weight:bold;color:#1a7f37;margin-bottom:8px;'>✅ ĐÃ CẬP NHẬT THƯ MỤC THEO DÕI NVL THÀNH CÔNG</div>"
                f"<div style='background:#f6f8fa;padding:12px;border-radius:6px;border-left:4px solid #1a7f37;font-size:13px;line-height:1.6;'>"
                f"📁 <b>Đường dẫn mới:</b> <code>{clean_folder}</code><br/>"
                f"📊 <b>Số file theo dõi tìm thấy:</b> <b>{len(found_files)} file</b> (*THEO DOI ODER NVL_P_*.xlsx)<br/>"
                f"📋 <b>File mới nhất:</b> <code>{file_list_str}</code><br/>"
                f"ℹ️ <i>Từ bây giờ, tất cả các yêu cầu kiểm tra tiến độ, đối soát cấp hàng và tính DOH sẽ tự động truy xuất từ thư mục này. Khi sang tháng mới, bạn chỉ cần gửi lệnh: <code>chọn folder: [đường dẫn mới]</code> để chuyển tháng.</i>"
                f"</div></div>"
            )

            return CommandResult(
                status="SUCCEEDED",
                summary=(
                    f"✅ Đã cập nhật thư mục theo dõi NVL thành công: `{clean_folder}`\n"
                    f"- Tìm thấy: {len(found_files)} file theo dõi (*THEO DOI ODER NVL_P_*.xlsx)\n"
                    f"- File mới nhất: {file_list_str}\n"
                    f"- Kể từ bây giờ, Bot sẽ tự động truy xuất dữ liệu từ thư mục này.\n"
                    f"- Khi sang tháng mới, bạn chỉ cần gõ: `chọn folder: [đường dẫn mới]`"
                ),
                records=[{"folder_path": clean_folder, "file_count": len(found_files), "latest_files": file_list_str}],
                headers={"folder_path": "Đường dẫn thư mục", "file_count": "Số lượng file", "latest_files": "File mới nhất"},
                custom_html=custom_html,
                audit_target=clean_folder,
                audit_message=f"Đổi thư mục theo dõi NVL sang {clean_folder} ({len(found_files)} files)"
            )

        # -------------------------------------------------------------
        # NHÁNH 7: CẤU HÌNH & BỔ SUNG MẬT KHẨU FILE ORDER NVL (.XLSX / .XLSB)
        # -------------------------------------------------------------
        elif cmd in ["SET_NVL_PASSWORD", "MAT_KHAU_NVL", "SET_PASSWORD", "PASSWORD_NVL"]:
            new_pwd = str(ctx.meta.get("nvlPassword") or ctx.meta.get("password") or "").strip()
            if not new_pwd:
                raw_text = str(ctx.meta.get("rawText") or "").strip()
                m_pwd = re.search(r"(?:m[aậ]t\s*kh[aẩ]u|pass(?:word)?)\s*(?:nvl|order|theo\s*d[oõ]i|tracking|excel)?[\s:=]+([^\s,;]+)", raw_text, re.IGNORECASE)
                if m_pwd:
                    new_pwd = m_pwd.group(1).strip()

            if not new_pwd:
                return CommandResult(
                    status="NEED_INPUT",
                    summary="Vui lòng cung cấp mật khẩu cần cài đặt theo cú pháp: `mật khẩu nvl: [mật_khẩu]`.",
                    records=[],
                    headers={},
                    audit_target="SET_PASSWORD_NO_PWD",
                    audit_message="Thiếu mật khẩu khi cấu hình mật khẩu NVL"
                )

            # Đăng ký vào runtime
            NvlTrackingService.register_password(new_pwd)

            # Lưu vào config.json
            config_file = Path(__file__).resolve().parent.parent.parent / "config.json"
            if config_file.exists():
                try:
                    with open(config_file, "r", encoding="utf-8") as f:
                        cfg_data = json.load(f)
                    current_pwds = cfg_data.get("tracking_excel_passwords", [])
                    if new_pwd not in current_pwds:
                        current_pwds.insert(0, new_pwd)
                        cfg_data["tracking_excel_passwords"] = current_pwds
                        with open(config_file, "w", encoding="utf-8") as f:
                            json.dump(cfg_data, f, indent=2, ensure_ascii=False)
                except Exception as save_err:
                    print(f"[SET_NVL_PASSWORD WARN] Không thể ghi config.json: {save_err}")

            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#ffffff;border:1px solid #d0d7de;border-radius:6px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.05);'>"
                f"<div style='font-size:16px;font-weight:bold;color:#1a7f37;margin-bottom:8px;'>✅ ĐÃ LƯU MẬT KHẨU FILE ORDER NVL THÀNH CÔNG</div>"
                f"<div style='background:#f6f8fa;padding:12px;border-radius:6px;border-left:4px solid #1a7f37;font-size:13px;line-height:1.6;'>"
                f"🔑 <b>Mật khẩu:</b> <code>{new_pwd}</code><br/>"
                f"🔒 <b>Áp dụng:</b> Tự động mở khóa file <code>.xlsx</code> và <code>.xlsb</code> trong thư mục Theo Dõi Order NVL.<br/>"
                f"ℹ️ <i>Bot đã lưu mật khẩu vào bộ nhớ và cấu hình hệ thống. Giờ sếp có thể gửi lại câu lệnh kiểm tra đơn để Bot tự động truy xuất dữ liệu!</i>"
                f"</div></div>"
            )

            return CommandResult(
                status="SUCCEEDED",
                summary=(
                    f"✅ Đã lưu mật khẩu `{new_pwd}` cho file theo dõi Order NVL!\n"
                    f"- Bot sẽ tự động sử dụng mật khẩu này để mở khóa các file `.xlsx` và `.xlsb` được bảo vệ.\n"
                    f"- Sếp có thể gửi lại câu lệnh tra cứu tiến độ cấp ngay bây giờ."
                ),
                records=[{"password": new_pwd}],
                headers={"password": "Mật khẩu NVL đã lưu"},
                custom_html=custom_html,
                audit_target=new_pwd,
                audit_message=f"Đã lưu mật khẩu NVL: {new_pwd}"
            )

        raise ValueError(f"NvlHandler không hỗ trợ lệnh: {cmd}")
