# -*- coding: utf-8 -*-
"""Handler chuyên trách nghiệp vụ ra đơn NVL MIGO 311 và đối soát Posting Change Kho 1005."""
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
            tracking_dir = ctx.config.get("tracking_files_dir")

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

            headers = {
                "material": "Mã NVL",
                "model": "Model xe",
                "shift_chay": "Ca chạy",
                "station": "Trạm cấp",
                "order_qty": "SL Order",
                "total_cap": "Đã cấp",
                "gap": "Còn thiếu (Gap)",
                "sp_display": "Số Phiếu / OB",
                "status_pick": "TT nhặt hàng"
            }

            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#ffffff;border:1px solid #d0d7de;"
                f"border-radius:6px;padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.05);'>"
                f"<div style='font-size:15px;font-weight:bold;color:#0969da;margin-bottom:10px;'>"
                f"🔍 KẾT QUẢ ĐỐI SOÁT & ĐỊNH DANH CẤP HÀNG NVL (Theo File Theo Dõi & SAP)"
                f"</div>"
                f"<div style='background:#f6f8fa;border-left:4px solid #0969da;padding:12px;border-radius:4px;font-size:13px;line-height:1.6;margin-bottom:12px;'>"
                f"{trace_res['summary']}"
                f"</div>"
                f"</div>"
            )

            return CommandResult(
                status="SUCCEEDED",
                summary=trace_res["summary"].replace("<br/>", " | ").replace("<b>", "").replace("</b>", "").replace("<code>", "").replace("</code>", "")[:250] + "...",
                records=trace_res["records"],
                headers=headers,
                custom_html=custom_html,
                audit_target=", ".join(mats[:5]),
                audit_message=f"Định danh cấp NVL đa Model ({len(trace_res['records'])} dòng khớp)"
            )

        # -------------------------------------------------------------
        # NHÁNH 4: TRA CỨU & HƯỚNG DẪN IN PHIẾU WHT / WT TRONG SAP
        # -------------------------------------------------------------
        elif cmd in ["IN_WHT", "PRINT_WHT", "HUONG_DAN_IN_WHT"]:
            doc_no = ctx.meta.get("obNo") or ctx.meta.get("materialDocument") or ""
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

        raise ValueError(f"NvlHandler không hỗ trợ lệnh: {cmd}")
