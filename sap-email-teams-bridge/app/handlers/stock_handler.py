# -*- coding: utf-8 -*-
"""Handler chuyên trách tra cứu tồn kho, vị trí Bin EWM, lịch sử MB51 và cảnh báo tồn thấp."""
import re
from datetime import datetime, timedelta
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler
from app.sap_mb51_service import SapMb51Service, parse_mb51_date_range, generate_mb51_table_html


class StockHandler(BaseHandler):
    """Xử lý nhóm lệnh tra cứu tồn kho & vị trí:
    - VITRI, BIN, POST (EWM /SCWM/POST + Fallback MB52)
    - TON_KHO, TONKHO (Kho 1005 & 3093)
    - TON, MB52 (Toàn bộ SLoc)
    - LICHSU, MB51 (Lịch sử giao dịch vật tư)
    - TON_THAP, LOW_STOCK (Cảnh báo tồn theo Model/Shop)
    - LICHSU_BIN, MON, WHT (Lịch sử chuyển bin /SCWM/MON)
    - TIM_GR_DOC, GR_DOC, CHUNGTU_GR (Chứng từ nhập kho GR MB51)
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
        # 1. TRA CỨU VỊ TRÍ BIN EWM (/SCWM/POST) & SMART COMBO CHECK
        # -------------------------------------------------------------
        if cmd in ["VITRI", "VỊ TRÍ", "BIN", "POST"]:
            records = ctx.sap.check_position_posting_change(resolved_mats, ctx.batch, plant=ctx.plant)
            headers = {
                "product": "Mã hàng", "batch": "Batch", "bin": "Vị trí (Bin)",
                "totalQty": "Tổng tồn", "availQty": "Khả dụng",
                "unit": "ĐVT", "grDate": "Ngày GR", "grTime": "Giờ GR",
                "cat": "Trạng thái"
            }
            summary = f"Vị trí lưu kho & Giờ GR (/SCWM/POST) cho mã: {', '.join(resolved_mats)}{mapping_info}"

            # Smart Combo Check: Nếu /SCWM/POST trả về 0 dòng, tự động fallback kiểm tra MB52 xem còn tồn ở SLoc ngoài EWM không
            if not records and resolved_mats:
                try:
                    mb52_fb = ctx.sap.check_stock_mb52(resolved_mats, plant=ctx.plant)
                    if mb52_fb:
                        sloc_summary = []
                        for r_fb in mb52_fb:
                            sloc_summary.append(f"SLoc {r_fb.get('sloc', '')}: {r_fb.get('unrest', '0')} {r_fb.get('unit', '')}")
                        sloc_str = ", ".join(sloc_summary[:4])
                        summary = f"Trên EWM (Kho {ctx.plant}) hiện không còn Bin nào, nhưng trên MB52 đang có tồn tại {sloc_str}!"
                        print(f"[COMBO CHECK] EWM rỗng nhưng MB52 có tồn: {sloc_str}")
                        for r_fb in mb52_fb:
                            records.append({
                                "product": r_fb.get("product", ""),
                                "batch": r_fb.get("batch", ""),
                                "bin": f"SLoc {r_fb.get('sloc', '')} (Ngoài EWM)",
                                "totalQty": r_fb.get("unrest", "0"),
                                "availQty": r_fb.get("unrest", "0"),
                                "unit": r_fb.get("unit", ""),
                                "grDate": "Xem MB52",
                                "grTime": "",
                                "cat": "Tồn SLoc"
                            })
                except Exception as fb_err:
                    print(f"[COMBO CHECK WARN] Không thể fallback sang MB52: {fb_err}")

        # -------------------------------------------------------------
        # 2. TRA CỨU TỒN KHO 1005 & 3093 (MB52 FILTERED)
        # -------------------------------------------------------------
        elif cmd in ["TON_KHO", "TONKHO"]:
            wh_slocs = ctx.config.get("warehouse_slocs", ["1005", "3093"])
            records = ctx.sap.check_stock_mb52(resolved_mats, ctx.plant, slocs=wh_slocs)
            headers = {
                "product": "Mã hàng", "plant": "Plant", "sloc": "Sloc",
                "batch": "Batch", "unrest": "Tồn", "unit": "ĐVT"
            }
            summary = f"Tồn kho (Kho 1005 & 3093 - Plant {ctx.plant}) cho mã: {', '.join(resolved_mats)}{mapping_info}"

        # -------------------------------------------------------------
        # 3. TRA CỨU TỒN TỔNG HỢP TOÀN BỘ SLOC (MB52)
        # -------------------------------------------------------------
        elif cmd in ["TON", "TỒN", "MB52"]:
            records = ctx.sap.check_stock_mb52(resolved_mats, ctx.plant)
            headers = {
                "product": "Mã hàng", "plant": "Plant", "sloc": "Sloc",
                "batch": "Batch", "unrest": "Tồn", "unit": "ĐVT"
            }
            summary = f"Tồn tổng hợp (Toàn bộ SLoc - Plant {ctx.plant}) cho mã: {', '.join(resolved_mats)}{mapping_info}"

        # -------------------------------------------------------------
        # 4. TRA CỨU LỊCH SỬ GIAO DỊCH VẬT TƯ (MB51)
        # -------------------------------------------------------------
        elif cmd in ["LICHSU", "LỊCH SỬ", "MB51", "LICH_SU_MB51", "MB51_HISTORY"]:
            explicit_p = ctx.meta.get("explicitPlant", False)
            target_plant = ctx.plant if explicit_p else ""
            raw_text = ctx.meta.get("rawText", "")

            # 1. Trích xuất khoảng ngày (Date Range)
            date_from = ctx.meta.get("dateFrom")
            date_to = ctx.meta.get("dateTo")
            if not date_from or not date_to:
                dt_from, dt_to = parse_mb51_date_range(raw_text, default_days=ctx.days or 7)
            else:
                dt_from = datetime.strptime(date_from, "%d.%m.%Y") if isinstance(date_from, str) else date_from
                dt_to = datetime.strptime(date_to, "%d.%m.%Y") if isinstance(date_to, str) else date_to

            # 2. Tra cứu dữ liệu (Ưu tiên SapMb51Service chống Error 617, fallback sang ctx.sap)
            records = []
            mb51_service = SapMb51Service(default_plant=target_plant or "3093")
            sap_session = getattr(ctx.sap, "session", None) if ctx.sap else None

            try:
                records = mb51_service.fetch_movement_history(
                    materials=resolved_mats,
                    date_from=dt_from,
                    date_to=dt_to,
                    plant=target_plant,
                    session=sap_session
                )
            except Exception:
                records = []

            # Nếu service trả về rỗng (hoặc chạy trong môi trường mock/test nơi ctx.sap được mock):
            if not records and ctx.sap and hasattr(ctx.sap, "check_history_mb51"):
                try:
                    delta_days = max(1, (dt_to - dt_from).days + 1) if dt_from and dt_to else ctx.days
                    raw_recs = ctx.sap.check_history_mb51(resolved_mats, delta_days, target_plant, explicit_plant=explicit_p)
                    for r in raw_recs:
                        qty_val = str(r.get("qty", ""))
                        records.append({
                            "budat": r.get("postDate", ""),
                            "matnr": r.get("product", ""),
                            "werks": r.get("plant", target_plant or "3093"),
                            "lgort": r.get("sloc", ""),
                            "umlgo": r.get("umlgo", ""),
                            "bwart": r.get("mvtType", ""),
                            "qty": qty_val,
                            "meins": r.get("unit", ""),
                            "ebeln": r.get("ebeln", r.get("docNo", "")),
                            "usnam": r.get("usnam", ""),
                            "charg": r.get("batch", ""),
                            "mblnr": r.get("docNo", ""),
                        })
                except Exception:
                    pass

            # Đồng bộ các alias cho record để tương thích 100% với các component khác
            for r in records:
                if "product" not in r: r["product"] = r.get("matnr", "")
                if "docNo" not in r: r["docNo"] = r.get("mblnr", r.get("ebeln", ""))
                if "postDate" not in r: r["postDate"] = r.get("budat", "")
                if "mvtType" not in r: r["mvtType"] = r.get("bwart", "")
                if "unit" not in r: r["unit"] = r.get("meins", "")
                if "sloc" not in r: r["sloc"] = r.get("lgort", "")
                if "batch" not in r: r["batch"] = r.get("charg", "")
                if "plant" not in r: r["plant"] = r.get("werks", "")

            headers = {
                "budat": "Ngày HT",
                "matnr": "Mã NVL",
                "werks": "Plant",
                "lgort": "Kho",
                "umlgo": "Kho đích",
                "bwart": "Loại GD",
                "qty": "Số lượng",
                "meins": "ĐVT",
                "ebeln": "Số PO/Đơn",
                "usnam": "User SAP",
            }
            plant_info = f" (Plant {ctx.plant})" if explicit_p else " (Tất cả nhà máy)"
            range_info = f"từ {dt_from.strftime('%d/%m/%Y')} đến {dt_to.strftime('%d/%m/%Y')}" if dt_from.date() != dt_to.date() else f"ngày {dt_from.strftime('%d/%m/%Y')}"
            summary = f"Lịch sử giao dịch vật tư (MB51 - {range_info}){plant_info} cho mã: {', '.join(resolved_mats)}{mapping_info}"

            if records:
                custom_html = generate_mb51_table_html(records, summary)

        # -------------------------------------------------------------
        # 5. CẢNH BÁO TỒN THẤP & THIẾU DOH THEO CA / MODEL
        # -------------------------------------------------------------
        elif cmd in ["TON_THAP", "LOW_STOCK", "CANH_BAO_DOH", "DOH_CA", "TON_THAP_CA"]:
            raw_t = str(ctx.meta.get("rawText") or "").lower()
            is_shift_doh = (cmd in ["CANH_BAO_DOH", "DOH_CA", "TON_THAP_CA"]) or any(
                k in raw_t for k in ["ca ", "ca s", "doh", "chua du doh", "thieu doh", "canh bao ton ca"]
            )
            model_name = ctx.meta.get("model", "LIMO")
            shift_name = ctx.meta.get("shift") or "S6"

            if is_shift_doh:
                threshold = int(ctx.meta.get("threshold") or 200)
                limit_cnt = ctx.meta.get("limit", 50)
                records = ctx.sap.check_shift_doh_warning(
                    materials=resolved_mats,
                    model_name=model_name,
                    shift_name=shift_name,
                    plant=ctx.plant,
                    threshold=threshold,
                    limit=limit_cnt
                )
                headers = {
                    "product": "Mã hàng",
                    "ncc": "Nhà cung cấp",
                    "stock_3001": "Tồn Line 3001",
                    "stock_1005": "Tồn Kho 1005",
                    "doh": "DOH (Xe)",
                    "model": "Model"
                }
                summary = (
                    f"Cảnh báo tồn thấp & thiếu DOH ca {shift_name} (Model: {model_name}) - "
                    f"Có {len(records)} mã chưa đủ DOH (< {threshold} xe)"
                )
            else:
                shop_name = ctx.meta.get("shop", "BODY")
                target_plants = ctx.meta.get("plants", ["3093", "3001"] if shop_name == "BODY" else ["3093", "2001", "3001"])
                limit_cnt = ctx.meta.get("limit", 20)
                full_model = ctx.meta.get("fullModelList", False)
                records = ctx.sap.check_low_stock(resolved_mats, target_plants, limit=limit_cnt, model_name=model_name, shop_name=shop_name, full_model_list=full_model)
            if shop_name == "BODY":
                headers = {
                    "product": "Material",
                    "us": "US",
                    "line_body": "Line Body",
                    "bd_kho": "BD Kho",
                    "bd_carset": "BD Carset",
                    "transit": "Transit",
                    "kho_ma": "Tồn MA",
                    "line_ma": "Tồn Line",
                    "model": "Model",
                    "station": "Trạm",
                    "status": "Cảnh báo"
                }
            elif shop_name == "XMD":
                headers = {
                    "product": "Mã hàng", "station": "Trạm", "model": "Model",
                    "stock_3093": "Kho 3093", "stock_2001": "XMĐ 2001", "stock_3001": "Sơn 3001",
                    "total_stock": "Tổng tồn", "unit": "ĐVT", "status": "Cảnh báo"
                }
            else:
                headers = {
                    "product": "Mã hàng", "station": "Trạm", "model": "Model",
                    "stock_3093": "Kho FG 3093", "stock_3001": "Line Body 3001",
                    "total_stock": "Tổng tồn", "unit": "ĐVT", "status": "Cảnh báo"
                }
            list_mode = "Toàn bộ danh sách" if full_model else "Ưu tiên xuất"
            summary = f"Cảnh báo tồn kho Body - Model {model_name} ({len(records)} mã {list_mode} theo chuẩn Stock_FG_MA_BODY)"

        # -------------------------------------------------------------
        # 6. LỊCH SỬ CHUYỂN BIN-TO-BIN & CANCEL WHT (/SCWM/MON)
        # -------------------------------------------------------------
        elif cmd in ["LICHSU_BIN", "MON", "WHT"]:
            bin_name = ctx.meta.get("bin", "")
            records = ctx.sap.check_bin_history_mon(resolved_mats, bin_name=bin_name, days=ctx.days)
            headers = {
                "wtNo": "Số WT", "product": "Mã hàng", "batch": "Số lô",
                "srcBin": "Bin xuất", "destBin": "Bin đích",
                "qty": "SL chuyển", "unit": "ĐVT", "status": "Trạng thái",
                "createDate": "Ngày tạo", "user": "Người xử lý"
            }
            target_label = f"Bin {bin_name}" if bin_name else f"Mã {', '.join(resolved_mats)}"
            summary = f"Lịch sử chuyển Bin-to-Bin & Cancel WHT (/SCWM/MON) cho: {target_label}"

        # -------------------------------------------------------------
        # 7. TÌM CHỨNG TỪ NHẬP KHO GR (MB51)
        # -------------------------------------------------------------
        elif cmd in ["TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR"]:
            inbound_no = ctx.meta.get("inboundNo", "")
            ref_no = ctx.meta.get("refNo", "")
            mat_doc = ctx.meta.get("materialDocument", "")
            records = ctx.sap.check_gr_doc_mb51(inbound_no=inbound_no, ref_no=ref_no, mat_doc=mat_doc, materials=resolved_mats, plant=ctx.plant, days=ctx.days)
            headers = {
                "mblnr": "Số CT GR", "mjahr": "Năm", "item": "Item", "bwart": "Loại GD",
                "matnr": "Mã hàng", "qty": "SL", "unit": "ĐVT",
                "xblnr": "Inbound / Ref", "postDate": "Ngày HT", "sloc": "Kho", "batch": "Số lô", "user": "Người tạo"
            }
            target_lbl = inbound_no or ref_no or mat_doc or (', '.join(resolved_mats) if resolved_mats else 'Tất cả')
            summary = f"Chứng từ nhập kho GR (MB51) cho: {target_lbl}"

        return CommandResult(
            status=result_status,
            summary=summary,
            records=records,
            headers=headers,
            custom_html=custom_html,
            audit_target=", ".join(resolved_mats[:5]) if resolved_mats else (ctx.meta.get("bin", "") or ""),
            audit_message=f"{cmd}: {summary}"
        )
