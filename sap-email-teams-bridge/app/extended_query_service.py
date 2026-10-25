# -*- coding: utf-8 -*-
"""Dịch vụ Tra cứu mở rộng IB / BOM / Queue / ASN (Extended Query Service).

Kế thừa và chuẩn hóa toàn bộ nghiệp vụ tra cứu từ các tool VBA:
1. CS03 BOM Component Explosion:
   - Tra cứu cấu trúc linh kiện định mức, tỷ lệ tiêu hao, phân biệt thành phẩm có BOM hay linh kiện đơn.
2. /SCWM/MON Queue Diagnostics:
   - Bóc tách nguyên nhân kẹt hàng đợi EWM (SYSFAIL, STOP, RETRY), đưa ra hướng xử lý trực quan cho quản lý kho.
3. ZMME240 / VL06I ASN & Provenance Tracking:
   - Theo dõi xuất xứ lô hàng: Vận đơn (Bill of Lading), Số Container, Nhà cung cấp (Vendor), PO và tiến độ hải quan.
"""

import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from app.number_utils import parse_sap_float, format_qty_display


class ExtendedQueryService:
    """Xử lý tra cứu nâng cao cấu trúc BOM, chuẩn đoán Queue và nguồn gốc ASN."""

    @classmethod
    def explode_bom_structure(
        cls,
        parent_mat: str,
        plant: str,
        raw_bom_items: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Phân tích và định dạng danh sách linh kiện định mức (BOM Structure).

        raw_bom_items: [{component/idnrk, description/ojtxp, qty/menge, unit/meins, scrap/ausfm, ...}]
        """
        parent_clean = str(parent_mat).strip().upper()
        plant_clean = str(plant).strip()

        if not raw_bom_items:
            return {
                "parent_material": parent_clean,
                "plant": plant_clean,
                "has_bom": False,
                "total_components": 0,
                "components": [],
                "summary": f"Vật tư {parent_clean} tại Plant {plant_clean} không có cấu trúc BOM trên SAP."
            }

        formatted_components: List[Dict[str, Any]] = []
        tot_qty = 0.0

        for idx, it in enumerate(raw_bom_items, 1):
            comp = str(it.get("component") or it.get("idnrk") or it.get("matnr") or "").strip().upper()
            desc = str(it.get("description") or it.get("ojtxp") or it.get("maktx") or "").strip()
            qty = parse_sap_float(it.get("qty") or it.get("menge") or 0.0)
            unit = str(it.get("unit") or it.get("meins") or "PC").strip()
            scrap = parse_sap_float(it.get("scrap") or it.get("ausfm") or 0.0)

            tot_qty += qty
            formatted_components.append({
                "stt": idx,
                "parent": parent_clean,
                "component": comp,
                "description": desc,
                "qty": qty,
                "unit": unit,
                "scrap_percent": scrap,
                "status": "CÓ BOM"
            })

        return {
            "parent_material": parent_clean,
            "plant": plant_clean,
            "has_bom": True,
            "total_components": len(formatted_components),
            "total_base_qty": tot_qty,
            "components": formatted_components,
            "summary": f"Định mức BOM {parent_clean} gồm {len(formatted_components)} linh kiện tại Plant {plant_clean}."
        }

    @classmethod
    def diagnose_ewm_queues(
        cls,
        raw_queue_rows: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Chuẩn đoán tình trạng lỗi kẹt hàng đợi EWM (/SCWM/MON Queue Monitor).

        raw_queue_rows: [{queue_name/qname, error_message/errtext, document/docno, user/uname, ...}]
        """
        if not raw_queue_rows:
            return {
                "is_clean": True,
                "stuck_count": 0,
                "diagnostics": [],
                "summary": "Tất cả hàng đợi EWM hoạt động bình thường, không có Queue bị kẹt."
            }

        diagnostics: List[Dict[str, Any]] = []

        for idx, q in enumerate(raw_queue_rows, 1):
            q_name = str(q.get("queue_name") or q.get("qname") or "").strip()
            err_msg = str(q.get("error_message") or q.get("errtext") or "").strip()
            doc_no = str(q.get("document") or q.get("docno") or "").strip()

            # Phân loại mức độ và nguyên nhân lỗi
            err_upper = err_msg.upper()
            if "SYSFAIL" in err_upper or "SYSTEM FAILURE" in err_upper:
                category = "CRITICAL_SYSFAIL"
                advice = "Lỗi hệ thống RFC/EWM. Cần kiểm tra SMQ2 và liên hệ đội BASIS SAP."
            elif "STOP" in err_upper or "LOCKED" in err_upper or "BLOCKED" in err_upper:
                category = "QUEUE_STOPPED"
                advice = "Queue đang bị dừng (Locked). Cần Unlock và Restart Queue trong SMQ2."
            elif "PACKAGING" in err_upper or "PACK" in err_upper or "QUAN" in err_upper:
                category = "MASTER_DATA_PACKAGING"
                advice = "Sai lệch dữ liệu bao bì hoặc đơn vị đóng gói (Packaging Spec / Master Data)."
            elif "POSTING" in err_upper or "PERIOD" in err_upper:
                category = "POSTING_PERIOD_ERROR"
                advice = "Kỳ kế toán chưa mở hoặc ngày chứng từ không hợp lệ."
            else:
                category = "OPERATIONAL_ERROR"
                advice = "Thao tác kho bị xung đột hoặc thiếu vị trí Bin đích."

            diagnostics.append({
                "stt": idx,
                "queue_name": q_name,
                "document": doc_no,
                "error_category": category,
                "error_message": err_msg,
                "action_advice": advice
            })

        return {
            "is_clean": False,
            "stuck_count": len(diagnostics),
            "diagnostics": diagnostics,
            "summary": f"Cảnh báo: Phát hiện {len(diagnostics)} Queue EWM đang bị tắc nghẽn!"
        }

    @classmethod
    def track_asn_provenance(
        cls,
        inbound_no: str,
        asn_items: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Tổng hợp hồ sơ nguồn gốc lô hàng ASN (Inbound / Container / Bill of Lading).

        asn_items: [{product, batch, qty, refNo, poNo, vendor, eta, container, bill, customs_status, ...}]
        """
        ib_clean = str(inbound_no).strip()
        if not asn_items:
            return {
                "inbound_no": ib_clean,
                "found": False,
                "items": [],
                "summary": f"Không tìm thấy thông tin ASN cho số Inbound {ib_clean}."
            }

        first = asn_items[0]
        container_no = str(first.get("container") or first.get("cont_no") or "N/A").strip()
        bill_of_lading = str(first.get("bill") or first.get("refNo") or "N/A").strip()
        vendor = str(first.get("vendor") or first.get("supplier") or "N/A").strip()
        eta = str(first.get("eta") or first.get("arrival_date") or "N/A").strip()
        customs = str(first.get("customs_status") or "ĐÃ THÔNG QUAN").strip()

        rows = []
        tot_qty = 0.0
        for it in asn_items:
            mat = str(it.get("product") or it.get("material") or "").strip()
            bat = str(it.get("batch") or "").strip()
            qty = parse_sap_float(it.get("qty") or 0.0)
            unit = str(it.get("unit") or "PC").strip()
            po = str(it.get("poNo") or "").strip()
            tot_qty += qty

            rows.append({
                "material": mat,
                "batch": bat,
                "qty": qty,
                "unit": unit,
                "po_no": po,
                "status": it.get("status", "IN_TRANSIT")
            })

        return {
            "inbound_no": ib_clean,
            "found": True,
            "container_no": container_no,
            "bill_of_lading": bill_of_lading,
            "vendor": vendor,
            "eta": eta,
            "customs_status": customs,
            "total_items": len(rows),
            "total_qty": tot_qty,
            "items": rows,
            "summary": f"Hồ sơ Inbound {ib_clean} (Cont: {container_no} | Bill: {bill_of_lading}) gồm {len(rows)} mã hàng."
        }
