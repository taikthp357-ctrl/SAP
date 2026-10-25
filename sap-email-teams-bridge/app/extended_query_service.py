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
import math
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
        raw_bom_items: List[Dict[str, Any]],
        *, source_complete: bool = True, source_error: str = "",
    ) -> Dict[str, Any]:
        """Phân tích và định dạng danh sách linh kiện định mức (BOM Structure).

        raw_bom_items: [{component/idnrk, description/ojtxp, qty/menge, unit/meins, scrap/ausfm, ...}]
        """
        parent_clean = str(parent_mat).strip().upper()
        plant_clean = str(plant).strip()

        if source_error or not source_complete:
            return {"parent_material": parent_clean, "plant": plant_clean, "has_bom": None,
                    "total_components": len(raw_bom_items), "components": [], "status": "UNKNOWN",
                    "summary": "Nguồn BOM chưa đầy đủ hoặc tra cứu lỗi; chưa thể kết luận có/không BOM."}
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
        raw_queue_rows: List[Dict[str, Any]],
        *, source_complete: bool = True, source_error: str = "",
    ) -> Dict[str, Any]:
        """Chuẩn đoán tình trạng lỗi kẹt hàng đợi EWM (/SCWM/MON Queue Monitor).

        raw_queue_rows: [{queue_name/qname, error_message/errtext, document/docno, user/uname, ...}]
        """
        if source_error or not source_complete:
            return {"is_clean": None, "stuck_count": None, "diagnostics": [], "status": "UNKNOWN",
                    "summary": "Nguồn Queue chưa đầy đủ hoặc tra cứu lỗi; chưa thể kết luận hàng đợi sạch."}
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
        asn_items: List[Dict[str, Any]],
        *, source_complete: bool = False,
    ) -> Dict[str, Any]:
        """Return only evidence explicitly associated with the requested inbound.

        VL06I is not evidence of customs clearance. Missing fields stay UNKNOWN;
        rows without an inbound identifier are excluded, never assigned to the
        requested inbound just because they arrived in the same result list.
        """
        ib_clean = str(inbound_no or "").strip()
        def norm(value):
            text = str(value or "").strip().upper()
            return (text.lstrip("0") or "0") if text.isdigit() else text
        def value(row, *keys):
            for key in keys:
                if key in row and row[key] is not None and str(row[key]).strip() != "":
                    return row[key]
            return ""
        selected = [r for r in asn_items if ib_clean and norm(value(r, "ibNo", "inboundNo", "inbound_no", "vbeln")) == norm(ib_clean)]
        result = {
            "inbound_no": ib_clean, "found": bool(selected), "items": [],
            "container_no": "UNKNOWN", "bill_of_lading": "UNKNOWN", "vendor": "UNKNOWN",
            "eta": "UNKNOWN", "customs_status": "UNKNOWN", "total_items": 0,
            "total_qty": None, "totals_by_unit": {}, "source_complete": source_complete,
            "excluded_rows": len(asn_items) - len(selected), "status": "UNKNOWN",
            "summary": f"Chưa tìm thấy dòng nguồn xác minh được thuộc Inbound {ib_clean}.",
            "field_values": {},
        }
        if not selected:
            return result
        for field, aliases in {
            "container_no": ("container", "cont_no"), "bill_of_lading": ("bill", "bill_of_lading", "refNo"),
            "vendor": ("vendor", "supplier"), "eta": ("eta", "arrival_date"), "customs_status": ("customs_status",),
        }.items():
            values = [str(value(row, *aliases)).strip() for row in selected]
            nonempty = sorted(set(v for v in values if v))
            result["field_values"][field] = nonempty
            # Partial/conflicting evidence must not present one row as the whole consignment.
            if len(nonempty) == 1 and all(values):
                result[field] = nonempty[0]
            elif len(nonempty) > 1:
                result[field] = "MULTIPLE_VALUES"
        totals = {}
        invalid_quantity = False
        for row in selected:
            qty = parse_sap_float(value(row, "qty", "actualQty"), default=float("nan"))
            valid = math.isfinite(qty) and qty >= 0
            unit = str(value(row, "unit", "uom", "meins")).strip().upper() or "UNKNOWN"
            if valid:
                totals[unit] = totals.get(unit, 0.0) + qty
            else:
                invalid_quantity = True
            result["items"].append({
                "inbound_no": ib_clean,
                "material": str(value(row, "product", "material", "matnr")).strip(),
                "batch": str(value(row, "batch", "charg")).strip(),
                "qty": qty if valid else None, "unit": unit,
                "po_no": str(value(row, "poNo", "ebeln")).strip(),
                "status": str(value(row, "status")) or "UNKNOWN",
            })
        result.update({
            "total_items": len(selected), "totals_by_unit": totals,
            "total_qty": next(iter(totals.values())) if len(totals) == 1 and "UNKNOWN" not in totals and not invalid_quantity else None,
            "status": "FOUND" if source_complete and not invalid_quantity else "PARTIAL",
            "summary": f"Có {len(selected)} dòng xác minh thuộc Inbound {ib_clean}; trường thiếu nguồn được ghi UNKNOWN.",
        })
        return result
