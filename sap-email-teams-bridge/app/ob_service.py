# -*- coding: utf-8 -*-
"""Dịch vụ Điều phối Outbound Delivery RC65 (Outbound Coordination Service).

Kế thừa và chuẩn hóa toàn bộ nghiệp vụ từ 261010_OB_TOOL_V1_RC65.xlsm:
1. Whole-batch Gate Validation:
   - Kiểm tra toàn bộ danh sách item của đợt giao hàng.
   - Nếu có bất kỳ item nào thiếu tồn hoặc bị thâm hụt (deficit), chặn toàn bộ đợt xuất (DEFICIT_BLOCKED),
     tuyệt đối không cho xuất dở dang làm lệch số liệu đối tác.
2. Điều phối theo 3 phân hệ kho (Multi-Plant Routing):
   - Kho 3001: Điều phối VL02N (Picking/Packing).
   - Kho 4007: Điều phối STO / VL10D (Giao hàng chuyển xưởng).
   - Kho 2001: Điều phối EWM /SCWM/MON (Warehouse Task WT confirmation).
3. Đọc ngược kiểm chứng (Readback Verification) và ghi sổ cái Audit.
"""

from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

from app.number_utils import parse_sap_float, format_qty_display


class ObService:
    """Xử lý kiểm soát cổng lô hàng (Batch Gate) và điều phối Outbound."""

    @classmethod
    def validate_whole_batch_gate(
        cls,
        delivery_no: str,
        delivery_items: List[Dict[str, Any]],
        stock_records: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Kiểm tra điều kiện cổng toàn bộ lô hàng (Whole-batch Gate).

        delivery_items: [{item, material/product, qty/orderQty, batch?}]
        stock_records: [{material/product, batch?, availQty/totalQty, ...}]
        """
        # 1. Tính tổng tồn khả dụng theo material (hoặc theo material + batch nếu có batch chỉ định)
        stock_pool: Dict[str, float] = {}
        stock_by_mat_batch: Dict[Tuple[str, str], float] = {}

        for st in stock_records:
            mat = str(st.get("product") or st.get("material") or "").strip().upper()
            bat = str(st.get("batch") or "").strip().upper()
            qty = parse_sap_float(st.get("availQty") or st.get("totalQty") or st.get("qty") or 0.0)

            if mat:
                stock_pool[mat] = stock_pool.get(mat, 0.0) + qty
                if bat:
                    stock_by_mat_batch[(mat, bat)] = stock_by_mat_batch.get((mat, bat), 0.0) + qty

        evaluated_items: List[Dict[str, Any]] = []
        deficits: List[Dict[str, Any]] = []
        has_deficit = False

        # 2. Đánh giá từng dòng trong delivery
        # Lưu ý: Nếu nhiều dòng trùng material, phải trừ lùi để kiểm tra chính xác tổng nhu cầu
        consumed_stock: Dict[str, float] = {}
        consumed_mat_batch: Dict[Tuple[str, str], float] = {}

        for it in delivery_items:
            item_no = str(it.get("item") or it.get("item_no") or "").strip()
            mat = str(it.get("material") or it.get("product") or "").strip().upper()
            req_qty = abs(parse_sap_float(it.get("qty") or it.get("orderQty") or 0.0))
            bat = str(it.get("batch") or "").strip().upper()

            if bat:
                # Kiểm tra theo lô chỉ định
                total_lot = stock_by_mat_batch.get((mat, bat), 0.0)
                used_lot = consumed_mat_batch.get((mat, bat), 0.0)
                avail_for_line = max(0.0, total_lot - used_lot)
                allocated = min(req_qty, avail_for_line)
                consumed_mat_batch[(mat, bat)] = used_lot + allocated
            else:
                # FIFO / General pool
                total_mat = stock_pool.get(mat, 0.0)
                used_mat = consumed_stock.get(mat, 0.0)
                avail_for_line = max(0.0, total_mat - used_mat)
                allocated = min(req_qty, avail_for_line)
                consumed_stock[mat] = used_mat + allocated

            deficit = max(0.0, req_qty - allocated)
            line_status = "READY" if deficit == 0 else "DEFICIT"

            eval_item = {
                "item_no": item_no,
                "material": mat,
                "batch": bat or "AUTO_FIFO",
                "required_qty": req_qty,
                "available_qty": allocated,
                "deficit_qty": deficit,
                "status": line_status
            }
            evaluated_items.append(eval_item)

            if deficit > 0:
                has_deficit = True
                deficits.append({
                    "item_no": item_no,
                    "material": mat,
                    "batch": bat,
                    "required_qty": req_qty,
                    "shortage_qty": deficit,
                    "reason": f"Thiếu {format_qty_display(deficit)} trên tổng tồn khả dụng"
                })

        gate_passed = not has_deficit
        gate_status = "GATE_PASSED" if gate_passed else "DEFICIT_BLOCKED"

        return {
            "delivery_no": delivery_no,
            "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "gate_passed": gate_passed,
            "gate_status": gate_status,
            "total_items": len(delivery_items),
            "deficit_count": len(deficits),
            "items": evaluated_items,
            "deficits": deficits,
            "action_advice": (
                "Đủ điều kiện xuất hàng. Cho phép điều phối xuất kho."
                if gate_passed
                else f"CHẶN XUẤT HÀNG TOÀN LÔ (DEFICIT_BLOCKED): Có {len(deficits)} dòng thiếu tồn. Không cho xuất dở dang."
            )
        }

    @classmethod
    def dispatch_outbound_simulated(
        cls,
        delivery_no: str,
        plant: str,
        evaluated_gate: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Mô phỏng thực thi điều phối theo phân hệ kho tương ứng."""
        if not evaluated_gate.get("gate_passed"):
            return {
                "success": False,
                "delivery_no": delivery_no,
                "plant": plant,
                "status": "BLOCKED",
                "error": "Lô hàng không vượt qua kiểm tra cổng tồn kho (Whole-batch Gate).",
                "deficits": evaluated_gate.get("deficits", [])
            }

        plant_str = str(plant).strip()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if plant_str == "3001":
            channel = "VL02N_PICK_PACK"
            action_desc = "Xác nhận nhặt hàng (Pick) và đóng gói qua VL02N tại Kho 3001"
            readback_status = "C"  # Completely picked
        elif plant_str == "4007":
            channel = "VL10D_STO_TRANSFER"
            action_desc = "Giao hàng điều chuyển nội bộ STO qua VL10D tại Kho 4007"
            readback_status = "C"
        elif plant_str == "2001":
            channel = "EWM_MON_WT_CONFIRM"
            action_desc = "Tạo và xác nhận Warehouse Task (WT) qua EWM /SCWM/MON tại Kho 2001"
            readback_status = "CONFIRMED"
        else:
            channel = "STANDARD_VL02N"
            action_desc = f"Điều phối tiêu chuẩn tại Plant {plant_str}"
            readback_status = "C"

        return {
            "success": True,
            "delivery_no": delivery_no,
            "plant": plant_str,
            "channel": channel,
            "action": action_desc,
            "readback_status": readback_status,
            "items_processed": evaluated_gate.get("total_items", 0),
            "dispatched_at": now_str,
            "status": "DISPATCHED"
        }
