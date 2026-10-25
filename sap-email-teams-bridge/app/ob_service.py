# -*- coding: utf-8 -*-
"""Outbound stock preview. No SAP dispatch or confirmation is implemented."""
from datetime import datetime, timezone
from decimal import Decimal

from app.nvl_service import _text, _first, validated_stock_lot
from app.number_utils import parse_sap_decimal, format_qty_display

ZERO = Decimal("0")


class ObService:
    @classmethod
    def validate_whole_batch_gate(cls, delivery_no, delivery_items, stock_records,
                                  source_complete=False, source_plant=None, source_sloc=None):
        """Evaluate all lines against one shared stock pool, preserving UoM/scope.

        Batch-specific requirements are allocated first, then unrestricted-batch
        requirements consume the same remaining lots. A passed gate is a preview,
        never authorization/evidence of an SAP operation.
        """
        errors, excluded, demands, lots = [], [], [], {}
        if not source_complete:
            errors.append("Nguồn tồn/nhu cầu chưa chứng minh đầy đủ")
        if not delivery_no or not delivery_items:
            errors.append("Thiếu delivery hoặc danh sách dòng giao hàng")
        for index, item in enumerate(delivery_items):
            try:
                material = _text(_first(item, "material", "product"))
                unit = _text(_first(item, "unit", "uom"))
                plant = _text(_first(item, "plant", "source_plant") or source_plant)
                sloc = _text(_first(item, "sloc", "source_sloc") or source_sloc)
                quantity = parse_sap_decimal(_first(item, "qty", "orderQty", "openQty"), fmt=item.get("number_format"))
                if not all((material, unit, plant, sloc)) or quantity <= ZERO:
                    raise ValueError("Cần material/ĐVT/plant/sloc và số lượng > 0")
                demands.append(dict(index=index, item_no=str(_first(item, "item", "item_no") or index + 1),
                                    material=material, batch=_text(item.get("batch")), unit=unit,
                                    plant=plant, sloc=sloc, required=quantity))
            except ValueError as exc:
                errors.append(f"Dòng yêu cầu {index + 1}: {exc}")
        scopes = {(item["plant"], item["sloc"]) for item in demands}
        for index, stock in enumerate(stock_records):
            try:
                plant = _text(_first(stock, "plant", "source_plant"))
                sloc = _text(_first(stock, "sloc", "source_sloc"))
                if (plant, sloc) not in scopes:
                    raise ValueError("Tồn không thuộc plant/sloc của delivery")
                lot = validated_stock_lot(stock, plant, sloc, require_fifo=False)
                if lot["stock_key"] in lots:
                    errors.append(f"Dòng tồn {index + 1}: khóa tồn trùng; cần stock_id riêng")
                    continue
                lot["remaining"] = lot["raw_qty"]
                lots[lot["stock_key"]] = lot
            except ValueError as exc:
                excluded.append({"row": index + 1, "reason": str(exc)})
        evaluated, deficits = {}, []
        for demand in sorted(demands, key=lambda d: (not bool(d["batch"]), d["index"])):
            remaining = demand["required"]
            if not errors:
                for key in sorted(lots):
                    lot = lots[key]
                    if any(lot[field] != demand[field] for field in ("material", "unit", "plant", "sloc")):
                        continue
                    if demand["batch"] and demand["batch"] != lot["batch"]:
                        continue
                    take = min(remaining, lot["remaining"])
                    lot["remaining"] -= take
                    remaining -= take
                    if remaining <= ZERO:
                        break
            line = dict(item_no=demand["item_no"], material=demand["material"],
                        batch=demand["batch"] or "AUTO_BATCH", unit=demand["unit"],
                        plant=demand["plant"], sloc=demand["sloc"], required_qty=float(demand["required"]),
                        available_qty=float(demand["required"] - remaining), deficit_qty=float(remaining),
                        status="BLOCKED" if errors else ("READY" if remaining == ZERO else "DEFICIT"))
            evaluated[demand["index"]] = line
            if remaining > ZERO:
                deficits.append(dict(item_no=line["item_no"], material=line["material"], batch=demand["batch"],
                                     unit=demand["unit"], required_qty=line["required_qty"], shortage_qty=float(remaining),
                                     reason=f"Thiếu {format_qty_display(remaining)} tồn đủ điều kiện trong cùng scope/ĐVT"))
        passed = bool(demands and not errors and not deficits)
        status = "GATE_PASSED" if passed else ("VALIDATION_BLOCKED" if errors else "DEFICIT_BLOCKED")
        return dict(delivery_no=delivery_no, checked_at=datetime.now(timezone.utc).isoformat(),
                    gate_passed=passed, gate_status=status, total_items=len(delivery_items),
                    deficit_count=len(deficits), items=[evaluated[i] for i in sorted(evaluated)],
                    deficits=deficits, errors=errors, excluded_stock=excluded, source_complete=bool(source_complete),
                    action_advice="Đủ tồn trong snapshot kiểm tra; chưa thực thi hoặc xác nhận SAP." if passed else
                    f"CHẶN XUẤT HÀNG TOÀN LÔ ({status}): {len(deficits)} dòng thiếu tồn; {len(errors)} lỗi xác minh.")

    @classmethod
    def dispatch_outbound_simulated(cls, delivery_no, plant, evaluated_gate):
        """Preview routing only: never fabricate readback or dispatched evidence."""
        if not evaluated_gate.get("gate_passed"):
            return dict(success=False, executed=False, delivery_no=delivery_no, plant=plant,
                        status="BLOCKED", readback_status=None,
                        error="Lô hàng chưa vượt qua kiểm tra tồn kho", deficits=evaluated_gate.get("deficits", []))
        routes = {"3001": "VL02N_PICK_PACK", "4007": "VL10D_STO_TRANSFER", "2001": "EWM_MON_WT_CONFIRM"}
        plant = _text(plant)
        if plant not in routes:
            return dict(success=False, executed=False, delivery_no=delivery_no, plant=plant, status="BLOCKED",
                        readback_status=None, error="Plant chưa có quy trình điều phối được hỗ trợ", deficits=[])
        return dict(success=False, executed=False, simulation_valid=True, delivery_no=delivery_no, plant=plant,
                    channel=routes[plant], action="Bản xem trước điều phối; chưa gửi lệnh hoặc xác nhận trên SAP.",
                    readback_status=None, items_processed=0, preview_items=evaluated_gate.get("total_items", 0),
                    status="SIMULATED")
