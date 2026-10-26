# -*- coding: utf-8 -*-
"""Outbound stock preview and RC65 baseline architecture field tokens / table controls."""
from datetime import datetime, timezone
from decimal import Decimal
from typing import Any, Dict, List, Optional

from app.nvl_service import _text, _first, validated_stock_lot
from app.number_utils import parse_sap_decimal, format_qty_display

ZERO = Decimal("0")

# ==============================================================================
# BASELINE 261010_OB_TOOL_V1_RC65 (RC65_ModuleSAP.bas) ARCHITECTURE CONSTANTS
# ==============================================================================

# Khóa cứng các field ID token theo chuẩn RC65 (tuyệt đối không quét đoán mò)
RC65_FIELD_TOKENS = {
    "QTY_OVERVIEW": "LIPSD-G_LFIMG",      # Số lượng xuất trên Overview Table
    "STORAGE_LOCATION": "LIPS-LGORT",     # Kho xuất / nhập (SLoc)
    "BATCH": "LIPS-CHARG",                # Số lô (Batch)
    "ITEM_NO": "LIPS-POSNR",              # Số dòng giao hàng (Item)
    "MATERIAL": "LIPS-MATNR",             # Mã vật tư / thành phẩm (Material)
    "QTY_SPLIT": "LIPS-LFIMG",            # Số lượng trên bảng tách lô (Batch Split Table)
    "RECEIVING_PLANT": "LIPS-UMWRK",      # Nhà máy nhận (Receiving Plant)
    "REF_STO": "LIPS-VGBEL",              # Đơn STO tham chiếu (Reference Document)
    "REF_STO_ITEM": "LIPS-VGPOS",         # Dòng STO tham chiếu (Reference Item)
}

# Khóa cứng mã kỹ thuật cột ALV VL06O chuẩn RC65 (OB_ReadVL06O)
RC65_VL06O_COLUMNS = {
    "DELIVERY": "VBELN",                  # Số đơn Outbound Delivery
    "ITEM": "POSNR",                      # Dòng Delivery
    "MATERIAL": "MATNR",                  # Mã vật tư / thành phẩm
    "QTY": "LFIMG",                       # Số lượng giao hàng
    "UOM": "VRKME",                       # Đơn vị tính bán hàng / giao hàng
    "SLOC": "LGORT",                      # Vị trí lưu kho
    "BATCH": "CHARG",                     # Số lô
    "SHIP_TO": "KUNWE",                   # Nơi nhận hàng (hoặc KUNNR)
    "GM_STATUS": "WBSTA",                 # Trạng thái xuất hàng Goods Movement (hoặc WBSTK)
    "REF_STO": "VGBEL",                   # Số đơn STO tham chiếu
    "CREATE_DATE": "ERDAT",               # Ngày tạo chứng từ
}

# Danh mục cột ALV VL06O chuẩn hóa theo RC65 (khóa cứng ưu tiên mã kỹ thuật SAP)
RC65_VL06O_COLUMN_MAPPING = {
    "obNo":       ["VBELN", "DELIVERY", "OUTBOUND DELIVERY"],
    "item":       ["POSNR", "DELIVERY ITEM", "ITEM"],
    "product":    ["MATNR", "MATERIAL", "PRODUCT"],
    "qty":        ["LFIMG", "DELIVERY QTY", "DLV.QTY", "ACTUAL DELIVERED QTY", "MENGE", "QUANTITY"],
    "unit":       ["VRKME", "SALES UNIT", "DELIVERY UNIT", "MEINS", "BASE UNIT", "UNIT"],
    "sloc":       ["LGORT", "STORAGE LOCATION", "SLOC"],
    "batch":      ["S_CHARG", "CHARG", "BATCH"],
    "shipTo":     ["KUNNR", "KUNWE", "SHIP-TO PARTY", "SHIP TO"],
    "gmStatus":   ["WBSTA", "WBSTK", "GOODS MOVEMENT STATUS", "GM STATUS"],
    "refSTO":     ["VGBEL", "REFERENCE DOCUMENT", "STO"],
    "createDate": ["ERDAT", "CREATED ON", "DATE"]
}

# ID và Token bảng TableControl chuẩn RC65
RC65_TABLE_OVERVIEW_ID = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER"
)
RC65_TABLE_OVERVIEW_TOKEN = "tblSAPMV50ATC_LIPS_OVER"
RC65_TABLE_SPLIT_TOKEN = "tblSAPMV50ATC_LIPS_CHND"


class ObService:
    _field_paths: Dict[str, str] = {}

    @classmethod
    def resolve_rc65_tokens(cls) -> Dict[str, str]:
        """Trả về toàn bộ các field token đã khóa cứng theo baseline RC65."""
        return dict(RC65_FIELD_TOKENS)

    @classmethod
    def get_table_field(cls, tbl, token: str, visible_row: int):
        """Truy xuất control ô trong SAP TableControl theo token chuẩn RC65 (TableField).

        Khóa cứng theo token relative ID (LIPSD-G_LFIMG, LIPS-LGORT, LIPS-CHARG, LIPS-POSNR, LIPS-MATNR)
        thay vì đoán mò toạ độ hoặc chỉ số cột có thể thay đổi giữa các cấu hình màn hình SAP.
        """
        if tbl is None:
            return None

        tbl_id = str(getattr(tbl, "Id", getattr(tbl, "id", "")) or "")
        cache_key = f"{tbl_id}|{token}"

        if cache_key in cls._field_paths:
            relative_prefix = cls._field_paths[cache_key]
            rel_id = f"{relative_prefix}{visible_row}]"
            try:
                ch = tbl.findById(rel_id)
                if ch is not None:
                    return ch
            except Exception:
                pass
            cls._field_paths.pop(cache_key, None)

        children = getattr(tbl, "Children", getattr(tbl, "children", None))
        if children is None:
            return None

        cnt = getattr(children, "Count", getattr(children, "count", 0))
        target_token = f"{token.upper()}["

        for i in range(cnt):
            try:
                ch = children(i) if callable(children) else children[i]
            except Exception:
                try:
                    fn = getattr(children, "ElementAt", getattr(children, "Item", None))
                    ch = fn(i) if callable(fn) else None
                except Exception:
                    continue

            cid = str(getattr(ch, "Id", getattr(ch, "id", "")) or "")
            if target_token in cid.upper():
                at = cid.rfind("[")
                if at != -1 and cid.endswith("]"):
                    coord_str = cid[at + 1:-1]
                    parts = coord_str.split(",")
                    if len(parts) == 2:
                        try:
                            row_coord = int(parts[1])
                            if row_coord == visible_row:
                                if tbl_id and cid.upper().startswith(tbl_id.upper()):
                                    rel_part = cid[len(tbl_id) + 1:]
                                    comma_idx = rel_part.rfind(",")
                                    if comma_idx != -1:
                                        cls._field_paths[cache_key] = rel_part[:comma_idx + 1]
                                return ch
                        except (ValueError, TypeError):
                            pass

        return None

    @classmethod
    def set_table_text(cls, tbl, token: str, visible_row: int, text: str) -> bool:
        """Ghi giá trị text vào ô theo token chuẩn RC65 (SetTableText) kèm xác nhận readback."""
        fld = cls.get_table_field(tbl, token, visible_row)
        if fld is None:
            return False
        val_str = str(text or "").strip()
        for prop in ("Text", "text"):
            try:
                setattr(fld, prop, val_str)
                break
            except Exception:
                pass
        readback = str(getattr(fld, "Text", getattr(fld, "text", "")) or "").strip()
        return readback == val_str

    @classmethod
    def locate_delivery_item(cls, tbl, item_no: str, overview_index: Optional[Dict[str, Any]] = None) -> int:
        """Định vị dòng hiển thị của item trong bảng Delivery Overview (LocateItem) theo chuẩn RC65."""
        if tbl is None:
            return -1
        try:
            visible_rows = int(getattr(tbl, "VisibleRowCount", getattr(tbl, "visibleRowCount", 15)))
        except Exception:
            visible_rows = 15

        target_item = str(item_no).strip()

        # Step 1: Kiểm tra expected row từ index nếu có
        if overview_index and target_item in overview_index:
            expected_abs_row = overview_index[target_item].get("abs_row", -1)
            try:
                pos = int(getattr(tbl.VerticalScrollbar, "Position", 0))
                cand_row = expected_abs_row - pos
                if 0 <= cand_row < visible_rows:
                    fld = cls.get_table_field(tbl, RC65_FIELD_TOKENS["ITEM_NO"], cand_row)
                    if fld is not None:
                        val = str(getattr(fld, "Text", getattr(fld, "text", "")) or "").strip()
                        if (val.lstrip("0") or "0") == (target_item.lstrip("0") or "0"):
                            return cand_row
            except Exception:
                pass

        # Step 2: Quét các dòng đang hiển thị
        for r in range(visible_rows):
            fld = cls.get_table_field(tbl, RC65_FIELD_TOKENS["ITEM_NO"], r)
            if fld is not None:
                val = str(getattr(fld, "Text", getattr(fld, "text", "")) or "").strip()
                if (val.lstrip("0") or "0") == (target_item.lstrip("0") or "0"):
                    return r
        return -1

    @classmethod
    def snapshot_delivery_overview(cls, tbl, index_only: bool = False) -> Dict[str, Dict[str, Any]]:
        """Chụp snapshot toàn bộ các dòng trên Delivery Overview (OverviewSnapshot) chuẩn RC65."""
        if tbl is None:
            return {}

        try:
            visible_rows = int(getattr(tbl, "VisibleRowCount", getattr(tbl, "visibleRowCount", 15)))
        except Exception:
            visible_rows = 15

        result = {}
        for r in range(visible_rows):
            fld_item = cls.get_table_field(tbl, RC65_FIELD_TOKENS["ITEM_NO"], r)
            if fld_item is None:
                continue
            item_no = str(getattr(fld_item, "Text", getattr(fld_item, "text", "")) or "").strip()
            if not item_no:
                continue

            fld_mat = cls.get_table_field(tbl, RC65_FIELD_TOKENS["MATERIAL"], r)
            mat = str(getattr(fld_mat, "Text", getattr(fld_mat, "text", "")) or "").strip() if fld_mat else ""

            row_data = {
                "item_no": item_no,
                "material": mat,
                "abs_row": r
            }

            if not index_only:
                fld_qty = cls.get_table_field(tbl, RC65_FIELD_TOKENS["QTY_OVERVIEW"], r)
                fld_sloc = cls.get_table_field(tbl, RC65_FIELD_TOKENS["STORAGE_LOCATION"], r)
                fld_batch = cls.get_table_field(tbl, RC65_FIELD_TOKENS["BATCH"], r)

                row_data["qty"] = str(getattr(fld_qty, "Text", getattr(fld_qty, "text", "")) or "").strip() if fld_qty else ""
                row_data["sloc"] = str(getattr(fld_sloc, "Text", getattr(fld_sloc, "text", "")) or "").strip() if fld_sloc else ""
                row_data["batch"] = str(getattr(fld_batch, "Text", getattr(fld_batch, "text", "")) or "").strip() if fld_batch else ""

            clean_key = item_no.lstrip("0") or "0"
            result[clean_key] = row_data

        return result
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
