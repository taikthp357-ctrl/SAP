# -*- coding: utf-8 -*-
"""Read-only NVL planning and explicit, local SQLite reservations.

No function in this module posts SAP or fabricates SAP documents. Reservation
scope is material/batch/bin/plant/sloc/UoM/stock identity, never material alone.
"""
import hashlib
import json
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from app.number_utils import parse_sap_decimal, parse_sap_datetime_sort_key, format_qty_display
from app.nvl_master import NvlMasterService

ZERO = Decimal("0")


def _text(value):
    return str(value or "").strip().upper()


def _first(row, *names):
    for name in names:
        if name in row and row[name] is not None and row[name] != "":
            return row[name]
    return None


def _json(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _now():
    return datetime.now(timezone.utc)


def _scope_key(row):
    return _json([_text(row.get(k)) for k in ("material", "batch", "bin", "plant", "sloc", "unit", "stock_id")])


def validated_stock_lot(row, source_plant, source_sloc, require_fifo=True):
    """Return an eligible, typed lot; reject missing evidence and zero availability."""
    material = _text(_first(row, "product", "material"))
    batch, bin_no = _text(row.get("batch")), _text(row.get("bin"))
    unit = _text(_first(row, "unit", "uom", "meins") or "PC")
    plant = _text(_first(row, "plant", "source_plant", "werks") or source_plant)
    sloc = _text(_first(row, "sloc", "source_sloc", "lgort") or source_sloc)
    cat = _text(_first(row, "cat", "stock_type", "stock_category") or "FT")

    if not require_fifo:
        if not bin_no:
            bin_no = "N/A"
        if not cat:
            cat = "UNRESTRICTED"
    if not all((material, batch, bin_no, unit, plant, sloc)):
        raise ValueError("Thiếu material/batch/bin/plant/sloc/ĐVT được xác minh từ nguồn SAP")
    if not cat:
        raise ValueError("Thiếu Stock Category (Cat/Stock Type) được xác minh từ nguồn SAP")
    is_glue = NvlMasterService.is_henkel_glue(material)
    if plant != _text(source_plant):
        raise ValueError("Tồn kho ngoài phạm vi plant yêu cầu")
    if sloc != _text(source_sloc):
        if not (is_glue and sloc == "200C"):
            raise ValueError("Tồn kho ngoài phạm vi plant/sloc yêu cầu")
    if cat not in {"FT", "F1", "F2", "UNRESTRICTED"} or row.get("eligible") is False:
        raise ValueError("Tồn QC/blocked/category chưa xác định không được phân bổ")
    # Zero must not fall back to total stock. Availability must be provided.
    qty = parse_sap_decimal(_first(row, "availQty", "available_qty", "unrest", "qty"), fmt=row.get("number_format"))
    if qty < ZERO:
        raise ValueError("Số lượng tồn khả dụng âm")
    gr_date = _first(row, "grDate", "gr_date")
    gr_time = _first(row, "grTime", "gr_time")
    gr_key = parse_sap_datetime_sort_key(gr_date, gr_time)
    if gr_key[0] in ("99999999", "INVALID_DATE") or gr_key[1] in ("999999", "INVALID_TIME"):
        if require_fifo:
            raise ValueError("Thời điểm GR không hợp lệ trên lịch (INVALID_FIFO_TIMESTAMP)")
    hsd = _first(row, "hsd", "expiry_date")
    hsd_key = parse_sap_datetime_sort_key(hsd, "00:00:00")[0]
    if hsd and hsd_key == "99999999":
        raise ValueError("Hạn sử dụng không hợp lệ")
    if hsd and hsd_key < datetime.now().strftime("%Y%m%d"):
        raise ValueError("Lô đã hết hạn sử dụng")
    # V31 BIN BLOCK: Bỏ qua các Bin block hệ thống (APFL-BL* hoặc Zone BIN BLOCK)
    if NvlMasterService.is_bin_blocked(bin_no):
        raise ValueError(f"BIN BLOCK HE THONG - BO QUA FIFO/OB (Bin {bin_no})")

    # Tự động gán Zone, Box và Area từ Master Data nếu chưa có
    zone_val = str(row.get("zone") or "").strip()
    if not zone_val:
        zone_val = NvlMasterService.get_bin_zone(bin_no)

    box_val = str(row.get("box") or "").strip()
    packing = 0
    b_info = NvlMasterService.get_box_info(material)
    if b_info:
        if not box_val:
            box_val = f"{b_info['model']}/{b_info['box_type']}/Pack {b_info['packing']}"
        packing = int(b_info.get("packing") or 0)

    area_val = str(row.get("area") or "").strip()
    if not area_val:
        area_val = NvlMasterService.get_bin_area(bin_no)

    is_broken = bool(packing > 1 and 0 < qty < packing)

    lot = dict(material=material, batch=batch, bin=bin_no, unit=unit, plant=plant, sloc=sloc,
               stock_id=_text(_first(row, "stock_id", "stockId", "quant")), raw_qty=qty,
               gr_date=str(gr_date or ""), gr_time=str(gr_time or ""), gr_key=gr_key,
               hsd=str(hsd or ""), hsd_score=hsd_key, zone=zone_val, box=box_val,
               station=str(row.get("station") or "").strip(),
               area=area_val, packing=packing, is_broken=is_broken,
               orig_index=int(row.get("orig_index", 0)))
    lot["stock_key"] = _scope_key(lot)
    return lot


class NvlService:
    DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "data" / "bridge_state.db"

    @classmethod
    @contextmanager
    def _get_connection(cls, db_path=None):
        path = Path(db_path or cls.DEFAULT_DB_PATH)
        path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(path), timeout=15.0)
        conn.row_factory = sqlite3.Row
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    @classmethod
    def init_tables(cls, db_path=None):
        with cls._get_connection(db_path) as conn:
            conn.execute("""CREATE TABLE IF NOT EXISTS nvl_plans (
                plan_id TEXT PRIMARY KEY, version INTEGER DEFAULT 1, status TEXT DEFAULT 'DRAFT',
                source_plant TEXT, source_sloc TEXT, dest_plant TEXT, dest_sloc TEXT,
                created_by TEXT, created_at TEXT, total_ordered REAL, total_allocated REAL,
                mat_doc TEXT, payload_json TEXT)""")
            conn.execute("""CREATE TABLE IF NOT EXISTS nvl_reservations (
                res_id TEXT PRIMARY KEY, plan_id TEXT, material TEXT, batch TEXT, qty REAL,
                reserved_at TEXT, status TEXT DEFAULT 'ACTIVE',
                FOREIGN KEY(plan_id) REFERENCES nvl_plans(plan_id))""")
            for table, columns in {
                "nvl_plans": {"request_id": "TEXT", "intent_hash": "TEXT", "expires_at": "TEXT"},
                "nvl_reservations": {"stock_key": "TEXT", "expires_at": "TEXT"},
            }.items():
                present = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
                for column, kind in columns.items():
                    if column not in present:
                        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {kind}")
            conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS nvl_request_key ON nvl_plans(request_id) WHERE request_id IS NOT NULL AND request_id != ''")
            conn.execute("CREATE INDEX IF NOT EXISTS nvl_reservation_scope ON nvl_reservations(stock_key, status)")
            # Legacy active reservations retain a blocking footprint until reviewed.
            conn.execute("UPDATE nvl_reservations SET status='REVIEW_REQUIRED' WHERE status='ACTIVE' AND (stock_key IS NULL OR stock_key='')")

    @staticmethod
    def parse_date_score(value):
        return parse_sap_datetime_sort_key(value, "00:00:00")[0]

    @classmethod
    def _expire(cls, conn):
        now = _now().isoformat()
        conn.execute("UPDATE nvl_reservations SET status='EXPIRED' WHERE status='ACTIVE' AND expires_at <= ?", (now,))
        conn.execute("UPDATE nvl_plans SET status='EXPIRED' WHERE status='LOCAL_RESERVED' AND expires_at <= ?", (now,))

    @classmethod
    def get_active_reservations(cls, db_path=None, by_stock_key=False):
        """Compatibility aggregate; allocation always uses the full stock key."""
        cls.init_tables(db_path)
        result = {}
        with cls._get_connection(db_path) as conn:
            cls._expire(conn)
            for row in conn.execute("SELECT material,batch,stock_key,qty FROM nvl_reservations WHERE status IN ('ACTIVE','REVIEW_REQUIRED')"):
                key = row["stock_key"] if by_stock_key else (_text(row["material"]), _text(row["batch"]))
                result[key] = result.get(key, 0.0) + float(row["qty"])
        return result

    @classmethod
    def allocate_orders(cls, orders, stock_records, source_plant="3093", source_sloc="1005",
                        default_dest_plant="3093", default_dest_sloc="3001", db_path=None,
                        request_id=None, source_complete=False, reservation_ttl_seconds=900):
        """Preview from a complete snapshot; no reservation is created here.

        A source adapter must prove scope and completeness. Missing proof returns
        BLOCKED. Pass a durable request_id so crash/replay returns the same plan.
        """
        cls.init_tables(db_path)
        intent = {"orders": orders, "source_plant": source_plant, "source_sloc": source_sloc,
                  "dest_plant": default_dest_plant, "dest_sloc": default_dest_sloc}
        intent_hash = hashlib.sha256(_json(intent).encode()).hexdigest()
        with cls._get_connection(db_path) as conn:
            cls._expire(conn)
            if request_id:
                old = conn.execute("SELECT payload_json,intent_hash,status FROM nvl_plans WHERE request_id=?", (request_id,)).fetchone()
                if old:
                    if old["intent_hash"] != intent_hash:
                        raise ValueError("Request ID đã dùng cho nhu cầu khác")
                    old_plan = json.loads(old["payload_json"])
                    old_plan["status"] = old["status"]
                    old_plan["replayed"] = True
                    return old_plan
            reservations = list(conn.execute("SELECT * FROM nvl_reservations WHERE status IN ('ACTIVE','REVIEW_REQUIRED')"))
        reserved, legacy_materials = {}, set()
        for row in reservations:
            if row["status"] == "REVIEW_REQUIRED":
                legacy_materials.add(_text(row["material"]))
            else:
                key = row["stock_key"]
                reserved[key] = reserved.get(key, ZERO) + Decimal(str(row["qty"]))
        errors, exclusions, validated_orders, lots = [], [], [], {}
        if not source_complete:
            errors.append("Nguồn tồn kho chưa chứng minh đầy đủ; chưa được giữ tồn")
        if not orders:
            errors.append("Danh sách nhu cầu rỗng")
        for index, order in enumerate(orders):
            try:
                mat = _text(_first(order, "material", "product"))
                unit = _text(_first(order, "unit", "uom") or "PC")
                qty = parse_sap_decimal(_first(order, "qty", "orderQty"), fmt=order.get("number_format"))
                station = str(_first(order, "station", "tram") or "").strip()
                if not mat or not unit or qty <= ZERO or not station:
                    raise ValueError("Cần mã vật tư, ĐVT, trạm và số lượng > 0")
                if mat in legacy_materials:
                    raise ValueError("Có reservation bản cũ chưa xác minh phạm vi; cần đối chiếu")
                validated_orders.append((index, order, mat, unit, qty, station))
            except ValueError as exc:
                errors.append(f"Nhu cầu {index + 1}: {exc}")
        for index, stock in enumerate(stock_records):
            try:
                stock_copy = dict(stock)
                stock_copy.setdefault("orig_index", index)
                lot = validated_stock_lot(stock_copy, source_plant, source_sloc)
                key = lot["stock_key"]
                if key in lots:
                    # Nhiều HU/Pallet trong cùng 1 Bin & Batch trên /SCWM/POST (không có stock_id)
                    # Cộng dồn tồn khả dụng thay vì chặn toàn bộ kế hoạch
                    lots[key]["raw_qty"] += lot["raw_qty"]
                    lots[key]["available_qty"] = max(ZERO, lots[key]["raw_qty"] - reserved.get(key, ZERO))
                    p = lots[key].get("packing", 0)
                    lots[key]["is_broken"] = bool(p > 1 and 0 < lots[key]["available_qty"] < p)
                    continue
                lot["available_qty"] = max(ZERO, lot["raw_qty"] - reserved.get(key, ZERO))
                p = lot.get("packing", 0)
                lot["is_broken"] = bool(p > 1 and 0 < lot["available_qty"] < p)
                lots[key] = lot
            except ValueError as exc:
                exclusions.append({"row": index + 1, "reason": str(exc)})
        items, shortages = [], []
        total_ordered = sum((order[4] for order in validated_orders), ZERO)
        total_allocated = ZERO
        # If any intent or source validation failed, do not partially reserve another order.
        if not errors:
            for index, order, material, unit, qty, station in validated_orders:
                remaining = qty
                st_zone = NvlMasterService.extract_zone_from_station(station)
                target_bin = str(order.get("target_bin") or order.get("bin") or "").strip().upper()

                def _lot_sort_key(lot):
                    # 1. FEFO (HSD earliest first) - Tuyệt đối số 1
                    hsd = lot.get("hsd_score", "99999999")
                    # 2. FIFO (GR Date & Time earliest first) - Tuyệt đối số 2
                    gr = lot.get("gr_key", ("99999999", "999999"))
                    # 3. Target Bin Matching: nếu chỉ định vị trí cụ thể, ưu tiên tuyệt đối bin đó
                    bin_str = str(lot.get("bin") or "").upper()
                    bin_match = 0 if (target_bin and target_bin == bin_str) else (1 if target_bin else 0)
                    # 4. Zone Matching: nếu trạm có Zone, ưu tiên bin có chứa số Zone hoặc khu vực tương thích
                    zone_str = str(lot.get("zone") or "").upper()
                    zone_match = 0 if (st_zone and (st_zone in bin_str or st_zone in zone_str)) else 1
                    # 5. Vét hàng lẻ (Broken Box / Partial Quantity First):
                    # Ô bin có tồn lẻ (0 < available_qty < packing) được ưu tiên vét trước
                    broken_priority = 0 if lot.get("is_broken") else 1
                    # 6. Ưu tiên mặt đất (Ground Level First): sàn/mặt đất lấy dễ hơn tầng cao (HR1-*-*-2+)
                    ground_priority = 0 if NvlMasterService.is_ground_bin(bin_str) else 1
                    # 7. Giữ nguyên thứ tự dòng gốc từ /SCWM/POST (như Excel V31 origRow)
                    orig_idx = lot.get("orig_index", 0)
                    # 8. Fallback cố định
                    stock_k = lot.get("stock_key", "")
                    return (hsd, gr, bin_match, zone_match, broken_priority, ground_priority, orig_idx, stock_k)

                candidate_lots = [lot for lot in lots.values() if lot["material"] == material]
                candidate_lots.sort(key=_lot_sort_key)

                for lot in candidate_lots:
                    if lot["unit"] != unit and not ({lot["unit"], unit} <= {"PC", "EA"}):
                        continue
                    requested_batch = _text(order.get("batch"))
                    if requested_batch and lot["batch"] != requested_batch:
                        continue
                    take = min(remaining, lot["available_qty"])
                    if take <= ZERO:
                        continue
                    lot["available_qty"] -= take
                    p = lot.get("packing", 0)
                    lot["is_broken"] = bool(p > 1 and 0 < lot["available_qty"] < p)
                    remaining -= take
                    total_allocated += take
                    item = {key: lot[key] for key in ("material", "batch", "bin", "zone", "area", "box", "hsd", "gr_date", "stock_key", "stock_id")}
                    item.update(item_no=len(items) + 1, order_index=index, station=station, unit=lot["unit"],
                                allocated_qty=float(take), snapshot_qty=float(lot["raw_qty"]),
                                source_plant=lot["plant"], source_sloc=lot["sloc"],
                                dest_plant=_text(order.get("dest_plant") or default_dest_plant),
                                dest_sloc=_text(order.get("dest_sloc") or default_dest_sloc), status="PLANNED")
                    items.append(item)
                    if remaining <= ZERO:
                        break
                if remaining > ZERO:
                    shortages.append(dict(material=material, station=station, requested_qty=float(qty),
                                          allocated_qty=float(qty - remaining), shortage_qty=float(remaining), unit=unit,
                                          reason="Không đủ tồn đủ điều kiện sau giữ chỗ cục bộ; kiểm tra nguồn và dòng bị loại"))
        now = _now()
        return dict(plan_id="PLAN-NVL-" + (uuid.uuid5(uuid.NAMESPACE_URL, str(request_id)).hex if request_id else uuid.uuid4().hex).upper(),
                    request_id=request_id, intent_hash=intent_hash, version=1, status="BLOCKED" if errors else "DRAFT",
                    source_plant=source_plant, source_sloc=source_sloc, dest_plant=default_dest_plant,
                    dest_sloc=default_dest_sloc, created_at=now.isoformat(),
                    expires_at=(now + timedelta(seconds=max(1, min(int(reservation_ttl_seconds), 3600)))).isoformat(),
                    source_complete=bool(source_complete), total_ordered=float(total_ordered), total_allocated=float(total_allocated),
                    is_full_allocation=bool(orders and not errors and not shortages and total_ordered > ZERO and total_allocated == total_ordered),
                    allocated_items=items, shortages=shortages, errors=errors, excluded_stock=exclusions,
                    material_document=None, reservation_scope="LOCAL_ONLY")

    @classmethod
    def save_plan_and_reserve(cls, plan, created_by="BOT", db_path=None):
        """Atomically reserve against the snapshot after rechecking concurrent claims."""
        cls.init_tables(db_path)
        plan_id = str(plan["plan_id"])
        with cls._get_connection(db_path) as conn:
            conn.execute("BEGIN IMMEDIATE")
            cls._expire(conn)
            old = conn.execute("SELECT payload_json FROM nvl_plans WHERE plan_id=?", (plan_id,)).fetchone()
            if old:
                old_payload = json.loads(old["payload_json"])
                def immutable_payload(payload):
                    return {k: v for k, v in payload.items() if k not in {"status", "replayed"}}
                if immutable_payload(old_payload) != immutable_payload(plan):
                    raise ValueError("Plan ID đã tồn tại với nội dung khác")
                return plan_id
            if plan.get("status") != "DRAFT" or not plan.get("source_complete") or not plan.get("allocated_items"):
                raise ValueError("Kế hoạch bị chặn, rỗng hoặc nguồn chưa đầy đủ")
            if not plan.get("expires_at") or datetime.fromisoformat(plan["expires_at"]) <= _now():
                raise ValueError("Snapshot hết hạn; cần yêu cầu mới")
            quantities, capacities = {}, {}
            for item in plan["allocated_items"]:
                key = item["stock_key"]
                actual_key = _scope_key(dict(item, plant=item["source_plant"], sloc=item["source_sloc"]))
                if key != actual_key:
                    raise ValueError("Khóa scope không khớp item")
                qty, capacity = parse_sap_decimal(item["allocated_qty"]), parse_sap_decimal(item["snapshot_qty"])
                if qty <= ZERO or capacity < qty:
                    raise ValueError("Số lượng giữ chỗ không hợp lệ")
                quantities[key] = quantities.get(key, ZERO) + qty
                if key in capacities and capacities[key] != capacity:
                    raise ValueError("Snapshot cùng khóa nhưng khác số lượng")
                capacities[key] = capacity
                legacy = conn.execute("SELECT 1 FROM nvl_reservations WHERE status='REVIEW_REQUIRED' AND material=? LIMIT 1", (item["material"],)).fetchone()
                if legacy:
                    raise ValueError("Reservation bản cũ cần đối chiếu trước khi giữ tồn")
            for key, qty in quantities.items():
                current = sum((Decimal(str(row[0])) for row in conn.execute("SELECT qty FROM nvl_reservations WHERE stock_key=? AND status='ACTIVE'", (key,))), ZERO)
                if current + qty > capacities[key]:
                    raise ValueError("Tồn đã được yêu cầu khác giữ chỗ; lập lại kế hoạch")
            stored_plan = dict(plan, status="LOCAL_RESERVED")
            conn.execute("""INSERT INTO nvl_plans (plan_id,version,status,source_plant,source_sloc,dest_plant,dest_sloc,
                created_by,created_at,total_ordered,total_allocated,mat_doc,payload_json,request_id,intent_hash,expires_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (plan_id, plan.get("version", 1), "LOCAL_RESERVED", plan["source_plant"], plan["source_sloc"],
                 plan["dest_plant"], plan["dest_sloc"], created_by, plan["created_at"], plan["total_ordered"],
                 plan["total_allocated"], None, _json(stored_plan), plan.get("request_id"), plan.get("intent_hash"), plan["expires_at"]))
            by_key = {item["stock_key"]: item for item in plan["allocated_items"]}
            for key, qty in quantities.items():
                item = by_key[key]
                res_id = "RES-" + uuid.uuid5(uuid.NAMESPACE_URL, plan_id + key).hex
                conn.execute("""INSERT INTO nvl_reservations (res_id,plan_id,material,batch,qty,reserved_at,status,stock_key,expires_at)
                    VALUES (?,?,?,?,?,?,'ACTIVE',?,?)""", (res_id, plan_id, item["material"], item["batch"], float(qty), _now().isoformat(), key, plan["expires_at"]))
        return plan_id

    @classmethod
    def confirm_plan_reservation(cls, plan_id, mat_doc, db_path=None):
        """Xác nhận reservation sau khi SAP Post thành công chứng từ."""
        cls.init_tables(db_path)
        with cls._get_connection(db_path) as conn:
            conn.execute("UPDATE nvl_plans SET status='CONFIRMED', mat_doc=? WHERE plan_id=?", (mat_doc, str(plan_id)))
            conn.execute("UPDATE nvl_reservations SET status='CONFIRMED' WHERE plan_id=?", (str(plan_id),))

    @classmethod
    def release_plan_reservation(cls, plan_id, db_path=None):
        """Hủy reservation khi SAP Post thất bại hoặc bị hủy."""
        cls.init_tables(db_path)
        with cls._get_connection(db_path) as conn:
            conn.execute("UPDATE nvl_plans SET status='RELEASED' WHERE plan_id=?", (str(plan_id),))
            conn.execute("UPDATE nvl_reservations SET status='RELEASED' WHERE plan_id=?", (str(plan_id),))

    @classmethod
    def validate_migo_precheck(cls, plan):
        errors, warnings = list(plan.get("errors", [])), []
        if not plan.get("source_complete"):
            errors.append("Nguồn chưa được xác nhận đầy đủ")
        if not plan.get("allocated_items"):
            errors.append("Kế hoạch không có dòng phân bổ")
        if plan.get("shortages"):
            warnings.append("Kế hoạch phân bổ một phần; chưa phải chứng từ SAP")
        for item in plan.get("allocated_items", []):
            if not all(item.get(k) for k in ("material", "batch", "unit", "source_plant", "source_sloc", "dest_plant", "dest_sloc")):
                errors.append("Thiếu thông tin scope/ĐVT/chứng từ")
            try:
                if parse_sap_decimal(item.get("allocated_qty")) <= ZERO:
                    errors.append("Số lượng phải > 0")
            except ValueError:
                errors.append("Số lượng không hợp lệ")
            if (item.get("source_plant"), item.get("source_sloc")) == (item.get("dest_plant"), item.get("dest_sloc")):
                errors.append("Kho xuất và nhận trùng nhau")
        return not errors, errors, warnings

    @classmethod
    def execute_migo_311_simulated(cls, plan_id, db_path=None):
        """Preview only: never change plan/reservations or generate a document."""
        cls.init_tables(db_path)
        with cls._get_connection(db_path) as conn:
            row = conn.execute("SELECT payload_json FROM nvl_plans WHERE plan_id=?", (plan_id,)).fetchone()
            if not row:
                raise ValueError(f"Không tìm thấy kế hoạch {plan_id}")
            plan = json.loads(row["payload_json"])
        valid, errors, warnings = cls.validate_migo_precheck(plan)
        return dict(success=False, simulation_valid=valid, executed=False, plan_id=plan_id,
                    material_document=None, executed_items=0, preview_items=len(plan.get("allocated_items", [])),
                    errors=errors, warnings=warnings, status="SIMULATED" if valid else "VALIDATION_FAILED")

    @classmethod
    def release_reservations(cls, plan_id, db_path=None):
        cls.init_tables(db_path)
        with cls._get_connection(db_path) as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("UPDATE nvl_plans SET status='CANCELLED' WHERE plan_id=? AND status IN ('DRAFT','LOCAL_RESERVED','EXPIRED')", (plan_id,))
            conn.execute("UPDATE nvl_reservations SET status='RELEASED' WHERE plan_id=? AND status IN ('ACTIVE','EXPIRED','REVIEW_REQUIRED')", (plan_id,))

    @classmethod
    def group_into_ob_groups(
        cls,
        allocated_items: List[Dict[str, Any]],
        shortages: Optional[List[Dict[str, Any]]] = None,
        delivery_no: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Gom cụm theo Trạm/Zone và cộng dồn các dòng phân bổ NVL thành các nhóm OB (OB1, OB2, ...).
        - Đánh số nhóm OB rút gọn: 'OB1', 'OB2', ... (chuẩn 3 ký tự).
        - Gán số chứng từ Outbound Delivery (10xxxxxxxx, VD: 1012924052) hoặc mô phỏng.
        - Cộng dồn số lượng nếu trùng (nhóm OB, Trạm, Mã NVL, Batch, Bin).
        - Thêm các dòng cảnh báo nếu có thiếu tồn (Shortage).
        """
        results: List[Dict[str, Any]] = []
        if not allocated_items and not shortages:
            return results

        # 1. Đánh số nhóm OB theo Trạm / Cụm Zone
        station_to_ob: Dict[str, str] = {}
        ob_counter = 1
        for item in allocated_items:
            station = str(item.get("station") or "KV-LINE").strip()
            if station not in station_to_ob:
                station_to_ob[station] = f"OB{ob_counter}"
                ob_counter += 1

        # 2. Cộng dồn các dòng có cùng (OB group, Trạm, Mã NVL, Batch, Bin)
        consolidated: Dict[Tuple[str, str, str, str, str], Dict[str, Any]] = {}
        for item in allocated_items:
            st = str(item.get("station") or "KV-LINE").strip()
            ob_tag = station_to_ob.get(st, f"OB{ob_counter}")
            mat = str(item.get("material") or "").strip().upper()
            batch = str(item.get("batch") or "-").strip().upper()
            bin_no = str(item.get("bin") or "-").strip().upper()
            qty = Decimal(str(item.get("allocated_qty", 0)))
            unit = str(item.get("unit") or "PC").strip().upper()
            key = (ob_tag, st, mat, batch, bin_no)
            if key not in consolidated:
                deliv_clean = str(delivery_no or "").strip()
                if deliv_clean and deliv_clean != "---" and not deliv_clean.startswith("Mô phỏng"):
                    deliv_str = deliv_clean
                    status_text = "✅ Đã tạo OB"
                elif deliv_clean.startswith("Mô phỏng"):
                    deliv_str = deliv_clean
                    status_text = "📋 Kế hoạch (Preview)"
                else:
                    deliv_str = "---"
                    status_text = "⚠️ Chưa tạo OB"
                consolidated[key] = {
                    "stt": ob_tag,
                    "delivery_no": deliv_str,
                    "material": mat,
                    "allocated_qty": qty,
                    "batch": batch,
                    "bin": bin_no,
                    "station": st,
                    "unit": unit,
                    "status": status_text,
                    "hsd": item.get("hsd", ""),
                    "gr_date": item.get("gr_date", "")
                }
            else:
                consolidated[key]["allocated_qty"] += qty

        for row in consolidated.values():
            q_val = float(row["allocated_qty"])
            row["allocated_qty"] = int(q_val) if q_val.is_integer() else q_val
            results.append(row)

        # 3. Bổ sung các dòng thiếu tồn (Shortages) nếu có
        for s in (shortages or []):
            st = str(s.get("station") or "KV-LINE").strip()
            ob_tag = station_to_ob.get(st, f"OB{ob_counter}")
            mat = str(s.get("material") or "").strip().upper()
            short_qty = s.get("shortage_qty", 0)
            req_qty = s.get("requested_qty", 0)
            alloc_qty = s.get("allocated_qty", 0)
            q_display = int(short_qty) if isinstance(short_qty, (int, float)) and float(short_qty).is_integer() else short_qty
            a_display = int(alloc_qty) if isinstance(alloc_qty, (int, float)) and float(alloc_qty).is_integer() else alloc_qty
            r_display = int(req_qty) if isinstance(req_qty, (int, float)) and float(req_qty).is_integer() else req_qty
            results.append({
                "stt": ob_tag,
                "delivery_no": "---",
                "material": mat,
                "allocated_qty": f"Thiếu {q_display}",
                "batch": "---",
                "bin": "---",
                "station": st,
                "unit": s.get("unit", "PC"),
                "status": f"⚠️ Thiếu tồn ({a_display}/{r_display})",
                "hsd": "",
                "gr_date": ""
            })

        return results
