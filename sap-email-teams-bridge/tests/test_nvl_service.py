# -*- coding: utf-8 -*-
"""Unit tests for NvlService (Đợt 4: Ra đơn NVL V31 Zone & MIGO 311)."""

import os
from pathlib import Path
import pytest

from app.nvl_service import NvlService
from app.nvl_master import NvlMasterService


@pytest.fixture
def tmp_db(tmp_path):
    db_file = tmp_path / "test_nvl.db"
    NvlService.init_tables(db_file)
    return db_file


def test_nvl_allocation_fefo_fifo(tmp_db):
    orders = [
        {"material": "BIW32013765AA", "qty": 150, "station": "OP10", "unit": "PC"}
    ]

    # Stock records with different HSD and GR Date
    stock_records = [
        # Lot 1: Far future HSD
        {"product": "BIW32013765AA", "batch": "LOT_LATE", "bin": "B-01-01", "plant": "3093", "sloc": "1005", "unit": "PC", "cat": "FT", "grTime": "08:00:00", "availQty": 100, "hsd": "30/12/2099", "grDate": "01/10/2026"},
        # Lot 2: Near HSD (should be allocated FIRST by FEFO)
        {"product": "BIW32013765AA", "batch": "LOT_EARLY", "bin": "B-01-02", "plant": "3093", "sloc": "1005", "unit": "PC", "cat": "FT", "grTime": "08:00:00", "availQty": 100, "hsd": "01/11/2099", "grDate": "05/10/2026"},
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)

    assert plan["total_ordered"] == 150
    assert plan["total_allocated"] == 150
    assert plan["is_full_allocation"] is True
    assert len(plan["allocated_items"]) == 2

    # Verify Lot 2 (EARLY HSD) was consumed first
    first_item = plan["allocated_items"][0]
    assert first_item["batch"] == "LOT_EARLY"
    assert first_item["allocated_qty"] == 100

    # Remaining 50 taken from Lot 1
    second_item = plan["allocated_items"][1]
    assert second_item["batch"] == "LOT_LATE"
    assert second_item["allocated_qty"] == 50


def test_nvl_stock_reservation_prevents_double_booking(tmp_db):
    orders1 = [
        {"material": "BIW32013765AA", "qty": 80, "station": "OP10", "unit": "PC"}
    ]
    stock_records = [
        {"product": "BIW32013765AA", "batch": "LOT_SINGLE", "bin": "B-01-01", "plant": "3093", "sloc": "1005", "unit": "PC", "cat": "FT", "grTime": "08:00:00", "availQty": 100, "hsd": "15/12/2099", "grDate": "01/10/2026"}
    ]

    # Plan 1 allocates 80 and reserves it
    plan1 = NvlService.allocate_orders(orders1, stock_records, db_path=tmp_db, source_complete=True)
    assert plan1["total_allocated"] == 80
    plan_id_1 = NvlService.save_plan_and_reserve(plan1, db_path=tmp_db)

    # Check active reservations
    res_map = NvlService.get_active_reservations(tmp_db)
    assert res_map[("BIW32013765AA", "LOT_SINGLE")] == 80.0

    # Plan 2 asks for 50 of the same material
    # Only 20 should be available! (100 - 80 reserved = 20)
    orders2 = [
        {"material": "BIW32013765AA", "qty": 50, "station": "OP20", "unit": "PC"}
    ]
    plan2 = NvlService.allocate_orders(orders2, stock_records, db_path=tmp_db, source_complete=True)
    assert plan2["total_allocated"] == 20.0
    assert len(plan2["shortages"]) == 1
    assert plan2["shortages"][0]["shortage_qty"] == 30.0


def test_nvl_migo_precheck_and_execution(tmp_db):
    orders = [{"material": "BIW32013765AA", "qty": 50, "station": "OP10", "unit": "PC"}]
    stock = [{"product": "BIW32013765AA", "batch": "B01", "bin": "BIN-A", "plant": "3093", "sloc": "1005", "unit": "PC", "cat": "FT", "grTime": "08:00:00", "availQty": 50, "grDate": "01/10/2026"}]

    plan = NvlService.allocate_orders(orders, stock, db_path=tmp_db, source_complete=True)
    plan_id = NvlService.save_plan_and_reserve(plan, db_path=tmp_db)

    # Precheck should pass
    is_valid, hard_errors, warnings = NvlService.validate_migo_precheck(plan)
    assert is_valid is True
    assert len(hard_errors) == 0

    # Simulated execution
    exec_res = NvlService.execute_migo_311_simulated(plan_id, db_path=tmp_db)
    assert exec_res["success"] is False
    assert exec_res["simulation_valid"] is True
    assert exec_res["material_document"] is None
    assert exec_res["status"] == "SIMULATED"

    # Simulation must not consume an operational reservation
    res_map = NvlService.get_active_reservations(tmp_db)
    assert res_map.get(("BIW32013765AA", "B01"), 0.0) == 50.0


def test_group_into_ob_groups_consolidation_and_shortages():
    """Kiểm tra gom cụm OB, cộng dồn lô/bin trùng, gán số chứng từ 10xxxxxxxx và hiển thị thiếu tồn."""
    allocated = [
        {"material": "BIW75111244AA", "batch": "240801", "bin": "01-02-03", "allocated_qty": 100, "station": "KV08-Thức/Vũ"},
        {"material": "BIW75111244AA", "batch": "240801", "bin": "01-02-03", "allocated_qty": 100, "station": "KV08-Thức/Vũ"},
        {"material": "BIW75111244AA", "batch": "240801", "bin": "01-02-03", "allocated_qty": 100, "station": "KV08-Thức/Vũ"},
        {"material": "BIW75111244AA", "batch": "240802", "bin": "01-02-04", "allocated_qty": 100, "station": "KV08-Thức/Vũ"},
    ]
    shortages = [
        {"material": "BIW75111244AA", "requested_qty": 100, "allocated_qty": 0, "shortage_qty": 100, "station": "KV08-Thức/Vũ", "unit": "PC"}
    ]

    grouped = NvlService.group_into_ob_groups(allocated, shortages=shortages, delivery_no="1012924052")
    assert len(grouped) == 3

    # Dòng 1: Đã cộng dồn 3 lô 240801 thành 300
    assert grouped[0]["stt"] == "OB1"
    assert grouped[0]["delivery_no"] == "1012924052"
    assert grouped[0]["allocated_qty"] == 300
    assert grouped[0]["batch"] == "240801"
    assert grouped[0]["status"] == "✅ Đã tạo OB"

    # Dòng 2: Lô 240802 số lượng 100
    assert grouped[1]["stt"] == "OB1"
    assert grouped[1]["delivery_no"] == "1012924052"
    assert grouped[1]["allocated_qty"] == 100
    assert grouped[1]["batch"] == "240802"

    # Dòng 3: Dòng thiếu tồn
    assert grouped[2]["stt"] == "OB1"
    assert grouped[2]["delivery_no"] == "---"
    assert "Thiếu 100" in str(grouped[2]["allocated_qty"])
    assert "Thiếu tồn" in grouped[2]["status"]


def test_group_into_ob_groups_no_fake_ob_when_posting_fails():
    """Kiểm tra: Khi chưa tạo được OB hoặc MIGO 311 thất bại, delivery_no='---' và status='⚠️ Chưa tạo OB', tuyệt đối không fake số OB cũ."""
    allocated = [
        {"material": "BIW75111244AA", "batch": "261023-15", "bin": "APFL-P23-01-12", "allocated_qty": 100, "station": "KV08-Thức/Vũ", "unit": "EA"}
    ]
    grouped = NvlService.group_into_ob_groups(allocated, delivery_no="")
    assert len(grouped) == 1
    assert grouped[0]["stt"] == "OB1"
    assert grouped[0]["delivery_no"] == "---"
    assert grouped[0]["status"] == "⚠️ Chưa tạo OB"
    assert grouped[0]["allocated_qty"] == 100


def test_nvl_allocation_scwm_post_fallback_and_unit_coercion(tmp_db):
    """Kiểm tra: /SCWM/POST không có cột plant/sloc và đơn vị EA vẫn khớp với đơn hàng mặc định PC."""
    from app.nvl_service import validated_stock_lot

    # 1. Kiểm tra trực tiếp validated_stock_lot
    raw_scwm_row = {
        "product": "BIW75111244AA",
        "batch": "LOT_P23",
        "bin": "APFL-P23-01-12",
        "availQty": "170",
        "grDate": "20.10.2026",
        "grTime": "08:00:00"
        # Hoàn toàn không có plant, sloc, unit, cat
    }
    lot = validated_stock_lot(raw_scwm_row, source_plant="3093", source_sloc="1005")
    assert lot["plant"] == "3093"
    assert lot["sloc"] == "1005"
    assert lot["unit"] == "PC"
    assert lot["raw_qty"] == 170

    # 2. Kiểm tra phân bổ với 2 dòng tồn thực tế (170 + 297, ĐVT EA) và đơn hàng yêu cầu 100 PC
    orders = [
        {"material": "BIW75111244AA", "qty": 100, "station": "KV08-Thức/Vũ"} # Mặc định PC
    ]
    stock_records = [
        {"product": "BIW75111244AA", "batch": "LOT_P23_1", "bin": "APFL-P23-01-12", "availQty": 170, "unit": "EA", "grDate": "20.10.2026", "grTime": "08:00:00"},
        {"product": "BIW75111244AA", "batch": "LOT_P23_2", "bin": "APFL-P23-01-14", "availQty": 297, "unit": "EA", "grDate": "21.10.2026", "grTime": "09:00:00"},
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)
    assert plan["is_full_allocation"] is True
    assert plan["total_ordered"] == 100
    assert plan["total_allocated"] == 100
    assert len(plan["allocated_items"]) == 1

    item0 = plan["allocated_items"][0]
    assert item0["batch"] == "LOT_P23_1"
    assert item0["bin"] == "APFL-P23-01-12"
    assert item0["allocated_qty"] == 100
    assert item0["unit"] == "EA"  # Lấy theo ĐVT thực tế của lô trên SAP
    assert len(plan["shortages"]) == 0
    assert len(plan["excluded_stock"]) == 0


def test_nvl_duplicate_hu_same_bin_batch_aggregation(tmp_db):
    """Lỗi 1.2: Trên /SCWM/POST nhiều Pallet/HU nằm chung Bin và Batch không bị báo lỗi khóa tồn trùng, mà được cộng dồn."""
    orders = [
        {"material": "BIW32025210AA", "qty": 200, "station": "KV-1/Duy/Bằng ZONE 21", "unit": "EA"}
    ]
    # 2 dòng tồn cùng Material, Batch, Bin nhưng không có stock_id (đặc thù /SCWM/POST)
    stock_records = [
        {"product": "BIW32025210AA", "batch": "LOT_SAME", "bin": "APFL-P21-01-01", "availQty": 120, "unit": "EA", "grDate": "20.10.2026", "grTime": "08:00:00"},
        {"product": "BIW32025210AA", "batch": "LOT_SAME", "bin": "APFL-P21-01-01", "availQty": 100, "unit": "EA", "grDate": "20.10.2026", "grTime": "08:00:00"},
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)
    assert plan["status"] == "DRAFT"
    assert len(plan["errors"]) == 0
    assert plan["total_ordered"] == 200
    assert plan["total_allocated"] == 200
    assert plan["is_full_allocation"] is True
    assert len(plan["allocated_items"]) == 1
    assert plan["allocated_items"][0]["allocated_qty"] == 200.0


def test_nvl_bin_block_system_rejection(tmp_db):
    """Kiểm tra Bin thuộc mẫu BIN BLOCK hệ thống (APFL-BL*) bị loại khỏi danh sách cấp hàng."""
    orders = [
        {"material": "BIW80004410", "qty": 40, "station": "LINE_01", "unit": "PC"}
    ]
    stock_records = [
        {"product": "BIW80004410", "batch": "240130-02", "bin": "APFL-BL01-01-01", "availQty": 46, "unit": "PC", "grDate": "20.10.2026", "grTime": "08:00:00"}
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)
    assert len(plan["excluded_stock"]) == 1
    assert "BIN BLOCK" in plan["excluded_stock"][0]["reason"]
    assert plan["total_allocated"] == 0
    assert len(plan["shortages"]) == 1


def test_nvl_master_v31_zone_and_box_enrichment(tmp_db):
    """Kiểm tra tự động điền Zone và Hàng thùng từ Master Data V31."""
    from app.nvl_master import NvlMasterService

    orders = [
        {"material": "BIW80001331", "qty": 48, "station": "LINE_VF9", "unit": "EA"}
    ]
    stock_records = [
        {"product": "BIW80001331", "batch": "LOT_BOX", "bin": "APFL-P00-01-01", "availQty": 50, "unit": "EA", "grDate": "20.10.2026", "grTime": "08:00:00"}
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)
    assert plan["is_full_allocation"] is True
    item = plan["allocated_items"][0]
    assert item["zone"] == "Hàng Box"
    assert "VF9" in item["box"]
    assert "Pack 24" in item["box"]

    # Kiểm tra hàm tính toán chia hàng thùng / hàng lẻ
    split = NvlMasterService.calculate_box_split("BIW80001331", 100)
    assert split["is_box"] is True
    assert split["packing"] == 24
    assert split["full_boxes"] == 4
    assert split["box_qty"] == 96.0
    assert split["loose_qty"] == 4.0
    assert split["v31_max_boxes"] == 3
    assert split["v31_capped_qty"] == 72.0

def test_station_zone_extraction_and_bin_area():
    """Kiểm tra trích xuất Zone & Team Leader từ chuỗi Trạm và phân loại Bin Area chuẩn Excel V31."""
    assert NvlMasterService.extract_zone_from_station("KV-1/Duy/Bằng ZONE 21") == "21"
    assert NvlMasterService.extract_zone_from_station("KV-12-Thức/Vũ") == "12"
    assert NvlMasterService.extract_zone_from_station("Line Zone 3.1 Chuyền A") == "3.1"
    assert NvlMasterService.extract_zone_from_station("KV 04") == "04"
    assert NvlMasterService.extract_zone_from_station("Chuyền ráp không có số") == ""

    assert NvlMasterService.extract_tl_from_station("KV-1/Duy/Bằng ZONE 21") == "DUY/BẰNG"
    assert NvlMasterService.extract_tl_from_station("KV-12-Thức/Vũ") == "THỨC/VŨ"

    assert NvlMasterService.get_bin_area("APFL-ZF30-01-08") == "ZF"
    assert NvlMasterService.get_bin_area("APFL-P21-01-01") == "P"
    assert NvlMasterService.get_bin_area("SPECIAL-RACK") == "SPECIAL"


def test_broken_box_and_zone_affinity_allocation(tmp_db):
    """Kiểm tra quy tắc vét hàng lẻ (Broken Box First) và ưu tiên Zone matching khớp chuẩn SAP EWM & Excel V31.
    Tái hiện đúng ca thực tế:
    - Nhu cầu: 128 EA mã BIW32025210AA (packing = 128) cho trạm KV-1/Duy/Bằng ZONE 21
    - Tồn kho:
      + Bin 1: APFL-ZF30-01-08: 10 EA (thùng lẻ dở dang)
      + Bin 2: APFL-ZF30-02-01: 150 EA (nguyên thùng cùng khu vực)
      + Bin 3: APFL-ZF06-01-06: 200 EA (nguyên thùng khu vực ZF06 khác)
    - Kết quả mong đợi:
      + Dòng 1: Bin APFL-ZF30-01-08 lấy trọn 10 EA (vét lẻ trước)
      + Dòng 2: Bin APFL-ZF30-02-01 lấy tiếp 118 EA (cùng khu vực)
      + Tổng 10 + 118 = 128 EA; Bin ZF06-01-06 không bị bốc nhầm.
    """
    orders = [
        {"material": "BIW32025210AA", "qty": 128, "station": "KV-1/Duy/Bằng ZONE 21", "unit": "EA"}
    ]

    stock_records = [
        # Dòng 1 trong /SCWM/POST: Bin ZF30-01-08 có 10 EA (thùng dở)
        {"product": "BIW32025210AA", "batch": "260926-10", "bin": "APFL-ZF30-01-08",
         "plant": "3093", "sloc": "1005", "unit": "EA", "cat": "FT",
         "availQty": 10, "grDate": "01/10/2026", "grTime": "08:00:00"},
        # Dòng 2 trong /SCWM/POST: Bin ZF30-02-01 có 150 EA (thùng nguyên)
        {"product": "BIW32025210AA", "batch": "260926-10", "bin": "APFL-ZF30-02-01",
         "plant": "3093", "sloc": "1005", "unit": "EA", "cat": "FT",
         "availQty": 150, "grDate": "01/10/2026", "grTime": "08:00:00"},
        # Dòng 3 trong /SCWM/POST: Bin ZF06-01-06 có 200 EA (ở khu vực xa hơn ZF06)
        {"product": "BIW32025210AA", "batch": "260926-10", "bin": "APFL-ZF06-01-06",
         "plant": "3093", "sloc": "1005", "unit": "EA", "cat": "FT",
         "availQty": 200, "grDate": "01/10/2026", "grTime": "08:00:00"},
    ]

    plan = NvlService.allocate_orders(orders, stock_records, db_path=tmp_db, source_complete=True)

    assert plan["total_ordered"] == 128
    assert plan["total_allocated"] == 128
    assert plan["is_full_allocation"] is True
    assert len(plan["allocated_items"]) == 2

    # Dòng 1 phải là bin lẻ ZF30-01-08 với 10 EA
    item1 = plan["allocated_items"][0]
    assert item1["bin"] == "APFL-ZF30-01-08"
    assert item1["allocated_qty"] == 10

    # Dòng 2 phải là bin ZF30-02-01 với 118 EA
    item2 = plan["allocated_items"][1]
    assert item2["bin"] == "APFL-ZF30-02-01"
    assert item2["allocated_qty"] == 118

    # Xác nhận khi gom OB: tạo đúng 2 dòng với số lượng tương ứng
    records = NvlService.group_into_ob_groups(plan["allocated_items"], delivery_no="1012948236")
    assert len(records) == 2
    assert records[0]["bin"] == "APFL-ZF30-01-08"
    assert records[0]["allocated_qty"] == 10
    assert records[1]["bin"] == "APFL-ZF30-02-01"
    assert records[1]["allocated_qty"] == 118

