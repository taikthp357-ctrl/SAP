"""Unit tests for v1.0.95 - Fix Ra Đơn V31 Zone (Hàng Thùng & Tách Dòng Gộp Teams)."""
import pytest
from app.email_parser import EmailCommandParser
from app.nvl_master import NvlMasterService
from app.nvl_service import NvlService
from app.nvl_tracking import NvlTrackingService


def test_parse_merged_table_rows_from_teams():
    """Kiểm tra bóc tách chính xác 4 mã vật tư khi người dùng gửi bảng bị gộp dòng trong Teams."""
    text = (
        "BIW32025210AA 420 KV-1/Duy/Bằng ZONE 21 BIW32026215AA 420 KV-1/Duy/Bằng ZONE 21\n"
        "BIW32012580AA 420 KV-1/Duy/Bằng ZONE 24 BIW32012585AA 420 KV-1/Duy/Bằng ZONE 24"
    )
    rows = EmailCommandParser.parse_nvl_table_rows(text)
    assert len(rows) == 4, f"Phải bóc tách đủ 4 dòng, nhưng chỉ nhận được {len(rows)}"

    assert rows[0]["material"] == "BIW32025210AA"
    assert rows[0]["qty"] == 420
    assert rows[0]["station"] == "KV-1/Duy/Bằng ZONE 21"

    assert rows[1]["material"] == "BIW32026215AA"
    assert rows[1]["qty"] == 420
    assert rows[1]["station"] == "KV-1/Duy/Bằng ZONE 21"

    assert rows[2]["material"] == "BIW32012580AA"
    assert rows[2]["qty"] == 420
    assert rows[2]["station"] == "KV-1/Duy/Bằng ZONE 24"

    assert rows[3]["material"] == "BIW32012585AA"
    assert rows[3]["qty"] == 420
    assert rows[3]["station"] == "KV-1/Duy/Bằng ZONE 24"


def test_parse_single_unbroken_line_multiple_items():
    """Kiểm tra bóc tách khi toàn bộ 4 mã dính trên 1 dòng đơn không có ký tự xuống dòng."""
    text = (
        "BIW32025210AA 420 KV-1/Duy/Bằng ZONE 21 "
        "BIW32026215AA 420 KV-1/Duy/Bằng ZONE 21 "
        "BIW32012580AA 420 KV-1/Duy/Bằng ZONE 24 "
        "BIW32012585AA 420 KV-1/Duy/Bằng ZONE 24"
    )
    res = EmailCommandParser.parse(text)
    assert res["command"] == "RA_DON_NVL"
    assert len(res["orders"]) == 4
    assert sum(o["qty"] for o in res["orders"]) == 1680


def test_station_trailing_stt_cleanup():
    """Kiểm tra làm sạch số STT của dòng sau bị dính vào đuôi Trạm của dòng trước."""
    text = "BIW32025210AA 420 KV-1/Duy/Bằng ZONE 21 2 BIW32026215AA 420 KV-1/Duy/Bằng ZONE 21"
    rows = EmailCommandParser.parse_nvl_table_rows(text)
    assert len(rows) == 2
    assert rows[0]["station"] == "KV-1/Duy/Bằng ZONE 21"
    assert rows[1]["station"] == "KV-1/Duy/Bằng ZONE 21"


def test_hang_thung_packing_breakdown():
    """Kiểm tra tính toán quy cách đóng gói Hàng Thùng (Packing) chuẩn Master Data VinFast."""
    # BIW32025210AA có packing = 128 (Thùng giấy Limo/128)
    split_128 = NvlMasterService.calculate_box_split("BIW32025210AA", 420)
    assert split_128["is_box"] is True
    assert split_128["packing"] == 128
    assert split_128["full_boxes"] == 3
    assert split_128["box_qty"] == 384.0
    assert split_128["loose_qty"] == 36.0

    # BIW32012580AA có packing = 48 (Thùng giấy Limo/48)
    split_48 = NvlMasterService.calculate_box_split("BIW32012580AA", 420)
    assert split_48["is_box"] is True
    assert split_48["packing"] == 48
    assert split_48["full_boxes"] == 8
    assert split_48["box_qty"] == 384.0
    assert split_48["loose_qty"] == 36.0


def test_hang_thung_description_in_ob_groups():
    """Kiểm tra thông tin Hàng Thùng hiển thị chi tiết trong nhóm OB (OB_GROUPS)."""
    allocated = [
        {
            "material": "BIW32025210AA",
            "allocated_qty": 420,
            "station": "KV-1/Duy/Bằng ZONE 21",
            "batch": "B2610A",
            "bin": "APFL-ZF01",
            "unit": "PC"
        },
        {
            "material": "BIW32012580AA",
            "allocated_qty": 420,
            "station": "KV-1/Duy/Bằng ZONE 24",
            "batch": "B2610B",
            "bin": "APFL-TH02",
            "unit": "PC"
        }
    ]
    records = NvlService.group_into_ob_groups(allocated, delivery_no="1013158278")
    assert len(records) == 2
    r1 = next(r for r in records if r["material"] == "BIW32025210AA")
    assert "3 thùng (384)" in r1["packing_desc"]
    assert "lẻ 36 EA" in r1["packing_desc"]
    assert r1["packing"] == 128

    r2 = next(r for r in records if r["material"] == "BIW32012580AA")
    assert "8 thùng (384)" in r2["packing_desc"]
    assert "lẻ 36 EA" in r2["packing_desc"]
    assert r2["packing"] == 48


def test_xlsb_sync_safety(tmp_path):
    """Kiểm tra hàm đồng bộ tracking không bị crash khi gặp file .xlsb."""
    fake_xlsb = tmp_path / "261028_THEO DOI ODER NVL_P_S6.xlsb"
    fake_xlsb.write_bytes(b"FAKE_XLSB_CONTENT")

    res = NvlTrackingService.sync_dispatch_to_tracking_file(
        orders_dispatched=[{"material": "BIW32025210AA", "qty": 420}],
        delivery_no="1013158278",
        tracking_file=fake_xlsb
    )
    # Không được ném ngoại lệ crash openpyxl
    assert res["status"] in ["SKIPPED_XLSB", "LOCKED", "ERROR", "SUCCEEDED", "INVALID_FILE"]


def test_box_limit_and_capping():
    """Kiểm tra quy tắc giới hạn thùng V31:
    - Dưới 100 là tối đa 3 thùng
    - Trên 100 là tối đa 2 thùng
    - Trên 200 là tối đa 1 thùng
    """
    assert NvlMasterService.get_box_limit(24) == 3
    assert NvlMasterService.get_box_limit(48) == 3
    assert NvlMasterService.get_box_limit(100) == 3
    assert NvlMasterService.get_box_limit(101) == 2
    assert NvlMasterService.get_box_limit(128) == 2
    assert NvlMasterService.get_box_limit(200) == 2
    assert NvlMasterService.get_box_limit(201) == 1
    assert NvlMasterService.get_box_limit(500) == 1

    # Kiểm tra cờ capBoxes khi người dùng yêu cầu 'theo thùng'
    res = EmailCommandParser.parse("ra đơn nvl theo thùng BIW32025210AA 420 KV-1")
    assert res.get("capBoxes") is True
