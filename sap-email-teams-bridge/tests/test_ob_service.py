# -*- coding: utf-8 -*-
"""Unit tests for ObService (Đợt 5: Điều phối Outbound Delivery RC65)."""

import pytest
from app.ob_service import ObService


def test_ob_whole_batch_gate_passed():
    delivery_items = [
        {"item": "10", "material": "BIW32013765AA", "unit": "PC", "plant": "3001", "sloc": "1003", "qty": 50, "batch": ""},
        {"item": "20", "material": "BIW32024300AB", "unit": "PC", "plant": "3001", "sloc": "1003", "qty": 30, "batch": "LOT_SPECIFIC"},
    ]

    stock_records = [
        {"product": "BIW32013765AA", "batch": "B01", "bin": "BIN-A", "unit": "PC", "plant": "3001", "sloc": "1003", "cat": "FT", "availQty": 100},
        {"product": "BIW32024300AB", "batch": "LOT_SPECIFIC", "bin": "BIN-A", "unit": "PC", "plant": "3001", "sloc": "1003", "cat": "FT", "availQty": 30},
    ]

    gate_res = ObService.validate_whole_batch_gate("8000123456", delivery_items, stock_records, source_complete=True)

    assert gate_res["gate_passed"] is True
    assert gate_res["gate_status"] == "GATE_PASSED"
    assert gate_res["deficit_count"] == 0

    # Dispatch to plant 3001
    disp_res = ObService.dispatch_outbound_simulated("8000123456", "3001", gate_res)
    assert disp_res["success"] is False
    assert disp_res["status"] == "SIMULATED"
    assert disp_res["channel"] == "VL02N_PICK_PACK"
    assert disp_res["readback_status"] is None


def test_ob_whole_batch_gate_deficit_blocked():
    delivery_items = [
        {"item": "10", "material": "BIW32013765AA", "unit": "PC", "plant": "3001", "sloc": "1003", "qty": 50},
        {"item": "20", "material": "BIW32024300AB", "unit": "PC", "plant": "3001", "sloc": "1003", "qty": 100},  # Deficit
    ]

    stock_records = [
        {"product": "BIW32013765AA", "batch": "B01", "bin": "BIN-A", "unit": "PC", "plant": "3001", "sloc": "1003", "cat": "FT", "availQty": 100},
        {"product": "BIW32024300AB", "batch": "B02", "bin": "BIN-A", "unit": "PC", "plant": "3001", "sloc": "1003", "cat": "FT", "availQty": 20},  # Only 20 available!
    ]

    gate_res = ObService.validate_whole_batch_gate("8000999999", delivery_items, stock_records, source_complete=True)

    assert gate_res["gate_passed"] is False
    assert gate_res["gate_status"] == "DEFICIT_BLOCKED"
    assert gate_res["deficit_count"] == 1
    assert "DEFICIT_BLOCKED" in gate_res["action_advice"]

    # Dispatch should be rejected
    disp_res = ObService.dispatch_outbound_simulated("8000999999", "3001", gate_res)
    assert disp_res["success"] is False
    assert disp_res["status"] == "BLOCKED"
    assert len(disp_res["deficits"]) == 1


def test_ob_multi_plant_routing():
    gate_ok = {"gate_passed": True, "total_items": 1}

    # Plant 4007 STO
    res_4007 = ObService.dispatch_outbound_simulated("80001", "4007", gate_ok)
    assert res_4007["channel"] == "VL10D_STO_TRANSFER"

    # Plant 2001 EWM WT
    res_2001 = ObService.dispatch_outbound_simulated("80002", "2001", gate_ok)
    assert res_2001["channel"] == "EWM_MON_WT_CONFIRM"
    assert res_2001["readback_status"] is None
    assert res_2001["executed"] is False


def test_rc65_locked_field_tokens():
    """Kiểm tra toàn bộ các field ID token đã khóa cứng 100% theo baseline RC65."""
    tokens = ObService.resolve_rc65_tokens()
    assert tokens["QTY_OVERVIEW"] == "LIPSD-G_LFIMG"
    assert tokens["STORAGE_LOCATION"] == "LIPS-LGORT"
    assert tokens["BATCH"] == "LIPS-CHARG"
    assert tokens["ITEM_NO"] == "LIPS-POSNR"
    assert tokens["MATERIAL"] == "LIPS-MATNR"
    assert tokens["QTY_SPLIT"] == "LIPS-LFIMG"
    assert tokens["RECEIVING_PLANT"] == "LIPS-UMWRK"
    assert tokens["REF_STO"] == "LIPS-VGBEL"
    assert tokens["REF_STO_ITEM"] == "LIPS-VGPOS"


def test_rc65_table_field_and_set_table_text_mock():
    """Kiểm tra TableField và SetTableText chuẩn RC65 qua đối tượng COM mock."""
    class MockChild:
        def __init__(self, cid, text=""):
            self.Id = cid
            self.id = cid
            self.Text = text
            self.text = text

    class MockTable:
        def __init__(self):
            self.Id = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER"
            self.id = self.Id
            self.VisibleRowCount = 5
            self.RowCount = 2
            # Giả lập các ô theo token relative ID
            self._children = [
                MockChild(f"{self.Id}/txtLIPS-POSNR[0,0]", "10"),
                MockChild(f"{self.Id}/txtLIPS-MATNR[1,0]", "BIW32013765AA"),
                MockChild(f"{self.Id}/txtLIPSD-G_LFIMG[2,0]", "100.000"),
                MockChild(f"{self.Id}/txtLIPS-LGORT[3,0]", "3093"),
                MockChild(f"{self.Id}/txtLIPS-CHARG[4,0]", "LOT-01"),
                MockChild(f"{self.Id}/txtLIPS-POSNR[0,1]", "20"),
                MockChild(f"{self.Id}/txtLIPS-MATNR[1,1]", "BIW32024300AB"),
                MockChild(f"{self.Id}/txtLIPSD-G_LFIMG[2,1]", "50.000"),
                MockChild(f"{self.Id}/txtLIPS-LGORT[3,1]", ""),
                MockChild(f"{self.Id}/txtLIPS-CHARG[4,1]", ""),
            ]

        @property
        def Children(self):
            class Coll:
                Count = len(self._children)
                count = len(self._children)
                def __call__(_s, idx):
                    return self._children[idx]
                def ElementAt(_s, idx):
                    return self._children[idx]
            return Coll()

        def findById(self, rel_id):
            full_id = f"{self.Id}/{rel_id}" if not rel_id.startswith("wnd[0]") else rel_id
            for ch in self._children:
                if ch.Id == full_id or ch.Id.endswith(rel_id):
                    return ch
            return None

    tbl = MockTable()

    # 1. Test get_table_field (TableField)
    f_pos = ObService.get_table_field(tbl, "LIPS-POSNR", 0)
    assert f_pos is not None
    assert f_pos.Text == "10"

    f_qty = ObService.get_table_field(tbl, "LIPSD-G_LFIMG", 0)
    assert f_qty is not None
    assert f_qty.Text == "100.000"

    f_sloc = ObService.get_table_field(tbl, "LIPS-LGORT", 0)
    assert f_sloc is not None
    assert f_sloc.Text == "3093"

    f_batch = ObService.get_table_field(tbl, "LIPS-CHARG", 0)
    assert f_batch is not None
    assert f_batch.Text == "LOT-01"

    # Row 1
    f_pos1 = ObService.get_table_field(tbl, "LIPS-POSNR", 1)
    assert f_pos1 is not None
    assert f_pos1.Text == "20"

    # Cache hit check
    f_qty_cached = ObService.get_table_field(tbl, "LIPSD-G_LFIMG", 1)
    assert f_qty_cached is not None
    assert f_qty_cached.Text == "50.000"

    # 2. Test set_table_text (SetTableText)
    ok = ObService.set_table_text(tbl, "LIPS-LGORT", 1, "3093")
    assert ok is True
    assert ObService.get_table_field(tbl, "LIPS-LGORT", 1).Text == "3093"

    ok = ObService.set_table_text(tbl, "LIPS-CHARG", 1, "LOT-NEW")
    assert ok is True
    assert ObService.get_table_field(tbl, "LIPS-CHARG", 1).Text == "LOT-NEW"

    # 3. Test locate_delivery_item (LocateItem)
    loc0 = ObService.locate_delivery_item(tbl, "10")
    assert loc0 == 0
    loc1 = ObService.locate_delivery_item(tbl, "20")
    assert loc1 == 1
    loc_none = ObService.locate_delivery_item(tbl, "999")
    assert loc_none == -1

    # 4. Test snapshot_delivery_overview (OverviewSnapshot)
    snap = ObService.snapshot_delivery_overview(tbl)
    assert "10" in snap
    assert snap["10"]["material"] == "BIW32013765AA"
    assert snap["10"]["qty"] == "100.000"
    assert snap["10"]["sloc"] == "3093"
    assert snap["10"]["batch"] == "LOT-01"

    assert "20" in snap
    assert snap["20"]["material"] == "BIW32024300AB"
    assert snap["20"]["qty"] == "50.000"
    assert snap["20"]["sloc"] == "3093"
    assert snap["20"]["batch"] == "LOT-NEW"

