# -*- coding: utf-8 -*-
"""Unit tests for ObService (Đợt 5: Điều phối Outbound Delivery RC65)."""

import pytest
from app.ob_service import ObService


def test_ob_whole_batch_gate_passed():
    delivery_items = [
        {"item": "10", "material": "BIW32013765AA", "qty": 50, "batch": ""},
        {"item": "20", "material": "BIW32024300AB", "qty": 30, "batch": "LOT_SPECIFIC"},
    ]

    stock_records = [
        {"product": "BIW32013765AA", "batch": "B01", "availQty": 100},
        {"product": "BIW32024300AB", "batch": "LOT_SPECIFIC", "availQty": 30},
    ]

    gate_res = ObService.validate_whole_batch_gate("8000123456", delivery_items, stock_records)

    assert gate_res["gate_passed"] is True
    assert gate_res["gate_status"] == "GATE_PASSED"
    assert gate_res["deficit_count"] == 0

    # Dispatch to plant 3001
    disp_res = ObService.dispatch_outbound_simulated("8000123456", "3001", gate_res)
    assert disp_res["success"] is True
    assert disp_res["channel"] == "VL02N_PICK_PACK"
    assert disp_res["readback_status"] == "C"


def test_ob_whole_batch_gate_deficit_blocked():
    delivery_items = [
        {"item": "10", "material": "BIW32013765AA", "qty": 50},
        {"item": "20", "material": "BIW32024300AB", "qty": 100},  # Deficit
    ]

    stock_records = [
        {"product": "BIW32013765AA", "batch": "B01", "availQty": 100},
        {"product": "BIW32024300AB", "batch": "B02", "availQty": 20},  # Only 20 available!
    ]

    gate_res = ObService.validate_whole_batch_gate("8000999999", delivery_items, stock_records)

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
    assert res_2001["readback_status"] == "CONFIRMED"
