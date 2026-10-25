# -*- coding: utf-8 -*-
"""Unit tests for ExtendedQueryService (Đợt 6: BOM, Queue Diagnostics, ASN Provenance)."""

import pytest
from app.extended_query_service import ExtendedQueryService


def test_bom_explosion_has_bom():
    raw_bom = [
        {"component": "CHS3201001", "description": "Khung gầm phụ", "qty": "1.0", "unit": "PC", "scrap": "0"},
        {"component": "BIW3202002", "description": "Ốc siết thân xe", "qty": "12.0", "unit": "PC", "scrap": "2.5"},
    ]
    res = ExtendedQueryService.explode_bom_structure("BIW32013765AA", "3093", raw_bom)

    assert res["has_bom"] is True
    assert res["total_components"] == 2
    assert res["total_base_qty"] == 13.0
    assert res["components"][0]["component"] == "CHS3201001"
    assert "Định mức BOM" in res["summary"]


def test_bom_explosion_no_bom():
    res = ExtendedQueryService.explode_bom_structure("NON_BOM_PART", "3093", [])
    assert res["has_bom"] is False
    assert res["total_components"] == 0
    assert "không có cấu trúc BOM" in res["summary"]


def test_diagnose_ewm_queues():
    raw_queues = [
        {"queue_name": "DLV_18001122", "error_message": "SYSFAIL: Connection reset by peer", "document": "18001122"},
        {"queue_name": "WT_30004455", "error_message": "Queue locked by user TAIPT8", "document": "30004455"},
        {"queue_name": "PACK_9988", "error_message": "Missing packaging specification in EWM", "document": "9988"}
    ]
    res = ExtendedQueryService.diagnose_ewm_queues(raw_queues)

    assert res["is_clean"] is False
    assert res["stuck_count"] == 3
    assert res["diagnostics"][0]["error_category"] == "CRITICAL_SYSFAIL"
    assert res["diagnostics"][1]["error_category"] == "QUEUE_STOPPED"
    assert res["diagnostics"][2]["error_category"] == "MASTER_DATA_PACKAGING"


def test_track_asn_provenance():
    raw_asn = [
        {"product": "BIW32013765AA", "batch": "B01", "qty": 100, "container": "TCLU1234567", "bill": "BL998877", "vendor": "Gestamp", "eta": "28/10/2026", "poNo": "4500123456"},
        {"product": "BIW32024300AB", "batch": "B02", "qty": 200, "container": "TCLU1234567", "bill": "BL998877", "vendor": "Gestamp", "eta": "28/10/2026", "poNo": "4500123456"}
    ]
    res = ExtendedQueryService.track_asn_provenance("18009988", raw_asn)

    assert res["found"] is True
    assert res["container_no"] == "TCLU1234567"
    assert res["bill_of_lading"] == "BL998877"
    assert res["vendor"] == "Gestamp"
    assert res["total_items"] == 2
    assert res["total_qty"] == 300.0
