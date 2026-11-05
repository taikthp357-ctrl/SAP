# -*- coding: utf-8 -*-
"""Unit tests for v1.1.19: MB52 Stock QI/Blocked, SLoc 9004 defect warning, MB51 polling wait & smart date expansion, and slang parsing."""

import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta

from app.email_parser import SmartCommandParser
from app.teams_dispatcher import TeamsDispatcher
from app.handlers.stock_handler import StockHandler
from app.core.context import RequestContext
from app.sap_mb51_service import SapMb51Service
from app.number_utils import set_global_sap_number_format


class TestV1119StockQiBlockedAndMb51:
    """Test suite covering all v1.1.19 enhancements."""

    @classmethod
    def setup_class(cls):
        set_global_sap_number_format("EU")

    # ----------------------------------------------------------------------
    # 1. INTENT PARSING TESTS
    # ----------------------------------------------------------------------
    def test_ls_material_resolves_to_lichsu(self):
        """'ls <material>' must resolve to LICHSU (not VITRI)."""
        res = SmartCommandParser.parse("ls BIW32012214AA")
        assert res.get("command") == "LICHSU"
        assert "BIW32012214AA" in res.get("materials", [])
        assert res.get("days") == 7

    def test_ls_bin_resolves_to_lichsu_bin_with_bin_value(self):
        """'ls bin <bin>' must resolve to LICHSU_BIN and extract the bin coordinate cleanly."""
        res = SmartCommandParser.parse("ls bin APFL-HR08-08-01")
        assert res.get("command") == "LICHSU_BIN"
        assert res.get("bin") == "APFL-HR08-08-01"
        assert "APFL-HR08-08-01" not in res.get("materials", [])

    def test_tra_lich_su_and_mb51_intents(self):
        """Standard history commands resolve to LICHSU."""
        res1 = SmartCommandParser.parse("tra lịch sử BIW32012214AA")
        assert res1.get("command") == "LICHSU"
        assert "BIW32012214AA" in res1.get("materials", [])

        res2 = SmartCommandParser.parse("mb51 BIW32012214AA trong 30 ngay")
        assert res2.get("command") == "LICHSU"
        assert res2.get("days") == 30

    # ----------------------------------------------------------------------
    # 2. TEAMS DISPATCHER BADGE TESTS (MB52 STOCK DISCREPANCY RESOLUTION)
    # ----------------------------------------------------------------------
    def test_stock_badges_with_qi_blocked_and_sloc_9004(self):
        """Simulate BIW32012214AA from real SAP MB52 screenshot:
        1005: 0 unrest, 149 QI
        3001: 60 unrest, 1111 QI
        9004: 16 unrest, 0 QI (defective stock)
        Total Unrest = 76, Total QI = 1260
        """
        records = [
            {"product": "BIW32012214AA", "plant": "3093", "sloc": "1005", "batch": "261102-10", "unrest": "0", "qual": "149", "blocked": "0", "unit": "EA"},
            {"product": "BIW32012214AA", "plant": "3093", "sloc": "3001", "batch": "PROD_BATCH", "unrest": "60", "qual": "1.111", "blocked": "0", "unit": "EA"},
            {"product": "BIW32012214AA", "plant": "3093", "sloc": "9004", "batch": "PROD_BATCH", "unrest": "16", "qual": "0", "blocked": "0", "unit": "EA"},
        ]
        badges_html = TeamsDispatcher.build_kpi_badges(records, "TON")

        # 1. Total unrestricted stock badge is labeled 'Tồn khả dụng'
        assert "Tồn khả dụng: 76 EA" in badges_html

        # 2. Total Quality Inspection (QI) badge is surfaced
        assert "Chờ KCS (QI): 1260 EA" in badges_html

        # 3. SLoc 1005 breakdown surfaces 0 (+149 QI) EA
        assert "Tồn kho 1005: 0 (+149 QI) EA" in badges_html

        # 4. SLoc 3001 breakdown surfaces 60 (+1111 QI) EA
        assert "Tồn line nhỏ 3001: 60 (+1111 QI) EA" in badges_html

        # 5. SLoc 9004 is explicitly highlighted as defective/NCC stock
        assert "SLoc 9004 (Hàng lỗi/NCC): 16 EA" in badges_html

    def test_clean_stock_badges(self):
        """Standard item without QI/Blocked retains clean, standard badge."""
        records = [
            {"product": "BIW70011279", "plant": "3093", "sloc": "1005", "batch": "260923-10", "unrest": "100", "qual": "0", "blocked": "0", "unit": "EA"}
        ]
        badges_html = TeamsDispatcher.build_kpi_badges(records, "TON")
        assert "Tồn: 100 EA" in badges_html
        assert "Chờ KCS" not in badges_html
        assert "Khóa" not in badges_html

    # ----------------------------------------------------------------------
    # 3. STOCK HANDLER HEADERS & SMART DATE EXPANSION
    # ----------------------------------------------------------------------
    def test_stock_handler_headers_expansion_on_qi(self):
        """StockHandler must dynamically include QI and Blocked columns when present."""
        records = [
            {"product": "BIW32012214AA", "plant": "3093", "sloc": "1005", "batch": "261102-10", "unrest": "0", "qual": "149", "blocked": "0", "unit": "EA"}
        ]
        mock_sap = MagicMock()
        mock_sap.check_stock_mb52.return_value = records

        ctx = RequestContext(
            request_id="REQ-TEST-1",
            command="TON",
            raw_materials=["BIW32012214AA"],
            resolved_materials=["BIW32012214AA"],
            plant="3093",
            sap=mock_sap,
            config={},
            meta={}
        )

        handler = StockHandler()
        resp = handler.handle(ctx)
        assert resp is not None
        headers = resp.headers
        assert headers["unrest"] == "Khả dụng"
        assert headers["qual"] == "Chờ KCS (QI)"
        assert "blocked" not in headers  # blocked is 0, so not added unnecessarily

    def test_stock_handler_mb51_smart_date_expansion(self):
        """If 7 days returns 0 records and user used default range, auto-expand to 30 days."""
        ctx = RequestContext(
            request_id="REQ-TEST-2",
            command="LICHSU",
            raw_materials=["BIW32012214AA"],
            resolved_materials=["BIW32012214AA"],
            plant="3093",
            days=7,
            meta={"rawText": "ls BIW32012214AA"}
        )

        # Simulate 7-day query returning empty, but 30-day query returning records
        expanded_records = [
            {"budat": "29.08.2026", "matnr": "BIW32012214AA", "werks": "3093", "lgort": "1005", "bwart": "101", "qty": "+1400", "meins": "EA", "ebeln": "5006815109", "usnam": "TAIPT8"}
        ]

        with patch.object(SapMb51Service, "fetch_movement_history") as mock_fetch:
            # First call (7 days) returns [], second call (30 days) returns expanded_records
            mock_fetch.side_effect = [[], expanded_records]

            handler = StockHandler()
            resp = handler.handle(ctx)

            assert resp is not None
            assert len(resp.records) == 1
            assert "Tự động mở rộng 30 ngày" in resp.summary

    # ----------------------------------------------------------------------
    # 4. SAP MB51 SERVICE POLLING WAIT
    # ----------------------------------------------------------------------
    def test_mb51_service_polling_wait(self):
        """SapMb51Service must poll for cntlGRID1 up to 15s instead of giving up immediately."""
        service = SapMb51Service()
        mock_session = MagicMock()

        # Mock status bar (normal)
        mock_sbar = MagicMock()
        mock_sbar.MessageType = "S"
        mock_sbar.Text = ""

        # Mock grid becoming available on 3rd attempt
        mock_grid = MagicMock()
        mock_grid.RowCount = 1
        mock_grid.ColumnOrder.Count = 3
        mock_grid.ColumnOrder.ElementAt.side_effect = ["BUDAT", "MATNR", "BWART"]
        mock_grid.GetCellValue.side_effect = lambda row, col: {"BUDAT": "29.08.2026", "MATNR": "BIW32012214AA", "BWART": "101"}.get(col, "")

        attempts = [0]
        def find_by_id_side_effect(id_str):
            if "cntlGRID1" in id_str:
                attempts[0] += 1
                if attempts[0] < 3:
                    raise Exception("Grid not loaded yet")
                return mock_grid
            if "sbar" in id_str:
                return mock_sbar
            return MagicMock()

        mock_session.findById.side_effect = find_by_id_side_effect

        with patch("time.sleep", return_value=None):
            records = service.fetch_movement_history(
                materials=["BIW32012214AA"],
                date_from=datetime.now() - timedelta(days=7),
                date_to=datetime.now(),
                session=mock_session
            )

        assert len(records) == 1
        assert records[0]["matnr"] == "BIW32012214AA"
        assert attempts[0] == 3
