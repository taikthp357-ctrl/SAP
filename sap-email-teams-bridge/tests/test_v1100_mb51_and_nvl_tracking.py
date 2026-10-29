# -*- coding: utf-8 -*-
"""Unit tests for v1.1.00 features:
1. Dynamic NVL Tracking Layout & Excel Serial Date Handling
2. Robust SAP MB51 Transaction History & Wildcard Normalization
3. StockHandler 10-column ALV extraction & HTML Table Generation
"""

import os
from datetime import datetime, date, timedelta
from unittest.mock import MagicMock, patch

import pytest

from app.nvl_tracking import NvlTrackingService
from app.sap_mb51_service import SapMb51Service, parse_mb51_date_range, generate_mb51_table_html
from app.email_parser import SmartCommandParser
from app.core.context import RequestContext
from app.handlers.stock_handler import StockHandler
from app.router import CommandRouter


class TestNvlTrackingV1100:
    """Kiểm thử tính năng NVL Tracking phiên bản v1.1.00."""

    def test_parse_excel_date_serial_float(self):
        """Kiểm tra parse_excel_date xử lý số Serial Float từ pyxlsb."""
        # 46324.0 tương ứng ngày 29-10-2026
        dt = NvlTrackingService.parse_excel_date(46324.0)
        assert dt == date(2026, 10, 29)

        # 46323.0 tương ứng ngày 28-10-2026
        dt2 = NvlTrackingService.parse_excel_date(46323.0)
        assert dt2 == date(2026, 10, 28)

    def test_parse_excel_date_string_and_date(self):
        """Kiểm tra parse chuỗi ngày dạng '29-Oct' hoặc date object."""
        d_obj = date(2026, 10, 29)
        assert NvlTrackingService.parse_excel_date(d_obj) == d_obj

        d_str = NvlTrackingService.parse_excel_date("29-Oct")
        assert d_str.day == 29
        assert d_str.month == 10

    def test_resolve_column_map_20_columns(self):
        """Kiểm tra phát hiện layout 20 cột mới từ tiêu đề thực tế."""
        header_sample = [
            ["STT", "Model", "Mã NVL", "SL oder", "Tên NVL", "Khu vực cấp", "DVT", 
             "Khung giờ cấp", "Vị trí cấp", "Loại xe cấp", "Khu vực nhận", "OB", "OB1",
             "Thời gian nhận", "Ghi chú", "Tình Trạng cần đơn", "Tình trạng cấp", "KHSX", "Xưởng nhận", "Ngày order"]
        ]
        hdr_idx, col_map = NvlTrackingService.resolve_column_map(header_sample)
        assert col_map["mat_code"] == 2      # Cột Mã NVL
        assert col_map["qty_order"] == 3     # Cột SL oder
        assert col_map["ob"] == 11           # Cột OB (khác OB1)
        assert col_map["ob1"] == 12          # Cột OB1 (khác OB)
        assert col_map["status_pick"] == 16  # Cột Tình trạng cấp (không nhầm Tình Trạng cần đơn)
        assert col_map["date_order"] == 19   # Cột Ngày order

    def test_find_tracking_files_sorting_and_temp_file_filter(self, tmp_path):
        """Kiểm tra tìm kiếm file lọc bỏ ~$* và sắp xếp theo YYMMDD mới nhất."""
        service = NvlTrackingService()

        # Tạo các file test
        f1 = tmp_path / "261028_THEO DOI ODER NVL_P_S6.xlsx"
        f2 = tmp_path / "261029_THEO DOI ORDER NVL_P_S6.xlsb"  # Có chữ ORDER với chữ R
        f_temp = tmp_path / "~$261029_THEO DOI ORDER NVL_P_S6.xlsb"

        f1.touch()
        f2.touch()
        f_temp.touch()

        files = service.find_tracking_files(directory=str(tmp_path), shift="S6")
        file_names = [os.path.basename(f) for f in files]

        # File temp ~$ bắt buộc phải bị loại bỏ
        assert not any(fn.startswith("~$") for fn in file_names)

        # File ngày 29 phải đứng trước ngày 28
        assert len(file_names) == 2
        assert file_names[0] == "261029_THEO DOI ORDER NVL_P_S6.xlsb"
        assert file_names[1] == "261028_THEO DOI ODER NVL_P_S6.xlsx"


class TestSapMb51ServiceV1100:
    """Kiểm thử tính năng tra cứu lịch sử MB51 và xử lý lỗi Error 617."""

    def test_normalize_wildcard_mat(self):
        """Kiểm tra chuẩn hóa wildcard '*' cho mã 11 ký tự và 8 số."""
        # 11 ký tự chưa có hậu tố -> thêm *
        assert SapMb51Service.normalize_wildcard_mat("BIW32011434") == "BIW32011434*"
        # 8 chữ số -> bọc *
        assert SapMb51Service.normalize_wildcard_mat("32011434") == "*32011434*"
        # Đã có hậu tố HA/HB -> giữ nguyên
        assert SapMb51Service.normalize_wildcard_mat("BIW32011434HA") == "BIW32011434HA"
        # Đã có * -> giữ nguyên
        assert SapMb51Service.normalize_wildcard_mat("BIW75122581*") == "BIW75122581*"

    def test_parse_mb51_date_range(self):
        """Kiểm tra bóc tách khoảng ngày linh hoạt từ câu lệnh."""
        now = datetime.now()

        # Hôm nay
        d_from, d_to = parse_mb51_date_range("Sap check lịch sử BIW32011434 hôm nay")
        assert d_from.date() == now.date()
        assert d_to.date() == now.date()

        # Hôm qua
        d_from_y, d_to_y = parse_mb51_date_range("check lich su mb51 hom qua")
        assert d_from_y.date() == (now - timedelta(days=1)).date()

        # N ngày
        d_from_n, d_to_n = parse_mb51_date_range("mb51 32011434 trong 15 ngày")
        assert (d_to_n.date() - d_from_n.date()).days == 14

        # Ngày cụ thể: ngày 27/10
        d_from_sp, d_to_sp = parse_mb51_date_range("Sap check lịch sử BIW75122581 ngày 27/10")
        assert d_from_sp.day == 27
        assert d_from_sp.month == 10

    def test_generate_mb51_table_html(self):
        """Kiểm tra sinh mã HTML bảng 10 cột có badge loại GD và màu số lượng."""
        records = [
            {
                "budat": "27.10.2026",
                "matnr": "BIW32011434HA",
                "werks": "3093",
                "lgort": "1005",
                "umlgo": "3093",
                "bwart": "311",
                "qty": "+100",
                "meins": "PC",
                "ebeln": "PO123456",
                "usnam": "NGUYENVANA",
            },
            {
                "budat": "28.10.2026",
                "matnr": "BIW32011434HB",
                "werks": "3093",
                "lgort": "3093",
                "umlgo": "",
                "bwart": "261",
                "qty": "-50",
                "meins": "PC",
                "ebeln": "",
                "usnam": "TRANVANB",
            }
        ]
        html = generate_mb51_table_html(records, "Lịch sử biến động vật tư")
        assert "311 (Chuyển kho)" in html
        assert "261 (Xuất SX)" in html
        assert "color:#16a34a;font-weight:bold;'>+100" in html
        assert "color:#dc2626;font-weight:bold;'>-50" in html
        assert "BIW32011434HA" in html
        assert "NGUYENVANA" in html


class TestStockHandlerAndParserV1100:
    """Kiểm thử tích hợp Parser, Router và StockHandler cho lệnh MB51."""

    def test_parser_detects_mb51_command_and_dates(self):
        """Kiểm tra SmartCommandParser nhận diện lệnh LICHSU với ngày cụ thể."""
        parser = SmartCommandParser()
        parsed = parser.parse("Sap check lịch sử BIW32011434 ngày 27/10")
        assert parsed["command"] == "LICHSU"
        assert "BIW32011434" in parsed["materials"]
        assert parsed.get("dateFrom") is not None
        assert parsed.get("dateTo") is not None

    def test_stock_handler_mb51_dispatch_and_custom_html(self):
        """Kiểm tra StockHandler xử lý lệnh MB51_HISTORY với fallback mock."""
        mock_sap = MagicMock()
        mock_sap.session = None
        mock_sap.check_history_mb51.return_value = [
            {
                "product": "BIW32011434HA",
                "docNo": "5000123456",
                "mvtType": "311",
                "postDate": "27.10.2026",
                "qty": "+200",
                "unit": "PC",
                "sloc": "1005",
                "batch": "BATCH01",
                "plant": "3093",
                "umlgo": "3093",
                "usnam": "TESTUSER",
            }
        ]

        ctx = RequestContext(
            request_id="REQ-TEST-MB51",
            command="LICH_SU_MB51",
            raw_materials=["BIW32011434"],
            resolved_materials=["BIW32011434"],
            plant="3093",
            sap=mock_sap,
            meta={"rawText": "check lịch sử BIW32011434"}
        )

        handler = StockHandler()
        result = handler.handle(ctx)

        assert result.status == "SUCCEEDED"
        assert len(result.records) == 1
        rec = result.records[0]
        assert rec["matnr"] == "BIW32011434HA"
        assert rec["bwart"] == "311"
        assert rec["qty"] == "+200"
        # Kiểm tra alias
        assert rec["product"] == "BIW32011434HA"
        assert rec["mvtType"] == "311"
        # Kiểm tra custom HTML
        assert result.custom_html is not None
        assert "311 (Chuyển kho)" in result.custom_html
        assert "TESTUSER" in result.custom_html

    def test_command_router_routes_all_mb51_aliases(self):
        """Kiểm tra CommandRouter định tuyến đúng tất cả alias MB51 tới StockHandler."""
        router = CommandRouter()
        assert isinstance(router.get_handler("LICHSU"), StockHandler)
        assert isinstance(router.get_handler("LỊCH SỬ"), StockHandler)
        assert isinstance(router.get_handler("MB51"), StockHandler)
        assert isinstance(router.get_handler("LICH_SU_MB51"), StockHandler)
        assert isinstance(router.get_handler("MB51_HISTORY"), StockHandler)
