# -*- coding: utf-8 -*-
"""Unit tests cho logic xử lý ALV Grid và phân giải cột trong SapEngine."""

from unittest.mock import MagicMock
from app.sap_engine import SapEngine


class FakeColumnCollection:
    def __init__(self, cols):
        self._cols = cols
        self.Count = len(cols)

    def ElementAt(self, idx):
        return self._cols[idx]


def test_resolve_grid_columns_by_tech_name():
    """Kiểm tra phân giải cột theo tên kỹ thuật (Technical Column ID)."""
    engine = SapEngine.__new__(SapEngine)
    alv_mock = MagicMock()
    alv_mock.ColumnOrder = FakeColumnCollection(["MATNR", "CHARG", "LGPLA", "WDATU", "WTIME", "AQUAN"])
    alv_mock.GetColumnTitle.return_value = ""
    alv_mock.ColumnTitle.return_value = ""

    col_map = {
        "product": ["MATNR", "PRODUCT"],
        "batch":   ["CHARG", "BATCH"],
        "bin":     ["LGPLA", "VLPLA"],
        "qty":     ["AQUAN", "AQTY"]
    }
    resolved = engine._resolve_grid_columns(alv_mock, col_map)
    assert resolved["product"] == "MATNR"
    assert resolved["batch"] == "CHARG"
    assert resolved["bin"] == "LGPLA"
    assert resolved["qty"] == "AQUAN"


def test_resolve_grid_columns_by_caption():
    """Kiểm tra phân giải cột theo Caption / Tiêu đề hiển thị (khi mã kỹ thuật khác)."""
    engine = SapEngine.__new__(SapEngine)
    alv_mock = MagicMock()
    alv_mock.ColumnOrder = FakeColumnCollection(["COL1", "COL2", "COL3", "COL4"])

    titles = {
        "COL1": "MATERIAL",
        "COL2": "BATCH",
        "COL3": "SOURCE STORAGE BIN",
        "COL4": "AVAILABLE QTY"
    }
    alv_mock.GetDisplayedColumnTitle.side_effect = lambda c: titles.get(c, "")
    alv_mock.GetColumnTitles.return_value = []
    alv_mock.GetColumnTooltip.return_value = ""
    alv_mock.GetColumnTitle.side_effect = lambda c: titles.get(c, "")
    alv_mock.ColumnTitle.side_effect = lambda c: titles.get(c, "")

    col_map = {
        "product": ["MATNR", "MATERIAL"],
        "batch":   ["CHARG", "BATCH"],
        "bin":     ["SOURCE STORAGE BIN", "LGPLA"],
        "qty":     ["AVAILABLE QTY", "AQUAN"]
    }
    resolved = engine._resolve_grid_columns(alv_mock, col_map)
    assert resolved["product"] == "COL1"
    assert resolved["batch"] == "COL2"
    assert resolved["bin"] == "COL3"
    assert resolved["qty"] == "COL4"


def test_resolve_real_scwm_post_headers():
    """Kiểm tra phân giải chính xác 100% các cột thực tế trên /SCWM/POST từ ảnh chụp SAP GUI."""
    engine = SapEngine.__new__(SapEngine)
    alv_mock = MagicMock()
    alv_mock.ColumnOrder = FakeColumnCollection(["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8"])

    titles = {
        "C1": "Product",
        "C2": "Σ A.Qty AUoM",
        "C3": "Batch",
        "C4": "Source Storage Bin",
        "C5": "TST",
        "C6": "Destination Bin",
        "C7": "GR Date",
        "C8": "GR Time"
    }
    alv_mock.GetDisplayedColumnTitle.side_effect = lambda c: titles.get(c, "")
    alv_mock.GetColumnTitles.return_value = []
    alv_mock.GetColumnTooltip.return_value = ""

    col_map = {
        "product": ["MATNR", "PRODUCT", "MATERIAL"],
        "batch":   ["CHARG", "BATCH"],
        "bin":     ["LGPLA", "VLPLA", "NLPLA", "SOURCE STORAGE BIN", "SOURCE BIN", "STORAGE BIN", "BIN"],
        "qty":     ["AQUAN", "AQTY", "AVL_ALT_QTY", "A.QTY AUOM", "A.QTY", "AVAILABLE QTY", "AVAILABLE QUANTITY", "AVLQTY", "QUAN", "TOTAL STK", "TOTAL STOCK", "GESME"],
        "unit":    ["MEINS", "ALTME", "UNIT", "UOM"],
        "cat":     ["CAT", "TST", "STOCK TYPE", "ST", "TO_CAT"],
        "grDate":  ["WDATU", "GR_DATE", "WEDAT", "GR DATE", "GOODS RECEIPT DATE"],
        "grTime":  ["WTIME", "GR_TIME", "UZEIT", "GR TIME", "GOODS RECEIPT TIME"]
    }
    resolved = engine._resolve_grid_columns(alv_mock, col_map)
    assert resolved["product"] == "C1"
    assert resolved["qty"] == "C2"
    assert resolved["batch"] == "C3"
    assert resolved["bin"] == "C4"
    assert resolved["cat"] == "C5"
    assert resolved["grDate"] == "C7"
    assert resolved["grTime"] == "C8"


def test_extract_alv_grid_with_forward_fill():
    """Kiểm tra trích xuất dòng và forward-fill khi SAP ALV gộp ô Material."""
    engine = SapEngine.__new__(SapEngine)
    engine.max_alv_rows = 50

    alv_mock = MagicMock()
    alv_mock.RowCount = 3
    alv_mock.ColumnOrder = FakeColumnCollection(["MATNR", "CHARG", "LGPLA"])
    alv_mock.GetColumnTitle.return_value = ""
    alv_mock.ColumnTitle.return_value = ""

    # Dòng 0: MATNR='BIW100', CHARG='B1'
    # Dòng 1: MATNR='' (gộp hiển thị), CHARG='B2'
    # Dòng 2: MATNR='BIW200', CHARG='B3'
    data = [
        {"MATNR": "BIW100", "CHARG": "B1", "LGPLA": "A-01"},
        {"MATNR": "",       "CHARG": "B2", "LGPLA": "A-02"},
        {"MATNR": "BIW200", "CHARG": "B3", "LGPLA": "B-01"},
    ]
    alv_mock.GetCellValue.side_effect = lambda r, col: data[r].get(col, "")

    col_map = {
        "product": ["MATNR"],
        "batch":   ["CHARG"],
        "bin":     ["LGPLA"]
    }
    records = engine._extract_alv_grid(alv_mock, col_map)
    assert len(records) == 3
    assert records[0]["product"] == "BIW100"
    assert records[0]["batch"] == "B1"
    # Dòng 1 phải tự động forward-fill mã BIW100 từ dòng 0
    assert records[1]["product"] == "BIW100"
    assert records[1]["batch"] == "B2"
    assert records[2]["product"] == "BIW200"
    assert records[2]["batch"] == "B3"
