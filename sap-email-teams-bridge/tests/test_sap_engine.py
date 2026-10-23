# -*- coding: utf-8 -*-
"""Unit tests cho logic xử lý ALV Grid và phân giải cột trong SapEngine."""

from unittest.mock import MagicMock
import pytest
from app.sap_engine import SapEngine, SapExecutionError


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


def test_resolve_vl06i_columns():
    """Kiểm tra phân giải các cột của màn hình VL06I (Item ALV)."""
    engine = SapEngine.__new__(SapEngine)
    alv_mock = MagicMock()
    alv_mock.ColumnOrder = FakeColumnCollection(["VBELN", "LIFEX", "MATNR", "S_CHARG", "LFIMG", "VRKME", "WERKS", "LGORT", "WBSTK"])
    alv_mock.GetDisplayedColumnTitle.return_value = ""
    alv_mock.GetColumnTitles.return_value = []
    alv_mock.GetColumnTooltip.return_value = ""

    col_map = {
        "ibNo":    ["VBELN", "DELIVERY", "INBOUND DELIVERY"],
        "refNo":   ["LIFEX", "BILL OF LADING", "REF. DOC."],
        "product": ["MATNR", "MATERIAL", "PRODUCT"],
        "batch":   ["CHARG", "S_CHARG", "BATCH"],
        "qty":     ["LFIMG", "DELIVERY QUANTITY", "MENGE"],
        "unit":    ["VRKME", "MEINS", "UNIT"],
        "plant":   ["WERKS", "PLANT"],
        "sloc":    ["LGORT", "STORAGE LOCATION", "SLOC"],
        "status":  ["WBSTK", "GM STATUS", "STATUS"]
    }
    resolved = engine._resolve_grid_columns(alv_mock, col_map)
    assert resolved["ibNo"] == "VBELN"
    assert resolved["refNo"] == "LIFEX"
    assert resolved["product"] == "MATNR"
    assert resolved["batch"] == "S_CHARG"
    assert resolved["qty"] == "LFIMG"
    assert resolved["unit"] == "VRKME"
    assert resolved["plant"] == "WERKS"
    assert resolved["sloc"] == "LGORT"
    assert resolved["status"] == "WBSTK"


def test_resolve_me2n_columns():
    """Kiểm tra phân giải các cột của màn hình ME2N (ALV Grid)."""
    engine = SapEngine.__new__(SapEngine)
    alv_mock = MagicMock()
    alv_mock.ColumnOrder = FakeColumnCollection(["EBELN", "EBELP", "EMATN", "MENGE", "MGLIEF", "MEINS", "RESWK", "WERKS", "BEDAT"])
    alv_mock.GetDisplayedColumnTitle.return_value = ""
    alv_mock.GetColumnTitles.return_value = []
    alv_mock.GetColumnTooltip.return_value = ""

    col_map = {
        "stoNo":          ["EBELN", "PURCHASING DOCUMENT", "PO NUMBER"],
        "item":           ["EBELP", "ITEM"],
        "product":        ["EMATN", "MATNR", "MATERIAL"],
        "orderQty":       ["MENGE", "ORDER QUANTITY", "QTY"],
        "deliveredQty":   ["MGLIEF", "DELIVERED QUANTITY", "ISSUED QTY"],
        "unit":           ["MEINS", "BPRME", "UNIT"],
        "supplyingPlant": ["RESWK", "SUPPLYING PLANT"],
        "receivingPlant": ["WERKS", "RECEIVING PLANT", "PLANT"],
        "docDate":        ["BEDAT", "DOCUMENT DATE", "PO DATE"]
    }
    resolved = engine._resolve_grid_columns(alv_mock, col_map)
    assert resolved["stoNo"] == "EBELN"
    assert resolved["item"] == "EBELP"
    assert resolved["product"] == "EMATN"
    assert resolved["orderQty"] == "MENGE"
    assert resolved["deliveredQty"] == "MGLIEF"
    assert resolved["unit"] == "MEINS"
    assert resolved["supplyingPlant"] == "RESWK"
    assert resolved["receivingPlant"] == "WERKS"
    assert resolved["docDate"] == "BEDAT"


def test_extract_alv_grid_preserves_sto_and_filters_totals():
    """Đảm bảo _extract_alv_grid không loại bỏ các dòng STO (chỉ có stoNo/item) và lọc đúng dòng tổng."""
    engine = SapEngine.__new__(SapEngine)
    engine.max_alv_rows = 50

    alv_mock = MagicMock()
    alv_mock.RowCount = 3
    alv_mock.ColumnOrder = FakeColumnCollection(["EBELN", "EBELP", "EMATN", "MENGE"])
    alv_mock.GetColumnTitle.return_value = ""
    alv_mock.ColumnTitle.return_value = ""

    # Dòng 0: Dòng dữ liệu STO hợp lệ (không có batch, bin, sloc, docNo)
    # Dòng 1: Dòng dữ liệu STO hợp lệ thứ hai
    # Dòng 2: Dòng subtotal/total chỉ có MENGE (không có stoNo, item, product,...)
    data = [
        {"EBELN": "4500012345", "EBELP": "10", "EMATN": "BIW32013765AA", "MENGE": "1,000"},
        {"EBELN": "4500012346", "EBELP": "20", "EMATN": "BIW32013765AA", "MENGE": "500"},
        {"EBELN": "",           "EBELP": "",   "EMATN": "",               "MENGE": "1,500"},
    ]
    alv_mock.GetCellValue.side_effect = lambda r, col: data[r].get(col, "")

    col_map = {
        "stoNo":    ["EBELN"],
        "item":     ["EBELP"],
        "product":  ["EMATN"],
        "orderQty": ["MENGE"]
    }

    records = engine._extract_alv_grid(alv_mock, col_map)
    # Phải giữ lại 2 dòng hợp lệ và loại bỏ dòng tổng số 3
    assert len(records) == 2
    assert records[0]["stoNo"] == "4500012345"
    assert records[0]["item"] == "10"
    assert records[0]["orderQty"] == "1,000"
    assert records[1]["stoNo"] == "4500012346"
    assert records[1]["item"] == "20"


def test_status_bar_error_raises_sap_execution_error():
    """Kiểm tra khi thanh trạng thái SAP báo Error (E) hoặc Abort (A) thì bắn SapExecutionError."""
    engine = SapEngine.__new__(SapEngine)
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("No data found for selection", "E"))

    sess_mock = MagicMock()
    engine.session = sess_mock

    with pytest.raises(SapExecutionError) as exc_info:
        engine.check_sto_me2n(["BIW100"], plant="3093")
    assert "Lỗi SAP ME2N (E)" in str(exc_info.value)


def test_validate_grid_selection_screen_stuck_raises_error():
    """Kiểm tra màn hình bị kẹt ở Selection Screen (btn[8] còn) nhưng không có thông báo không có dữ liệu."""
    engine = SapEngine.__new__(SapEngine)
    # Giả lập còn nút btn[8] (selection screen)
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())

    with pytest.raises(SapExecutionError) as exc_info:
        engine._validate_grid_or_raise(None, "MB52", "Input timeout or lock error", "")
    assert "Không thể mở bảng kết quả SAP MB52" in str(exc_info.value)


def test_validate_grid_no_data_normal_returns_none():
    """Kiểm tra thông báo 'No data found' bình thường thì không văng ngoại lệ."""
    engine = SapEngine.__new__(SapEngine)
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())

    # Không được raise ngoại lệ khi status là No data
    engine._validate_grid_or_raise(None, "MB52", "No data found for criteria", "")


def test_strict_session_safety_guards_user_work():
    """Kiểm tra chế độ strict_session_safety từ chối chiếm phiên nếu user đang mở MIGO."""
    engine = SapEngine.__new__(SapEngine)
    engine.strict_session_safety = True
    engine.connection_index = 0
    engine.session_index = 0

    app_mock = MagicMock()
    app_mock.Children.Count = 1

    conn_mock = MagicMock()
    conn_mock.Children.Count = 1

    sess_user = MagicMock()
    sess_user.Info.Transaction = "MIGO"  # Người dùng đang nhập MIGO!

    conn_mock.Children.side_effect = lambda idx: sess_user
    app_mock.Children.side_effect = lambda idx: conn_mock

    # Giả lập win32com trả về app_mock
    import win32com.client
    with pytest.raises(RuntimeError) as exc_info:
        with pytest.MonkeyPatch.context() as mp:
            mp.setattr(win32com.client, "GetObject", lambda _: MagicMock(GetScriptingEngine=app_mock))
            engine._connect_sap()
    assert "Tất cả các phiên SAP đều đang bận thao tác của người dùng" in str(exc_info.value)



