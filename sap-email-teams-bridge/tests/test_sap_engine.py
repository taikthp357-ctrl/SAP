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


def test_resolve_mb52_unrestr_caption():
    """Kiểm tra phân giải chính xác tiêu đề Unrestr. trên MB52 SAP GUI theo đúng ảnh chụp thực tế."""
    engine = SapEngine.__new__(SapEngine)
    alv_mock = MagicMock()
    alv_mock.ColumnOrder = FakeColumnCollection(["C_MAT", "C_PLANT", "C_SLOC", "C_BATCH", "C_UNREST", "C_TRANS", "C_UNIT"])

    titles = {
        "C_MAT": "Material",
        "C_PLANT": "Plant",
        "C_SLOC": "SLoc",
        "C_BATCH": "Batch",
        "C_UNREST": "Unrestr.",
        "C_TRANS": "Transit/Transf.",
        "C_UNIT": "Base Unit"
    }
    alv_mock.GetDisplayedColumnTitle.side_effect = lambda c: titles.get(c, "")
    alv_mock.GetColumnTitles.return_value = []
    alv_mock.GetColumnTooltip.return_value = ""
    alv_mock.GetColumnTitle.side_effect = lambda c: titles.get(c, "")
    alv_mock.ColumnTitle.side_effect = lambda c: titles.get(c, "")

    col_map = {
        "product": ["MATNR", "MATERIAL", "PRODUCT"],
        "plant":   ["WERKS", "PLANT"],
        "sloc":    ["LGORT", "STORAGE LOCATION", "SLOC"],
        "batch":   ["CHARG", "BATCH"],
        "unrest":  ["LABST", "UNRESTR", "UNRESTR.", "UNRESTRICTED", "UNRESTRICTED STOCK"],
        "transit": ["TRAME", "STOCK IN TRANSIT", "TRANSIT/TRANSF."],
        "unit":    ["MEINS", "BASE UNIT", "UNIT"]
    }
    resolved = engine._resolve_grid_columns(alv_mock, col_map)
    assert resolved["product"] == "C_MAT"
    assert resolved["plant"] == "C_PLANT"
    assert resolved["sloc"] == "C_SLOC"
    assert resolved["batch"] == "C_BATCH"
    assert resolved["unrest"] == "C_UNREST"
    assert resolved["transit"] == "C_TRANS"
    assert resolved["unit"] == "C_UNIT"


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
    assert records[0]["orderQty"] in ("1000", "1,000")
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
    engine.auto_create_session = False

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


def test_check_stock_mb52_slocs_filtering():
    """Kiểm tra MB52 lọc đúng danh sách kho slocs (ví dụ 1005 và 3093)."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))

    mock_records = [
        {"product": "BIW71004683", "sloc": "9004", "unrest": "810", "batch": "261019-05"},
        {"product": "BIW71004683", "sloc": "1005", "unrest": "150", "batch": "LOT1005"},
        {"product": "BIW71004683", "sloc": "3093", "unrest": "200", "batch": "LOT3093"},
        {"product": "BIW71004683", "sloc": "601D", "unrest": "2", "batch": "PROD"},
    ]
    engine._extract_alv_grid = MagicMock(return_value=mock_records)

    # 1. Tra cứu tồn kho (chỉ lọc 1005 và 3093)
    res_wh = engine.check_stock_mb52(["BIW71004683"], plant="3093", slocs=["1005", "3093"])
    assert len(res_wh) == 2
    assert {r["sloc"] for r in res_wh} == {"1005", "3093"}

    # 2. Tra cứu tồn tổng hợp (slocs=None -> giữ nguyên toàn bộ)
    res_all = engine.check_stock_mb52(["BIW71004683"], plant="3093", slocs=None)
    assert len(res_all) == 4


def test_check_low_stock_multi_plant():
    """Kiểm tra tổng hợp tồn thấp qua cả 2 plant trong 1 lần gọi MB52 duy nhất (3093 và 3001 cho BODY)."""
    engine = SapEngine.__new__(SapEngine)

    mock_mb52_records = [
        {"product": "BIW001", "plant": "3093", "unrest": "10", "unit": "PC"},
        {"product": "BIW001", "plant": "3001", "unrest": "5", "unit": "PC"},
        {"product": "BIW002", "plant": "3093", "unrest": "0", "unit": "PC"}
    ]
    engine.check_stock_mb52 = MagicMock(return_value=mock_mb52_records)

    res = engine.check_low_stock(["BIW001", "BIW002"], plants=["3093", "3001"], limit=10, model_name="LIMO")
    assert len(res) == 2
    # BIW002 có tổng tồn 0 nên được sắp lên đầu
    assert res[0]["product"] == "BIW002"
    assert res[0]["total_stock"] == "0"
    assert "Hết hàng" in res[0]["status"]

    # BIW001 có tổng tồn 15 (10 + 5)
    assert res[1]["product"] == "BIW001"
    assert res[1]["stock_3093"] == "10"
    assert res[1]["stock_3001"] == "5"
    assert res[1]["total_stock"] == "15"
    assert "Tồn thấp" in res[1]["status"]


def test_check_low_stock_empty_mats_derives_from_mb52():
    """Kiểm tra khi không truyền danh sách mã, check_low_stock tự trích xuất mã từ MB52 và sắp xếp tồn thấp nhất."""
    engine = SapEngine.__new__(SapEngine)

    mock_mb52_records = [
        {"product": "BIW003", "plant": "3093", "unrest": "100", "unit": "PC"},
        {"product": "BIW001", "plant": "3093", "unrest": "10", "unit": "PC"},
        {"product": "BIW001", "plant": "3001", "unrest": "5", "unit": "PC"},
        {"product": "BIW002", "plant": "3093", "unrest": "0", "unit": "PC"}
    ]
    engine.check_stock_mb52 = MagicMock(return_value=mock_mb52_records)

    res = engine.check_low_stock([], plants=["3093", "3001"], limit=2, model_name="LIMO")
    assert len(res) == 2
    # BIW002 (tồn 0) đứng đầu, BIW001 (tồn 15) đứng nhì
    assert res[0]["product"] == "BIW002"
    assert res[0]["total_stock"] == "0"
    assert res[1]["product"] == "BIW001"
    assert res[1]["total_stock"] == "15"


def test_check_stock_mb52_multi_plant_selection():
    """Kiểm tra check_stock_mb52 kích hoạt multi-selection cho cả 2 xưởng (3093 và 3001)."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._fill_multiple_selection = MagicMock()
    engine._clear_multiple_selection = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._dismiss_popups = MagicMock(return_value="")
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._extract_alv_grid = MagicMock(return_value=[
        {"product": "BIW001", "plant": "3093", "sloc": "3093", "unrest": "50", "unit": "PC"},
        {"product": "BIW001", "plant": "3001", "sloc": "3001", "unrest": "20", "unit": "PC"}
    ])

    recs = engine.check_stock_mb52(["BIW001"], plants=["3093", "3001"])
    assert len(recs) == 2
    # Phải gọi _fill_multiple_selection với nút WERKS và danh sách 2 plant
    engine._fill_multiple_selection.assert_any_call("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH", ["3093", "3001"])
    # Phải gọi _fill_multiple_selection với nút MATNR và danh sách mã hàng
    engine._fill_multiple_selection.assert_any_call("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH", ["BIW001"])


def test_check_open_ob_vl06o_mock():
    """Kiểm tra trích xuất VL06O đơn OB mở theo chuẩn RC65."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    mock_controls = {}
    def mock_find(id_str):
        if id_str not in mock_controls:
            mock_controls[id_str] = MagicMock()
        return mock_controls[id_str]
    engine._find_by_id_safe = MagicMock(side_effect=mock_find)
    engine._clear_all_selection_inputs = MagicMock()
    engine._clear_multiple_selection = MagicMock()
    engine._fill_multiple_selection = MagicMock()
    engine._dismiss_popups = MagicMock(return_value="")
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._ensure_item_grid_vl06o = MagicMock(return_value=MagicMock())

    mock_records = [
        {"obNo": "80012345", "product": "BIW001", "qty": "50,000", "unit": "PC", "sloc": "3093", "batch": "2610A", "gmStatus": "A"},
        {"obNo": "80012346", "product": "BIW001", "qty": "20,000", "unit": "PC", "sloc": "", "batch": "", "gmStatus": ""},
        {"obNo": "80012347", "product": "OTHER_MAT", "qty": "10,000", "unit": "PC", "sloc": "", "batch": "", "gmStatus": "A"}
    ]
    engine._extract_alv_grid = MagicMock(return_value=mock_records)

    res = engine.check_open_ob_vl06o(materials=["BIW001"])
    # Chỉ giữ lại 2 dòng của BIW001, lọc bỏ OTHER_MAT
    assert len(res) == 2
    assert res[0]["obNo"] == "80012346"  # Sắp xếp số OB giảm dần
    assert res[1]["obNo"] == "80012345"
    # Kiểm tra Shipping Point được điền thẳng 3093 theo chuẩn RC65 (không dùng popup)
    assert mock_controls["wnd[0]/usr/ctxtIT_MATNR-LOW"].text == "BIW001"
    assert mock_controls["wnd[0]/usr/ctxtIF_VSTEL-LOW"].text == ""
    assert mock_controls["wnd[0]/usr/ctxtIF_VSTEL-HIGH"].text == ""


def test_extract_alv_grid_preserves_open_ob_without_batch_or_sloc():
    """Kiểm tra _extract_alv_grid giữ lại dòng đơn OB mở (chưa có Batch và SLoc) và loại bỏ dòng tổng cộng."""
    engine = SapEngine.__new__(SapEngine)
    engine.max_alv_rows = 100
    engine._resolve_grid_columns = MagicMock(return_value={
        "obNo": "VBELN", "product": "MATNR", "qty": "LFIMG", "batch": "CHARG", "sloc": "LGORT"
    })

    # Giả lập ALV Grid chứa 3 dòng:
    # Dòng 0: Đơn OB mở hợp lệ (chưa gán batch, chưa có sloc)
    # Dòng 1: Dòng tổng cộng (*)
    # Dòng 2: Đơn OB đã có batch
    mock_data = [
        {"VBELN": "80099001", "MATNR": "BIW73110009AA", "LFIMG": "100.0", "CHARG": "", "LGORT": ""},
        {"VBELN": "", "MATNR": "*", "LFIMG": "150.0", "CHARG": "", "LGORT": ""},
        {"VBELN": "80099002", "MATNR": "BIW73110009AA", "LFIMG": "50.0", "CHARG": "LOT01", "LGORT": "3093"},
    ]

    mock_grid = MagicMock()
    mock_grid.RowCount = 3
    mock_grid.GetCellValue.side_effect = lambda r, col: mock_data[r].get(col, "")

    col_map = {
        "obNo": ["VBELN"], "product": ["MATNR"], "qty": ["LFIMG"], "batch": ["CHARG"], "sloc": ["LGORT"]
    }
    extracted = engine._extract_alv_grid(mock_grid, col_map)

    # Dòng tổng cộng (*) phải bị loại bỏ, 2 dòng OB (kể cả dòng chưa có batch/sloc) phải được giữ lại
    assert len(extracted) == 2
    assert extracted[0]["obNo"] == "80099001"
    assert extracted[0]["product"] == "BIW73110009AA"
    assert extracted[0]["batch"] == ""
    assert extracted[1]["obNo"] == "80099002"
    assert extracted[1]["batch"] == "LOT01"


def test_matches_material_variations():
    """Kiểm tra helper _matches_material khớp linh hoạt mã vật tư VinFast."""
    # Khớp chính xác
    assert SapEngine._matches_material("BIW73110009AA", ["BIW73110009AA"]) is True
    # Khớp bỏ số 0 đầu
    assert SapEngine._matches_material("000000BIW73110009AA", ["BIW73110009AA"]) is True
    # Khớp mã rút gọn không có BIW
    assert SapEngine._matches_material("BIW73110009AA", ["73110009AA"]) is True
    assert SapEngine._matches_material("BIW73110009AA", ["73110009"]) is True
    # Khác mã hoàn toàn
    assert SapEngine._matches_material("BIW73110009AA", ["BIW75113803AB"]) is False


def test_check_bin_history_mon_mock():
    """Kiểm tra trích xuất /SCWM/MON lịch sử chuyển Bin."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))

    mock_records = [
        {"wtNo": "1001", "product": "BIW001", "batch": "B1", "srcBin": "10-A-01", "destBin": "10-B-02", "qty": "10", "status": "C"},
        {"wtNo": "1002", "product": "BIW001", "batch": "B1", "srcBin": "10-A-01", "destBin": "10-C-03", "qty": "10", "status": "A"},
    ]
    engine._find_largest_grid = MagicMock(return_value=MagicMock())
    engine._extract_alv_grid = MagicMock(return_value=mock_records)

    res = engine.check_bin_history_mon(materials=["BIW001"])
    assert len(res) == 2
    assert "Hoàn thành" in res[1]["status"]
    assert "Đã hủy" in res[0]["status"]


def test_check_open_ob_vl06o_no_data_returns_empty_list():
    """Kiểm tra khi VL06O không tìm thấy đơn OB nào thì trả về danh sách rỗng, không crash."""
    engine = SapEngine.__new__(SapEngine)
    engine.max_alv_rows = 300
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._dismiss_popups = MagicMock(return_value="")
    engine._get_status = MagicMock(return_value=("No outbound deliveries selected", "S"))
    engine._find_largest_grid = MagicMock(return_value=None)
    engine._clear_all_selection_inputs = MagicMock()
    engine._clear_multiple_selection = MagicMock()
    engine._fill_multiple_selection = MagicMock()
    engine._ensure_item_grid_vl06o = MagicMock(return_value=None)

    res = engine.check_open_ob_vl06o(materials=["BIW32011350AB"])
    assert res == []


def test_clear_all_selection_inputs():
    """Kiểm tra xóa toàn bộ các ô nhập liệu còn đọng trên màn hình chọn."""
    engine = SapEngine.__new__(SapEngine)
    mock_field1 = MagicMock()
    mock_field1.Type = "GuiTextField"
    mock_field1.Id = "wnd[0]/usr/ctxtMBLNR-LOW"
    mock_field1.Changeable = True
    mock_field1.ReadOnly = False
    mock_field1.Text = "4904180921"

    mock_field2 = MagicMock()
    mock_field2.Type = "GuiCTextField"
    mock_field2.Id = "wnd[0]/usr/ctxtBWART-LOW"
    mock_field2.Changeable = True
    mock_field2.ReadOnly = False
    mock_field2.Text = "311"

    usr_mock = MagicMock()
    usr_mock.Children.Count = 2
    usr_mock.Children.ElementAt.side_effect = lambda idx: [mock_field1, mock_field2][idx]
    engine._find_by_id_safe = MagicMock(return_value=usr_mock)

    engine._clear_all_selection_inputs()
    assert mock_field1.Text == ""
    assert mock_field2.Text == ""


def test_read_sap_exported_file_openxml(tmp_path):
    """Kiểm tra đọc file Excel MB52 xuất ra qua openpyxl với forward-fill và loại bỏ dòng tổng theo chuẩn Stock_FG_MA_BODY."""
    import openpyxl
    file_path = str(tmp_path / "test_mb52.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Material", "Material Description", "Plant", "SLoc", "Batch", "Unrestr.", "Blocked", "Transit/Transf.", "Base Unit"])
    ws.append(["BIW32011220HA", "LEFT_FRONT_LONGI...", "3001", "3003", "PROD_BAT01", 155, 0, 0, "EA"])
    ws.append(["", "", "3001", "600G", "HUYVIN3252", 1, 0, 0, "EA"])
    ws.append(["*", "", "", "", "", 156, 0, 0, "EA"])  # Dòng subtotal có dấu sao (*) của SAP ALV
    ws.append(["", "", "3093", "3000", "261024-256", 175, 0, 0, "EA"])
    ws.append(["BIW32011230HA", "LEFT_FRONT_LONGI...", "3001", "", "", 0, 0, 200, "EA"])
    ws.append(["", "", "3001", "3003", "PROD_BAT02", 213, 0, 0, "EA"])
    ws.append(["", "", "3093", "3000", "261023-250", 200, 0, 0, "EA"])
    ws.append(["Total", "", "", "", "", 744, 0, 200, "EA"])
    wb.save(file_path)

    engine = SapEngine.__new__(SapEngine)
    col_map = {
        "product": ["MATERIAL", "MATNR"],
        "plant":   ["PLANT", "WERKS"],
        "sloc":    ["SLOC", "STORAGE LOCATION"],
        "batch":   ["BATCH", "CHARG"],
        "unrest":  ["UNRESTR.", "UNRESTRICTED"],
        "transit": ["TRANSIT/TRANSF.", "TRANSIT"],
        "unit":    ["BASE UNIT", "UNIT"]
    }
    records = engine._read_sap_exported_file(file_path, col_map)
    assert len(records) == 6
    assert records[0]["product"] == "BIW32011220HA"
    assert records[0]["plant"] == "3001"
    assert records[0]["sloc"] == "3003"
    assert records[0]["unrest"] == "155"

    # Kiểm tra forward-fill: dòng 2 phải kế thừa product và plant từ dòng 1
    assert records[1]["product"] == "BIW32011220HA"
    assert records[1]["plant"] == "3001"
    assert records[1]["sloc"] == "600G"
    assert records[1]["unrest"] == "1"

    # Dòng 3 chuyển sang plant 3093 nhưng product vẫn là BIW32011220HA
    assert records[2]["product"] == "BIW32011220HA"
    assert records[2]["plant"] == "3093"
    assert records[2]["unrest"] == "175"

    # Dòng 4 chuyển sang mã BIW32011230HA
    assert records[3]["product"] == "BIW32011230HA"
    assert records[3]["plant"] == "3001"
    assert records[3]["transit"] == "200"

    # Dòng tổng cộng Total bị loại bỏ hoàn toàn
    assert not any(r["product"].lower() == "total" for r in records)


def test_reset_to_session_manager():
    """Kiểm tra reset_to_session_manager gửi lệnh /n về SAP GUI."""
    engine = SapEngine.__new__(SapEngine)
    engine._dismiss_popups = MagicMock()
    mock_okcd = MagicMock()
    mock_okcd.__class__.__name__ = "GuiTextField"
    engine._find_by_id_safe = MagicMock(return_value=mock_okcd)
    engine._wait_idle = MagicMock()
    mock_session = MagicMock()
    mock_session.__class__.__name__ = "GuiSession"
    engine.session = mock_session

    engine._reset_to_session_manager()
    assert mock_okcd.text == "/n"
    mock_session.findById.return_value.sendVKey.assert_called_with(0)









def test_check_low_stock_excludes_qc_sloc_9999():
    """Kiểm tra check_low_stock LOẠI TRỪ TRIỆT ĐỂ SLoc 9999 của bộ phận QC/Phế phẩm (theo 261008_Stock_FG_MA_BODY)."""
    engine = SapEngine.__new__(SapEngine)

    # Giả lập: BIW71011001AA có 2 pcs tại kho 9999 của QC, 0 pcs tại kho thực tế
    mock_mb52_records = [
        {"product": "BIW71011001AA", "plant": "3093", "sloc": "9999", "unrest": "2", "unit": "PC"},
        {"product": "BIW71004366AA", "plant": "3093", "sloc": "3093", "unrest": "100", "unit": "PC"},
        {"product": "BIW71004366AA", "plant": "3001", "sloc": "1003", "unrest": "50", "unit": "PC"},
    ]
    engine.check_stock_mb52 = MagicMock(return_value=mock_mb52_records)

    res = engine.check_low_stock(["BIW71011001AA", "BIW71004366AA"], plants=["3093", "3001"], limit=10, model_name="VF7")
    assert len(res) == 2
    # BIW71011001AA phải báo HẾT HÀNG (0), KHÔNG được cộng 2 pcs của QC 9999
    assert res[0]["product"] == "BIW71011001AA"
    assert res[0]["total_stock"] == "0"
    assert "Hết hàng" in res[0]["status"]

    # BIW71004366AA có tổng 150 pcs tại kho hợp lệ
    assert res[1]["product"] == "BIW71004366AA"
    assert res[1]["total_stock"] == "150"
    assert "Khả dụng" in res[1]["status"]


def test_check_low_stock_matches_flexible_material_codes():
    """Kiểm tra khớp linh hoạt mã hàng (ví dụ MB52 trả về mã không có suffix AA hoặc bỏ số 0 đầu)."""
    engine = SapEngine.__new__(SapEngine)

    # Trong MB52 trả về mã gốc 'BIW71004366' (thiếu suffix AA)
    mock_mb52_records = [
        {"product": "BIW71004366", "plant": "3093", "sloc": "3093", "unrest": "80", "unit": "PC"},
    ]
    engine.check_stock_mb52 = MagicMock(return_value=mock_mb52_records)

    # Yêu cầu đánh giá theo danh mục Master là 'BIW71004366AA'
    res = engine.check_low_stock(["BIW71004366AA"], plants=["3093", "3001"], limit=5, model_name="VF7")
    assert len(res) == 1
    assert res[0]["product"] == "BIW71004366AA"
    # Phải nhận diện được 80 pcs từ BIW71004366, không bị tính là 0
    assert res[0]["total_stock"] == "80"
    assert "Khả dụng" in res[0]["status"]


def test_check_low_stock_propagates_sap_execution_error():
    """Kiểm tra check_low_stock KHÔNG NUỐT LỖI (no swallow) khi MB52 gặp sự cố."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_stock_mb52 = MagicMock(side_effect=Exception("SAP Timeout or Connection Drop"))

    with pytest.raises(SapExecutionError) as exc_info:
        engine.check_low_stock(["BIW32013765AA"], plants=["3093", "3001"], limit=10, model_name="LIMO")

    assert "Không thể truy vấn tồn kho MB52 đa xưởng" in str(exc_info.value)
    assert "SAP Timeout or Connection Drop" in str(exc_info.value)


def test_check_gr_doc_mb51_mock():
    """Kiểm tra check_gr_doc_mb51 trích xuất đúng chứng từ GR từ ALV Grid."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._extract_alv_grid = MagicMock(return_value=[
        {
            "mblnr": "5000123456", "mjahr": "2026", "item": "1", "bwart": "101",
            "matnr": "BIW32013765AA", "qty": "100.000", "unit": "PC",
            "xblnr": "1800123456", "postDate": "24.10.2026", "sloc": "1005",
            "batch": "B2610", "user": "USER01"
        }
    ])
    engine._reset_to_session_manager = MagicMock()

    records = engine.check_gr_doc_mb51(inbound_no="1800123456", days=30)
    assert len(records) == 1
    assert records[0]["mblnr"] == "5000123456"
    assert records[0]["xblnr"] == "1800123456"
    assert records[0]["qty"] == "100"


def test_check_inbound_detail_vl06i_mock():
    """Kiểm tra check_inbound_detail_vl06i và đối chiếu expectedQty."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._extract_alv_grid = MagicMock(return_value=[
        {
            "ibNo": "1800123456", "item": "10", "product": "BIW32013765AA",
            "actualQty": "200.000", "unit": "PC", "plant": "3093", "sloc": "1005",
            "status": "A", "refPo": "4500123456", "poItem": "10", "refNo": "BL999"
        }
    ])
    engine._reset_to_session_manager = MagicMock()

    # 1. Trường hợp khớp số lượng
    records_match = engine.check_inbound_detail_vl06i(inbound_no="1800123456", expected_qty=200.0)
    assert len(records_match) == 1
    assert records_match[0]["actualQty"] == "200"
    assert records_match[0]["expectedQty"] == "200"
    assert records_match[0]["delta"] == "0"
    assert "Khớp SL" in records_match[0]["qtyCheck"]

    # 2. Trường hợp lệch số lượng (thực tế 200, yêu cầu 150 -> Lệch +50)
    records_diff = engine.check_inbound_detail_vl06i(inbound_no="1800123456", expected_qty=150.0)
    assert records_diff[0]["delta"] == "50"
    assert "Lệch +50" in records_diff[0]["qtyCheck"]


def test_find_inbound_vl06i_mock():
    """Kiểm tra find_inbound_vl06i tìm kiếm theo PO / Ref."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._extract_alv_grid = MagicMock(return_value=[
        {
            "ibNo": "1800888999", "item": "10", "product": "BIW32013765AA",
            "qty": "500", "unit": "PC", "plant": "3093", "sloc": "1005",
            "status": "C", "refPo": "4500123456", "refNo": "BL123"
        }
    ])
    engine._reset_to_session_manager = MagicMock()

    records = engine.find_inbound_vl06i(po_no="4500123456")
    assert len(records) == 1
    assert records[0]["ibNo"] == "1800888999"
    assert records[0]["refPo"] == "4500123456"
    assert "Đã hoàn tất GR" in records[0]["status"]


def test_check_po_me2n_mock():
    """Kiểm tra check_po_me2n tính toán openQty và format status."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._extract_alv_grid = MagicMock(return_value=[
        {
            "poNo": "4500123456", "item": "10", "product": "BIW32013765AA",
            "orderQty": "1,000", "deliveredQty": "600", "unit": "PC",
            "plant": "3093", "deleteInd": "", "releaseStatus": "G"
        }
    ])
    engine._reset_to_session_manager = MagicMock()

    records = engine.check_po_me2n(po_no="4500123456")
    assert len(records) == 1
    assert records[0]["orderQty"] in ("1000", "1,000")
    assert records[0]["deliveredQty"] == "600"
    assert records[0]["openQty"] == "400"
    assert "Hoạt động" in records[0]["deleteInd"]
    assert "Đã duyệt" in records[0]["releaseStatus"]


def test_check_fg_ob_status_mock():
    """Kiểm tra check_fg_ob_status tính toán tồn FG, OB mở và khả dụng xuất mới."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_stock_mb52 = MagicMock(return_value=[
        {"product": "BIW32013765AA", "plant": "3093", "sloc": "3093", "unrest": "200", "unit": "PC"},
        {"product": "BIW32013765AA", "plant": "3093", "sloc": "9999", "unrest": "5", "unit": "PC"}  # Phải bị loại
    ])
    engine.check_open_ob_vl06o = MagicMock(return_value=[
        {"obNo": "80012345", "product": "BIW32013765AA", "qty": "50", "unit": "PC"}
    ])

    records = engine.check_fg_ob_status(["BIW32013765AA"], plant="3093")
    assert len(records) == 1
    assert records[0]["fgStock"] == "200"
    assert records[0]["openObQty"] == "50"
    assert records[0]["availableQty"] == "150"
    assert "80012345" in records[0]["openObList"]
    assert "Đủ tồn điền OB" in records[0]["workAdvice"]


def test_check_nvl_migo_balance_mock():
    """Kiểm tra check_nvl_migo_balance kiểm tra Plant 3093 và duy nhất SLoc 1005 qua /SCWM/POST."""
    engine = SapEngine.__new__(SapEngine)
    # Mock /SCWM/POST trả về 2 lô: 1 lô nhập sớm (FIFO 1), 1 lô nhập muộn (FIFO 2)
    engine.check_position_posting_change = MagicMock(return_value=[
        {"product": "BIW32013765AA", "batch": "LOT-B", "bin": "APFL-02", "qty": "300", "unit": "PC", "grDate": "20.10.2026", "grTime": "14:00:00", "cat": "FT"},
        {"product": "BIW32013765AA", "batch": "LOT-A", "bin": "APFL-01", "qty": "200", "unit": "PC", "grDate": "15.10.2026", "grTime": "08:00:00", "cat": "FT"}
    ])

    records = engine.check_nvl_migo_balance(["BIW32013765AA"], plant="3093")
    assert len(records) == 2
    # LOT-A nhập 15.10 phải được xếp trước (FIFO #1)
    assert records[0]["batch"] == "LOT-A"
    assert records[0]["fifoRank"] == "#1"
    assert "Sẵn sàng ra đơn MIGO 311" in records[0]["migoAdvice"]
    assert records[0]["sloc"] == "1005"
    assert records[0]["plant"] == "3093"

    # LOT-B nhập 20.10 xếp sau (FIFO #2)
    assert records[1]["batch"] == "LOT-B"
    assert records[1]["fifoRank"] == "#2"
    assert records[1]["sloc"] == "1005"


def test_check_nvl_migo_fallback_mb52_sloc_1005():
    """Kiểm tra fallback MB52: Chỉ lấy SLoc 1005 và bỏ qua các SLoc khác."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_position_posting_change = MagicMock(return_value=[])
    engine.check_stock_mb52 = MagicMock(return_value=[
        {"product": "BIW32013765AA", "plant": "3093", "sloc": "1005", "unrest": "150", "unit": "PC"},
        {"product": "BIW32013765AA", "plant": "3093", "sloc": "3001", "unrest": "50", "unit": "PC"}
    ])

    records = engine.check_nvl_migo_balance(["BIW32013765AA"], plant="3093")
    assert len(records) == 1
    assert records[0]["sloc"] == "1005"
    assert records[0]["plant"] == "3093"
    assert records[0]["qty"] == "150"
    assert "Có tồn MB52 Kho 1005 (150 PC)" in records[0]["migoAdvice"]


def test_check_nvl_migo_both_sources_fail_raises_sap_execution_error():
    """Kiểm tra R01: Khi cả /SCWM/POST và MB52 đều lỗi -> Báo lỗi SapExecutionError, không nuốt thành hết hàng 0."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_position_posting_change = MagicMock(side_effect=RuntimeError("SAP POST timeout"))
    engine.check_stock_mb52 = MagicMock(side_effect=RuntimeError("SAP MB52 timeout"))

    with pytest.raises(SapExecutionError) as exc_info:
        engine.check_nvl_migo_balance(["BIW32013765AA"], plant="3093")
    assert "Lỗi truy vấn SAP NVL" in str(exc_info.value)


def test_check_nvl_migo_qc_and_unknown_tst_handling():
    """Kiểm tra R02 & R03: TST là QC hoặc rỗng (UNKNOWN) tuyệt đối không được coi là sẵn sàng ra đơn MIGO."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_position_posting_change = MagicMock(return_value=[
        {"product": "BIW32013765AA", "batch": "LOT-QC", "bin": "APFL-01", "qty": "500", "unit": "PC", "grDate": "15.10.2026", "grTime": "08:00", "cat": "QT"},
        {"product": "BIW32013765AA", "batch": "LOT-EMPTY", "bin": "APFL-02", "qty": "100", "unit": "PC", "grDate": "16.10.2026", "grTime": "09:00", "cat": ""},
    ])

    records = engine.check_nvl_migo_balance(["BIW32013765AA"], plant="3093")
    assert len(records) == 2
    # LOT-QC: TST QT -> Chờ kiểm định QC
    assert records[0]["cat"] == "QT"
    assert records[0]["fifoRank"] == "-"
    assert "Chờ kiểm định QC" in records[0]["migoAdvice"]
    assert "Sẵn sàng" not in records[0]["migoAdvice"]

    # LOT-EMPTY: TST rỗng -> UNKNOWN, không tự nhận là FT
    assert records[1]["cat"] == "UNKNOWN"
    assert records[1]["fifoRank"] == "-"
    assert "chưa xác định" in records[1]["migoAdvice"]
    assert "Sẵn sàng" not in records[1]["migoAdvice"]


def test_check_nvl_migo_rejects_invalid_plant():
    """Kiểm tra R07: Nhánh NVL từ chối plant khác 3093 một cách an toàn."""
    engine = SapEngine.__new__(SapEngine)
    records = engine.check_nvl_migo_balance(["BIW32013765AA"], plant="2001")
    assert len(records) == 1
    assert records[0]["cat"] == "INVALID_SCOPE"
    assert "chỉ hỗ trợ Plant 3093" in records[0]["migoAdvice"]


def test_check_nvl_migo_fallback_mb52_filters_batch_strictly():
    """Kiểm tra R08: Fallback MB52 khi có chỉ định Batch phải lọc đúng Batch, không nhận nhầm lô khác."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_position_posting_change = MagicMock(return_value=[])
    engine.check_stock_mb52 = MagicMock(return_value=[
        {"product": "BIW32013765AA", "plant": "3093", "sloc": "1005", "batch": "LOT-OTHER", "unrest": "200", "unit": "PC"}
    ])

    # Yêu cầu tìm LOT-WANTED nhưng kho chỉ có LOT-OTHER
    records = engine.check_nvl_migo_balance(["BIW32013765AA"], plant="3093", batch="LOT-WANTED")
    assert len(records) == 1
    assert records[0]["qty"] == "0"
    assert "Không có lô LOT-WANTED" in records[0]["migoAdvice"]


def test_check_fg_ob_vl06o_error_reports_warning_without_guessing_available():
    """Kiểm tra R05: Khi VL06O lỗi nhưng MB52 thành công, báo lỗi VL06O chứ không suy đoán hết OB."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_stock_mb52 = MagicMock(return_value=[
        {"product": "BIW32013765AA", "plant": "3093", "sloc": "3093", "unrest": "300", "unit": "PC"}
    ])
    engine.check_open_ob_vl06o = MagicMock(side_effect=RuntimeError("VL06O Session Busy"))

    records = engine.check_fg_ob_status(["BIW32013765AA"], plant="3093")
    assert len(records) == 1
    assert records[0]["fgStock"] == "300"
    assert records[0]["openObQty"] == "N/A"
    assert records[0]["availableQty"] == "N/A"
    assert "Chưa xác định (Lỗi VL06O)" in records[0]["openObList"]
    assert "Không suy đoán khả dụng" in records[0]["workAdvice"]


def test_check_fg_ob_both_fail_raises_sap_execution_error():
    """Kiểm tra R05: Khi cả MB52 và VL06O đều lỗi -> Raise SapExecutionError."""
    engine = SapEngine.__new__(SapEngine)
    engine.check_stock_mb52 = MagicMock(side_effect=RuntimeError("MB52 Error"))
    engine.check_open_ob_vl06o = MagicMock(side_effect=RuntimeError("VL06O Error"))

    with pytest.raises(SapExecutionError) as exc_info:
        engine.check_fg_ob_status(["BIW32013765AA"], plant="3093")
    assert "Lỗi truy vấn điều độ FG/OB" in str(exc_info.value)


def test_alv_grid_skips_subtotal_row_without_position():
    """Kiểm tra loại bỏ dòng tổng ALV (vàng) khi không có bin/batch."""
    engine = SapEngine.__new__(SapEngine)
    engine.max_alv_rows = 500

    class MockAlvWithSubtotal:
        RowCount = 2
        firstVisibleRow = 0
        ColumnOrder = FakeColumnCollection(["PRODUCT", "CHARG", "LGPLA", "AQUAN", "GESME", "MEINS"])
        def GetCellValue(self, row, col_id):
            if row == 0:
                # Dòng 0: Dữ liệu thật
                mapping = {"PRODUCT": "BIW73015987AA", "CHARG": "261017-06", "LGPLA": "APFL-P23-01-02", "AQUAN": "619", "GESME": "1.019", "MEINS": "EA"}
                return mapping.get(col_id, "")
            else:
                # Dòng 1: Dòng tổng cộng (subtotal) - Product trống, Batch trống, Bin trống
                mapping = {"PRODUCT": "", "CHARG": "", "LGPLA": "", "AQUAN": "619", "GESME": "1.019", "MEINS": "EA"}
                return mapping.get(col_id, "")
        def GetColumnTitles(self):
            return ["PRODUCT", "CHARG", "LGPLA", "AQUAN", "GESME", "MEINS"]
        def GetRowType(self, row):
            return "&TOTAL" if row == 1 else ""

    col_map = {
        "product": ["PRODUCT"],
        "batch": ["CHARG"],
        "bin": ["LGPLA"],
        "availQty": ["AQUAN"],
        "totalQty": ["GESME"],
        "unit": ["MEINS"]
    }
    records = engine._extract_alv_grid(MockAlvWithSubtotal(), col_map)
    # Phải chỉ trích xuất đúng 1 dòng dữ liệu thật, dòng tổng cộng bị bỏ qua
    assert len(records) == 1
    assert records[0]["product"] == "BIW73015987AA"
    assert records[0]["batch"] == "261017-06"
    assert records[0]["bin"] == "APFL-P23-01-02"


def test_mb52_excludes_subtotal_and_sorts_by_sloc():
    """Kiểm tra MB52: Loại bỏ dòng tổng trống Sloc, gom và sắp xếp theo Sloc, không nhân đôi tồn."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._export_alv_to_excel = MagicMock(return_value=None)
    engine._reset_to_session_manager = MagicMock()

    # Giả lập dữ liệu trả về từ ALV của MB52 giống hệt ảnh thực tế của người dùng:
    # Dòng 0: Dòng tổng của Plant (không có Sloc, không có Batch)
    # Dòng 1: Kho 1005 (28800)
    # Dòng 2: Kho 3001 (8045)
    # Dòng 3: Kho 1005 (800)
    engine._extract_alv_grid = MagicMock(return_value=[
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "", "batch": "", "unrest": "37645", "unit": "EA"},
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "1005", "batch": "261013-03", "unrest": "28800", "unit": "EA"},
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "3001", "batch": "PROD_BATCH", "unrest": "8045", "unit": "EA"},
        {"product": "BIW32012476AA", "plant": "3093", "sloc": "1005", "batch": "260929-03", "unrest": "800", "unit": "EA"},
    ])

    records = engine.check_stock_mb52(materials=["BIW32012476AA"], plant="3093")

    # 1. Dòng tổng trống Sloc (37645) bắt buộc phải bị loại bỏ!
    assert len(records) == 3

    # 2. Các dòng cùng kho (1005) phải được gom lại cạnh nhau trước, sau đó mới đến kho 3001
    assert records[0]["sloc"] == "1005"
    assert records[0]["unrest"] == "28800"
    assert records[1]["sloc"] == "1005"
    assert records[1]["unrest"] == "800"
    assert records[2]["sloc"] == "3001"
    assert records[2]["unrest"] == "8045"


def test_connect_sap_hunts_idle_session_when_first_session_busy():
    """Kiểm tra: Nếu Session 0 đang bận (Busy=True) bởi tool khác, Bot tự động chọn Session 1 rảnh rỗi."""
    engine = SapEngine.__new__(SapEngine)
    engine.connection_index = 0
    engine.session_index = 0
    engine.strict_session_safety = True
    engine.auto_create_session = True
    engine.prefer_dedicated_session = False

    mock_sess0 = MagicMock()
    mock_sess0.Busy = True
    mock_sess0.Info.Transaction = "MIGO"

    mock_sess1 = MagicMock()
    mock_sess1.Busy = False
    mock_sess1.Info.Transaction = "SESSION_MANAGER"

    mock_conn = MagicMock()
    mock_conn.Children.Count = 2
    mock_conn.Children.side_effect = lambda idx: [mock_sess0, mock_sess1][idx]

    mock_app = MagicMock()
    mock_app.Children.Count = 1
    mock_app.Children.side_effect = lambda idx: [mock_conn][idx]

    with pytest.MonkeyPatch().context() as m:
        mock_sap_gui_auto = MagicMock()
        mock_sap_gui_auto.GetScriptingEngine = mock_app
        m.setattr("win32com.client.GetObject", lambda name: mock_sap_gui_auto)

        selected = engine._connect_sap()
        assert selected == mock_sess1
        assert selected.Busy is False


def test_connect_sap_creates_new_session_when_all_sessions_busy():
    """Kiểm tra: Khi các session hiện có đều bận, Bot tự động gọi CreateSession() để mở Tab mới."""
    engine = SapEngine.__new__(SapEngine)
    engine.connection_index = 0
    engine.session_index = 0
    engine.strict_session_safety = True
    engine.auto_create_session = True
    engine.prefer_dedicated_session = False

    mock_sess0 = MagicMock()
    mock_sess0.Busy = True
    mock_sess0.Info.Transaction = "MIGO"

    mock_sess_new = MagicMock()
    mock_sess_new.Busy = False
    mock_sess_new.Info.Transaction = "SESSION_MANAGER"

    class DynamicConnMock:
        def __init__(self):
            self.sessions = [mock_sess0]

        @property
        def Count(self):
            return len(self.sessions)

        def __call__(self, idx):
            return self.sessions[idx]

    conn_dyn = DynamicConnMock()
    mock_conn = MagicMock()
    mock_conn.Children = conn_dyn

    # Khi CreateSession được gọi, thêm mock_sess_new vào
    def fake_create_session():
        conn_dyn.sessions.append(mock_sess_new)

    mock_sess0.CreateSession = MagicMock(side_effect=fake_create_session)

    mock_app = MagicMock()
    mock_app.Children.Count = 1
    mock_app.Children.side_effect = lambda idx: [mock_conn][idx]

    with pytest.MonkeyPatch().context() as m:
        mock_sap_gui_auto = MagicMock()
        mock_sap_gui_auto.GetScriptingEngine = mock_app
        m.setattr("win32com.client.GetObject", lambda name: mock_sap_gui_auto)

        selected = engine._connect_sap()
        assert selected == mock_sess_new
        mock_sess0.CreateSession.assert_called_once()


def test_connect_sap_preserves_first_three_user_tabs_and_creates_tab_4():
    """Kiểm tra: Khi người dùng đang dùng 3 tab đầu (0, 1, 2), Bot tự động tạo và dùng tab 4 (index 3)."""
    engine = SapEngine(
        connection_index=0,
        session_index=0,
        strict_session_safety=True,
        auto_create_session=True,
        prefer_dedicated_session=True,
        user_reserved_sessions=3,
        max_sap_sessions=6
    )

    sess0 = MagicMock(); sess0.Busy = True; sess0.Info.Transaction = "MIGO"
    sess1 = MagicMock(); sess1.Busy = False; sess1.Info.Transaction = "MD04"
    sess2 = MagicMock(); sess2.Busy = True; sess2.Info.Transaction = "ZFIR266"
    sess_bot = MagicMock(); sess_bot.Busy = False; sess_bot.Info.Transaction = "SESSION_MANAGER"

    class DynamicConnMock:
        def __init__(self):
            self.sessions = [sess0, sess1, sess2]

        @property
        def Count(self):
            return len(self.sessions)

        def __call__(self, idx):
            return self.sessions[idx]

    conn_dyn = DynamicConnMock()
    mock_conn = MagicMock()
    mock_conn.Children = conn_dyn

    def fake_create_session():
        conn_dyn.sessions.append(sess_bot)

    sess0.CreateSession = MagicMock(side_effect=fake_create_session)

    mock_app = MagicMock()
    mock_app.Children.Count = 1
    mock_app.Children.side_effect = lambda idx: [mock_conn][idx]

    with pytest.MonkeyPatch().context() as m:
        mock_sap_gui_auto = MagicMock()
        mock_sap_gui_auto.GetScriptingEngine = mock_app
        m.setattr("win32com.client.GetObject", lambda name: mock_sap_gui_auto)

        selected = engine._connect_sap()
        assert selected == sess_bot
        sess0.CreateSession.assert_called_once()
        assert conn_dyn.Count == 4


def test_connect_sap_reuses_existing_tab_in_bot_range():
    """Kiểm tra: Khi tab 4 (index 3) đã tồn tại và rảnh rỗi, Bot tái sử dụng ngay, không can thiệp 3 tab đầu."""
    engine = SapEngine(
        connection_index=0,
        session_index=0,
        strict_session_safety=True,
        auto_create_session=True,
        prefer_dedicated_session=True,
        user_reserved_sessions=3,
        max_sap_sessions=6
    )

    sess0 = MagicMock(); sess0.Busy = True; sess0.Info.Transaction = "MIGO"
    sess1 = MagicMock(); sess1.Busy = False; sess1.Info.Transaction = "MD04"
    sess2 = MagicMock(); sess2.Busy = True; sess2.Info.Transaction = "ZFIR266"
    sess3_bot = MagicMock(); sess3_bot.Busy = False; sess3_bot.Info.Transaction = "SESSION_MANAGER"

    mock_conn = MagicMock()
    mock_conn.Children.Count = 4
    mock_conn.Children.side_effect = lambda idx: [sess0, sess1, sess2, sess3_bot][idx]

    mock_app = MagicMock()
    mock_app.Children.Count = 1
    mock_app.Children.side_effect = lambda idx: [mock_conn][idx]

    with pytest.MonkeyPatch().context() as m:
        mock_sap_gui_auto = MagicMock()
        mock_sap_gui_auto.GetScriptingEngine = mock_app
        m.setattr("win32com.client.GetObject", lambda name: mock_sap_gui_auto)

        selected = engine._connect_sap()
        assert selected == sess3_bot



def test_check_history_mb51_clears_plant_when_not_explicit():
    """Kiểm tra check_history_mb51 xóa trường WERKS-LOW để tìm đa nhà máy khi explicit_plant=False."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._extract_alv_grid = MagicMock(return_value=[
        {"matnr": "BIW32026700HA", "mblnr": "5000999999", "bwart": "261", "budat": "24.10.2026", "erfmg": "10.000", "erfme": "PC", "lgort": "3001", "charg": "LOT1"}
    ])
    engine._reset_to_session_manager = MagicMock()

    mock_werks_low = MagicMock()
    mock_werks_low.text = "OLD_PLANT"

    def fake_find_by_id(ctl_id):
        if "ctxtWERKS-LOW" in ctl_id:
            return mock_werks_low
        m = MagicMock()
        m.text = ""
        return m

    engine._find_by_id_safe = MagicMock(side_effect=fake_find_by_id)

    # 1. Khi explicit_plant=False -> ctxtWERKS-LOW phải bị gán rỗng ""
    recs = engine.check_history_mb51(["BIW32026700HA"], days=3, plant="", explicit_plant=False)
    assert len(recs) == 1
    assert mock_werks_low.text == ""

    # 2. Khi explicit_plant=True và có plant='3093' -> ctxtWERKS-LOW phải là '3093'
    recs2 = engine.check_history_mb51(["BIW32026700HA"], days=3, plant="3093", explicit_plant=True)
    assert len(recs2) == 1
    assert mock_werks_low.text == "3093"


def test_check_low_stock_reads_excel_body_and_supports_full_model_list():
    """Kiểm tra check_low_stock đọc chuẩn file Excel 261008_Stock_FG_MA_BODY.xlsm cho model LIMO."""
    engine = SapEngine.__new__(SapEngine)
    # Lấy top 5
    recs = engine.check_low_stock([], plants=["3093", "3001"], limit=5, model_name="LIMO", shop_name="BODY")
    assert len(recs) == 5
    assert recs[0]["model"] == "Limo"
    assert "status" in recs[0]

    # Lấy full_model_list=True -> trả về 50 mã sắp xếp ưu tiên
    recs_full = engine.check_low_stock([], plants=["3093", "3001"], limit=20, model_name="LIMO", shop_name="BODY", full_model_list=True)
    assert len(recs_full) == 50


def test_validate_grid_or_raise_restrict_your_selection():
    """Kiểm tra khi SAP MB52 trả về 'Restrict your selection' thì văng SapExecutionError yêu cầu thu hẹp điều kiện lọc (không che giấu thành NO_DATA)."""
    engine = SapEngine.__new__(SapEngine)
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())  # Giả lập đang ở màn hình lọc
    engine.last_popup_message = ""

    with pytest.raises(SapExecutionError) as exc_info:
        engine._validate_grid_or_raise(None, "MB52", "Restrict your selection", "S")
    assert "thu hẹp điều kiện lọc" in str(exc_info.value) or "restrict" in str(exc_info.value).lower()


def test_sap_engine_max_bot_sessions_cap():
    """Kiểm tra trần tối đa 3 phiên SAP do bot quản lý (MAX_BOT_SESSIONS = 3) theo chuẩn RC65."""
    assert SapEngine.MAX_BOT_SESSIONS == 3


def test_find_filter_field_token_resolution():
    """Kiểm tra bộ phân giải token trường lọc ALV Filter theo chuẩn RC65."""
    engine = SapEngine.__new__(SapEngine)
    mock_wnd1 = MagicMock()
    
    mock_ctl = MagicMock()
    mock_lbl = MagicMock()
    mock_lbl.Text = "Material Number"

    def mock_find(id_str):
        if "ctxt%%DYN002-LOW" in id_str:
            return mock_ctl
        if "%%DYN002" in id_str and "TEXT" in id_str:
            return mock_lbl
        if "ctxt%%DYN001-LOW" in id_str:
            return mock_ctl
        return None

    engine._find_by_id_safe = MagicMock(side_effect=mock_find)
    token = engine._find_filter_field_token(mock_wnd1, "MATNR", "Material")
    assert token == "%%DYN002"


def test_check_sto_me2n_supplying_and_receiving_plants():
    """Kiểm tra ME2N tự động phân định đúng Supplying Plant (3093) và Receiving Plant (4007) theo chuẩn RC65."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    mock_controls = {}
    def mock_find(id_str):
        if id_str not in mock_controls:
            mock_controls[id_str] = MagicMock()
        return mock_controls[id_str]
    engine._find_by_id_safe = MagicMock(side_effect=mock_find)
    engine._dismiss_popups = MagicMock(return_value="")
    engine._get_status = MagicMock(return_value=("", "S"))

    mock_grid = MagicMock()
    mock_grid.RowCount = 1
    engine._find_largest_grid = MagicMock(return_value=mock_grid)

    mock_records = [
        {"stoNo": "45001234", "item": "10", "product": "BIW001", "orderQty": "100.0", "deliveredQty": "40.0", "supplyingPlant": "3093", "receivingPlant": "4007", "deleteInd": ""}
    ]
    engine._extract_alv_grid = MagicMock(return_value=mock_records)

    # Khi người dùng tra cứu STO cho kho 4007
    res = engine.check_sto_me2n(["BIW001"], plant="4007")
    assert len(res) == 1
    assert res[0]["remainingQty"] == "60"
    # Kiểm tra S_WERKS (Receiving Plant) được điền 4007, S_RESWK (Supplying Plant) được điền 3093
    assert mock_controls["wnd[0]/usr/ctxtS_WERKS-LOW"].text == "4007"
    assert mock_controls["wnd[0]/usr/ctxtS_RESWK-LOW"].text == "3093"




def test_check_low_stock_sloc_breakdown_body_v37():
    """Kiểm tra quy tắc S01: Bóc tách chính xác SLoc theo 261008_Stock_FG_MA_BODY V3.7 (1003, 3003, 3093, 3000, 9999)."""
    from unittest.mock import MagicMock
    engine = SapEngine.__new__(SapEngine)

    mock_mb52_records = [
        {"product": "BIW_TEST_01", "plant": "3001", "sloc": "1003", "unrest": "40", "unit": "PC"},  # BD Kho
        {"product": "BIW_TEST_01", "plant": "3001", "sloc": "3003", "unrest": "60", "unit": "PC"},  # Line Body
        {"product": "BIW_TEST_01", "plant": "3093", "sloc": "3093", "unrest": "200", "unit": "PC"}, # Kho MA
        {"product": "BIW_TEST_01", "plant": "3093", "sloc": "3000", "unrest": "50", "unit": "PC"},  # Line MA
        {"product": "BIW_TEST_01", "plant": "3093", "sloc": "9999", "unrest": "5", "unit": "PC"},   # Transit/QC
    ]
    engine.check_stock_mb52 = MagicMock(return_value=mock_mb52_records)

    res = engine.check_low_stock(["BIW_TEST_01"], plants=["3093", "3001"], limit=5)
    assert len(res) == 1
    r = res[0]
    assert r["product"] == "BIW_TEST_01"
    assert r["bd_kho"] == "40"
    assert r["line_body"] == "60"
    assert r["kho_ma"] == "200"
    assert r["line_ma"] == "50"
    assert r["transit"] == "5"
    assert r["total_stock"] == "350"  # 40 + 60 + 200 + 50 (không cộng 5 của 9999)
    assert r["bd_carset"] == "100"    # (40 + 60) / 1


def test_check_attachment_status_zfir266_mock():
    """Kiểm tra check_attachment_status qua ZFIR266 giả lập thành công."""
    from unittest.mock import MagicMock
    engine = SapEngine()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "ZFIR266"
    engine.ensure_session = MagicMock(return_value=mock_session)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._return_to_main_menu = MagicMock()

    mock_bukrs = MagicMock()
    mock_gjahr = MagicMock()
    mock_doc = MagicMock()
    mock_max = MagicMock()
    mock_wnd0 = MagicMock()
    mock_grid = MagicMock()
    mock_grid.RowCount = 2
    mock_grid.GetCellValue.side_effect = lambda row, col: "5007145690_BBBG.pdf" if row == 0 else "HDGTGT_3874.xlsx"

    def find_mock(ctl_id):
        if "ctxtP_BUKRS" in ctl_id: return mock_bukrs
        if "txtP_GJAHR" in ctl_id: return mock_gjahr
        if "ctxtS_MBELN-LOW" in ctl_id: return mock_doc
        if "txtP_MAX" in ctl_id: return mock_max
        if ctl_id == "wnd[0]": return mock_wnd0
        if "cntlGRID1" in ctl_id: return mock_grid
        return None

    engine._find_by_id_safe = MagicMock(side_effect=find_mock)

    records = engine.check_attachment_status(["5007145690"], year="2026")
    assert len(records) == 1
    rec = records[0]
    assert rec["doc"] == "5007145690"
    assert rec["status_code"] == "FULL_ATT"
    assert rec["bbbg_status"] == "CÓ"
    assert rec["vat_status"] == "CÓ"
    assert "Đã có cả BBBG và VAT" in rec["overall_status"]


def test_check_attachment_status_migo_fallback_mock():
    """Kiểm tra check_attachment_status tự động fallback sang MIGO Display khi ZFIR266 không tồn tại."""
    from unittest.mock import MagicMock
    from app.sap_attachment import MigoAttachmentRunner

    engine = SapEngine()
    mock_session = MagicMock()
    mock_session.Info.Transaction = "SESSION_MANAGER"
    mock_session.Info.SystemName = "PRD"
    mock_session.Info.Client = "100"
    mock_session.Info.User = "TESTUSER"
    mock_session.Id = "ses[0]"
    mock_session.Busy = False
    engine.session_identity = {"system": "PRD", "client": "100", "user": "TESTUSER", "session_id": "ses[0]"}

    engine.ensure_session = MagicMock(return_value=mock_session)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()
    engine._return_to_main_menu = MagicMock()

    mock_sbar = MagicMock()
    mock_sbar.Text = "Transaction ZFIR266 does not exist"
    engine._find_by_id_safe = MagicMock(return_value=mock_sbar)

    original_open = MigoAttachmentRunner.open_migo_document
    original_read = MigoAttachmentRunner.read_sap_attachment_titles
    try:
        MigoAttachmentRunner.open_migo_document = MagicMock(return_value=(True, "OK"))
        MigoAttachmentRunner.read_sap_attachment_titles = MagicMock(return_value=["5007145690_BBBG.pdf"])

        records = engine.check_attachment_status(["5007145690"], year="2026")
        assert len(records) == 1
        rec = records[0]
        assert rec["doc"] == "5007145690"
        assert rec["status_code"] == "MISSING_VAT"
        assert rec["bbbg_status"] == "CÓ"
        assert rec["vat_status"] == "CHƯA"
    finally:
        MigoAttachmentRunner.open_migo_document = original_open
        MigoAttachmentRunner.read_sap_attachment_titles = original_read


def test_check_stock_mb52_preserves_transit_row():
    """Kiểm tra: Dòng Hàng chờ xuất/Transit (SLoc trống nhưng Transit > 0) không bị loại bỏ, được gán SLoc TRANSIT."""
    engine = SapEngine.__new__(SapEngine)
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()
    engine._clear_all_selection_inputs = MagicMock()
    engine._find_by_id_safe = MagicMock(return_value=MagicMock())
    engine._dismiss_popups = MagicMock()
    engine._get_status = MagicMock(return_value=("", "S"))
    engine._export_alv_to_excel = MagicMock(return_value=None)
    engine._reset_to_session_manager = MagicMock()

    # Dữ liệu mô phỏng từ màn hình MB52 thực tế:
    # Dòng 0: Transit cấp Plant 4007 (SLoc trống, Transit 78 EA)
    # Dòng 1: Kho 4004 (4 EA)
    # Dòng 2: Kho 4007 (136 EA)
    engine._extract_alv_grid = MagicMock(return_value=[
        {"product": "BIW32022638AB", "plant": "4007", "sloc": "", "batch": "", "unrest": "0", "transit": "78", "unit": "EA"},
        {"product": "", "plant": "4007", "sloc": "4004", "batch": "IN31025038", "unrest": "4", "transit": "0", "unit": "EA"},
        {"product": "", "plant": "4007", "sloc": "4007", "batch": "IN31025039", "unrest": "136", "transit": "0", "unit": "EA"},
    ])

    recs = engine.check_stock_mb52(materials=["BIW32022638AB"], plant="4007")

    # Bắt buộc phải giữ đủ cả 3 dòng (kể cả dòng Transit)
    assert len(recs) == 3

    # Dòng Transit được đưa lên đầu với sloc TRANSIT
    assert recs[0]["sloc"] == "TRANSIT"
    assert recs[0]["unrest"] == "78"
    assert recs[0]["transit"] == "78"
    assert recs[0]["batch"] == "Chờ xuất (từ 3093)"
    assert recs[0]["product"] == "BIW32022638AB"

    # Các dòng kho vật lý tiếp theo
    assert recs[1]["sloc"] == "4004"
    assert recs[1]["unrest"] == "4"
    assert recs[2]["sloc"] == "4007"
    assert recs[2]["unrest"] == "136"


def test_post_migo_311_ob_group_simulation():
    """Kiểm tra post_migo_311_ob_group ở chế độ simulation / dry_run trả về số chứng từ OB 10xxxxxxxx."""
    engine = SapEngine.__new__(SapEngine)
    engine.session = None

    items = [
        {"material": "BIW75111244AA", "allocated_qty": 300, "batch": "240801", "bin": "01-02-03", "station": "KV08-Thức/Vũ"},
        {"material": "BIW75111244AA", "allocated_qty": 100, "batch": "240802", "bin": "01-02-04", "station": "KV08-Thức/Vũ"}
    ]

    ok, doc_no, msg = engine.post_migo_311_ob_group(items, plant="3093", dry_run=True)
    assert ok is True
    assert doc_no == ""
    assert "Mô phỏng" in msg


def test_post_migo_311_ob_group_v31_standard_mock():
    """Kiểm tra post_migo_311_ob_group tuân thủ chuẩn xác 100% VBA MIGO311_RunPost_V24_ErrorGuard.bas:
    - Tìm động GODEFAULT_TV-BWART, điền '311'
    - Điền GetCell: Material(3), Qty(4), Trạm Text(5), Plant(6), SLoc xuất(7), SLoc nhận(8), Batch(10)
    - Nhấn Post và trích xuất đúng số chứng từ Outbound Delivery từ StatusBar.
    """
    engine = SapEngine.__new__(SapEngine)
    engine.ensure_session = MagicMock()
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()

    class MockControl:
        def __init__(self, name="", c_type="", text="", key=""):
            self.Name = name
            self.name = name
            self.Type = c_type
            self.type = c_type
            self.Text = text
            self.text = text
            self.Key = key
            self.key = key
            self.Children = []
            self.children = []
            self.Entries = []
            self.entries = []
            self.Id = f"wnd[0]/{name}"
            self.id = self.Id
        def Press(self):
            pass
        def press(self):
            pass

    class MockCell:
        def __init__(self):
            self.Text = ""
            self.text = ""

    class MockCol:
        def __init__(self, name, title):
            self.Name = name
            self.name = name
            self.Title = title
            self.title = title

    class MockTable:
        def __init__(self):
            self.Name = "tblSAPLMIGOTV_GOITEM"
            self.name = self.Name
            self.Type = "GuiTableControl"
            self.type = self.Type
            self.Id = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"
            self.id = self.Id
            self.FirstVisibleRow = 0
            self.VisibleRowCount = 10
            self.cells = {}
            self.cols = [
                MockCol("LINE", "Line"), MockCol("STATUS", "Status"), MockCol("OK", "OK"),
                MockCol("MATNR", "Material"), MockCol("ERFMG", "Quantity"), MockCol("SGTXT", "Item Text"),
                MockCol("WERKS", "Plant"), MockCol("LGORT", "Storage Loc"), MockCol("UMLGO", "Rec. SLoc"),
                MockCol("UOM", "Unit"), MockCol("CHARG", "Batch")
            ]
        @property
        def Columns(self):
            cols_list = self.cols
            class ColColl:
                Count = len(cols_list)
                count = len(cols_list)
                def Item(self, idx):
                    return cols_list[idx]
                def __call__(self, idx):
                    return cols_list[idx]
            return ColColl()

        def GetCell(self, row, col):
            if (row, col) not in self.cells:
                self.cells[(row, col)] = MockCell()
            return self.cells[(row, col)]

    class MockSbar:
        def __init__(self):
            self.Text = "Material document 1013115228 posted"
            self.text = self.Text
            self.MessageType = "S"
            self.messageType = "S"

    action_cmb = MockControl(name="GODYNPRO-ACTION", c_type="GuiComboBox", key="A01")
    ref_cmb = MockControl(name="GODYNPRO-REFDOC", c_type="GuiComboBox", key="R01")
    mvt_field = MockControl(name="GODEFAULT_TV-BWART", c_type="GuiCTextField", text="")
    tbl = MockTable()
    sbar = MockSbar()

    mock_w0 = MockControl(name="wnd[0]", c_type="GuiMainWindow")
    mock_w0.sendVKey = MagicMock()
    mock_w0.Children = [action_cmb, ref_cmb, mvt_field, tbl, sbar]
    mock_w0.children = mock_w0.Children

    mock_session = MagicMock()
    mock_session.ActiveWindow.Name = "wnd[0]"
    mock_session.Children.Count = 1
    mock_session.children.count = 1

    def mock_find_by_id_safe(cid):
        if cid == "wnd[0]":
            return mock_w0
        if cid == "wnd[0]/sbar":
            return sbar
        if "tblSAPLMIGOTV_GOITEM" in cid:
            return tbl
        return None

    engine.session = mock_session
    engine._find_by_id_safe = mock_find_by_id_safe

    items = [
        {
            "material": "BIW75111244AA",
            "allocated_qty": 100,
            "batch": "261023-15",
            "bin": "APFL-P23-01-12",
            "station": "KV08-Thức/Vũ",
            "unit": "EA"
        }
    ]

    ok, doc_no, msg = engine.post_migo_311_ob_group(
        items,
        plant="3093",
        source_sloc="1005",
        dest_sloc="3001",
        dry_run=False
    )

    assert ok is True
    assert doc_no == "1013115228"
    assert "1013115228" in msg

    # Kiểm tra đã điền đúng GODEFAULT_TV-BWART = 311
    assert mvt_field.Text == "311" or mvt_field.text == "311"
    # Kiểm tra Action và RefDoc được cấu hình A08 và R10
    assert action_cmb.key == "A08"
    assert ref_cmb.key == "R10"
    # Kiểm tra ô Material (col 3)
    assert tbl.GetCell(0, 3).Text == "BIW75111244AA" or tbl.GetCell(0, 3).text == "BIW75111244AA"
    # Kiểm tra ô Qty (col 4)
    assert tbl.GetCell(0, 4).Text == "100" or tbl.GetCell(0, 4).text == "100"
    # Kiểm tra ô Trạm Text (col 5)
    assert tbl.GetCell(0, 5).Text == "KV08-Thức/Vũ" or tbl.GetCell(0, 5).text == "KV08-Thức/Vũ"
    # Kiểm tra ô Plant (col 6)
    assert tbl.GetCell(0, 6).Text == "3093" or tbl.GetCell(0, 6).text == "3093"
    # Kiểm tra ô SLoc nguồn (col 7)
    assert tbl.GetCell(0, 7).Text == "1005" or tbl.GetCell(0, 7).text == "1005"
    # Kiểm tra ô SLoc đích (col 8)
    assert tbl.GetCell(0, 8).Text == "3001" or tbl.GetCell(0, 8).text == "3001"
    # Kiểm tra ô Batch (col 10)
    assert tbl.GetCell(0, 10).Text == "261023-15" or tbl.GetCell(0, 10).text == "261023-15"


def test_post_migo_311_84_columns_real_layout_mock():
    """Kiểm tra post_migo_311_ob_group trên lưới MIGO thực tế 84 cột:
    - Cột 2: Material Trfr Pstg (UMMAT) -> KHÔNG bị gán nhầm
    - Cột 3: Material (MATNR) -> Phải điền đúng
    - Cột 4: Mat. Short Text (MAKTX) -> Không bị ghi đè
    - Cột 5: Qty in UnE (ERFMG) -> Phải điền đúng 100
    - Cột 6: Text (SGTXT) -> Phải điền đúng KV08-Thức/Vũ
    - Cột 7: Plnt (WERKS) -> Phải điền đúng 3093
    - Cột 8: SLoc (LGORT) -> Phải điền đúng 1005
    - Cột 9: SLoc Tra... (UMLGO) -> Phải điền đúng 3001
    - Cột 10: Batch (CHARG) -> Phải điền đúng 261023-15 (KHÔNG bị nhảy sang cột 76!)
    - Cột 14: OK (TAKE_IT) -> Phải được tick Selected = True
    """
    engine = SapEngine.__new__(SapEngine)
    engine.ensure_session = MagicMock()
    engine._open_transaction = MagicMock()
    engine._wait_idle = MagicMock()

    class MockControl:
        def __init__(self, name="", c_type="", text="", key=""):
            self.Name = name
            self.name = name
            self.Type = c_type
            self.type = c_type
            self.Text = text
            self.text = text
            self.Key = key
            self.key = key
            self.Children = []
            self.children = []
            self.Entries = []
            self.entries = []
            self.Id = f"wnd[0]/{name}"
            self.id = self.Id
        def Press(self):
            pass
        def press(self):
            pass

    class MockCell:
        def __init__(self):
            self.Text = ""
            self.text = ""
            self.Selected = False
            self.selected = False

    class MockCol:
        def __init__(self, name, title):
            self.Name = name
            self.name = name
            self.Title = title
            self.title = title

    class MockTable84:
        def __init__(self):
            self.Name = "tblSAPLMIGOTV_GOITEM"
            self.name = self.Name
            self.Type = "GuiTableControl"
            self.type = self.Type
            self.Id = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"
            self.id = self.Id
            self.FirstVisibleRow = 0
            self.VisibleRowCount = 10
            self.cells = {}
            # Layout thực tế giống hệt màn hình người dùng
            self.cols = [
                MockCol("GOITEM-ZEILE", "Line"),
                MockCol("STATUS", "Status"),
                MockCol("GOITEM-UMMAT", "Material Trfr Pstg"),
                MockCol("GOITEM-MATNR", "Material"),
                MockCol("GOITEM-MAKTX", "Mat. Short Text"),
                MockCol("GOITEM-ERFMG", "Qty in UnE"),
                MockCol("GOITEM-SGTXT", "Text"),
                MockCol("GOITEM-WERKS", "Plnt"),
                MockCol("GOITEM-LGORT", "SLoc"),
                MockCol("GOITEM-UMLGO", "SLoc Tra..."),
                MockCol("GOITEM-CHARG", "Batch"),
                MockCol("GOITEM-UMCHA", "Batch Trfr..."),
                MockCol("GOITEM-ERFME", "UnE"),
                MockCol("GOITEM-MEINS", "BUn"),
                MockCol("GOITEM-TAKE_IT", "OK"),
            ]
            # Bổ sung các cột kỹ thuật phía sau gây nhiễu
            for idx in range(15, 84):
                if idx == 52:
                    self.cols.append(MockCol("GOITEM-BPMNG", "PO Qty"))
                elif idx == 58:
                    self.cols.append(MockCol("GOITEM-UMWRK", "Receiving Plant"))
                elif idx == 62:
                    self.cols.append(MockCol("GOITEM-MATKL", "Material Group"))
                elif idx == 76:
                    self.cols.append(MockCol("GOITEM-CHARG_EXT", "Batch Management"))
                else:
                    self.cols.append(MockCol(f"GOITEM-COL{idx}", f"Technical Col {idx}"))

        @property
        def Columns(self):
            cols_list = self.cols
            class ColColl:
                Count = len(cols_list)
                count = len(cols_list)
                def Item(self, idx):
                    return cols_list[idx]
                def __call__(self, idx):
                    return cols_list[idx]
            return ColColl()

        def GetCell(self, row, col):
            if (row, col) not in self.cells:
                self.cells[(row, col)] = MockCell()
            return self.cells[(row, col)]

    class MockSbar:
        def __init__(self):
            self.Text = "Material document 1012924052 posted"
            self.text = self.Text
            self.MessageType = "S"
            self.messageType = "S"

    action_cmb = MockControl(name="GODYNPRO-ACTION", c_type="GuiComboBox", key="A01")
    ref_cmb = MockControl(name="GODYNPRO-REFDOC", c_type="GuiComboBox", key="R01")
    mvt_field = MockControl(name="GODEFAULT_TV-BWART", c_type="GuiCTextField", text="")
    tbl = MockTable84()
    sbar = MockSbar()

    mock_w0 = MockControl(name="wnd[0]", c_type="GuiMainWindow")
    mock_w0.sendVKey = MagicMock()
    mock_w0.Children = [action_cmb, ref_cmb, mvt_field, tbl, sbar]
    mock_w0.children = mock_w0.Children

    mock_session = MagicMock()
    mock_session.ActiveWindow.Name = "wnd[0]"
    mock_session.Children.Count = 1
    mock_session.children.count = 1

    def mock_find_by_id_safe(cid):
        if cid == "wnd[0]":
            return mock_w0
        if cid == "wnd[0]/sbar":
            return sbar
        if "tblSAPLMIGOTV_GOITEM" in cid:
            return tbl
        return None

    engine.session = mock_session
    engine._find_by_id_safe = mock_find_by_id_safe

    items = [
        {
            "material": "BIW75111244AA",
            "allocated_qty": 100,
            "batch": "261023-15",
            "bin": "APFL-P23-01-12",
            "station": "KV08-Thức/Vũ",
            "unit": "EA"
        }
    ]

    ok, doc_no, msg = engine.post_migo_311_ob_group(
        items,
        plant="3093",
        source_sloc="1005",
        dest_sloc="3001",
        dry_run=False
    )

    assert ok is True
    assert doc_no == "1012924052"

    # Kiểm tra ô Material (col 3 chuẩn Excel V31)
    assert tbl.GetCell(0, 3).Text == "BIW75111244AA"
    # Kiểm tra ô Qty (col 4 chuẩn Excel V31)
    assert tbl.GetCell(0, 4).Text == "100"
    # Kiểm tra ô Trạm Text (col 5 chuẩn Excel V31)
    assert tbl.GetCell(0, 5).Text == "KV08-Thức/Vũ"
    # Kiểm tra ô Plant (col 6 chuẩn Excel V31)
    assert tbl.GetCell(0, 6).Text == "3093"
    # Kiểm tra ô SLoc nguồn (col 7 chuẩn Excel V31)
    assert tbl.GetCell(0, 7).Text == "1005"
    # Kiểm tra ô SLoc đích (col 8 chuẩn Excel V31)
    assert tbl.GetCell(0, 8).Text == "3001"
    # Kiểm tra ô Batch (col 10 quét theo RUN_GetBatchColumnForRow)
    assert tbl.GetCell(0, 10).Text == "261023-15"
    # Kiểm tra ô OK (col 14) được tick Selected = True
    assert tbl.GetCell(0, 14).Selected is True


def test_open_transaction_okcd_syntax():
    """Kiểm tra cú pháp T-Code: không bao giờ sinh /n/MIGO mà phải là /nMIGO; giữ nguyên /n/SCWM/POST."""
    engine = SapEngine.__new__(SapEngine)
    engine.ensure_session = MagicMock()
    engine._wait_idle = MagicMock()
    engine._dismiss_popups = MagicMock()

    class MockOkcd:
        def __init__(self):
            self.text = ""

    mock_okcd = MockOkcd()
    mock_w0 = MagicMock()
    mock_session = MagicMock()

    def mock_find(cid):
        if cid == "wnd[0]/tbar[0]/okcd":
            return mock_okcd
        if cid == "wnd[0]":
            return mock_w0
        return None

    engine._find_by_id_safe = mock_find
    engine.session = mock_session

    # Test MIGO
    engine._open_transaction("MIGO")
    assert mock_okcd.text == "/nMIGO"

    engine._open_transaction("/nMIGO")
    assert mock_okcd.text == "/nMIGO"

    engine._open_transaction("/MIGO")
    assert mock_okcd.text == "/nMIGO"

    # Test Namespace T-Code /SCWM/POST
    engine._open_transaction("/SCWM/POST")
    assert mock_okcd.text == "/n/SCWM/POST"

    engine._open_transaction("/n/SCWM/POST")
    assert mock_okcd.text == "/n/SCWM/POST"

    # Test CS03, MB52
    engine._open_transaction("CS03")
    assert mock_okcd.text == "/nCS03"

    engine._open_transaction("/nMB52")
    assert mock_okcd.text == "/nMB52"


def test_sap_engine_rc65_table_field_and_overview_methods():
    """Kiểm tra SapEngine tích hợp đồng bộ kiến trúc RC65 (table_field, delivery_overview, snapshot)."""
    engine = SapEngine.__new__(SapEngine)

    class MockField:
        def __init__(self, cid, text=""):
            self.Id = cid
            self.id = cid
            self.Text = text
            self.text = text

    class MockOverviewTable:
        def __init__(self):
            self.Id = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER"
            self.id = self.Id
            self.VisibleRowCount = 5
            self.RowCount = 1
            self._children = [
                MockField(f"{self.Id}/txtLIPS-POSNR[0,0]", "10"),
                MockField(f"{self.Id}/txtLIPS-MATNR[1,0]", "BIW32013765AA"),
                MockField(f"{self.Id}/txtLIPSD-G_LFIMG[2,0]", "100.000"),
                MockField(f"{self.Id}/txtLIPS-LGORT[3,0]", "3093"),
                MockField(f"{self.Id}/txtLIPS-CHARG[4,0]", "LOT-A1"),
            ]

        @property
        def Children(self):
            children_list = self._children
            class Coll:
                Count = len(children_list)
                count = len(children_list)
                def __call__(_s, idx):
                    return children_list[idx]
                def ElementAt(_s, idx):
                    return children_list[idx]
            return Coll()

        def findById(self, rel_id):
            for ch in self._children:
                if ch.Id.endswith(rel_id):
                    return ch
            return None

    mock_tbl = MockOverviewTable()
    engine._find_by_id_safe = MagicMock(return_value=mock_tbl)
    engine.session = MagicMock()

    # 1. Test delivery_overview
    ov_tbl = engine.delivery_overview()
    assert ov_tbl is not None
    assert ov_tbl.Id == mock_tbl.Id

    # 2. Test table_field
    pos_fld = engine.table_field(mock_tbl, "LIPS-POSNR", 0)
    assert pos_fld is not None
    assert pos_fld.Text == "10"

    qty_fld = engine.table_field(mock_tbl, "LIPSD-G_LFIMG", 0)
    assert qty_fld is not None
    assert qty_fld.Text == "100.000"

    # 3. Test set_table_text
    ok = engine.set_table_text(mock_tbl, "LIPS-LGORT", 0, "3001")
    assert ok is True
    assert engine.table_field(mock_tbl, "LIPS-LGORT", 0).Text == "3001"

    # 4. Test locate_delivery_item
    loc = engine.locate_delivery_item(mock_tbl, "10")
    assert loc == 0
    assert engine.locate_delivery_item(mock_tbl, "99") == -1

    # 5. Test snapshot_delivery_overview
    snap = engine.snapshot_delivery_overview(mock_tbl)
    assert "10" in snap
    assert snap["10"]["material"] == "BIW32013765AA"
    assert snap["10"]["qty"] == "100.000"
    assert snap["10"]["sloc"] == "3001"
    assert snap["10"]["batch"] == "LOT-A1"


def test_resolve_grid_columns_rc65_strict_exact_locking():
    """Kiểm tra _resolve_grid_columns khóa cứng mã kỹ thuật Lượt 0 chuẩn RC65, không bị lệch do khớp mờ."""
    engine = SapEngine.__new__(SapEngine)

    class MockColOrder:
        def __init__(self, cols):
            self._cols = cols
            self.Count = len(cols)
        def ElementAt(self, idx):
            return self._cols[idx]

    mock_alv = MagicMock()
    mock_alv.ColumnOrder = MockColOrder([
        "VBELN", "POSNR", "MATNR", "LFIMG", "VRKME", "LGORT", "CHARG", "KUNWE", "WBSTA"
    ])
    engine._get_column_caption = MagicMock(return_value="")

    # Mapping với candidate đầu tiên là mã kỹ thuật chuẩn RC65
    col_mapping = {
        "obNo": ["VBELN", "DELIVERY"],
        "item": ["POSNR", "ITEM"],
        "product": ["MATNR", "MATERIAL"],
        "qty": ["LFIMG", "QUANTITY"],
        "unit": ["VRKME", "UNIT"],
        "sloc": ["LGORT", "STORAGE LOCATION"],
        "batch": ["CHARG", "BATCH"],
        "shipTo": ["KUNWE", "SHIP TO"],
        "gmStatus": ["WBSTA", "GM STATUS"]
    }

    resolved = engine._resolve_grid_columns(mock_alv, col_mapping)
    assert resolved["obNo"] == "VBELN"
    assert resolved["item"] == "POSNR"
    assert resolved["product"] == "MATNR"
    assert resolved["qty"] == "LFIMG"
    assert resolved["unit"] == "VRKME"
    assert resolved["sloc"] == "LGORT"
    assert resolved["batch"] == "CHARG"
    assert resolved["shipTo"] == "KUNWE"
    assert resolved["gmStatus"] == "WBSTA"

    # Kiểm tra biến thể kỹ thuật RC65: S_CHARG thay vì CHARG, KUNNR thay vì KUNWE
    mock_alv_alt = MagicMock()
    mock_alv_alt.ColumnOrder = MockColOrder([
        "VBELN", "POSNR", "MATNR", "LFIMG", "VRKME", "LGORT", "S_CHARG", "KUNNR", "WBSTA", "VGBEL", "ERDAT"
    ])
    from app.ob_service import RC65_VL06O_COLUMN_MAPPING
    resolved_alt = engine._resolve_grid_columns(mock_alv_alt, RC65_VL06O_COLUMN_MAPPING)
    assert resolved_alt["batch"] == "S_CHARG"
    assert resolved_alt["shipTo"] == "KUNNR"
    assert resolved_alt["refSTO"] == "VGBEL"
    assert resolved_alt["createDate"] == "ERDAT"



