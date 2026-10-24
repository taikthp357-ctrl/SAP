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
    # Kiểm tra Multiple Selection cho Shipping Point (3093 & 3001)
    engine._fill_multiple_selection.assert_any_call("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH", ["3093", "3001"])


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






