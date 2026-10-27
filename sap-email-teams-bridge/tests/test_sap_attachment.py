# -*- coding: utf-8 -*-
"""Unit tests cho module MigoAttachmentRunner tự động đính kèm BBBG, VAT, Email & điền số phiếu vào HeadText."""

import tempfile
import shutil
import sys
from pathlib import Path
from unittest.mock import MagicMock
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.sap_attachment import (
    MigoAttachmentRunner,
    normalize_text,
    extract_search_tokens,
    parse_attachment_items_from_text,
    clean_folder_path,
    format_status_label,
    BBBG_REGEX,
    VAT_REGEX,
    MAIL_REGEX
)


@pytest.fixture
def sample_attachment_dir():
    """Tạo thư mục tạm chứa các file mẫu BBBG, VAT và Email."""
    temp_dir = tempfile.mkdtemp(prefix="test_migo_att_")
    base = Path(temp_dir)

    # Tạo các file mẫu
    (base / "in 5794_BBBG.pdf").write_text("sample content bbbg 5794")
    (base / "in 5794_VAT.pdf").write_text("sample content vat 5794")
    (base / "in 5794_Order_Mail.msg").write_text("sample outlook mail file")
    
    sub = base / "subfolder"
    sub.mkdir()
    (sub / "Thanhcong in 5799_BBGH.pdf").write_text("sample content 5799")
    (sub / "Thanhcong in 5799_HoaDon.pdf").write_text("sample content vat 5799")

    (base / "5007072362_VAT.pdf").write_text("sample vat 6157")

    yield base

    shutil.rmtree(temp_dir, ignore_errors=True)


def owned_runner(session, engine=None):
    session.Info.SystemName = "TEST"
    session.Info.Client = "100"
    session.Info.User = "TESTUSER"
    session.Info.Transaction = "MIGO"
    session.Id = "/app/con[0]/ses[1]"
    session.Busy = False
    runner = MigoAttachmentRunner(engine, {"system": "TEST", "client": "100", "user": "TESTUSER", "session_id": session.Id})
    runner._active_document = ("5000000001", "2026")
    return runner


def test_normalize_and_token_extraction():
    """Kiểm tra bóc tách token và chuẩn hóa chuỗi."""
    assert normalize_text("Thành Đạt IN- 6157") == "thanh dat in- 6157"
    nums, compact = extract_search_tokens("Thành Đạt IN- 6157")
    assert "6157" in nums
    assert "thanhdatin6157" == compact

    nums_doc, _ = extract_search_tokens("5006995954")
    assert "5006995954" in nums_doc


def test_parse_attachment_items_from_text():
    """Kiểm tra bóc tách danh sách chứng từ và số phiếu từ nội dung email / tin nhắn Teams."""
    sample_text = """
    Thư mục nguồn: \\\\10.128.10.72\\fvn$\\03.Khối sản xuất\\260802_ File ATT BBGH
    5006995954: Thành Công in 5794
    5006995985: Thành Công in 5794
    5006995500 \t Thanhcong in 5799
    5006995502 - Thành Công in - 4929
    5007072362 Thành Đạt in 6157
    """
    items = parse_attachment_items_from_text(sample_text)
    assert len(items) == 5
    assert items[0]["doc"] == "5006995954"
    assert "in 5794" in items[0]["ticket_no"]
    assert items[2]["doc"] == "5006995500"
    assert "5799" in items[2]["ticket_no"]
    assert items[4]["doc"] == "5007072362"
    assert "6157" in items[4]["ticket_no"]


def test_index_source_folder(sample_attachment_dir):
    """Kiểm tra lập chỉ mục file trong thư mục nguồn và thư mục con (bao gồm cả .msg)."""
    runner = MigoAttachmentRunner()
    index = runner.index_source_folder(str(sample_attachment_dir), recursive=True)

    assert index["total_files"] == 6  # 5 file pdf + 1 file .msg
    assert "5794" in index["token_map"]
    assert "5799" in index["token_map"]
    assert "5007072362" in index["token_map"]


def test_match_files_exact(sample_attachment_dir):
    """Kiểm tra khớp file chính xác theo số phiếu và loại file (BBBG, VAT, Mail)."""
    runner = MigoAttachmentRunner()
    index = runner.index_source_folder(str(sample_attachment_dir), recursive=True)

    # Test khớp 5794 (Có đủ BBBG, VAT, và Email .msg)
    b_path, v_path, m_path, notes = runner.match_files("5006995954", "Thành Công in 5794", index)
    assert b_path is not None
    assert "5794_BBBG.pdf" in b_path.name
    assert v_path is not None
    assert "5794_VAT.pdf" in v_path.name
    assert m_path is not None
    assert "Order_Mail.msg" in m_path.name

    # Test khớp 5799 trong thư mục con
    b_path2, v_path2, m_path2, _ = runner.match_files("5006995500", "Thanhcong in 5799", index)
    assert b_path2 is not None
    assert "5799_BBGH.pdf" in b_path2.name
    assert v_path2 is not None
    assert "5799_HoaDon.pdf" in v_path2.name
    assert m_path2 is None

    # Test thiếu file BBBG nhưng có file VAT (như case 5007072362)
    b_path3, v_path3, m_path3, notes3 = runner.match_files("5007072362", "Thành Đạt in 6157", index)
    assert b_path3 is None
    assert v_path3 is not None
    assert "5007072362_VAT.pdf" in v_path3.name
    assert "BBBG: FILE_NOT_FOUND" in notes3


def test_process_attachment_batch_dry_run(sample_attachment_dir):
    """Kiểm tra xử lý danh sách hàng loạt ở chế độ dry_run."""
    runner = MigoAttachmentRunner()
    items = [
        {"doc": "5006995954", "ticket_no": "Thành Công in 5794"},
        {"doc": "5006995500", "ticket_no": "Thanhcong in 5799"},
        {"doc": "5007072362", "ticket_no": "Thành Đạt in 6157"},
        {"doc": "9999999999", "ticket_no": "Khong Ton Tai"}
    ]

    summary = runner.process_attachment_batch(str(sample_attachment_dir), items, fill_headtext=True, dry_run=True)
    assert summary["total_items"] == 4
    assert summary["total_success"] == 0
    assert summary["total_preview_ready"] == 3
    assert summary["total_missing"] == 4  # no SAP writes in preview

    report = runner.format_teams_report(summary)
    assert "`5006995954`: SIMULATED" in report
    assert "`9999999999`: FILE_NOT_FOUND" in report
    assert "| Số DOC | Số phiếu | Điền HeadText | BBBG | VAT | Trạng thái |" in report
    assert "Mail (.msg)" not in report
    assert "Chi tiết" not in report


def test_handle_migo_headtext_mock():
    """Kiểm tra logic đọc và điền số phiếu vào Header Text trên SAP MIGO."""
    mock_session = MagicMock()
    mock_field = MagicMock()
    mock_field.Text = ""  # Ban đầu Header Text rỗng
    mock_wnd0 = MagicMock()
    mock_wnd0.findByName.return_value = mock_field
    mock_sbar = MagicMock()
    mock_sbar.MessageType = "S"
    mock_sbar.Text = "Material document changed"

    def find_by_id(path):
        if "sbar" in path:
            return mock_sbar
        elif "wnd[0]" in path:
            return mock_wnd0
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    runner = owned_runner(mock_session)
    runner.open_migo_document = lambda *args, **kwargs: (True, "VERIFIED")
    from unittest.mock import patch
    with patch("app.sap_attachment.smart_find", return_value=mock_field):
        status, msg = runner.handle_migo_headtext(mock_session, "Thành Đạt in 6157")
    assert status == "HEADTEXT_DONE"
    assert "Đã điền và lưu mới HeadText" in msg
    assert mock_field.Text == "Thành Đạt in 6157"
    mock_wnd0.SendVKey.assert_called_with(11)  # Ctrl+S Save


def test_sap_session_gos_mock():
    """Kiểm tra tương tác mock SAP GUI với GOS."""
    mock_session = MagicMock()
    mock_sbar = MagicMock()
    mock_sbar.Text = "The attachment was successfully created"
    mock_sbar.MessageType = "S"

    mock_shell = MagicMock()
    wnd1_active = [False]
    mock_shell.SelectContextMenuItem.side_effect = lambda *a, **k: wnd1_active.__setitem__(0, True)
    mock_btn = MagicMock()
    mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_btn.Press.side_effect = lambda: wnd1_active.__setitem__(0, False)
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.side_effect = lambda name, *args: mock_btn if "btn" in name else MagicMock()

    def find_by_id(path):
        path_str = str(path)
        if "shell" in path_str:
            return mock_shell
        elif "sbar" in path_str:
            return mock_sbar
        elif "btn" in path_str:
            return mock_btn
        elif path_str == "wnd[1]" or path_str.startswith("wnd[1]"):
            if wnd1_active[0]:
                return mock_wnd1
            raise Exception("Control wnd[1] not found")
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    runner = owned_runner(mock_session)
    with tempfile.NamedTemporaryFile(suffix="_VAT.pdf", delete=False) as tf:
        tf.write(b"dummy pdf")
        dummy_file = Path(tf.name)

    try:
        runner.read_sap_attachment_titles = lambda session: [dummy_file.name]
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is True
        # Nhờ Fast-path Sbar Confirmation, khi sbar trả về 'S' thì xác nhận thành công ngay DONE_SAP_SBAR
        assert "DONE_SAP_SBAR" in msg or "DONE_SAP_READBACK" in msg
        mock_shell.PressContextButton.assert_called_with("%GOS_TOOLBOX")
        mock_shell.SelectContextMenuItem.assert_called_with("%GOS_PCATTA_CREA")
    finally:
        dummy_file.unlink(missing_ok=True)


def test_process_attachment_batch_headtext_only():
    """Kiểm tra xử lý điền số phiếu HeadText khi không cung cấp thư mục đính kèm file."""
    runner = MigoAttachmentRunner()
    batch_res = runner.process_attachment_batch(
        folder_path="",
        items=[{"doc": "4995880054", "ticket_no": "OUT_P1_23581"}],
        dry_run=True
    )
    assert batch_res["total_success"] == 0
    assert batch_res["total_preview_ready"] == 1
    assert batch_res["total_missing"] == 1
    res0 = batch_res["results"][0]
    assert res0["doc"] == "4995880054"
    assert res0["ticket_no"] == "OUT_P1_23581"
    assert res0["headtext_status"] == "SIMULATED"
    assert res0["bbbg_status"] == "SKIP_NO_ATTACH"
    assert res0["overall_status"] == "SIMULATED"


def test_parse_attachment_multi_row_list_and_table():
    """Kiểm tra bóc tách nhiều dòng chứng từ và số phiếu từ tin nhắn Teams (bảng / list dài)."""
    text_teams = """
    SAP điền số phiếu

    4995880054 | OUT_P1_23581
    4995877314 | OUT_P1_23582
    4904351086 | OUT_P1_23583
    """
    items = parse_attachment_items_from_text(text_teams)
    assert len(items) == 3
    assert items[0] == {"doc": "4995880054", "ticket_no": "OUT_P1_23581"}
    assert items[1] == {"doc": "4995877314", "ticket_no": "OUT_P1_23582"}
    assert items[2] == {"doc": "4904351086", "ticket_no": "OUT_P1_23583"}

    # Test trường hợp chuỗi bị dồn thành 1 dòng (space separated)
    text_single_line = "4995880054 OUT_P1_23581 4995877314 OUT_P1_23582 4904351086 OUT_P1_23583"
    items_sl = parse_attachment_items_from_text(text_single_line)
    assert len(items_sl) == 3
    assert items_sl[0]["doc"] == "4995880054"
    assert items_sl[0]["ticket_no"] == "OUT_P1_23581"
    assert items_sl[1]["doc"] == "4995877314"
    assert items_sl[1]["ticket_no"] == "OUT_P1_23582"
    assert items_sl[2]["doc"] == "4904351086"
    assert items_sl[2]["ticket_no"] == "OUT_P1_23583"

    # Test trường hợp HTML table từ Teams/Outlook
    text_html = "<table><tr><td>4995880054</td><td>OUT_P1_23581</td></tr><tr><td>4995877314</td><td>OUT_P1_23582</td></tr></table>"
    items_html = parse_attachment_items_from_text(text_html)
    assert len(items_html) == 2
    assert items_html[0] == {"doc": "4995880054", "ticket_no": "OUT_P1_23581"}
    assert items_html[1] == {"doc": "4995877314", "ticket_no": "OUT_P1_23582"}


def test_clean_folder_path():
    """Kiểm tra hàm làm sạch đường dẫn thư mục: xóa nbsp \xa0, zero-width, nháy, và chuẩn hóa Unicode NFC."""
    raw = "  \"\\\\10.128.10.72\\fvn$\\03.Khối\xa0sản\xa0xuất\\260802_\xa0File\xa0ATT\xa0BBGH\\\"  "
    cleaned = clean_folder_path(raw)
    assert "\xa0" not in cleaned
    assert '"' not in cleaned
    assert not cleaned.endswith("\\")
    assert "03.Khối sản xuất" in cleaned
    assert "260802_ File ATT BBGH" in cleaned


def test_format_status_label():
    """Kiểm tra format_status_label trả về DONE khi thành công và FILE_NOT_FOUND khi thất bại."""
    assert format_status_label("Thành Công in 3874", True) == "DONE"
    assert format_status_label("thanh cong 1234", True) == "DONE"
    assert format_status_label("in 3874", True) == "DONE"
    assert format_status_label("", True) == "DONE"
    assert format_status_label("in 3874", False) == "FILE_NOT_FOUND"


def test_unreachable_folder_handling():
    """Kiểm tra khi người dùng nhập thư mục không tồn tại / không truy cập được thì báo lỗi rõ ràng chứ không âm thầm skip."""
    runner = MigoAttachmentRunner()
    fake_folder = r"\\10.999.999.999\not_exist_share\fake_dir"
    res = runner.process_attachment_batch(
        folder_path=fake_folder,
        items=[{"doc": "5006458947", "ticket_no": "Thành Công in 3874"}],
        dry_run=True
    )
    assert res["folder_path"] == fake_folder
    assert "Thư mục không tồn tại" in res["folder_error"]
    assert len(res["results"]) == 1
    r0 = res["results"][0]
    assert r0["bbbg_status"] == "LỖI_THƯ_MỤC"
    assert r0["vat_status"] == "LỖI_THƯ_MỤC"
    assert r0["overall_status"] == "LỖI_THƯ_MỤC"

    report = runner.format_teams_report(res)
    assert "Cảnh báo thư mục" in report
    assert "| Số DOC | Số phiếu | Điền HeadText | BBBG | VAT | Trạng thái |" in report
    assert "Mail (.msg)" not in report
    assert "Chi tiết" not in report


def test_clean_folder_path_auto_complete_truncated_prefix():
    """Kiểm tra tự động hoàn thiện đường dẫn nếu bị Outlook ngắt dòng ở ký tự 73 (chữ 'Hải')."""
    from app.sap_attachment import DEFAULT_MIGO_ATT_FOLDER
    # Đoạn bị Outlook ngắt dòng ở ảnh của người dùng:
    truncated = r"\\10.128.10.72\fvn$\03.Khối sản xuất\01.Nhà máy VinFast Hải"
    resolved = clean_folder_path(truncated)
    assert resolved == DEFAULT_MIGO_ATT_FOLDER
    assert "260802_ File ATT BBGH" in resolved


def test_email_parser_att_mail_without_folder_uses_default():
    """Kiểm tra lệnh 'SAP att mail' khi không truyền link UNC thì tự động dùng thư mục mặc định từ cell B3."""
    from app.email_parser import SmartCommandParser
    from app.sap_attachment import DEFAULT_MIGO_ATT_FOLDER
    text = "SAP att mail\n5006458947 | Thành Công in 3874"
    meta = SmartCommandParser.parse(text)
    assert meta["command"] == "ATT_MIGO"
    assert meta["folderPath"] == DEFAULT_MIGO_ATT_FOLDER
    assert len(meta["attachmentItems"]) == 1
    assert meta["attachmentItems"][0]["doc"] == "5006458947"


def test_parse_attachment_large_list_111_rows():
    """Kiểm tra xử lý danh sách lớn 111 dòng từ người dùng gửi (tab, dấu cách, trailing backslash)."""
    large_input = """5007135407\tAUTOMECH IN 6456
5007135435\tAUTOMECH IN 6456
5007131679\tTÂM HỢP in - 6434
5007131612\tTÂM HỢP in - 6434
5007129150\ttâm hợp in - 6434
5007129128\ttâm hợp in - 6434
5007129127\ttâm hợp in - 6434
5007129129\ttâm hợp in - 6434
5007125132\tLTV in 6419
5007125130\tLTV in 6419
5007125109\tLTV in 6419
5007125142\tLTV in 6419
5007125102\tLTV in 6419
5007125096\tLTV in 6419
5007125095\tLTV in 6419
5007125141\tLTV in 6419
5007125098\tLTV in 6419
5007125105\tLTV in 6419
5007125112\tLTV in 6419
5007125111\tLTV in 6419
5007120919\tLTV in 6397
5007120918\tLTV in 6397
5007120915\tLTV in 6397
5007120760\tltv in  6397
5007120751\tltv in  6397
5007120718\tltv in  6397
5007120699\tltv in  6397
5007120738\tltv in  6397
5007120728\tltv in  6397
5007120727\tltv in  6397
5007117582\tTÂM HỢP in 6376
5007117581\tTÂM HỢP in 6376
5007117303\tLTV in 6371
5007117242\tLTV IN-6371
5007117241\tLTV IN-6371
5007117240\tLTV IN-6371
5007117216\tLTV IN-6371
5007117215\tLTV IN-6371
5007117213\tLTV IN-6371
5007117208\tLTV IN-6371
5007117220\tLTV IN-6371
5007117214\tLTV IN-6371
5007117209\tLTV IN-6371
5007117195\tLTV IN-6371
5007117196\tLTV IN-6371
5007116499\ttâm hợp in 6376
5007116498\ttâm hợp in 6376
5007116489\ttâm hợp in 6376
5007116477\ttâm hợp in 6376
5007116488\ttâm hợp in 6376
5007116509\ttâm hợp in 6376
5007116478\ttâm hợp in 6376
5007116487\ttâm hợp in 6376
5007116507\ttâm hợp in 6376
5007116506\ttâm hợp in 6376
5007115777\tLTV in 6371
5007115774\tLTV in 6371
5007115784\tLTV in 6371
5007113497\tLTV in - 6353
5007113496\tLTV in - 6353
5007113495\tLTV in - 6353
5007113494\tLTV in - 6353
5007113493\tLTV in - 6353
5007113492\tLTV in - 6353
5007113491\tLTV in - 6353
5007113490\tLTV in - 6353
5007113417\tLTV in - 6353
5007113425\tLTV in - 6353
5007113431\tLTV in - 6353
5007113434\tLTV in - 6353
5007113416\tLTV in - 6353
5007113424\tLTV in - 6353
5007113415\tLTV in - 6353
5007113430\tLTV in - 6353
5007113423\tLTV in - 6353
5007113433\tLTV in - 6353
5007113432\tLTV in - 6353
5007109449\tTâm Hợp in 6346
5007109456\tTâm Hợp IN-6346
5007109427\tTâm Hợp IN-6346
5007109458\tTâm Hợp IN-6346
5007109452\tTâm Hợp IN-6346
5007109455\tTâm Hợp IN-6346
5007109408\tTâm Hợp IN-6346
5007109453\tTâm Hợp IN-6346
5007109454\tTâm Hợp IN-6346
5007109451\tTâm Hợp IN-6346
5007109445\tTâm Hợp IN-6346
5007109426\tTâm Hợp IN-6346
5007109409\tTâm Hợp IN-6346
5007104421\tTâm Hợp in - 6322
5007104387\tTâm Hợp in - 6322
5007104388\tTâm Hợp in - 6322
5007104405\tTâm Hợp in - 6322
5007104410\tTâm Hợp in - 6322
5007104349\tTâm Hợp in - 6322
5007104404\tTâm Hợp in - 6322
5007104384\tTâm Hợp in - 6322
5007104347\tTâm Hợp in - 6322
5007104385\tTâm Hợp in - 6322
5007104386\tTâm Hợp in - 6322
5007104403\tTâm Hợp in - 6322
5007104395\tTâm Hợp in - 6322
5007104391\tTâm Hợp in - 6322
5007104394\tTâm Hợp in - 6322
5007104383\tTâm Hợp in - 6322
5006965415\tLTV in 5663
5006965374\tLTV in 5663
5006965384\tLTV in 5663
5006965371\tLTV in 5663
5006458947\tThành Công in 3874\\"""
    items = parse_attachment_items_from_text(large_input)
    assert len(items) == 111
    assert items[0]["doc"] == "5007135407"
    assert items[0]["ticket_no"] == "AUTOMECH IN 6456"
    assert items[-1]["doc"] == "5006458947"
    # Dấu gạch chéo ngược ở cuối đã được làm sạch
    assert items[-1]["ticket_no"] == "Thành Công in 3874"


def test_parse_attachment_power_automate_wrapped_text():
    """Kiểm tra xử lý văn bản bị ngắt dòng bởi Power Automate / Outlook ở cột 80 ký tự không bị cắt cụt số phiếu."""
    wrapped_text = """RAW_TEXT: SAP att mail  5007136566 thành đạt in - 6461 5007136565 thành đạt in - 6461
5007136564 thành
đạt in - 6461 5007136563 thành đạt in - 6461 5007136619 thành đạt in - 6461
5007136634 thành đạt in - 6461 5007136633 thành đạt in - 6461 5007136632 thành
đạt in - 6461 5007136621 thành đạt in - 6461 5007136612 thành đạt in - 6461
5007136630 thành đạt in - 6461 5007136620 thành đạt in - 6461 5007136631 thành
đạt in - 6461 5007136589 thành đạt in - 6461 5007136599 thành đạt in - 6461
5007133359 Thành Đạt in - 6449 5007133223 Thành Đạt in - 6449 5007133301 Thành
Đạt in - 6449 5007133292 Thành Đạt in - 6449 5007133293 Thành Đạt in -
6449 5007133257 Thành Đạt in - 6449"""
    items = parse_attachment_items_from_text(wrapped_text)
    item_map = {it["doc"]: it["ticket_no"] for it in items}

    # Các chứng từ từng bị cắt thành 'thành' do ngắt dòng nay phải được ghép hoàn chỉnh
    assert item_map["5007136564"] == "thành đạt in - 6461"
    assert item_map["5007136632"] == "thành đạt in - 6461"
    assert item_map["5007136631"] == "thành đạt in - 6461"
    assert item_map["5007133301"] == "Thành Đạt in - 6449"
    assert item_map["5007133293"] == "Thành Đạt in - 6449"
    assert item_map["5007133257"] == "Thành Đạt in - 6449"


def test_parse_attachment_last_item_with_signature():
    """Kiểm tra chứng từ cuối cùng có chữ ký email phía sau thì không lấy nhầm chữ ký vào số phiếu."""
    text = """5006458947 Thành Công in 3874
Người yêu cầu: Phạm Tiến Tài (FVN-KSX-NMVFHP)"""
    items = parse_attachment_items_from_text(text)
    assert len(items) == 1
    assert items[0]["doc"] == "5006458947"
    assert items[0]["ticket_no"] == "Thành Công in 3874"


def test_check_existing_sap_attachments():

    """Kiểm tra đối chiếu các file đính kèm đã có trên SAP MIGO với số phiếu."""
    titles = [
        "3874 BB",
        "VAT 3874",
        "Email_trao_doi.msg",
        "Other_Document.pdf"
    ]
    # Khớp số phiếu 3874 từ text 'Thành Công in 3874'
    bbbg, vat = MigoAttachmentRunner.check_existing_sap_attachments(titles, "Thành Công in 3874")
    assert bbbg == "3874 BB"
    assert vat == "VAT 3874"

    # Trường hợp số phiếu 6456
    titles_6456 = ["6456 BBBG.pdf", "HD_VAT_6456.xlsx"]
    bbbg, vat = MigoAttachmentRunner.check_existing_sap_attachments(titles_6456, "AUTOMECH IN 6456")
    assert bbbg == "6456 BBBG.pdf"
    assert vat == "HD_VAT_6456.xlsx"

    # Trường hợp trên SAP chỉ có BBBG, chưa có VAT
    titles_only_bb = ["3874 BB"]
    bbbg, vat = MigoAttachmentRunner.check_existing_sap_attachments(titles_only_bb, "Thành Công in 3874")
    assert bbbg == "3874 BB"
    assert vat is None

    # Không khớp số phiếu
    bbbg, vat = MigoAttachmentRunner.check_existing_sap_attachments(titles, "LTV in 6397")
    assert bbbg is None
    assert vat is None


def test_process_attachment_batch_with_sap_existing_attachments():
    """Kiểm tra khi BBBG đã có sẵn trên SAP MIGO thì bot đánh dấu Thành Công in 3874, không báo FILE_NOT_FOUND."""
    from unittest.mock import MagicMock, create_autospec
    from app.sap_engine import SapEngine
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_session = MagicMock()
    mock_sap.ensure_session.return_value = mock_session

    runner = owned_runner(mock_session, mock_sap)
    
    # Mock read_sap_attachment_titles và open_migo_document
    runner.open_migo_document = lambda session, doc, year=None, mode="DISPLAY": (True, "OK")
    runner.read_sap_attachment_titles = lambda session: ["3874 BB"]
    
    # Giả lập thư mục không có file nào (để chứng minh không upload lại từ thư mục)
    import tempfile
    with tempfile.TemporaryDirectory() as tmp_dir:
        res = runner.process_attachment_batch(
            folder_path=tmp_dir,
            items=[{"doc": "5006458947", "ticket_no": "Thành Công in 3874", "year": "2026"}],
            fill_headtext=False, dry_run=False
        )
        assert len(res["results"]) == 1
        item = res["results"][0]
        # BBBG đã có trên SAP -> đánh dấu DONE
        assert item["bbbg_status"] == "DONE"
        # VAT chưa có trên SAP và thư mục không có -> FILE_NOT_FOUND
        assert item["vat_status"] == "FILE_NOT_FOUND"
        # Overall status thành công vì đã có BBBG hợp lệ (PARTIAL)
        assert item["overall_status"] in ("DONE", "PARTIAL (Thiếu VAT)")
        assert res["total_partial"] == 1
        assert res["total_success"] == 0


def test_multiple_vat_files_matching():
    """Kiểm tra một chứng từ có nhiều file VAT trong thư mục thì đính kèm tất cả và báo DONE."""
    import tempfile
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        # Tạo 2 file VAT cho số phiếu 6461
        (tmp_path / "6461_VAT_1.pdf").write_text("vat1")
        (tmp_path / "6461_VAT_2.pdf").write_text("vat2")
        (tmp_path / "6461_BBBG.pdf").write_text("bbbg")

        runner = MigoAttachmentRunner()
        idx_data = runner.index_source_folder(str(tmp_path))
        all_matches = runner.match_all_files("5007136566", "thành đạt in - 6461", idx_data)
        assert len(all_matches["vat"]) == 2
        assert len(all_matches["bbbg"]) == 1

        res = runner.process_attachment_batch(
            folder_path=str(tmp_path),
            items=[{"doc": "5007136566", "ticket_no": "thành đạt in - 6461"}],
            dry_run=True
        )
        assert len(res["results"]) == 1
        r0 = res["results"][0]
        assert r0["bbbg_status"] == "SIMULATED"
        assert r0["vat_status"] == "SIMULATED"
        assert r0["overall_status"] == "SIMULATED"


def test_custom_vat_naming_fallback():
    """Kiểm tra hóa đơn không có chữ VAT mà dùng ký hiệu riêng như file Excel thì vẫn nhận diện được."""
    import tempfile
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        # File mang số phiếu 6461 nhưng không có chữ BBBG hoặc Mail
        (tmp_path / "6461_CT01.pdf").write_text("invoice")

        runner = MigoAttachmentRunner()
        idx_data = runner.index_source_folder(str(tmp_path))
        all_matches = runner.match_all_files("5007136566", "thành đạt in - 6461", idx_data)
        assert len(all_matches["vat"]) == 1
        assert all_matches["vat"][0].name == "6461_CT01.pdf"


def test_sanitize_migo_session():
    """Kiểm tra hàm dọn dẹp tab SAP khi gặp popup kẹt và reset về /nMIGO."""
    from app.sap_attachment import sanitize_migo_session
    mock_session = MagicMock()
    mock_session.findById.side_effect = lambda id_str, *args: MagicMock() if ('wnd[2]' in id_str or 'wnd[1]' in id_str or 'wnd[0]' in id_str) else None
    
    sanitize_migo_session(mock_session)
    assert mock_session.findById.called



def test_sap_com_retry_logic():
    """Kiểm tra cơ chế retry khi gặp lỗi COM tạm thời."""
    from app.sap_attachment import sap_com_retry
    import pywintypes

    attempts = 0
    def flaky_func():
        nonlocal attempts
        attempts += 1
        if attempts < 3:
            raise pywintypes.com_error(-2147417846, 'The message filter indicated that the application is busy.', None, None)
        return 'SUCCESS'

    res = sap_com_retry(flaky_func, retries=5, delay=0.01)
    assert res == 'SUCCESS'
    assert attempts == 3









def test_open_migo_document_esc_cancel():
    """Reload validates document/year after Enter, not just successful navigation."""
    mock_session = MagicMock()
    runner = owned_runner(mock_session)
    stale = MagicMock(Changeable=False)
    active = MagicMock(Changeable=True, Text="")
    year_field = MagicMock(Text="")
    window = MagicMock()
    statusbar = MagicMock(MessageType="S", Text="Document loaded")
    mock_session.findById.side_effect = lambda cid, *a: statusbar if "sbar" in cid else window
    def enter(key):
        if key == 0:
            active.Changeable = False
    window.SendVKey.side_effect = enter
    from unittest.mock import patch
    with patch("app.sap_attachment.smart_find", side_effect=[stale, active, year_field, active, year_field]):
        ok, msg = runner.open_migo_document(mock_session, "5000000001", "2026")
    assert ok is True
    assert runner._active_document == ("5000000001", "2026")
    window.SendVKey.assert_any_call(12)



def test_batch_summary_success_count_consistent():
    """Kiểm tra tổng số thành công total_success và nhãn is_succ được đếm chuẩn xác."""
    runner = MigoAttachmentRunner()
    batch_res = runner.process_attachment_batch(
        folder_path="",
        items=[
            {"doc": "5000000001", "ticket_no": "TICK01"},
            {"doc": "5000000002", "ticket_no": "TICK02"},
            {"doc": "5000000003", "ticket_no": ""},
        ],
        dry_run=True
    )
    assert batch_res["total_success"] == 0
    assert batch_res["total_preview_ready"] == 2
    assert batch_res["total_missing"] == 3
    assert batch_res["results"][0]["is_succ"] is False
    assert batch_res["results"][1]["is_succ"] is False
    assert batch_res["results"][2]["is_succ"] is False



def test_handle_migo_headtext_different_text_returns_review_without_overwrite():
    """Kiểm tra quy tắc A01: Header Text đã có nội dung khác thì trả về HEADTEXT_REVIEW, KHÔNG ghi đè và KHÔNG bấm Save."""
    from unittest.mock import MagicMock
    mock_session = MagicMock()
    mock_field = MagicMock()
    mock_field.Text = "TICKET_OLD_123"  # Đã có sẵn số phiếu cũ khác
    mock_wnd0 = MagicMock()
    mock_wnd0.findByName.return_value = mock_field
    mock_session.findById.return_value = mock_wnd0

    runner = owned_runner(mock_session)
    status, msg = runner.handle_migo_headtext(mock_session, "TICKET_NEW_456")
    assert status == "HEADTEXT_REVIEW"
    assert "Cần Review thủ công, không ghi đè ngầm" in msg
    # Bảo đảm không gọi SendVKey(11) để Save
    mock_wnd0.SendVKey.assert_not_called()


def test_process_attachment_batch_multi_ticket_conflict_detection():
    """Kiểm tra quy tắc A02: Phát hiện xung đột MULTI_TICKET_DOC khi cùng 1 DOC có nhiều ticket khác nhau trong cùng lô."""
    runner = MigoAttachmentRunner()
    items = [
        {"doc": "5006995954", "ticket_no": "TICKET_A"},
        {"doc": "5006995954", "ticket_no": "TICKET_B"},
        {"doc": "5007072362", "ticket_no": "TICKET_C"}
    ]
    summary = runner.process_attachment_batch("", items, dry_run=True)
    # DOC 5006995954 phải được gắn cờ conflict
    conflict_items = [r for r in summary["results"] if r.get("doc") == "5006995954"]
    assert len(conflict_items) == 2


def test_classify_sap_attachments_full_logic():
    """Kiểm tra phân loại file đính kèm BBBG / VAT cho các trạng thái: FULL_ATT, MISSING_VAT, MISSING_BBBG, NO_ATT, REVIEW."""
    from app.sap_attachment import classify_sap_attachments

    # 1. Chưa có file nào
    r_empty = classify_sap_attachments([])
    assert r_empty["status_code"] == "NO_ATT"
    assert "Chưa có cả BBBG lẫn VAT" in r_empty["overall_status"]
    assert r_empty["has_bbbg"] is False and r_empty["has_vat"] is False

    # 2. Đã có cả BBBG và VAT
    r_full = classify_sap_attachments(["5007145690_BBBG.pdf", "HDGTGT_3874.xlsx"])
    assert r_full["status_code"] == "FULL_ATT"
    assert "Đã có cả BBBG và VAT" in r_full["overall_status"]
    assert r_full["has_bbbg"] is True and r_full["has_vat"] is True

    # 3. Có BBBG, thiếu VAT
    r_miss_vat = classify_sap_attachments(["5007145690_BBBG.pdf"])
    assert r_miss_vat["status_code"] == "MISSING_VAT"
    assert "thiếu VAT" in r_miss_vat["overall_status"]
    assert r_miss_vat["has_bbbg"] is True and r_miss_vat["has_vat"] is False

    # 4. Có VAT, thiếu BBBG
    r_miss_bb = classify_sap_attachments(["HD_VAT_5007145690.xlsx"])
    assert r_miss_bb["status_code"] == "MISSING_BBBG"
    assert "thiếu BBBG" in r_miss_bb["overall_status"]
    assert r_miss_bb["has_bbbg"] is False and r_miss_bb["has_vat"] is True

    # 5. Có file nhưng không nhận diện được loại
    r_rev = classify_sap_attachments(["tailieu_xacnhan.docx"])
    assert r_rev["status_code"] == "REVIEW"
    assert "chưa rõ loại" in r_rev["overall_status"]


def test_open_migo_document_default_year():
    """Kiểm tra open_migo_document tự động lấy năm hiện tại nếu không truyền year."""
    from unittest.mock import MagicMock, patch
    mock_session = MagicMock()
    runner = owned_runner(mock_session)
    active = MagicMock(Changeable=True, Text="")
    year_field = MagicMock(Text="")
    window = MagicMock()
    statusbar = MagicMock(MessageType="S", Text="Document loaded")
    mock_session.findById.side_effect = lambda cid, *a: statusbar if "sbar" in cid else window

    def enter(key):
        if key == 0:
            active.Changeable = False
    window.SendVKey.side_effect = enter

    with patch("app.sap_attachment.smart_find", side_effect=[active, year_field, active, year_field]):
        ok, msg = runner.open_migo_document(mock_session, "5007145690")
    assert ok is True
    from datetime import datetime
    assert runner._active_document == ("5007145690", str(datetime.now().year))


def test_open_migo_document_auto_switch_display_to_change():
    """Kiểm tra khi MIGO đang ở A04 Display, open_migo_document tự động chuyển thành A12 Change và R02."""
    from unittest.mock import MagicMock, patch
    mock_session = MagicMock()
    runner = owned_runner(mock_session)
    active = MagicMock(Changeable=True, Text="")
    year_field = MagicMock(Text="")
    window = MagicMock()
    statusbar = MagicMock(MessageType="S", Text="Document loaded")

    class MockComboBox:
        def __init__(self, key):
            self.Key = key

    act_cb = MockComboBox("A04")
    ref_cb = MockComboBox("R01")

    def find_by_name(name, ctype):
        if name == "GODYNPRO-ACTION":
            return act_cb
        elif name == "GODYNPRO-REFDOC":
            return ref_cb
        return MagicMock()

    window.findByName.side_effect = find_by_name
    mock_session.findById.side_effect = lambda cid, *a: statusbar if "sbar" in cid else window

    def enter(key):
        if key == 0:
            active.Changeable = False
    window.SendVKey.side_effect = enter

    with patch("app.sap_attachment.smart_find", side_effect=[active, year_field, active, year_field]):
        ok, msg = runner.open_migo_document(mock_session, "5007145690", mode="CHANGE")

    assert ok is True
    assert act_cb.Key == "A12"
    assert ref_cb.Key == "R02"


def test_multi_tab_session_identity_and_eligibility():
    """Kiểm tra xác thực phiên đa tab dựa trên system/client/user và tính toán is_multi_tab_eligible."""
    from unittest.mock import MagicMock, create_autospec
    from app.sap_engine import SapEngine
    mock_sap = create_autospec(SapEngine, instance=True)
    mock_session_tab0 = MagicMock()
    mock_session_tab0.Info.SystemName = "PRD"
    mock_session_tab0.Info.Client = "100"
    mock_session_tab0.Info.User = "USER_TEST"
    mock_session_tab0.Info.Transaction = "MIGO"
    mock_session_tab0.Id = "/app/con[0]/ses[0]"
    mock_session_tab0.Busy = False

    expected_identity = {
        "system": "PRD",
        "client": "100",
        "user": "USER_TEST",
        "session_id": "/app/con[0]/ses[0]"
    }
    runner = MigoAttachmentRunner(mock_sap, expected_session_identity=expected_identity)

    # Tab 0: Kiểm tra hợp lệ
    ok0, msg0 = runner.verify_session_identity(mock_session_tab0)
    assert ok0 is True

    # Tab 1: Cùng system/client/user nhưng khác session_id (/app/con[0]/ses[1])
    mock_session_tab1 = MagicMock()
    mock_session_tab1.Info.SystemName = "PRD"
    mock_session_tab1.Info.Client = "100"
    mock_session_tab1.Info.User = "USER_TEST"
    mock_session_tab1.Info.Transaction = "MIGO"
    mock_session_tab1.Id = "/app/con[0]/ses[1]"
    mock_session_tab1.Busy = False

    # Tab 1: Khác session_id -> Phải reject vì không khớp ownership được giao
    ok1, msg1 = runner.verify_session_identity(mock_session_tab1)
    assert ok1 is False
    assert msg1 == "SESSION_IDENTITY_MISMATCH"

    # Tab khác Client -> Phải reject
    mock_session_bad = MagicMock()
    mock_session_bad.Info.SystemName = "PRD"
    mock_session_bad.Info.Client = "200"
    mock_session_bad.Info.User = "USER_TEST"
    mock_session_bad.Busy = False
    ok_bad, msg_bad = runner.verify_session_identity(mock_session_bad)
    assert ok_bad is False
    assert msg_bad == "SESSION_IDENTITY_MISMATCH"



def test_deduplicate_files_by_content_and_claim_matching():
    """Kiểm tra tính năng loại bỏ file trùng lặp nội dung và thuật toán so khớp SAP GOS Claim-based."""
    import tempfile
    from app.sap_attachment import deduplicate_files_by_content, match_files_to_sap

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        # Tạo 2 file có nội dung y hệt nhau (duplicate)
        f_dup1 = tmp_path / "VAT_5122.pdf"
        f_dup2 = tmp_path / "VAT 5122.pdf"
        f_dup1.write_bytes(b"SAME_CONTENT_12345")
        f_dup2.write_bytes(b"SAME_CONTENT_12345")

        # Tạo file thứ 3 có nội dung khác (distinct)
        f_diff = tmp_path / "VAT 5122(1).pdf"
        f_diff.write_bytes(b"DIFFERENT_CONTENT_67890")

        # Deduplication check
        deduped = deduplicate_files_by_content([f_dup1, f_dup2, f_diff])
        assert len(deduped) == 2
        assert f_diff in deduped
        assert (f_dup1 in deduped) or (f_dup2 in deduped)

        # Claim matching check: SAP đã có sẵn 'VAT_5122'
        sap_titles = ["VAT_5122"]
        to_upload, already_on_sap, unclaimed = match_files_to_sap(
            deduped, sap_titles, ticket_nums=["5122"], file_type="vat"
        )
        # File f_diff (VAT 5122(1)) chưa có trên SAP -> phải nằm trong to_upload
        assert len(to_upload) == 1
        assert to_upload[0].name == "VAT 5122(1).pdf"
        # File f_dup1 đã có trên SAP
        assert len(already_on_sap) == 1
        assert already_on_sap[0][1] == "VAT_5122"


def test_incremental_multi_vat_attachment():
    """Kiểm tra kịch bản đính kèm lũy kế: DOC đã có sẵn 1 VAT trên SAP, đính thêm file VAT thứ 2."""
    import tempfile
    from unittest.mock import MagicMock, create_autospec
    from app.sap_attachment import MigoAttachmentRunner
    from app.sap_engine import SapEngine

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        f1 = tmp_path / "VAT_5122.pdf"
        f2 = tmp_path / "VAT 5122(1).pdf"
        f_bb = tmp_path / "BBBG 5122.pdf"
        f1.write_bytes(b"VAT1")
        f2.write_bytes(b"VAT2")
        f_bb.write_bytes(b"BBBG")

        mock_sap = create_autospec(SapEngine, instance=True)
        mock_session = MagicMock()
        mock_session.Info.SystemName = "PRD"
        mock_session.Info.Client = "100"
        mock_session.Info.User = "TEST_USER"
        mock_session.Info.Transaction = "MIGO"
        mock_session.Id = "/app/con[0]/ses[0]"
        mock_session.Busy = False
        mock_sap.ensure_session.return_value = mock_session
        identity = {"system": "PRD", "client": "100", "user": "TEST_USER", "session_id": "/app/con[0]/ses[0]"}
        mock_sap.session_identity = identity
        runner = MigoAttachmentRunner(mock_sap, expected_session_identity=identity)
        runner.open_migo_document = lambda session, doc, year=None, mode="CHANGE": (True, "OK")
        # Giả lập SAP đã có sẵn file 'VAT_5122' và 'BBBG 5122'
        runner.read_sap_attachment_titles = lambda session: ["VAT_5122", "BBBG 5122"]

        uploaded_files = []
        def mock_attach(session, file_path):
            uploaded_files.append(file_path.name)
            return True, f"DONE_SAP_READBACK: {file_path.name}"
        runner.attach_file_to_current_migo = mock_attach

        res = runner.process_attachment_batch(
            folder_path=str(tmp_path),
            items=[{"doc": "5007145690", "ticket_no": "5122", "year": "2026"}],
            fill_headtext=False,
            dry_run=False
        )

        assert len(res["results"]) == 1
        r0 = res["results"][0]
        # File BBBG 5122 đã có sẵn -> không upload lại
        assert "BBBG 5122.pdf" not in uploaded_files
        # File VAT_5122 đã có sẵn -> không upload lại
        assert "VAT_5122.pdf" not in uploaded_files
        # File VAT 5122(1).pdf chưa có -> ĐƯỢC UPLOAD THÊM!
        assert "VAT 5122(1).pdf" in uploaded_files
        assert r0["vat_status"] == "DONE (Đính thêm 1, đã có 1)"
        assert r0["overall_status"] == "DONE"
        assert res["total_success"] == 1


def test_handle_migo_headtext_substring_containment():
    """Kiểm tra: Khi SAP Header Text đã có 'Thành Công IN-6058' và lệnh truyền '6058', nhận diện là HEADTEXT_OK."""
    mock_session = MagicMock()
    mock_field = MagicMock()
    mock_field.Text = "Thành Công IN-6058"
    mock_wnd0 = MagicMock()
    mock_wnd0.findByName.return_value = mock_field
    mock_session.findById.return_value = mock_wnd0

    runner = owned_runner(mock_session)
    status, msg = runner.handle_migo_headtext(mock_session, "6058")
    assert status == "HEADTEXT_OK"
    assert "Đã có đúng số phiếu ('Thành Công IN-6058')" in msg
    mock_wnd0.SendVKey.assert_not_called()


def test_migo_runner_active_document_thread_isolation():
    """Kiểm tra: Multi-tab đa luồng có active_document được cách ly qua threading.local()."""
    import threading
    import time
    runner = MigoAttachmentRunner()

    results = {}

    def worker_1():
        runner._active_document = ("5007000001", "2026")
        time.sleep(0.05)
        results["t1"] = runner._active_document

    def worker_2():
        time.sleep(0.02)
        runner._active_document = ("5007000002", "2026")
        time.sleep(0.05)
        results["t2"] = runner._active_document

    t1 = threading.Thread(target=worker_1)
    t2 = threading.Thread(target=worker_2)
    t1.start()
    t2.start()
    t1.join()
    t2.join()

    assert results["t1"] == ("5007000001", "2026")
    assert results["t2"] == ("5007000002", "2026")
    assert runner._active_document is None


def test_fast_path_sbar_confirmation_skips_readback():
    """Kiểm tra Fast-path Sbar Confirmation không gọi read_sap_attachment_titles khi sbar trả về 'S'."""
    from unittest.mock import MagicMock
    import tempfile
    from pathlib import Path
    from app.sap_attachment import POPUP_LOCK

    mock_session = MagicMock()
    mock_sbar = MagicMock()
    mock_sbar.Text = "Attachment successfully created"
    mock_sbar.MessageType = "S"

    mock_shell = MagicMock()
    wnd1_active = [False]
    mock_shell.SelectContextMenuItem.side_effect = lambda *a, **k: wnd1_active.__setitem__(0, True)
    mock_btn = MagicMock()
    mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_btn.Press.side_effect = lambda: wnd1_active.__setitem__(0, False)
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.side_effect = lambda name, *args: mock_btn if "btn" in name else MagicMock()

    def find_by_id(path):
        path_str = str(path)
        if "shell" in path_str:
            return mock_shell
        elif "sbar" in path_str:
            return mock_sbar
        elif "btn" in path_str:
            return mock_btn
        elif path_str == "wnd[1]" or path_str.startswith("wnd[1]"):
            if wnd1_active[0]:
                return mock_wnd1
            raise Exception("Control wnd[1] not found")
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    runner = owned_runner(mock_session)
    mock_read_titles = MagicMock(return_value=["some_other_file.pdf"])
    runner.read_sap_attachment_titles = mock_read_titles

    with tempfile.NamedTemporaryFile(suffix="_VAT.pdf", delete=False) as tf:
        tf.write(b"dummy pdf")
        dummy_file = Path(tf.name)

    try:
        # P1-07: Xác minh vật lý qua readback là bắt buộc để chống stale sbar và fake success
        mock_read_titles.return_value = [dummy_file.name]
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is True
        assert "DONE_SAP_READBACK" in msg or "DONE_SAP_SBAR" in msg
        mock_read_titles.assert_called_once()
        # POPUP_LOCK phải được release, không bị kẹt
        assert not POPUP_LOCK.locked()
    finally:
        dummy_file.unlink(missing_ok=True)


def test_fallback_readback_when_sbar_uncertain():
    """Kiểm tra khi sbar không phải 'S' rõ ràng thì fallback sang read_sap_attachment_titles."""
    from unittest.mock import MagicMock
    import tempfile
    from pathlib import Path

    mock_session = MagicMock()
    mock_sbar = MagicMock()
    mock_sbar.Text = "Document 5000000001 changed"  # Không có SBAR MessageType = 'S', nhưng có chữ SUCCESS hoặc kiểm tra fallback
    mock_sbar.MessageType = "I"

    mock_shell = MagicMock()
    wnd1_active = [False]
    mock_shell.SelectContextMenuItem.side_effect = lambda *a, **k: wnd1_active.__setitem__(0, True)
    mock_btn = MagicMock()
    mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_btn.Press.side_effect = lambda: wnd1_active.__setitem__(0, False)
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.side_effect = lambda name, *args: mock_btn if "btn" in name else MagicMock()

    def find_by_id(path):
        path_str = str(path)
        if "shell" in path_str:
            return mock_shell
        elif "sbar" in path_str:
            return mock_sbar
        elif "btn" in path_str:
            return mock_btn
        elif path_str == "wnd[1]" or path_str.startswith("wnd[1]"):
            if wnd1_active[0]:
                return mock_wnd1
            raise Exception("Control wnd[1] not found")
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    runner = owned_runner(mock_session)

    with tempfile.NamedTemporaryFile(suffix="_VAT.pdf", delete=False) as tf:
        tf.write(b"dummy pdf")
        dummy_file = Path(tf.name)

    try:
        # Nếu sbar trả về I và không có từ khóa created/saved -> ATTACHMENT_UNCERTAIN
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is False
        assert "ATTACHMENT_UNCERTAIN" in msg

        # Giờ test nếu sbar là message bình thường có 'SAVED' nhưng MessageType='W' -> kích hoạt fallback readback
        mock_sbar.Text = "Attachment was SAVED with warning"
        mock_sbar.MessageType = "W"
        mock_read_titles = MagicMock(return_value=[dummy_file.name])
        runner.read_sap_attachment_titles = mock_read_titles

        ok2, msg2 = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok2 is True
        assert "DONE_SAP_READBACK" in msg2
        mock_read_titles.assert_called_once()
    finally:
        dummy_file.unlink(missing_ok=True)


def test_multi_tab_scheduler_dynamic_scaling():
    """Kiểm tra Multi-tab scheduler tự thích ứng 3-4 tabs khi có đủ sessions."""
    from unittest.mock import patch, MagicMock
    from app.sap_attachment import MigoAttachmentRunner

    runner = MigoAttachmentRunner()
    mock_engine = MagicMock()
    mock_engine.user_reserved_sessions = 2
    runner.sap_engine = mock_engine

    items = [{"doc": f"500700000{i}", "ticket_no": f"TICK_{i}"} for i in range(10)]

    # Giả lập SAP có 6 tabs khả dụng, user_reserved=2 -> bot_tabs = [2, 3, 4, 5] (4 tabs)
    with patch("app.sap_attachment.pythoncom", MagicMock()), \
         patch("app.sap_attachment.ensure_sap_tabs", return_value=6), \
         patch.object(runner, "process_single_migo_item", return_value=({"doc": "500"}, True)), \
         patch.object(runner, "prepare_migo_session", return_value=True), \
         patch("app.sap_attachment.win32com.client.GetObject") as mock_get_obj:

        mock_gui = MagicMock()
        mock_conn = MagicMock()
        mock_conn.Children.Count = 6
        mock_conn.Children.side_effect = lambda idx: MagicMock()
        mock_gui.GetScriptingEngine.Children.return_value = mock_conn
        mock_get_obj.return_value = mock_gui

        batch_res = runner.process_attachment_batch(
            folder_path="",
            items=items,
            dry_run=False
        )

        assert batch_res["total_items"] == 10
        assert batch_res["total_success"] == 10
        assert len(batch_res["results"]) == 10


def test_stale_sbar_message_falls_back_to_readback_and_verifies_exact_match():
    """Kiểm tra thông điệp cũ (stale sbar: Material document displayed) không kích hoạt Fast-path mà phải fallback."""
    from unittest.mock import MagicMock
    import tempfile
    from pathlib import Path

    mock_session = MagicMock()
    mock_sbar = MagicMock()
    # Thông điệp cũ còn sót lại từ lệnh open_migo_document, có type 'S' nhưng không phải GOS creation
    mock_sbar.Text = "Material document 5007000001 displayed"
    mock_sbar.MessageType = "S"

    mock_shell = MagicMock()
    wnd1_active = [False]
    mock_shell.SelectContextMenuItem.side_effect = lambda *a, **k: wnd1_active.__setitem__(0, True)
    mock_btn = MagicMock()
    mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
    mock_btn.Name = "btn[0]"
    mock_btn.Type = "GuiButton"
    mock_btn.Press.side_effect = lambda: wnd1_active.__setitem__(0, False)
    mock_wnd1 = MagicMock()
    mock_wnd1.FindByName.side_effect = lambda name, *args: mock_btn if "btn" in name else MagicMock()

    def find_by_id(path):
        path_str = str(path)
        if "shell" in path_str:
            return mock_shell
        elif "sbar" in path_str:
            return mock_sbar
        elif "btn" in path_str:
            return mock_btn
        elif path_str == "wnd[1]" or path_str.startswith("wnd[1]"):
            if wnd1_active[0]:
                return mock_wnd1
            raise Exception("Control wnd[1] not found")
        return MagicMock()

    mock_session.findById.side_effect = find_by_id
    runner = owned_runner(mock_session)

    with tempfile.NamedTemporaryFile(suffix="_BBBG.pdf", delete=False) as tf:
        tf.write(b"dummy pdf")
        dummy_file = Path(tf.name)

    try:
        # TH1: Fallback readback tìm thấy file thật trên SAP -> DONE_SAP_READBACK
        mock_read_titles = MagicMock(return_value=[dummy_file.name])
        runner.read_sap_attachment_titles = mock_read_titles

        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is True
        assert "DONE_SAP_READBACK" in msg
        mock_read_titles.assert_called_once()

        # TH2: Fallback readback KHÔNG tìm thấy file -> Không được báo DONE_SAP_SBAR giả tạo
        mock_read_titles.return_value = ["some_other_doc.pdf"]
        ok2, msg2 = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok2 is False
        assert "ATTACHMENT_UNCERTAIN" in msg2
    finally:
        dummy_file.unlink(missing_ok=True)




def test_verify_session_identity_multi_tab_bypass():
    """Lỗi 2: Runner ở chế độ Multi-Tab không từ chối session thuộc các tab khác trong cùng connection."""
    from unittest.mock import MagicMock
    from app.sap_attachment import MigoAttachmentRunner

    exp_identity = {
        "system": "PRD",
        "client": "100",
        "user": "ROBOT",
        "session_id": "/app/con[0]/ses[0]",
        "session_index": 0,
        "connection_index": 0
    }
    runner = MigoAttachmentRunner(expected_session_identity=exp_identity)

    # Session tab 4 (ses[3])
    tab4_session = MagicMock()
    tab4_session.Id = "/app/con[0]/ses[3]"
    tab4_session.Busy = False
    tab4_info = MagicMock()
    tab4_info.SystemName = "PRD"
    tab4_info.Client = "100"
    tab4_info.User = "ROBOT"
    tab4_info.SessionNumber = 3
    tab4_session.Info = tab4_info

    # 1. Khi chưa bật multi_tab: Bị từ chối do session_index mismatch
    ok, reason = runner.verify_session_identity(tab4_session)
    assert ok is False
    assert reason == "SESSION_IDENTITY_MISMATCH"

    # 2. Khi bật multi_tab (_multi_tab_mode = True): Chấp nhận tab 4 thuộc cùng user/client/connection
    runner._multi_tab_mode = True
    ok_mt, reason_mt = runner.verify_session_identity(tab4_session)
    assert ok_mt is True
    assert reason_mt == "SESSION_VERIFIED"


def test_handle_migo_headtext_underscore_ticket():
    """Lỗi 2.2: Số phiếu chứa ký tự gạch dưới (VD: IN_6058) phải khớp đúng với trường Header Text."""
    import re

    ticket = "IN_6058"
    headtext1 = "BBBG - IN_6058 - Ngay 26.10"
    headtext2 = "BBBG - IN_60580 - Sai so"

    # Regex mới: (?<![A-Za-z0-9]){ticket}(?![A-Za-z0-9])
    pattern = rf"(?<![A-Za-z0-9]){re.escape(ticket)}(?![A-Za-z0-9])"

    assert bool(re.search(pattern, headtext1)) is True
    assert bool(re.search(pattern, headtext2)) is False
