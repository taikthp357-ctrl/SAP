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
    assert summary["total_success"] == 3  # 3 items khớp được ít nhất 1 file
    assert summary["total_missing"] == 1  # 1 item hoàn toàn không có file

    report = runner.format_teams_report(summary)
    assert "`5006995954`: Thành Công in 5794" in report
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

    runner = MigoAttachmentRunner()
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
    def find_by_id(path):
        if "shell" in path:
            return mock_shell
        elif "sbar" in path:
            return mock_sbar
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    runner = MigoAttachmentRunner()
    with tempfile.NamedTemporaryFile(suffix="_VAT.pdf", delete=False) as tf:
        tf.write(b"dummy pdf")
        dummy_file = Path(tf.name)

    try:
        ok, msg = runner.attach_file_to_current_migo(mock_session, dummy_file)
        assert ok is True
        assert "DONE_SAP_ACK" in msg
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
    assert batch_res["total_success"] == 1
    assert batch_res["total_missing"] == 0
    res0 = batch_res["results"][0]
    assert res0["doc"] == "4995880054"
    assert res0["ticket_no"] == "OUT_P1_23581"
    assert res0["headtext_status"] == "HEADTEXT_DONE"
    assert res0["bbbg_status"] == "SKIP_NO_ATTACH"
    assert "Thành Công OUT_P1_23581" in res0["overall_status"]


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


def test_format_status_label_no_duplication():
    """Kiểm tra format_status_label không bị lặp từ 'Thành Công'."""
    assert format_status_label("Thành Công in 3874", True) == "Thành Công in 3874"
    assert format_status_label("thanh cong 1234", True) == "thanh cong 1234"
    assert format_status_label("in 3874", True) == "Thành Công in 3874"
    assert format_status_label("", True) == "Thành Công"
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
    assert "Lỗi link thư mục" in r0["overall_status"]

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
    from unittest.mock import MagicMock
    mock_sap = MagicMock()
    mock_session = MagicMock()
    mock_sap.ensure_session.return_value = mock_session

    runner = MigoAttachmentRunner(sap_engine=mock_sap)
    
    # Mock read_sap_attachment_titles và open_migo_document
    runner.open_migo_document = lambda session, doc, year=None, mode="DISPLAY": (True, "OK")
    runner.read_sap_attachment_titles = lambda session: ["3874 BB"]
    
    # Giả lập thư mục không có file nào (để chứng minh không upload lại từ thư mục)
    import tempfile
    with tempfile.TemporaryDirectory() as tmp_dir:
        res = runner.process_attachment_batch(
            folder_path=tmp_dir,
            items=[{"doc": "5006458947", "ticket_no": "Thành Công in 3874"}],
            dry_run=False
        )
        assert len(res["results"]) == 1
        item = res["results"][0]
        # BBBG đã có trên SAP -> đánh dấu Thành Công in 3874
        assert item["bbbg_status"] == "Thành Công in 3874"
        # VAT chưa có trên SAP và thư mục không có -> FILE_NOT_FOUND
        assert item["vat_status"] == "FILE_NOT_FOUND"
        # Overall status thành công vì đã có BBBG hợp lệ
        assert item["overall_status"] == "Thành Công in 3874"
        assert res["total_success"] == 1




