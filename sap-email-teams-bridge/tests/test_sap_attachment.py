# -*- coding: utf-8 -*-
"""Unit tests cho module MigoAttachmentRunner tự động đính kèm BBBG & VAT vào SAP MIGO."""

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
    BBBG_REGEX,
    VAT_REGEX
)


@pytest.fixture
def sample_attachment_dir():
    """Tạo thư mục tạm chứa các file mẫu BBBG và VAT."""
    temp_dir = tempfile.mkdtemp(prefix="test_migo_att_")
    base = Path(temp_dir)

    # Tạo các file mẫu
    (base / "in 5794_BBBG.pdf").write_text("sample content bbbg 5794")
    (base / "in 5794_VAT.pdf").write_text("sample content vat 5794")
    
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


def test_index_source_folder(sample_attachment_dir):
    """Kiểm tra lập chỉ mục file trong thư mục nguồn và thư mục con."""
    runner = MigoAttachmentRunner()
    index = runner.index_source_folder(str(sample_attachment_dir), recursive=True)

    assert index["total_files"] == 5
    assert "5794" in index["token_map"]
    assert "5799" in index["token_map"]
    assert "5007072362" in index["token_map"]


def test_match_files_exact(sample_attachment_dir):
    """Kiểm tra khớp file chính xác theo số phiếu và loại file."""
    runner = MigoAttachmentRunner()
    index = runner.index_source_folder(str(sample_attachment_dir), recursive=True)

    # Test khớp 5794
    b_path, v_path, notes = runner.match_files("5006995954", "Thành Công in 5794", index)
    assert b_path is not None
    assert "5794_BBBG.pdf" in b_path.name
    assert v_path is not None
    assert "5794_VAT.pdf" in v_path.name

    # Test khớp 5799 trong thư mục con
    b_path2, v_path2, _ = runner.match_files("5006995500", "Thanhcong in 5799", index)
    assert b_path2 is not None
    assert "5799_BBGH.pdf" in b_path2.name
    assert v_path2 is not None
    assert "5799_HoaDon.pdf" in v_path2.name

    # Test thiếu file BBBG nhưng có file VAT (như case 5007072362)
    b_path3, v_path3, notes3 = runner.match_files("5007072362", "Thành Đạt in 6157", index)
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

    summary = runner.process_attachment_batch(str(sample_attachment_dir), items, dry_run=True)
    assert summary["total_items"] == 4
    assert summary["total_success"] == 3  # 3 items khớp được ít nhất 1 file
    assert summary["total_missing"] == 1  # 1 item hoàn toàn không có file

    report = runner.format_teams_report(summary)
    assert "`5006995954`: Thành Công in 5794" in report
    assert "`9999999999`: FILE_NOT_FOUND" in report
    assert "| DOC | Số phiếu | BBBG | VAT |" in report


def test_sap_session_gos_mock():
    """Kiểm tra tương tác mock SAP GUI với GOS."""
    mock_session = MagicMock()
    # Mock sbar
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
