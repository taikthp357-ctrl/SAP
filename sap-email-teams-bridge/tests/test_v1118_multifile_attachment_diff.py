# -*- coding: utf-8 -*-
"""Unit tests for v1.1.18:
Multi-file Attachment Smart Diffing, Comma-separated ZFIR266 parsing,
and Anti-Short-Circuit Gating for files with suffix like (1), _Dot2.
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
from app.sap_attachment import (
    classify_sap_attachments,
    match_files_to_sap,
    MigoAttachmentRunner
)


class TestV1118MultiFileAttachmentDiff:
    """Kiểm thử bộ tính năng thông minh v1.1.18:
    1. Bóc tách chuỗi gộp phân cách bằng dấu phẩy trong ZFIR266.
    2. Nhận diện và đính kèm bổ sung file mang số phụ (1), _Dot2 kể cả khi SAP đã có 1 file VAT.
    3. Ngăn chặn triệt để hiện tượng ép to_upload = [] tại Phase 2.
    """

    def test_classify_sap_attachments_splits_comma_and_semicolon(self):
        """classify_sap_attachments phải tự động bóc tách chuỗi gộp có dấu phẩy/chấm phẩy."""
        raw_titles = ["5047 BB.PDF,VAT 5047 +5050.PDF;Extra_Doc.pdf"]
        res = classify_sap_attachments(raw_titles, ticket_no="5047")
        assert res["has_bbbg"] is True
        assert res["has_vat"] is True
        assert res["total_files"] == 3
        assert "5047 BB.PDF" in res["files"]
        assert "VAT 5047 +5050.PDF" in res["files"]
        assert "Extra_Doc.pdf" in res["files"]

    def test_match_files_to_sap_detects_vat_5046_1_when_sap_has_vat_5046(self):
        """Khi local có VAT 5046(1).pdf và SAP đã có VAT 5046.pdf:
        VAT 5046(1).pdf phải nằm trong to_upload để được nạp bổ sung.
        """
        local_vat = [Path("VAT 5046(1).pdf")]
        sap_titles = ["5046 BB.PDF", "VAT 5046.PDF"]
        to_up, already, uncl = match_files_to_sap(
            local_vat, sap_titles, ticket_nums=["5046"], file_type="vat"
        )
        assert len(to_up) == 1
        assert to_up[0].name == "VAT 5046(1).pdf"
        assert len(already) == 0

    def test_match_files_to_sap_both_original_and_numbered_vat(self):
        """Khi local có cả VAT 5046.pdf và VAT 5046(1).pdf, SAP mới có VAT 5046.pdf:
        VAT 5046.pdf được nhận diện đã có, VAT 5046(1).pdf được nạp bổ sung.
        """
        local_vat = [Path("VAT 5046.pdf"), Path("VAT 5046(1).pdf")]
        sap_titles = ["5046 BB.PDF", "VAT 5046.PDF"]
        to_up, already, uncl = match_files_to_sap(
            local_vat, sap_titles, ticket_nums=["5046"], file_type="vat"
        )
        assert len(to_up) == 1
        assert to_up[0].name == "VAT 5046(1).pdf"
        assert len(already) == 1
        assert already[0][0].name == "VAT 5046.pdf"

    def test_smart_gating_phase1_updates_need_vat_flag_when_extra_file_found(self):
        """Phase 1 Smart Gating: Khi ZFIR266 báo need_vat=False nhưng local có file (1),
        cờ need_vat trong item phải được cập nhật thành True và can_skip_migo=False.
        """
        runner = MigoAttachmentRunner()
        runner.expected_session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        mock_sap = MagicMock()
        mock_sap.session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        mock_sap.check_attachments_zfir266.return_value = [{
            "doc": "5006815109",
            "found_in_zfir": True,
            "need_bbbg": False,
            "need_vat": False,  # ZFIR266 nghĩ là đã có VAT
            "need_header_text": False,
            "sap_titles": ["5046 BB.PDF,VAT 5046.PDF"],  # Chuỗi gộp từ ALV
            "bbbg_file": "5046 BB.PDF",
            "vat_file": "VAT 5046.PDF"
        }]
        mock_sap.ensure_session.return_value = MagicMock()
        runner.sap_engine = mock_sap
        runner.verify_session_identity = MagicMock(return_value=(True, "OK"))

        mock_bb = Path("G:/test/5046 BB.pdf")
        mock_vat1 = Path("G:/test/VAT 5046.pdf")
        mock_vat2 = Path("G:/test/VAT 5046(1).pdf")
        existing_folder = str(Path(__file__).parent)

        with patch.object(runner, "index_source_folder", return_value={
            "total_files": 3,
            "files": [mock_bb, mock_vat1, mock_vat2],
            "token_map": {}
        }), patch.object(runner, "match_all_files", return_value={
            "bbbg": [mock_bb],
            "vat": [mock_vat1, mock_vat2],
            "mail": []
        }), patch.object(runner, "open_migo_document", return_value=(True, "OK")), \
           patch.object(runner, "attach_file_to_current_migo", return_value=(True, "OK")):

            res = runner.process_attachment_batch(
                folder_path=existing_folder,
                items=[{"doc": "5006815109", "ticket_no": "5046"}],
                year="2026",
                dry_run=False
            )
            # Chứng từ không được bỏ qua mà phải được xử lý upload file VAT (1)
            assert res["total_items"] == 1
            item_res = res["results"][0]
            assert item_res["doc"] == "5006815109"
            assert item_res["is_succ"] is True
            assert runner.attach_file_to_current_migo.called
            attached_files = [call[0][1].name for call in runner.attach_file_to_current_migo.call_args_list]
            assert "VAT 5046(1).pdf" in attached_files

    def test_phase2_does_not_short_circuit_vat_when_vat_paths_present(self):
        """Phase 2: Dù item có need_vat=False từ trước, nếu vat_paths có file chưa có trên SAP,
        Phase 2 vẫn phải đối chiếu qua match_files_to_sap và upload file thiếu.
        """
        runner = MigoAttachmentRunner()
        runner.expected_session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        runner.verify_session_identity = MagicMock(return_value=(True, "OK"))

        mock_session = MagicMock()
        mock_attach = MagicMock(return_value=(True, "OK"))
        runner.attach_file_to_current_migo = mock_attach
        runner.open_migo_document = MagicMock(return_value=(True, "OK"))

        mock_vat_extra = Path("G:/test/VAT 5046(1).pdf")
        item = {
            "doc": "5006815109",
            "ticket_no": "5046",
            "year": "2026",
            "need_bbbg": False,
            "need_vat": False,  # Giả sử cờ là False từ trước
            "sap_titles": ["5046 BB.PDF", "VAT 5046.PDF"]
        }

        with patch.object(runner, "match_all_files", return_value={
            "bbbg": [],
            "vat": [mock_vat_extra],
            "mail": []
        }):
            item_res, succ = runner.process_single_migo_item(
                session=mock_session,
                item=item,
                index_data={"files": [mock_vat_extra], "token_map": {}},
                fill_headtext=False,
                year="2026",
                is_folder_requested=True,
                has_folder=True,
                dry_run=False
            )

            # File VAT 5046(1).pdf phải được gọi attach
            assert mock_attach.called
            assert mock_attach.call_args[0][1] == mock_vat_extra
            assert succ is True
            assert item_res["overall_status"] == "DONE"
