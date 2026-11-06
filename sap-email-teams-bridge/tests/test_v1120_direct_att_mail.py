# -*- coding: utf-8 -*-
"""Comprehensive verification suite for v1.1.20:
Direct folder-to-MIGO attachment workflow without GOS pre-check, ZFIR266 skip gating,
or strict BBBG vs VAT classification.
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
from app.sap_attachment import (
    MigoAttachmentRunner,
    extract_search_tokens,
    deduplicate_files_by_content
)


class TestV1120DirectAttMail:
    """Kiểm thử chuyên sâu cho kiến trúc đính kèm trực tiếp v1.1.20:
    1. Quét tìm tất cả các file trong thư mục khớp với số phiếu / số chứng từ (doc).
    2. Không phân biệt BBBG hay VAT, gom toàn bộ file khớp để đính kèm lên SAP MIGO.
    3. Không mở GOS attachment list popup hay ZFIR266 để skip file / skip DOC.
    4. Xử lý thành công (DONE) với 1 file duy nhất, 2 file, hoặc nhiều file (có số phụ (1), (2)...).
    """

    def test_match_ticket_files_matches_ticket_and_doc(self):
        """Khớp tất cả các file có số phiếu hoặc số chứng từ bất kể ký hiệu BBBG, VAT hay custom name."""
        index_data = {
            "folder": "G:/mock_folder",
            "files": [
                {"path": Path("G:/mock_folder/5046 BB.pdf"), "name": "5046 BB.pdf", "compact": "5046bb", "nums": {"5046"}},
                {"path": Path("G:/mock_folder/VAT 5046(1).pdf"), "name": "VAT 5046(1).pdf", "compact": "vat50461", "nums": {"5046"}},
                {"path": Path("G:/mock_folder/5047 BB.PDF"), "name": "5047 BB.PDF", "compact": "5047bb", "nums": {"5047"}},
                {"path": Path("G:/mock_folder/VAT 5047 +5050.PDF"), "name": "VAT 5047 +5050.PDF", "compact": "vat50475050", "nums": {"5047", "5050"}},
                {"path": Path("G:/mock_folder/5006815109_GiaoNhan.pdf"), "name": "5006815109_GiaoNhan.pdf", "compact": "5006815109giaonhan", "nums": {"5006815109"}},
                {"path": Path("G:/mock_folder/Unrelated_9999.pdf"), "name": "Unrelated_9999.pdf", "compact": "unrelated9999", "nums": {"9999"}},
            ],
            "token_map": {
                "5046": [0, 1],
                "5047": [2, 3],
                "5050": [3],
                "5006815109": [4],
                "9999": [5]
            }
        }

        # 1. Tìm theo số phiếu 5046
        matches_5046 = MigoAttachmentRunner.match_ticket_files("5006815109", "TamHop-in 5046", index_data)
        names_5046 = [p.name for p in matches_5046]
        assert "5046 BB.pdf" in names_5046
        assert "VAT 5046(1).pdf" in names_5046
        assert "5006815109_GiaoNhan.pdf" in names_5046  # Khớp qua số DOC
        assert "Unrelated_9999.pdf" not in names_5046

        # 2. Tìm theo số phiếu 5047
        matches_5047 = MigoAttachmentRunner.match_ticket_files("5006816681", "5047", index_data)
        names_5047 = [p.name for p in matches_5047]
        assert "5047 BB.PDF" in names_5047
        assert "VAT 5047 +5050.PDF" in names_5047

        # 3. Tìm theo số phiếu 5050 (gộp trong tên file VAT 5047 +5050.PDF)
        matches_5050 = MigoAttachmentRunner.match_ticket_files("5006816690", "5050", index_data)
        names_5050 = [p.name for p in matches_5050]
        assert "VAT 5047 +5050.PDF" in names_5050

    def test_process_single_migo_item_attaches_all_files_direct_without_gos_check(self):
        """Khi folder có 2 file (BBBG và VAT đổi tên (1)), bot đính kèm trực tiếp cả 2 file mà không skip."""
        runner = MigoAttachmentRunner()
        runner.expected_session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        runner.verify_session_identity = MagicMock(return_value=(True, "OK"))

        mock_session = MagicMock()
        mock_attach = MagicMock(return_value=(True, "OK"))
        runner.attach_file_to_current_migo = mock_attach
        runner.open_migo_document = MagicMock(return_value=(True, "OK"))
        runner.handle_migo_headtext = MagicMock(return_value=("HEADTEXT_DONE", "OK"))
        # Giả lập read_sap_attachment_titles sẽ KHÔNG được gọi khi có file thư mục
        runner.read_sap_attachment_titles = MagicMock(side_effect=AssertionError("Không được gọi GOS popup khi đã có file!"))

        mock_f1 = Path("G:/mock_dir/5046 BB.pdf")
        mock_f2 = Path("G:/mock_dir/VAT 5046(1).pdf")

        index_data = {
            "folder": "G:/mock_dir",
            "files": [
                {"path": mock_f1, "name": mock_f1.name, "compact": "5046bb", "nums": {"5046"}},
                {"path": mock_f2, "name": mock_f2.name, "compact": "vat50461", "nums": {"5046"}},
            ],
            "token_map": {"5046": [0, 1]}
        }

        item = {"doc": "5006815109", "ticket_no": "5046", "year": "2026"}
        res, succ = runner.process_single_migo_item(
            session=mock_session,
            item=item,
            index_data=index_data,
            fill_headtext=True,
            year="2026",
            is_folder_requested=True,
            has_folder=True,
            dry_run=False
        )

        assert succ is True
        assert res["overall_status"] == "DONE"
        assert res["is_succ"] is True
        assert res["headtext_status"] == "HEADTEXT_DONE"
        assert mock_attach.call_count == 2
        attached_names = [call[0][1].name for call in mock_attach.call_args_list]
        assert "5046 BB.pdf" in attached_names
        assert "VAT 5046(1).pdf" in attached_names

    def test_process_single_migo_item_succeeds_with_only_single_file(self):
        """Khi chỉ có 1 file duy nhất (không có VAT), bot đính kèm file đó và báo DONE (không báo PARTIAL)."""
        runner = MigoAttachmentRunner()
        runner.expected_session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        runner.verify_session_identity = MagicMock(return_value=(True, "OK"))

        mock_session = MagicMock()
        mock_attach = MagicMock(return_value=(True, "OK"))
        runner.attach_file_to_current_migo = mock_attach
        runner.open_migo_document = MagicMock(return_value=(True, "OK"))
        runner.handle_migo_headtext = MagicMock(return_value=("HEADTEXT_DONE", "OK"))

        mock_f1 = Path("G:/mock_dir/5046 BB.pdf")
        index_data = {
            "folder": "G:/mock_dir",
            "files": [
                {"path": mock_f1, "name": mock_f1.name, "compact": "5046bb", "nums": {"5046"}}
            ],
            "token_map": {"5046": [0]}
        }

        item = {"doc": "5006815109", "ticket_no": "5046", "year": "2026"}
        res, succ = runner.process_single_migo_item(
            session=mock_session,
            item=item,
            index_data=index_data,
            fill_headtext=True,
            year="2026",
            is_folder_requested=True,
            has_folder=True,
            dry_run=False
        )

        assert succ is True
        assert res["overall_status"] == "DONE"
        assert res["bbbg_status"] == "DONE"
        assert res["vat_status"] == "DONE"
        assert mock_attach.call_count == 1
        assert mock_attach.call_args[0][1].name == "5046 BB.pdf"

    def test_phase1_does_not_skip_doc_when_folder_is_present(self):
        """Khi có thư mục đính kèm, Phase 1 ZFIR266 KHÔNG BAO GIỜ gạch tên/skip DOC khỏi danh sách MIGO."""
        runner = MigoAttachmentRunner()
        runner.expected_session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        mock_sap = MagicMock()
        mock_sap.session_identity = runner.expected_session_identity
        mock_sap.ensure_session.return_value = MagicMock()
        # ZFIR266 báo đã có đủ 100% BBBG và VAT
        mock_sap.check_attachments_zfir266.return_value = [{
            "doc": "5006815109",
            "found_in_zfir": True,
            "need_bbbg": False,
            "need_vat": False,
            "need_header_text": False,
            "sap_titles": ["5046 BB.pdf", "VAT 5046.pdf"]
        }]
        runner.sap_engine = mock_sap
        runner.verify_session_identity = MagicMock(return_value=(True, "OK"))

        mock_f1 = Path("G:/mock_dir/5046 BB.pdf")
        mock_f2 = Path("G:/mock_dir/VAT 5046(1).pdf")
        index_mock = {
            "total_files": 2,
            "files": [
                {"path": mock_f1, "name": mock_f1.name, "compact": "5046bb", "nums": {"5046"}},
                {"path": mock_f2, "name": mock_f2.name, "compact": "vat50461", "nums": {"5046"}},
            ],
            "token_map": {"5046": [0, 1]}
        }

        with patch.object(runner, "index_source_folder", return_value=index_mock), \
             patch.object(runner, "open_migo_document", return_value=(True, "OK")), \
             patch.object(runner, "handle_migo_headtext", return_value=("HEADTEXT_DONE", "OK")), \
             patch.object(runner, "attach_file_to_current_migo", return_value=(True, "OK")) as mock_attach:

            summary = runner.process_attachment_batch(
                folder_path=str(Path(__file__).parent),
                items=[{"doc": "5006815109", "ticket_no": "5046"}],
                year="2026",
                dry_run=False
            )

            # DOC không được bỏ qua ở Phase 1 mà phải được nạp vào MIGO
            assert summary["total_items"] == 1
            assert summary["total_success"] == 1
            item_res = summary["results"][0]
            assert item_res["overall_status"] == "DONE"
            # Cả 2 file trong folder đều được attach
            assert mock_attach.call_count == 2
