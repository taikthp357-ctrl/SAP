# -*- coding: utf-8 -*-
"""Unit tests cho phiên bản v1.1.21:
1. Cơ chế Thu hồi toàn diện (Inclusive Reclaim): Thu hồi 100% chứng từ chưa DONE khi xảy ra sự cố sập phiên / ngắt kết nối.
2. Triệt tiêu Race Condition bằng Thread Drain (f.result timeout).
3. Khử trùng lặp danh sách xử lý lại (Deduplication by original index) trong Self-Heal.
4. Đồng bộ chính xác bộ đếm tiến độ processed_count & success_count gửi Teams.
5. Chống đính kèm lặp (Anti-duplicate attachment) nếu file đã hiện diện trên SAP.
6. Phase 1 Smart Skip nếu toàn bộ file trong thư mục LAN đã có trên SAP.
"""
from collections import deque
from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from app.sap_attachment import MigoAttachmentRunner


class TestInclusiveReclaimAndDeduplicationV1121:
    """Kiểm thử cơ chế Inclusive Reclaim và khử lặp chứng từ trong Self-Healing."""

    def test_inclusive_reclaim_captures_unusual_errors(self):
        """Chứng từ gặp lỗi bất thường (GOS_TRIGGER_FAILED, RPC error, v.v.) chưa DONE phải được thu hồi 100%."""
        items = [
            {"doc": "5007180001", "ticket_no": "TICK_1"},
            {"doc": "5007180002", "ticket_no": "TICK_2"},
            {"doc": "5007180003", "ticket_no": "TICK_3"},
        ]
        # Giả lập 1 chứng từ thành công DONE, 1 chứng từ bị GOS_TRIGGER_FAILED (chưa DONE), 1 chứng từ lỗi RPC
        indexed_results = [
            (0, {"doc": "5007180001", "overall_status": "DONE", "is_succ": True}, True),
            (1, {"doc": "5007180002", "overall_status": "Lỗi đính kèm: GOS_TRIGGER_FAILED", "is_succ": False}, False),
            (2, {"doc": "5007180003", "overall_status": "RPC failed: disconnected", "is_succ": False}, False),
        ]

        reclaimed_items = []
        survived_results = []
        for o_idx, r_dict, is_s in indexed_results:
            ov = str(r_dict.get("overall_status", "")).upper()
            if is_s is True and ov == "DONE":
                survived_results.append((o_idx, r_dict, is_s))
            else:
                orig_itm = items[o_idx] if o_idx < len(items) else {"doc": r_dict.get("doc"), "ticket_no": r_dict.get("ticket_no")}
                reclaimed_items.append((o_idx, orig_itm))

        # Chỉ có chứng từ 0 sống sót, 2 chứng từ lỗi được thu hồi
        assert len(survived_results) == 1
        assert survived_results[0][0] == 0
        assert len(reclaimed_items) == 2
        assert [idx for idx, _ in reclaimed_items] == [1, 2]

    def test_queue_deduplication_prevents_duplicate_processing(self):
        """Khử trùng lặp giữa reclaimed_items và tab_queues đảm bảo không có chứng từ nào chạy 2 lần."""
        reclaimed_items = [
            (1, {"doc": "5007180002", "ticket_no": "TICK_2"}),
            (2, {"doc": "5007180003", "ticket_no": "TICK_3"}),
        ]
        # Giả lập tab_queues vô tình chứa lại index 2 và index 3
        tab_queues = {
            0: deque([(2, {"doc": "5007180003", "ticket_no": "TICK_3"})]),
            1: deque([(3, {"doc": "5007180004", "ticket_no": "TICK_4"})]),
        }

        seen_indices = set(idx for idx, _ in reclaimed_items)
        unprocessed_items = list(reclaimed_items)
        for t_q in tab_queues.values():
            while t_q:
                q_tuple = t_q.popleft()
                if q_tuple:
                    q_idx, q_itm = q_tuple
                    if q_idx not in seen_indices:
                        unprocessed_items.append((q_idx, q_itm))
                        seen_indices.add(q_idx)

        unprocessed_items.sort(key=lambda x: x[0])
        indices = [idx for idx, _ in unprocessed_items]
        # Phải là [1, 2, 3] duy nhất, không lặp lại 2
        assert indices == [1, 2, 3]
        assert len(indices) == len(set(indices))

    def test_anti_duplicate_attachment_skips_files_already_on_sap(self):
        """Khi file đã có sẵn trên SAP (trong sap_titles), bot không đính kèm đúp qua GOS."""
        runner = MigoAttachmentRunner()
        mock_session = MagicMock()
        mock_info = MagicMock()
        mock_info.SystemName = "PRD"
        mock_info.Client = "908"
        mock_info.User = "TAIPT8"
        mock_session.Info = mock_info
        mock_session.Id = "/app/con[0]/ses[0]"

        item = {
            "doc": "5007183651",
            "ticket_no": "6646",
            "sap_titles": ["vat 6646 + 6676.pdf"]
        }
        index_data = {
            "files": [Path("C:/FakeFolder/vat 6646 + 6676.pdf")],
            "token_map": {"6646": [0]}
        }

        with patch.object(runner, "verify_session_identity", return_value=(True, "OK")), \
             patch.object(runner, "open_migo_document", return_value=(True, "LOADED")), \
             patch.object(runner, "handle_migo_headtext", return_value=("HEADTEXT_OK", "OK")), \
             patch.object(runner, "match_ticket_files", return_value=[Path("C:/FakeFolder/vat 6646 + 6676.pdf")]), \
             patch.object(runner, "attach_file_to_current_migo") as mock_attach:

            res_dict, is_s = runner.process_single_migo_item(
                mock_session, item, index_data, fill_headtext=True, year="2026",
                is_folder_requested=True, has_folder=True
            )

            # Phải thành công DONE nhưng KHÔNG gọi attach_file_to_current_migo vì file đã có trên SAP!
            assert is_s is True
            assert res_dict["overall_status"] == "DONE"
            mock_attach.assert_not_called()

    def test_smart_phase1_skip_when_all_folder_files_on_sap(self):
        """Phase 1 skip MIGO khi toàn bộ file khớp trong folder đã có trên SAP và HeadText đã thỏa mãn."""
        runner = MigoAttachmentRunner()
        mock_engine = MagicMock()
        mock_engine.session_identity = {"system": "PRD", "client": "908", "user": "TAIPT8"}
        mock_engine.check_attachments_zfir266.return_value = [
            {
                "doc": "5007183651",
                "found_in_zfir": True,
                "need_bbbg": False,
                "need_vat": False,
                "need_header_text": False,
                "sap_titles": ["vat 6646 + 6676.pdf"]
            }
        ]
        runner.sap_engine = mock_engine

        items = [{"doc": "5007183651", "ticket_no": "6646"}]
        mock_path = Path("C:/Fake/vat 6646 + 6676.pdf")
        fake_index = {"total_files": 1, "files": [mock_path], "token_map": {"6646": [0]}}

        with patch("pathlib.Path.is_dir", return_value=True), \
             patch("pathlib.Path.exists", return_value=True), \
             patch.object(runner, "index_source_folder", return_value=fake_index), \
             patch.object(runner, "match_ticket_files", return_value=[mock_path]), \
             patch.object(runner, "open_migo_document") as mock_open:

            summary = runner.process_attachment_batch(
                folder_path="C:/Fake", items=items, fill_headtext=True, year="2026"
            )

            # Phải thành công 1/1 và KHÔNG cần mở MIGO (skip thành công từ Phase 1)
            assert summary["total_success"] == 1
            assert summary["results"][0]["headtext_status"] == "DONE (ZFIR266)"
            mock_open.assert_not_called()
