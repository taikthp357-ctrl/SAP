# -*- coding: utf-8 -*-
"""Unit tests for v1.0.92 features:
1. TRA_CUU_CAP_NVL intent recognition with LIMO model and 'cấp giúp em' / 'check đơn cấp'
2. MaterialResolver: extraction of workshop prefix BIW + core 8 digits
3. Sửa số phiếu: allow_overwrite_headtext parsing and execution in handle_migo_headtext
4. SapWhoPrinterService: smart classification of WHO (6*) vs OB (101*)
"""

import pytest
from unittest.mock import MagicMock, patch

from app.email_parser import SmartCommandParser
from app.material_resolver import MaterialResolver
from app.sap_who_printer import SapWhoPrinterService
from app.sap_attachment import MigoAttachmentRunner


class TestV1092Features:

    def test_intent_check_don_cap_limo(self):
        """Kiểm tra lệnh 'SAP check đơn cấp limo Cấp giúp em mã ...' phải vào TRA_CUU_CAP_NVL và model LIMO."""
        raw_text = (
            "SAP check đơn cấp limo Cấp giúp em mã\n"
            "Biw32022789\n"
            "Biw32022628"
        )
        meta = SmartCommandParser.parse(raw_text)
        assert meta["command"] == "TRA_CUU_CAP_NVL", f"Expected TRA_CUU_CAP_NVL but got {meta['command']}"
        assert meta["model"] == "LIMO", f"Expected model LIMO but got {meta['model']}"
        assert any("32022789" in m for m in meta["materials"]), "Expected 32022789 in materials"
        assert any("32022628" in m for m in meta["materials"]), "Expected 32022628 in materials"

    def test_material_resolver_biw_prefix(self):
        """Kiểm tra bóc tách tiền tố BIW + 8 số gốc để vừa có mã full vừa có core cho Excel/SAP."""
        resolver = MaterialResolver()
        codes, note = resolver.resolve("Biw32022789")
        assert "32022789" in codes, "Core 8-digit must be present for Excel search"
        assert "BIW32022789AA" in codes, "Full inferred code must be present"
        assert "BIW32022789*" in codes, "Wildcard code must be present for SAP search"

    def test_overwrite_headtext_intent_parsing(self):
        """Kiểm tra lệnh 'sửa số phiếu' bật cờ allow_overwrite_headtext=True, còn 'điền số phiếu' thì False."""
        cmd_sua = "SAP sửa số phiếu\n5007135730 LTV - IN 6458"
        meta_sua = SmartCommandParser.parse(cmd_sua)
        assert meta_sua["command"] in ("MIGO_HEADTEXT", "ATT_MIGO")
        assert meta_sua.get("allow_overwrite_headtext") is True

        cmd_dien = "SAP điền số phiếu\n5007135730 LTV - IN 6458"
        meta_dien = SmartCommandParser.parse(cmd_dien)
        assert meta_dien["command"] in ("MIGO_HEADTEXT", "ATT_MIGO")
        assert meta_dien.get("allow_overwrite_headtext") is False

    def test_handle_migo_headtext_safety_vs_overwrite(self):
        """Kiểm tra: allow_overwrite=False thì báo REVIEW, allow_overwrite=True thì ghi đè và lưu."""
        runner = MigoAttachmentRunner()
        runner._active_document = "5007135730"

        session = MagicMock()
        head_field = MagicMock()
        head_field.Text = "OLD_IN_6457"
        head_field.text = "OLD_IN_6457"

        w0 = MagicMock()
        w0.findByName = MagicMock(return_value=head_field)
        sbar = MagicMock()
        sbar.MessageType = "S"
        sbar.Text = "Saved"

        def mock_safe_find(sess, p):
            if p == "wnd[0]":
                return w0
            if p == "wnd[0]/sbar":
                return sbar
            return None

        with patch.object(runner, "verify_session_identity", return_value=(True, "OK")), \
             patch("app.sap_attachment.safe_find_by_id", side_effect=mock_safe_find), \
             patch("app.sap_attachment.smart_find", return_value=head_field), \
             patch("app.sap_attachment.sap_wait_idle"), \
             patch("app.sap_attachment.dismiss_any_modal_popup"):

            # 1. Trường hợp bình thường (allow_overwrite=False): Báo HEADTEXT_REVIEW
            res_safe, msg_safe = runner.handle_migo_headtext(session, "LTV - IN 6458", allow_overwrite=False)
            assert res_safe == "HEADTEXT_REVIEW"
            assert "Cần Review thủ công" in msg_safe

            # 2. Trường hợp người dùng yêu cầu sửa (allow_overwrite=True): Ghi đè và báo HEADTEXT_DONE
            res_over, msg_over = runner.handle_migo_headtext(session, "LTV - IN 6458", allow_overwrite=True)
            assert res_over == "HEADTEXT_DONE"
            assert "Đã ghi đè số phiếu cũ" in msg_over
            assert head_field.Text == "LTV - IN 6458" or head_field.text == "LTV - IN 6458"
            w0.SendVKey.assert_called_with(11)

    def test_sap_who_printer_classification(self):
        """Kiểm tra nhận diện đầu số: 6* -> WHO, 101* -> OB."""
        assert SapWhoPrinterService.classify_document("6028447808") == "WHO"
        assert SapWhoPrinterService.classify_document("6028465380") == "WHO"
        assert SapWhoPrinterService.classify_document("1013144096") == "OB"
        assert SapWhoPrinterService.classify_document("1012957578") == "OB"
        assert SapWhoPrinterService.classify_document("32022789") == "UNKNOWN"
