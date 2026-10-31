# -*- coding: utf-8 -*-
"""Unit tests cho AttHandler nhánh MIGO_HEADTEXT chuyên trách."""
import pytest
from unittest.mock import MagicMock, patch
from app.core.context import RequestContext
from app.handlers.att_handler import AttHandler


def test_att_handler_migo_headtext_success():
    """Kiểm tra AttHandler xử lý lệnh MIGO_HEADTEXT gọi đúng MigoHeadtextWriter mà không mở ZFIR266."""
    handler = AttHandler()
    ctx = RequestContext(
        request_id="REQ-TEST-HEADTEXT-01",
        command="MIGO_HEADTEXT",
        meta={
            "dryRun": False,
            "attachmentItems": [
                {"doc": "5007145690", "ticket_no": "6381"}
            ],
            "year": "2026"
        },
        sap=MagicMock()
    )

    with patch("app.handlers.att_handler.assert_write_permission"), \
         patch("app.writers.migo_headtext_writer.MigoHeadtextWriter.fill", return_value=(True, "FILLED", "Lưu thành công")) as mock_fill:

        res = handler.handle(ctx)
        assert res.status == "SUCCEEDED"
        assert len(res.records) == 1
        assert res.records[0]["doc"] == "5007145690"
        assert res.records[0]["ticket_no"] == "6381"
        assert res.records[0]["headtext_status"] == "FILLED"
        assert mock_fill.called
        assert mock_fill.call_args[1]["ticket_no"] == "6381"


def test_att_handler_migo_headtext_preview():
    """Kiểm tra dryRun/preview không gọi lệnh ghi thật."""
    handler = AttHandler()
    ctx = RequestContext(
        request_id="REQ-TEST-HEADTEXT-02",
        command="MIGO_HEADTEXT",
        meta={
            "dryRun": True,
            "attachmentDocs": ["5007145690"],
            "ticket_no": "6381"
        },
        sap=MagicMock()
    )

    with patch("app.writers.migo_headtext_writer.MigoHeadtextWriter.fill") as mock_fill:
        res = handler.handle(ctx)
        assert res.status == "SUCCEEDED"
        assert len(res.records) == 1
        assert not mock_fill.called
