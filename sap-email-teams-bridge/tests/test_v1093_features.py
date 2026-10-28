# -*- coding: utf-8 -*-
"""Unit tests for v1.0.93 features:
Cơ chế SAP Keep-Alive Watchdog ('thở nhẹ') chống timeout tự động của SAP NetWeaver (rdisp/gui_auto_logout).
"""

import time
import pytest
from unittest.mock import MagicMock, patch

from app.sap_engine import SapEngine


class TestV1093KeepAliveWatchdog:

    def test_keep_alive_triggers_and_sends_slash_n_on_dedicated_session(self):
        """Kiểm tra: Khi vượt ngưỡng idle trên tab Bot riêng, tự động gửi tín hiệu 'thở nhẹ' /n."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=True
        )
        engine.is_dedicated_session = True
        engine.last_activity = time.time() - 650  # Đã idle 650s > 600s

        session = MagicMock()
        session.Busy = False
        session.Info.Transaction = "SESSION_MANAGER"
        engine.session = session

        okcd = MagicMock()
        w0 = MagicMock()

        def mock_find_safe(element_id):
            if element_id == "wnd[0]/tbar[0]/okcd":
                return okcd
            if element_id == "wnd[0]":
                return w0
            if element_id == "wnd[1]":
                return None
            return None

        with patch.object(engine, "_find_by_id_safe", side_effect=mock_find_safe), \
             patch.object(engine, "_is_session_alive", return_value=True), \
             patch.object(engine, "_wait_idle", return_value=True), \
             patch.object(engine, "_dismiss_popups", return_value=""):

            result = engine.keep_alive()
            assert result is True
            assert okcd.text == "/n"
            w0.sendVKey.assert_called_once_with(0)
            assert time.time() - engine.last_activity < 5.0

    def test_keep_alive_skips_when_below_threshold(self):
        """Kiểm tra: Khi thời gian idle chưa vượt ngưỡng, không can thiệp vào SAP."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=True
        )
        engine.is_dedicated_session = True
        engine.last_activity = time.time() - 100  # Mới idle 100s < 600s

        session = MagicMock()
        engine.session = session

        result = engine.keep_alive()
        assert result is False

    def test_keep_alive_skips_when_session_busy(self):
        """Kiểm tra: Khi SAP đang bận xử lý (Busy = True), watchdog không gửi lệnh để tránh đụng độ."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=True
        )
        engine.is_dedicated_session = True
        engine.last_activity = time.time() - 700

        session = MagicMock()
        session.Busy = True
        engine.session = session

        with patch.object(engine, "_is_session_alive", return_value=True):
            result = engine.keep_alive()
            assert result is False

    def test_keep_alive_shared_session_skips_slash_n_when_on_user_transaction(self):
        """Kiểm tra: Khi dùng chung tab với người dùng và đang ở màn hình ME21N, KHÔNG gửi /n để bảo vệ dữ liệu."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=True
        )
        engine.is_dedicated_session = False  # Tab dùng chung
        engine.last_activity = time.time() - 700

        session = MagicMock()
        session.Busy = False
        session.Info.Transaction = "ME21N"  # Người dùng đang nhập PO
        engine.session = session

        okcd = MagicMock()
        w0 = MagicMock()

        def mock_find_safe(element_id):
            if element_id == "wnd[0]/tbar[0]/okcd":
                return okcd
            if element_id == "wnd[0]":
                return w0
            return None

        with patch.object(engine, "_find_by_id_safe", side_effect=mock_find_safe), \
             patch.object(engine, "_is_session_alive", return_value=True), \
             patch.object(engine, "_check_identity", return_value={}):

            result = engine.keep_alive()
            assert result is False
            okcd.text = ""
            assert okcd.text != "/n"
            w0.sendVKey.assert_not_called()

    def test_keep_alive_shared_session_sends_slash_n_when_on_session_manager(self):
        """Kiểm tra: Tab dùng chung nhưng đang ở màn hình gốc SESSION_MANAGER thì gửi /n an toàn tuyệt đối."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=True
        )
        engine.is_dedicated_session = False
        engine.last_activity = time.time() - 700

        session = MagicMock()
        session.Busy = False
        session.Info.Transaction = "SESSION_MANAGER"
        engine.session = session

        okcd = MagicMock()
        w0 = MagicMock()

        def mock_find_safe(element_id):
            if element_id == "wnd[0]/tbar[0]/okcd":
                return okcd
            if element_id == "wnd[0]":
                return w0
            return None

        with patch.object(engine, "_find_by_id_safe", side_effect=mock_find_safe), \
             patch.object(engine, "_is_session_alive", return_value=True), \
             patch.object(engine, "_wait_idle", return_value=True), \
             patch.object(engine, "_dismiss_popups", return_value=""):

            result = engine.keep_alive()
            assert result is True
            assert okcd.text == "/n"
            w0.sendVKey.assert_called_once_with(0)

    def test_keep_alive_disabled_by_config(self):
        """Kiểm tra cờ enable_keep_alive=False tắt hoàn toàn watchdog."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=False
        )
        engine.last_activity = time.time() - 1000
        assert engine.keep_alive() is False

    def test_keep_alive_handles_exceptions_gracefully(self):
        """Kiểm tra watchdog bọc thép chống văng lỗi ra luồng chính khi SAP COM gặp sự cố."""
        engine = SapEngine(
            keep_alive_idle_seconds=600,
            enable_keep_alive=True
        )
        engine.is_dedicated_session = True
        engine.last_activity = time.time() - 700

        session = MagicMock()
        session.Busy = False
        session.Info.Transaction = "SESSION_MANAGER"
        engine.session = session

        okcd = MagicMock()
        w0 = MagicMock()
        w0.sendVKey.side_effect = RuntimeError("COM disconnected")

        def mock_find_safe(element_id):
            if element_id == "wnd[0]/tbar[0]/okcd":
                return okcd
            if element_id == "wnd[0]":
                return w0
            return None

        with patch.object(engine, "_find_by_id_safe", side_effect=mock_find_safe), \
             patch.object(engine, "_is_session_alive", return_value=False), \
             patch.object(engine, "_dismiss_popups", return_value=""):

            result = engine.keep_alive()
            assert result is False
            assert engine.session is None
