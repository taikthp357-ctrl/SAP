# -*- coding: utf-8 -*-
"""Kiểm thử tính năng Trạm gác độc lập 24/7 (SapGlobalWatchdog)."""
import time
from unittest.mock import MagicMock, patch
import pytest

from app.sap_recovery import (
    SapGlobalWatchdog,
    get_sap_process_pids,
    is_sap_gui_hung,
    start_global_sap_watchdog,
    stop_global_sap_watchdog,
    suppress_global_watchdog,
    record_global_watchdog_activity,
    recover_hung_sap_session,
    set_last_recovered_session,
)


def test_get_sap_process_pids():
    """Kiểm tra hàm bóc tách PIDs của saplogon.exe / sapfront.exe."""
    sample_tasklist = (
        '"Image Name","PID","Session Name","Session#","Mem Usage"\n'
        '"System Idle Process","0","Services","0","8 K"\n'
        '"saplogon.exe","12345","Console","1","50,000 K"\n'
        '"sapfront.exe","67890","Console","1","45,000 K"\n'
        '"notepad.exe","11111","Console","1","10,000 K"\n'
    )
    with patch("subprocess.check_output", return_value=sample_tasklist):
        pids = get_sap_process_pids()
        assert 12345 in pids
        assert 67890 in pids
        assert 11111 not in pids
        assert len(pids) == 2


def test_sap_global_watchdog_start_stop():
    """Kiểm tra vòng đời start/stop của SapGlobalWatchdog."""
    watchdog = SapGlobalWatchdog(
        connection_name="Test Connection",
        hung_threshold_sec=0.5,
        check_interval_sec=0.05,
    )
    assert not watchdog._stop_event.is_set()

    watchdog.start()
    assert watchdog._thread is not None
    assert watchdog._thread.is_alive()

    # Dừng watchdog
    watchdog.stop()
    assert watchdog._thread is None


def test_sap_global_watchdog_recovers_when_hung():
    """Kiểm tra kịch bản khi SAP bị đơ quá ngưỡng thời gian quy định -> Watchdog tự diệt & hồi sinh."""
    recovery_callback = MagicMock()
    mock_new_session = MagicMock()

    watchdog = SapGlobalWatchdog(
        connection_name="Test SSO",
        hung_threshold_sec=0.2,  # Ngưỡng ngắn 200ms để test nhanh
        check_interval_sec=0.05,
        on_recovery=recovery_callback,
    )

    with patch("app.sap_recovery.is_saplogon_running", return_value=True), \
         patch("app.sap_recovery.is_sap_gui_hung", return_value=True), \
         patch("app.sap_recovery.kill_hung_sap_processes") as mock_kill, \
         patch("app.sap_recovery.auto_connect_sso_session", return_value=mock_new_session) as mock_sso:

        watchdog.start()
        # Chờ vượt quá ngưỡng hung_threshold_sec (0.2s)
        time.sleep(0.4)
        watchdog.stop()

        # Kiểm tra kill_hung_sap_processes và auto_connect_sso_session được gọi
        assert mock_kill.called
        assert mock_sso.called
        mock_sso.assert_called_with(
            connection_name="Test SSO",
            timeout_sec=20.0,
            max_retries=3,
        )
        # Kiểm tra on_recovery callback được gọi với new_session
        recovery_callback.assert_called_with(mock_new_session)


def test_sap_global_watchdog_no_action_when_healthy():
    """Kiểm tra khi SAP hoạt động bình thường thì không bao giờ kích hoạt kill hoặc hồi sinh."""
    recovery_callback = MagicMock()

    watchdog = SapGlobalWatchdog(
        connection_name="Test SSO",
        hung_threshold_sec=0.2,
        check_interval_sec=0.05,
        on_recovery=recovery_callback,
    )

    with patch("app.sap_recovery.is_saplogon_running", return_value=True), \
         patch("app.sap_recovery.is_sap_gui_hung", return_value=False), \
         patch("app.sap_recovery.kill_hung_sap_processes") as mock_kill, \
         patch("app.sap_recovery.auto_connect_sso_session") as mock_sso:

        watchdog.start()
        time.sleep(0.3)
        watchdog.stop()

        assert not mock_kill.called
        assert not mock_sso.called
        assert not recovery_callback.called


def test_sap_global_watchdog_suppressed_no_recovery():
    """Kiểm tra khi được gọi suppress (đang mở Multi-Tab tải nặng) thì không bao giờ kích hoạt diệt SAP."""
    recovery_callback = MagicMock()

    watchdog = SapGlobalWatchdog(
        connection_name="Test SSO",
        hung_threshold_sec=0.2,
        check_interval_sec=0.05,
        on_recovery=recovery_callback,
    )

    with patch("app.sap_recovery.is_saplogon_running", return_value=True), \
         patch("app.sap_recovery.is_sap_gui_hung", return_value=True), \
         patch("app.sap_recovery.kill_hung_sap_processes") as mock_kill, \
         patch("app.sap_recovery.auto_connect_sso_session") as mock_sso:

        watchdog.suppress(duration_sec=1.0)  # Tạm hoãn 1.0s
        watchdog.start()
        time.sleep(0.3)
        watchdog.stop()

        # Không được kill hay sso vì đang suppressed
        assert not mock_kill.called
        assert not mock_sso.called
        assert not recovery_callback.called


def test_recover_hung_sap_session_deduplication():
    """Kiểm tra Deduplication Guard: Nếu vừa hồi phục <15s và SAP còn sống thì tái sử dụng, không kill lần 2."""
    mock_session = MagicMock()
    set_last_recovered_session(mock_session)

    with patch("app.sap_recovery.is_sap_gui_alive", return_value=True), \
         patch("app.sap_recovery.is_sap_gui_hung", return_value=False), \
         patch("app.sap_recovery.kill_hung_sap_processes") as mock_kill:

        recovered = recover_hung_sap_session(connection_name="Test SSO", force=False)
        assert recovered is mock_session
        assert not mock_kill.called  # KHÔNG được gọi kill lần 2!
