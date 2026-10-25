# -*- coding: utf-8 -*-
import os
import json
import tempfile
import pytest
from unittest.mock import MagicMock
from app.adaptive_cache import AdaptiveConfigCache


def test_adaptive_cache_init_and_save():
    with tempfile.TemporaryDirectory() as tmpdir:
        cache_file = os.path.join(tmpdir, "test_cache.json")
        cache = AdaptiveConfigCache(cache_path=cache_file)

        assert cache.data == {"HIGH_SPEED": {}, "LOW_SPEED": {}}

        # Set config value
        mock_session = MagicMock()
        mock_session.Parent.IsLowSpeedConnection = False
        assert cache.get_mode(mock_session) == "HIGH_SPEED"

        cache.set_cached_id(mock_session, "test_key", "wnd[0]/usr/ctxtTEST")
        assert cache.get_cached_id(mock_session, "test_key") == "wnd[0]/usr/ctxtTEST"

        # Verify saved to disk
        assert os.path.exists(cache_file)
        with open(cache_file, "r", encoding="utf-8") as f:
            disk_data = json.load(f)
            assert disk_data["HIGH_SPEED"]["test_key"] == "wnd[0]/usr/ctxtTEST"


def test_adaptive_cache_low_speed_detection():
    cache = AdaptiveConfigCache(cache_path=None)
    mock_session = MagicMock()
    mock_session.Parent.IsLowSpeedConnection = True
    assert cache.get_mode(mock_session) == "LOW_SPEED"


def test_adaptive_cache_resolve_from_cache():
    with tempfile.TemporaryDirectory() as tmpdir:
        cache_file = os.path.join(tmpdir, "test_cache.json")
        cache = AdaptiveConfigCache(cache_path=cache_file)

        mock_session = MagicMock()
        mock_session.Parent.IsLowSpeedConnection = False
        mock_el = MagicMock()
        mock_session.findById.return_value = mock_el

        # Manually seed cache
        cache.data["HIGH_SPEED"]["migo_doc"] = "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC"

        resolved = cache.resolve_element(mock_session, "migo_doc", "GODYNPRO-MAT_DOC")
        assert resolved == mock_el
        mock_session.findById.assert_called_with("wnd[0]/usr/ctxtGODYNPRO-MAT_DOC", False)


def test_adaptive_cache_resolve_from_find_by_name():
    with tempfile.TemporaryDirectory() as tmpdir:
        cache_file = os.path.join(tmpdir, "test_cache.json")
        cache = AdaptiveConfigCache(cache_path=cache_file)

        mock_session = MagicMock()
        mock_session.Parent.IsLowSpeedConnection = False
        # findById for cached_id returns None or raises
        mock_session.findById.side_effect = lambda cid, *args: mock_usr if cid == "wnd[0]/usr" else None

        mock_usr = MagicMock()
        mock_field = MagicMock()
        mock_field.Id = "wnd[0]/usr/ssubSUB:SAPLMIGO/ctxtGODYNPRO-MAT_DOC"
        mock_usr.FindByName.return_value = mock_field

        resolved = cache.resolve_element(mock_session, "migo_doc", "GODYNPRO-MAT_DOC")
        assert resolved == mock_field
        # Verify it learned and cached the ID
        assert cache.data["HIGH_SPEED"]["migo_doc"] == "wnd[0]/usr/ssubSUB:SAPLMIGO/ctxtGODYNPRO-MAT_DOC"
