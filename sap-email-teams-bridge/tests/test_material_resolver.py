# -*- coding: utf-8 -*-
"""Unit tests for MaterialResolver."""

import sys
from pathlib import Path
import pytest

CURRENT_DIR = Path(__file__).resolve().parent.parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from app.material_resolver import MaterialResolver


@pytest.fixture
def resolver():
    return MaterialResolver()


def test_resolve_exact_full_code(resolver):
    codes, note = resolver.resolve("BIW32013765AA")
    assert "BIW32013765AA" in codes
    assert "Khớp mã đầy đủ" in note


def test_resolve_short_code(resolver):
    codes, note = resolver.resolve("32013765")
    assert "BIW32013765AA" in codes
    assert "Quy đổi từ mã rút gọn" in note


def test_resolve_another_real_code(resolver):
    # Test mã 32024300 có sẵn trong DATA_MASTER
    codes, note = resolver.resolve("32024300")
    assert any("32024300" in c for c in codes)


def test_resolve_unknown_8_digits(resolver):
    # 8 chữ số chưa có trong file -> suy đoán BIW...AA
    codes, note = resolver.resolve("99998888")
    assert "BIW99998888AA" in codes
