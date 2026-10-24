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


def test_get_materials_by_model():
    resolver = MaterialResolver()
    limo_mats = resolver.get_materials_by_model("Limo", limit=5)
    assert isinstance(limo_mats, list)
    assert len(limo_mats) > 0
    assert any("32013765" in m or "BIW" in m for m in limo_mats)


def test_detect_model_and_shop():
    m1, s1, p1 = MaterialResolver.detect_model_and_shop("@tài check cho tôi 20 mã limo tồn thấp")
    assert m1 == "LIMO"
    assert s1 == "BODY"
    assert p1 == ["3093", "3001"]

    m2, s2, p2 = MaterialResolver.detect_model_and_shop("check tồn thấp xmd 10 mã")
    assert m2 == "XMD"
    assert s2 == "XMD"
    assert p2 == ["3093", "2001", "3001"]

