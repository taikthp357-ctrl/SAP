# -*- coding: utf-8 -*-
"""Kiểm thử tự động kiến trúc (Architecture-as-Code):
Đảm bảo ZERO CROSS-IMPORT giữa các tầng (core, readers, writers, parsers, handlers).
Không một ai có thể vô tình import chéo vi phạm ranh giới kiến trúc.
"""
import ast
from pathlib import Path
import pytest

RULES = {
    "core":       {"allowed": {"core"}},
    "readers":    {"allowed": {"core"}},
    "writers":    {"allowed": {"core"}},
    "parsers":    {"allowed": {"core", "parsers"}},
}

def test_zero_cross_import_boundaries():
    """Kiểm tra cây cú pháp trừu tượng AST của mọi file .py trong app/."""
    base_app_dir = Path(__file__).resolve().parent.parent / "app"
    assert base_app_dir.exists(), "Thư mục app/ phải tồn tại"

    violations = []
    for layer, config in RULES.items():
        layer_dir = base_app_dir / layer
        if not layer_dir.exists():
            continue
        for py_file in layer_dir.rglob("*.py"):
            try:
                tree = ast.parse(py_file.read_text(encoding="utf-8"))
            except Exception as e:
                violations.append(f"{py_file.name}: Không thể phân tích AST: {e}")
                continue

            for node in ast.walk(tree):
                # 1. from app.xyz import ...
                if isinstance(node, ast.ImportFrom) and node.module and node.module.startswith("app."):
                    parts = node.module.split(".")
                    if len(parts) >= 2:
                        target_layer = parts[1]
                        if target_layer not in config["allowed"]:
                            violations.append(
                                f"{py_file.relative_to(base_app_dir)} (line {node.lineno}): "
                                f"KHÔNG ĐƯỢC PHÉP import từ 'app.{target_layer}'!"
                            )
                # 2. import app.xyz
                elif isinstance(node, ast.Import):
                    for alias in node.names:
                        if alias.name.startswith("app."):
                            parts = alias.name.split(".")
                            if len(parts) >= 2:
                                target_layer = parts[1]
                                if target_layer not in config["allowed"]:
                                    violations.append(
                                        f"{py_file.relative_to(base_app_dir)} (line {node.lineno}): "
                                        f"KHÔNG ĐƯỢC PHÉP import 'app.{target_layer}'!"
                                    )

    assert not violations, "Phát hiện vi phạm ranh giới kiến trúc:\n" + "\n".join(violations)
