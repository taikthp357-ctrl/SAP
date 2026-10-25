# -*- coding: utf-8 -*-
"""
adaptive_cache.py - Bộ quản lý cấu hình tự học UI SAP GUI (AdaptiveConfigCache)

Tự động nhận diện chế độ kết nối (HIGH_SPEED vs LOW_SPEED), dò tìm các trường
nhập liệu và nút bấm ở lần chạy đầu tiên, sau đó lưu lại đường dẫn ID thực tế
vào file `data/sap_ui_cache.json` để các lần sau gọi trực tiếp `findById` đạt
tốc độ tối đa (~0.001s).
"""

import os
import json
import threading
from typing import Optional, Any, Dict, List

CACHE_LOCK = threading.Lock()


def get_default_cache_path() -> str:
    """Xác định đường dẫn mặc định tới sap_ui_cache.json (ưu tiên thư mục data/)."""
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, "data")
    if os.path.isdir(data_dir):
        return os.path.join(data_dir, "sap_ui_cache.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "sap_ui_cache.json")


class AdaptiveConfigCache:
    """Tự động nhận diện Low/High Speed Connection, dò ID lần đầu và lưu cache sap_ui_cache.json."""

    def __init__(self, cache_path: Optional[str] = None):
        self.cache_path = cache_path or get_default_cache_path()
        self.data: Dict[str, Dict[str, Any]] = self._load()

    def _load(self) -> Dict[str, Dict[str, Any]]:
        """Đọc cache từ file JSON trên đĩa."""
        if os.path.exists(self.cache_path):
            try:
                with open(self.cache_path, "r", encoding="utf-8") as f:
                    content = json.load(f)
                    if isinstance(content, dict):
                        return {
                            "HIGH_SPEED": content.get("HIGH_SPEED", {}),
                            "LOW_SPEED": content.get("LOW_SPEED", {})
                        }
            except Exception:
                pass
        return {"HIGH_SPEED": {}, "LOW_SPEED": {}}

    def save(self):
        """Ghi đè cache ra file JSON một cách an toàn đa luồng."""
        with CACHE_LOCK:
            try:
                parent = os.path.dirname(self.cache_path)
                if parent:
                    os.makedirs(parent, exist_ok=True)
                with open(self.cache_path, "w", encoding="utf-8") as f:
                    json.dump(self.data, f, indent=2, ensure_ascii=False)
            except Exception:
                pass

    def get_mode(self, session: Any) -> str:
        """Phát hiện kết nối hiện tại của SAP session là LOW_SPEED hay HIGH_SPEED."""
        try:
            parent = getattr(session, "Parent", None)
            is_low = bool(getattr(parent, "IsLowSpeedConnection", False))
            return "LOW_SPEED" if is_low else "HIGH_SPEED"
        except Exception:
            return "HIGH_SPEED"

    def get_cached_id(self, session: Any, key_alias: str) -> Optional[str]:
        """Lấy ID từ cache tương ứng với mode của session."""
        mode = self.get_mode(session)
        return self.data.get(mode, {}).get(key_alias)

    def set_cached_id(self, session: Any, key_alias: str, element_id: str):
        """Lưu ID vào cache tương ứng với mode của session."""
        mode = self.get_mode(session)
        mode_cache = self.data.setdefault(mode, {})
        mode_cache[key_alias] = element_id
        self.save()

    def get_config_value(self, session: Any, key: str, default: Any = None) -> Any:
        """Đọc giá trị cấu hình phụ trợ (ví dụ: gos_list_type, mb52_output_mode)."""
        mode = self.get_mode(session)
        return self.data.get(mode, {}).get(key, default)

    def set_config_value(self, session: Any, key: str, value: Any):
        """Lưu giá trị cấu hình phụ trợ."""
        mode = self.get_mode(session)
        mode_cache = self.data.setdefault(mode, {})
        mode_cache[key] = value
        self.save()

    def resolve_element(
        self,
        session: Any,
        key_alias: str,
        field_name: str,
        element_type: Optional[str] = None,
        window: str = "wnd[0]"
    ) -> Optional[Any]:
        """
        Ưu tiên lấy đối tượng từ ID đã lưu trong Cache.
        Nếu lần đầu chưa có hoặc ID thay đổi -> Tự dò bằng FindByName rồi lưu lại ID mới.
        """
        if session is None:
            return None

        mode = self.get_mode(session)
        mode_cache = self.data.setdefault(mode, {})

        # 1. Thử lấy nhanh bằng ID đã lưu trong cache
        cached_id = mode_cache.get(key_alias)
        if cached_id:
            try:
                el = session.findById(cached_id, False)
                if el is not None:
                    return el
            except TypeError:
                try:
                    el = session.findById(cached_id)
                    if el is not None:
                        return el
                except Exception:
                    pass
            except Exception:
                pass

        # 2. Quét tìm đối tượng trên màn hình qua FindByName
        try:
            root = session.findById(f"{window}/usr")
        except Exception:
            try:
                root = session.findById(window)
            except Exception:
                root = None

        if root is None:
            return None

        # Thử tìm trực tiếp với element_type nếu được chỉ định
        if element_type:
            try:
                el = root.FindByName(field_name, element_type)
                if el is not None:
                    el_id = getattr(el, "Id", getattr(el, "id", ""))
                    if el_id:
                        mode_cache[key_alias] = el_id
                        self.save()
                    return el
            except Exception:
                pass

        # Quét qua các control types phổ biến
        control_types = [
            "GuiCTextField", "GuiTextField", "GuiComboBox",
            "GuiButton", "GuiTab", "GuiCheckBox", "GuiRadioButton",
            "GuiOkCodeField", "GuiLabel", "GuiShell"
        ]
        for c_type in control_types:
            try:
                el = root.FindByName(field_name, c_type)
                if el is not None:
                    el_id = getattr(el, "Id", getattr(el, "id", ""))
                    if el_id:
                        mode_cache[key_alias] = el_id
                        self.save()
                    return el
            except Exception:
                continue

        # 3. Đệ quy quét con nếu FindByName không thấy
        def _find_recursive(obj, name: str, typ: Optional[str]):
            children = getattr(obj, "Children", None)
            if not children:
                return None
            try:
                count = getattr(children, "Count", 0)
            except Exception:
                count = 0
            for idx in range(count):
                try:
                    child = children(idx)
                    c_name = getattr(child, "Name", getattr(child, "name", ""))
                    c_type = getattr(child, "Type", getattr(child, "type", ""))
                    if c_name == name:
                        if typ is None or c_type == typ:
                            return child
                    sub = _find_recursive(child, name, typ)
                    if sub is not None:
                        return sub
                except Exception:
                    continue
            return None

        found = _find_recursive(root, field_name, element_type)
        if found is not None:
            el_id = getattr(found, "Id", getattr(found, "id", ""))
            if el_id:
                mode_cache[key_alias] = el_id
                self.save()
            return found

        return None


# Khởi tạo singleton dùng chung trong toàn bộ ứng dụng
UI_CACHE = AdaptiveConfigCache()
