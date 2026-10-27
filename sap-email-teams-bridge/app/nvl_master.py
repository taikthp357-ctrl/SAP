# -*- coding: utf-8 -*-
"""Master data management for NVL Ra Đơn (Kế thừa từ 260810_RA_DON_NVL_V31_ZONE.xlsm).

Bao gồm 3 khu vực chính:
1. MÃ CHÚ Ý / FEFO: Các vật tư có hạn sử dụng cần ưu tiên cấp hàng trước.
2. HÀNG THÙNG / PACKING: Danh mục vật tư đóng thùng, quy cách đóng gói (packing) và giới hạn số thùng.
3. BIN / ZONE: Danh mục ánh xạ mẫu Bin sang Zone/Khu vực lưu trữ và lọc BIN BLOCK hệ thống.
"""

import fnmatch
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


class NvlMasterService:
    DEFAULT_JSON_PATH = Path(__file__).resolve().parent.parent / "data" / "nvl_master_v31.json"
    EXCEL_V31_PATH = Path("G:/SAP/260810_RA_DON_NVL_V31_ZONE.xlsm")

    _master_cache: Optional[Dict[str, Any]] = None

    @classmethod
    def load_master_data(cls, force_reload: bool = False) -> Dict[str, Any]:
        """Tải dữ liệu danh mục từ cache JSON hoặc tự động trích xuất từ file V31 nếu cần."""
        if cls._master_cache is not None and not force_reload:
            return cls._master_cache

        data = None
        json_path = Path(cls.DEFAULT_JSON_PATH)

        if json_path.exists() and not force_reload:
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                print(f"[NVL_MASTER WARN] Không đọc được file cache JSON {json_path}: {e}")

        if data is None and cls.EXCEL_V31_PATH.exists():
            try:
                data = cls._sync_from_excel(cls.EXCEL_V31_PATH, json_path)
            except Exception as e:
                print(f"[NVL_MASTER WARN] Lỗi khi trích xuất từ Excel {cls.EXCEL_V31_PATH}: {e}")

        if data is None:
            # Dữ liệu dự phòng mặc định tối thiểu
            data = {
                "version": "DEFAULT_FALLBACK",
                "fefo_materials": {
                    "BIW30009362": "HSD",
                    "BIW30111041": "HSD",
                    "BIW71005512": "HSD",
                    "BIW71005612": "HSD",
                    "BIW32032012AA": "HSD",
                    "BIW32032018AA": "HSD",
                    "CON69000038": "HSD"
                },
                "box_catalog": {},
                "bin_patterns": [
                    {"pattern": "APFL-BL*", "zone": "BIN BLOCK"},
                    {"pattern": "APFL-CC*", "zone": "Sọt sắt-Box giấy"},
                    {"pattern": "APFL-HO*", "zone": "Hàng Hold"},
                    {"pattern": "APFL-HR*", "zone": "Box giấy-Hàng thùng"},
                    {"pattern": "APFL-HTK*", "zone": "Hàng tập kết"},
                    {"pattern": "APFL-MA*", "zone": "Hàng mới về"},
                    {"pattern": "APFL-OV*", "zone": "MA2"},
                    {"pattern": "APFL-P00*", "zone": "Hàng Box"},
                    {"pattern": "APFL-P01*", "zone": "Hàng Box"},
                    {"pattern": "APFL-P02*", "zone": "Hàng Box"},
                    {"pattern": "APFL-P*", "zone": "Hàng Box"}
                ],
                "box_rules": [
                    {"min": 1, "max": 100, "boxes": 3},
                    {"min": 101, "max": 200, "boxes": 2},
                    {"min": 201, "max": 999999, "boxes": 1}
                ]
            }

        cls._master_cache = data
        return data

    @classmethod
    def _sync_from_excel(cls, excel_path: Path, output_json_path: Path) -> Dict[str, Any]:
        """Trích xuất dữ liệu từ các sheet DANH_MUC và CONFIG trong file Excel V31."""
        import openpyxl

        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
        ws_dm = wb["DANH_MUC"]
        ws_cfg = wb["CONFIG"]

        fefo = {}
        for r in range(5, ws_dm.max_row + 1):
            p = ws_dm.cell(r, 1).value
            t = ws_dm.cell(r, 2).value
            if p:
                fefo[str(p).strip().upper()] = str(t or "HSD").strip()

        boxes = {}
        for r in range(5, ws_dm.max_row + 1):
            m = ws_dm.cell(r, 5).value
            if m:
                mat_key = str(m).strip().upper()
                if mat_key not in boxes:
                    boxes[mat_key] = {
                        "model": str(ws_dm.cell(r, 4).value or "").strip(),
                        "material": mat_key,
                        "box_type": str(ws_dm.cell(r, 6).value or "").strip(),
                        "packing": int(ws_dm.cell(r, 7).value) if ws_dm.cell(r, 7).value else 1
                    }

        bin_patterns = []
        for r in range(5, ws_dm.max_row + 1):
            bp = ws_dm.cell(r, 9).value
            bz = ws_dm.cell(r, 10).value
            if bp:
                bin_patterns.append({
                    "pattern": str(bp).strip().upper(),
                    "zone": str(bz or "KHAC").strip()
                })

        box_rules = []
        for r in range(3, ws_cfg.max_row + 1):
            pmin = ws_cfg.cell(r, 7).value
            pmax = ws_cfg.cell(r, 8).value
            nbox = ws_cfg.cell(r, 9).value
            act = ws_cfg.cell(r, 10).value
            if pmin is not None and str(act).upper() == "TRUE":
                box_rules.append({
                    "min": int(pmin),
                    "max": int(pmax),
                    "boxes": int(nbox)
                })

        master_data = {
            "version": "V31_ZONE",
            "source_file": excel_path.name,
            "fefo_materials": fefo,
            "box_catalog": boxes,
            "bin_patterns": bin_patterns,
            "box_rules": box_rules
        }

        try:
            output_json_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_json_path, "w", encoding="utf-8") as f:
                json.dump(master_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[NVL_MASTER WARN] Không thể lưu cache JSON {output_json_path}: {e}")

        return master_data

    @classmethod
    def get_bin_zone(cls, bin_name: str) -> str:
        """Xác định Zone từ tên Bin dựa trên các mẫu trong DANH_MUC BIN/ZONE. Mặc định 'KHAC'."""
        b_clean = str(bin_name or "").strip().upper()
        if not b_clean:
            return "KHAC"

        master = cls.load_master_data()
        patterns = master.get("bin_patterns", [])

        for item in patterns:
            pat = item.get("pattern", "")
            if fnmatch.fnmatchcase(b_clean, pat):
                return item.get("zone", "KHAC")

        return "KHAC"

    @classmethod
    def is_bin_blocked(cls, bin_name: str) -> bool:
        """Kiểm tra Bin có thuộc diện BIN BLOCK hệ thống không (APFL-BL* hoặc Zone BIN BLOCK)."""
        b_clean = str(bin_name or "").strip().upper()
        if not b_clean:
            return False
        if b_clean.startswith("APFL-BL") or "BLOCK" in b_clean:
            return True
        zone = cls.get_bin_zone(b_clean)
        return "BLOCK" in zone.upper()

    @classmethod
    def get_box_info(cls, material: str) -> Optional[Dict[str, Any]]:
        """Lấy thông tin quy cách đóng gói hàng thùng cho mã vật tư."""
        m_clean = str(material or "").strip().upper()
        if not m_clean:
            return None
        master = cls.load_master_data()
        catalog = master.get("box_catalog", {})
        return catalog.get(m_clean)

    @classmethod
    def is_fefo_material(cls, material: str) -> bool:
        """Kiểm tra mã vật tư có nằm trong danh mục FEFO / Hạn sử dụng không."""
        m_clean = str(material or "").strip().upper()
        if not m_clean:
            return False
        master = cls.load_master_data()
        fefo = master.get("fefo_materials", {})
        return m_clean in fefo

    @classmethod
    def get_box_limit(cls, packing: int) -> int:
        """Lấy số thùng tối đa cho phép dựa trên khoảng Packing từ CONFIG."""
        if packing <= 0:
            return 1
        master = cls.load_master_data()
        rules = master.get("box_rules", [])
        for r in rules:
            if r.get("min", 0) <= packing <= r.get("max", 999999):
                return int(r.get("boxes", 1))
        return 1

    @classmethod
    def calculate_box_split(cls, material: str, requested_qty: float) -> Dict[str, Any]:
        """Tính toán chia hàng thùng / hàng lẻ cho mã vật tư:
        - full_boxes: Số thùng nguyên
        - box_qty: Số lượng chẵn thùng
        - loose_qty: Số lượng lẻ dư ra
        - v31_max_boxes: Giới hạn số thùng theo CONFIG V31
        - v31_capped_qty: Số lượng đơn sau khi áp dụng giới hạn thùng
        """
        box_info = cls.get_box_info(material)
        if not box_info or box_info.get("packing", 0) <= 0:
            return {
                "is_box": False,
                "material": str(material).strip().upper(),
                "model": "",
                "box_type": "",
                "packing": 0,
                "requested_qty": float(requested_qty),
                "full_boxes": 0,
                "box_qty": 0.0,
                "loose_qty": float(requested_qty),
                "v31_max_boxes": 0,
                "v31_capped_qty": float(requested_qty)
            }

        packing = int(box_info["packing"])
        full_boxes = int(requested_qty // packing)
        box_qty = float(full_boxes * packing)
        loose_qty = float(requested_qty - box_qty)
        max_boxes = cls.get_box_limit(packing)
        v31_capped_qty = float(max_boxes * packing)

        return {
            "is_box": True,
            "material": box_info["material"],
            "model": box_info["model"],
            "box_type": box_info["box_type"],
            "packing": packing,
            "requested_qty": float(requested_qty),
            "full_boxes": full_boxes,
            "box_qty": box_qty,
            "loose_qty": loose_qty,
            "v31_max_boxes": max_boxes,
            "v31_capped_qty": v31_capped_qty
        }

    @classmethod
    def extract_zone_from_station(cls, station_text: str) -> str:
        """Trích xuất Zone từ chuỗi Trạm theo chuẩn ModuleLineZone.ZoneOf (Excel V31).
        Ưu tiên:
        1. ZONE <number> (vd: 'Zone 21', 'ZONE-12', 'Zone 3.1')
        2. KV-<number> (vd: 'KV-12', 'KV21')
        """
        text = str(station_text or "").strip()
        if not text:
            return ""
        m_zone = re.search(r"\bZONE[ -]*([0-9]+(?:\.[0-9]+)*)\b", text, re.IGNORECASE)
        if m_zone:
            return m_zone.group(1).strip()
        m_kv = re.search(r"\bKV[ -]*([0-9]+)\b", text, re.IGNORECASE)
        if m_kv:
            return m_kv.group(1).strip()
        return ""

    @classmethod
    def extract_tl_from_station(cls, station_text: str) -> str:
        """Trích xuất Tên Team Leader (TL) từ chuỗi Trạm theo chuẩn ModuleLineZone.TLOf (Excel V31)."""
        text = str(station_text or "").strip()
        if not text:
            return ""
        t = re.sub(r"^KV[ ]*-?[ ]*[0-9]+(?:[/,-][0-9]+)*(?:SV)?[ ]*[/\\-_:]*[ ]*", "", text, flags=re.IGNORECASE).strip()
        t = re.sub(r"\bZONE[ -]*([0-9]+(?:\.[0-9]+)*)\b", "", t, flags=re.IGNORECASE).strip()
        t = t.strip("-/:_ ")
        return re.sub(r"\s+", " ", t).upper() if t else text.upper()

    @classmethod
    def get_bin_area(cls, bin_name: str) -> str:
        """Phân loại khu vực Bin theo chuẩn ModuleHelper.GetBinArea (Excel V31):
        - 'P' nếu bắt đầu bằng APFL-P*
        - 'ZF' nếu bắt đầu bằng APFL-ZF*
        - 'SPECIAL' cho các trường hợp còn lại
        """
        b_clean = str(bin_name or "").strip().upper()
        if b_clean.startswith("APFL-P"):
            return "P"
        elif b_clean.startswith("APFL-ZF"):
            return "ZF"
        return "SPECIAL"
