# -*- coding: utf-8 -*-
"""Dịch vụ theo dõi cấp phát NVL, đồng bộ trạng thái đơn và báo cáo ca sản xuất.

Chức năng:
1. find_tracking_files: Tìm kiếm file theo dõi cấp NVL mới nhất hoặc theo ca.
2. get_shift_supply_status: Báo cáo tình trạng cấp hàng theo Model/Ca/Ngày, tính KPI và lọc mã nợ.
3. sync_dispatch_to_tracking_file: Ghi nhận số OB vào file theo dõi sau khi ra đơn MIGO/OB,
   cập nhật Total cấp, Gap, đổi trạng thái sang 'Đã ra đơn chờ cấp' hoặc 'Finished'.
4. parse_production_order_text: Bóc tách email order từ xưởng thành danh sách yêu cầu chuẩn hóa.
5. get_wht_print_guidance: Cung cấp hướng dẫn và cú pháp T-code in phiếu WHT (/SCWM/MON, /SCWM/PRDO, MB90, LT31).
6. generate_shift_kpi_report: Xuất file Excel báo cáo kết ca chuẩn form BC Cap NVL.
"""

import os
import re
import io
import json
import glob
import shutil
import datetime
from pathlib import Path
from decimal import Decimal
from typing import Dict, Any, List, Optional, Tuple

import openpyxl
from app.nvl_master import NvlMasterService


class NvlFilePasswordRequiredError(Exception):
    """Ngoại lệ khi file theo dõi Excel (.xlsx/.xlsb) bị cài mật khẩu mở file nhưng danh sách mật khẩu đã cấu hình không thể mở khóa."""

    def __init__(self, filename: str, message: str = ""):
        self.filename = filename
        self.message = message or f"File '{filename}' được bảo vệ bằng mật khẩu mở file. Vui lòng cung cấp mật khẩu qua lệnh 'mật khẩu nvl: <mật_khẩu>'."
        super().__init__(self.message)


class NvlTrackingService:
    """Dịch vụ điều phối và đồng bộ luồng theo dõi cấp NVL."""

    DEFAULT_TRACKING_DIR = r"G:\SAP"
    DEFAULT_PASSWORDS = [
        "Khop", "khop", "KHOP", "KhoP", "KhoP@123",
        "vinfast", "Vinfast", "VinFast", "Vinfast@123", "VinFast@123", "VINFAST",
        "Vipass", "vipass", "Vipass@123",
        "", "1", "12", "123", "1234", "12345", "123456", "12345678",
        "admin", "pass", "password", "3093", "1005", "VF", "vf",
        "LIMO", "limo", "Limo", "S6", "S7", "NVL", "nvl"
    ]
    RUNTIME_PASSWORDS: List[str] = []

    @classmethod
    def register_password(cls, password: str) -> None:
        """Đăng ký thêm mật khẩu giải mã Excel trong phiên chạy hiện tại."""
        p_clean = (password or "").strip()
        if p_clean and p_clean not in cls.RUNTIME_PASSWORDS:
            cls.RUNTIME_PASSWORDS.insert(0, p_clean)

    @classmethod
    def get_known_passwords(cls) -> List[str]:
        """Lấy toàn bộ danh sách mật khẩu đã biết (runtime + config.json + default)."""
        result = list(cls.RUNTIME_PASSWORDS)
        try:
            config_path = Path(__file__).resolve().parent.parent / "config.json"
            if config_path.is_file():
                with open(config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    cfg_pwds = cfg.get("tracking_excel_passwords", [])
                    for p in cfg_pwds:
                        p_str = str(p).strip()
                        if p_str and p_str not in result:
                            result.append(p_str)
        except Exception:
            pass
        for p in cls.DEFAULT_PASSWORDS:
            if p not in result:
                result.append(p)
        return result

    @classmethod
    def _read_tracking_sheet_rows(
        cls,
        file_path: Path,
        sheet_target: str = "ODER NVl",
        passwords: Optional[List[str]] = None
    ) -> List[List[Any]]:
        """Đọc toàn bộ các dòng của sheet theo dõi, hỗ trợ cả .xlsx và .xlsb,
        tự động giải mã in-memory nếu file bị cài mật khẩu mở file.
        
        Trả về danh sách các row (mỗi row là list các giá trị cell 0-based).
        """
        is_xlsb = file_path.suffix.lower() == ".xlsb"
        pw_list = passwords or cls.get_known_passwords()

        with open(file_path, "rb") as f:
            file_bytes = f.read()

        stream: Optional[io.BytesIO] = None
        # Kiểm tra OLE2 header (\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1)
        if len(file_bytes) > 8 and file_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            try:
                import msoffcrypto
                buf_in = io.BytesIO(file_bytes)
                try:
                    office_file = msoffcrypto.OfficeFile(buf_in)
                    if office_file.is_encrypted():
                        decrypted_buf = None
                        for pwd in pw_list:
                            try:
                                buf_in.seek(0)
                                temp_of = msoffcrypto.OfficeFile(buf_in)
                                temp_of.load_key(password=pwd)
                                out_buf = io.BytesIO()
                                temp_of.decrypt(out_buf)
                                out_buf.seek(0)
                                decrypted_buf = out_buf
                                break
                            except Exception:
                                continue
                        if decrypted_buf is None:
                            raise NvlFilePasswordRequiredError(file_path.name)
                        stream = decrypted_buf
                except NvlFilePasswordRequiredError:
                    raise
                except Exception:
                    pass
            except ImportError:
                pass

        if stream is None:
            stream = io.BytesIO(file_bytes)

        rows: List[List[Any]] = []
        if is_xlsb:
            try:
                import pyxlsb
                with pyxlsb.open_workbook(stream) as wb:
                    s_name = sheet_target if sheet_target in wb.sheets else wb.sheets[0]
                    sheet = wb.get_sheet(s_name)
                    for row in sheet.rows():
                        rows.append([c.v for c in row])
            except NvlFilePasswordRequiredError:
                raise
            except Exception as e:
                if "BadZipFile" in type(e).__name__ or "not a zip file" in str(e).lower():
                    raise NvlFilePasswordRequiredError(file_path.name)
                raise
        else:
            try:
                wb = openpyxl.load_workbook(stream, data_only=True)
                s_name = sheet_target if sheet_target in wb.sheetnames else wb.sheetnames[0]
                ws = wb[s_name]
                for row in ws.iter_rows(values_only=True):
                    rows.append(list(row))
            except NvlFilePasswordRequiredError:
                raise
            except Exception as e:
                if "InvalidFileException" in type(e).__name__ or "encrypted" in str(e).lower():
                    raise NvlFilePasswordRequiredError(file_path.name)
                raise

        return rows

    @classmethod
    def get_tracking_directory(cls, configured_dir: Optional[str] = None) -> Path:
        """Xác định thư mục chứa file theo dõi ca."""
        if configured_dir and os.path.isdir(configured_dir):
            return Path(configured_dir)
        if os.path.isdir(cls.DEFAULT_TRACKING_DIR):
            return Path(cls.DEFAULT_TRACKING_DIR)
        # Fallback về thư mục cha của project
        parent_dir = Path(__file__).resolve().parent.parent.parent
        return parent_dir

    @classmethod
    def find_tracking_files(
        cls,
        directory: Optional[str] = None,
        shift: Optional[str] = None,
        date_str: Optional[str] = None
    ) -> List[Path]:
        """Tìm các file theo dõi cấp NVL (THEO DOI ODER NVL_P_*.xlsx).
        
        Ưu tiên:
        1. Khớp chính xác Shift (ví dụ: *_S6.xlsx) và Date nếu có.
        2. Sắp xếp theo thời gian sửa đổi mới nhất (mtime).
        """
        search_dir = cls.get_tracking_directory(directory)
        pattern = str(search_dir / "*THEO DOI ODER NVL_P_*.xlsx")
        pattern_xlsb = str(search_dir / "*THEO DOI ODER NVL_P_*.xlsb")
        all_files = glob.glob(pattern)
        all_files.extend(glob.glob(pattern_xlsb))
        # Hỗ trợ tự động quét thêm thư mục con theo tháng (ví dụ G:\SAP\2026_10, G:\SAP\Thang_10...)
        sub_files = glob.glob(str(search_dir / "*/*THEO DOI ODER NVL_P_*.xlsx"))
        sub_files_xlsb = glob.glob(str(search_dir / "*/*THEO DOI ODER NVL_P_*.xlsb"))
        all_files.extend(sub_files)
        all_files.extend(sub_files_xlsb)
        if not all_files:
            # Tìm đệ quy toàn bộ thư mục
            all_files = glob.glob(str(search_dir / "**/*THEO DOI ODER NVL_P_*.xlsx"), recursive=True)
            all_files.extend(glob.glob(str(search_dir / "**/*THEO DOI ODER NVL_P_*.xlsb"), recursive=True))

        found_paths = [Path(f) for f in all_files if not os.path.basename(f).startswith("~$")]

        # Lọc theo ca nếu có (vd: S6, S7)
        if shift:
            shift_clean = shift.strip().upper()
            filtered = [
                p for p in found_paths
                if f"_{shift_clean}." in p.name.upper() or f"_{shift_clean}_" in p.name.upper()
            ]
            if filtered:
                found_paths = filtered

        # Lọc theo ngày nếu có (vd: 261028 hoặc 2026-10-28)
        if date_str:
            clean_d = re.sub(r"[^0-9]", "", date_str)
            if len(clean_d) == 8:
                short_d = clean_d[2:]  # 20261028 -> 261028
            else:
                short_d = clean_d
            date_filtered = [p for p in found_paths if short_d in p.name]
            if date_filtered:
                found_paths = date_filtered

        # Sắp xếp theo file mới nhất
        found_paths.sort(key=lambda p: (p.stat().st_mtime, p.name), reverse=True)
        return found_paths

    @classmethod
    def get_shift_supply_status(
        cls,
        tracking_file: Optional[Path] = None,
        shift: Optional[str] = None,
        model: Optional[str] = None,
        target_date: Optional[str] = None,
        directory: Optional[str] = None,
        sap_engine: Any = None
    ) -> Dict[str, Any]:
        """Báo cáo tình trạng cấp hàng NVL trong ca theo Model / Ca chạy / Ngày.
        
        Trả về:
        - summary_kpi: Tổng order, Hoàn thành, 1 phần, Pending, Tỷ lệ cấp (%)
        - shortages: Danh sách các mã thiếu hoặc chưa cấp kèm thông tin tồn kho 1005
        - records: Toàn bộ danh sách dòng theo dõi đã lọc
        """
        if not tracking_file:
            candidates = cls.find_tracking_files(directory=directory, shift=shift, date_str=target_date)
            if not candidates:
                return {
                    "status": "NOT_FOUND",
                    "summary": f"Không tìm thấy file theo dõi cấp NVL cho ca {shift or 'hiện tại'} tại {directory or cls.DEFAULT_TRACKING_DIR}",
                    "records": [],
                    "shortages": [],
                    "kpi": {}
                }
            tracking_file = candidates[0]

        try:
            sheet_rows = cls._read_tracking_sheet_rows(tracking_file)
        except NvlFilePasswordRequiredError as e:
            return {
                "status": "PASSWORD_REQUIRED",
                "summary": f"🔒 <b>FILE CÓ MẬT KHẨU:</b> File theo dõi ca <code>{tracking_file.name}</code> đang được bảo vệ bằng mật khẩu mở file.<br/>👉 Vui lòng chat: <code>mật khẩu nvl: [mật_khẩu]</code> để Bot tự động mở khóa.",
                "records": [],
                "shortages": [],
                "kpi": {},
                "file": tracking_file.name
            }

        model_clean = (model or "").strip().upper()
        shift_clean = (shift or "").strip().upper()

        total_orders = 0
        completed_count = 0
        partial_count = 0
        pending_count = 0

        total_qty_ordered = 0.0
        total_qty_supplied = 0.0

        records: List[Dict[str, Any]] = []
        shortages: List[Dict[str, Any]] = []

        # Cache tra cứu NCC từ Master
        vendor_cache: Dict[str, str] = {}

        for r_idx, r_vals in enumerate(sheet_rows):
            if r_idx == 0:
                continue
            r = r_idx + 1

            mat_val = r_vals[3] if len(r_vals) > 3 else None
            if not mat_val:
                continue

            mat = str(mat_val).strip().upper()
            row_model = str(r_vals[9] or "").strip().upper() if len(r_vals) > 9 else ""
            row_ca_nhap = str(r_vals[10] or "").strip().upper() if len(r_vals) > 10 else ""
            row_ca_chay = str(r_vals[11] or "").strip().upper() if len(r_vals) > 11 else ""

            # Bộ lọc Model (nếu được yêu cầu)
            if model_clean and model_clean not in row_model:
                continue

            # Bộ lọc Ca: Khớp ca chạy hoặc ca nhập
            if shift_clean and (shift_clean not in row_ca_chay and shift_clean not in row_ca_nhap):
                continue

            # Bộ lọc Ngày (nếu được yêu cầu)
            if target_date:
                date_cell = str((r_vals[31] if len(r_vals) > 31 else None) or (r_vals[29] if len(r_vals) > 29 else None) or "") if len(r_vals) > 29 else ""
                if target_date not in date_cell:
                    continue

            # Lấy thông tin số lượng
            try:
                sl_order = float(Decimal(str(r_vals[4] or 0))) if len(r_vals) > 4 else 0.0
            except Exception:
                sl_order = 0.0

            try:
                total_cap = float(Decimal(str(r_vals[25] or 0))) if len(r_vals) > 25 else 0.0
            except Exception:
                total_cap = 0.0

            try:
                gap_val = float(Decimal(str(r_vals[26] or 0))) if len(r_vals) > 26 else (total_cap - sl_order)
            except Exception:
                gap_val = total_cap - sl_order

            station = str(r_vals[5] or "").strip() if len(r_vals) > 5 else ""
            ob_val = str((r_vals[12] if len(r_vals) > 12 else None) or (r_vals[13] if len(r_vals) > 13 else None) or "").strip()
            status_order = str(r_vals[27] or "Pending").strip() if len(r_vals) > 27 else "Pending"
            status_pick = str(r_vals[28] or "Chưa Cấp").strip() if len(r_vals) > 28 else "Chưa Cấp"
            note_val = str(r_vals[30] or "").strip() if len(r_vals) > 30 else ""

            # Lấy NCC
            ncc_val = str(r_vals[0] or "").strip() if len(r_vals) > 0 else ""
            if not ncc_val or ncc_val in ["0x2a", "*", "None"]:
                if mat not in vendor_cache:
                    vendor_cache[mat] = NvlMasterService.get_vendor(mat) or "NCC Tiêu Chuẩn"
                ncc = vendor_cache[mat]
            else:
                ncc = ncc_val

            total_orders += 1
            total_qty_ordered += sl_order
            total_qty_supplied += total_cap

            # Đánh giá hoàn thành
            is_finished = (gap_val >= 0 and total_cap > 0) or (status_order.lower() in ["finished", "hoàn thành"])
            is_partial = (total_cap > 0 and gap_val < 0) or ("partial" in status_order.lower()) or ("cấp thiếu" in status_pick.lower())
            is_pending = not is_finished and not is_partial

            if is_finished:
                completed_count += 1
                current_status_label = "Hoàn thành"
            elif is_partial:
                partial_count += 1
                current_status_label = "Cấp 1 phần"
            else:
                pending_count += 1
                current_status_label = "Chưa cấp / Chờ đơn"

            item_dict = {
                "material": mat,
                "ncc": ncc,
                "station": station,
                "model": row_model,
                "shift_nhap": row_ca_nhap,
                "shift_chay": row_ca_chay,
                "qty_ordered": sl_order,
                "qty_supplied": total_cap,
                "gap": gap_val,
                "ob": ob_val,
                "status_order": status_order,
                "status_pick": status_pick,
                "status_label": current_status_label,
                "note": note_val
            }
            records.append(item_dict)

            # Thu thập danh sách nợ hàng / thiếu tồn
            if not is_finished:
                shortages.append(item_dict)

        completion_rate = round((completed_count / total_orders * 100), 1) if total_orders > 0 else 0.0

        # Nếu có SAP Engine, tra cứu bổ sung tồn kho 1005 cho các mã thiếu
        if sap_engine and shortages:
            shortage_mats = list(dict.fromkeys(s["material"] for s in shortages[:40]))
            try:
                stock_res = sap_engine.check_stock_balance(materials=shortage_mats, plant="3093", sloc="1005")
                stock_map = {r.get("product"): float(r.get("qty", 0.0)) for r in stock_res if r.get("product")}
                for s in shortages:
                    s["stock_1005"] = stock_map.get(s["material"], 0.0)
                    needed = abs(s["gap"])
                    if s["stock_1005"] >= needed:
                        s["advice"] = f"Kho 1005 còn {s['stock_1005']} PC -> Cần ra đơn bổ sung"
                    elif s["stock_1005"] > 0:
                        s["advice"] = f"Kho 1005 chỉ còn {s['stock_1005']}/{needed} PC -> Cấp 1 phần"
                    else:
                        s["advice"] = f"Kho 1005 hết hàng -> Cần đôn đốc NCC {s['ncc']}"
            except Exception as sap_err:
                print(f"[TRACKING WARN] Không thể tra cứu tồn kho SAP: {sap_err}")
                for s in shortages:
                    s["stock_1005"] = "N/A"
                    s["advice"] = "Chưa kết nối tồn SAP MB52"

        kpi = {
            "total_orders": total_orders,
            "completed_count": completed_count,
            "partial_count": partial_count,
            "pending_count": pending_count,
            "completion_rate": completion_rate,
            "total_qty_ordered": total_qty_ordered,
            "total_qty_supplied": total_qty_supplied,
            "tracking_file": tracking_file.name
        }

        filter_desc = []
        if model:
            filter_desc.append(f"Model {model.upper()}")
        if shift:
            filter_desc.append(f"Ca {shift.upper()}")
        if target_date:
            filter_desc.append(f"Ngày {target_date}")
        desc_str = " - ".join(filter_desc) if filter_desc else "Toàn ca"

        summary = (
            f"Báo cáo tiến độ cấp NVL ({desc_str}) từ file '{tracking_file.name}': "
            f"Tổng {total_orders} mã | Đã cấp đủ: {completed_count} ({completion_rate}%) | "
            f"Cấp 1 phần: {partial_count} | Chờ cấp/Pending: {pending_count}"
        )

        return {
            "status": "OK",
            "summary": summary,
            "kpi": kpi,
            "records": records,
            "shortages": shortages,
            "file_path": str(tracking_file)
        }

    @classmethod
    def sync_dispatch_to_tracking_file(
        cls,
        orders_dispatched: List[Dict[str, Any]],
        delivery_no: str,
        shift: Optional[str] = None,
        tracking_file: Optional[Path] = None,
        directory: Optional[str] = None
    ) -> Dict[str, Any]:
        """Tự động ghi nhận số OB và SL đã ra đơn vào file theo dõi ca.
        
        Quy tắc cập nhật:
        1. Tìm dòng khớp Mã NVL (và Trạm nếu có).
        2. Tìm ô OB còn trống tiếp theo: OB1 -> OB2 -> OB_3 -> OB_4.
        3. Điền delivery_no vào OBx, điền SL xuất vào SL_x.
        4. Tính lại Total cấp = sum(SL_1..SL_4).
        5. Tính lại Gap = Total cấp - SL oder.
        6. Đổi Cột 28 (Tình Trạng cấp đơn):
           - Nếu Gap == 0: 'Finished'
           - Nếu Gap < 0: 'Đã ra đơn chờ cấp'
        7. Đổi Cột 29 (TT nhặt Hàng): 'Chờ cấp'.
        8. Tạo bản sao lưu (.bak) và lưu file Excel an toàn.
        """
        if not orders_dispatched or not delivery_no:
            return {"status": "SKIPPED", "updated_count": 0, "message": "Không có dữ liệu đơn hoặc số OB"}

        if not tracking_file:
            candidates = cls.find_tracking_files(directory=directory, shift=shift)
            if not candidates:
                return {
                    "status": "FILE_NOT_FOUND",
                    "updated_count": 0,
                    "message": f"Không tìm thấy file theo dõi ca để đồng bộ OB {delivery_no}"
                }
            xlsx_cands = [p for p in candidates if p.suffix.lower() in [".xlsx", ".xlsm"]]
            tracking_file = xlsx_cands[0] if xlsx_cands else candidates[0]

        # Nếu file là .xlsb: kiểm tra xem có file .xlsx song hành không
        if tracking_file.suffix.lower() == ".xlsb":
            xlsx_alt = tracking_file.with_suffix(".xlsx")
            if xlsx_alt.exists():
                tracking_file = xlsx_alt
            else:
                return cls._sync_dispatch_to_xlsb(
                    tracking_file=tracking_file,
                    orders_dispatched=orders_dispatched,
                    delivery_no=delivery_no
                )

        wb = openpyxl.load_workbook(str(tracking_file))
        sheet_name = "ODER NVl" if "ODER NVl" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        updated_count = 0
        sync_details: List[Dict[str, Any]] = []
        now_str = datetime.datetime.now().strftime("%H:%M %d/%m")

        # Map danh sách xuất theo mã hàng
        disp_map: Dict[str, List[Dict[str, Any]]] = {}
        for o in orders_dispatched:
            m = str(o.get("material") or "").strip().upper()
            if m:
                disp_map.setdefault(m, []).append(o)

        for r in range(2, ws.max_row + 1):
            mat_val = ws.cell(r, 4).value
            if not mat_val:
                continue

            mat = str(mat_val).strip().upper()
            if mat not in disp_map or not disp_map[mat]:
                continue

            row_station = str(ws.cell(r, 6).value or "").strip().upper()
            
            # Khớp theo trạm hoặc lấy phần tử đầu tiên của mã này
            target_order = None
            for cand in disp_map[mat]:
                cand_st = str(cand.get("station") or "").strip().upper()
                if not cand_st or cand_st in row_station or row_station in cand_st:
                    target_order = cand
                    break
            if not target_order:
                target_order = disp_map[mat][0]

            disp_qty = float(target_order.get("allocated_qty") or target_order.get("qty") or 0.0)

            # Tìm khe OB trống tiếp theo (OB1, OB2, OB_3, OB_4)
            # Cột: OB1=14, SL_1=16, OB2=17, SL_2=19, OB_3=20, SL_3=22, OB_4=23, SL_4=25
            ob_slots = [(14, 16), (17, 19), (20, 22), (23, 25)]
            chosen_slot = None

            for ob_col, sl_col in ob_slots:
                cur_ob = ws.cell(r, ob_col).value
                cur_sl = ws.cell(r, sl_col).value
                if cur_ob is None or str(cur_ob).strip() in ["", "0", "0.0", "None"]:
                    chosen_slot = (ob_col, sl_col)
                    break
                elif str(cur_ob).strip() == str(delivery_no).strip():
                    # Đã ghi nhận số OB này trước đó, chỉ cần cập nhật lại SL
                    chosen_slot = (ob_col, sl_col)
                    break

            # Nếu cả 4 đợt đều đã có OB, ghi đè vào slot 4
            if not chosen_slot:
                chosen_slot = ob_slots[-1]

            ob_col, sl_col = chosen_slot
            ws.cell(r, ob_col).value = str(delivery_no)
            ws.cell(r, sl_col).value = disp_qty

            # Tính lại Total cấp
            total_cap = 0.0
            for _, sl_c in ob_slots:
                val = ws.cell(r, sl_c).value
                try:
                    if val is not None and str(val).strip() not in ["", "None"]:
                        total_cap += float(Decimal(str(val)))
                except Exception:
                    pass

            ws.cell(r, 26).value = total_cap

            # Tính lại Gap
            try:
                sl_order = float(Decimal(str(ws.cell(r, 5).value or 0)))
            except Exception:
                sl_order = 0.0

            gap = total_cap - sl_order
            ws.cell(r, 27).value = gap

            # Cập nhật Trạng thái
            if gap >= 0:
                ws.cell(r, 28).value = "Finished"
                ws.cell(r, 29).value = "Đã cấp"
            else:
                ws.cell(r, 28).value = "Đã ra đơn chờ cấp"
                ws.cell(r, 29).value = "Chờ cấp"

            # Ghi chú lịch sử
            old_note = str(ws.cell(r, 31).value or "").strip()
            add_note = f"MIGO {delivery_no} ({disp_qty} PC lúc {now_str})"
            ws.cell(r, 31).value = f"{old_note}; {add_note}".strip("; ")

            updated_count += 1
            sync_details.append({
                "row": r,
                "material": mat,
                "station": row_station,
                "qty": disp_qty,
                "total_cap": total_cap,
                "gap": gap,
                "delivery_no": delivery_no
            })

            # Loại order đã phân bổ xong khỏi map để tránh ghi đè lặp lại
            disp_map[mat].remove(target_order)

        # Lưu file an toàn kèm backup
        if updated_count > 0:
            try:
                bak_path = str(tracking_file) + f".bak_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copy2(str(tracking_file), bak_path)
            except Exception as bak_err:
                print(f"[SYNC BACKUP WARN] Không thể tạo backup file: {bak_err}")

            try:
                wb.save(str(tracking_file))
                print(f"[TRACKING SYNC OK] Đã cập nhật {updated_count} dòng trong '{tracking_file.name}' với số OB {delivery_no}")
            except Exception as save_err:
                print(f"[TRACKING SYNC ERROR] Lỗi khi lưu file theo dõi: {save_err}")
                return {
                    "status": "SAVE_ERROR",
                    "updated_count": 0,
                    "message": f"Không thể lưu file {tracking_file.name}: {save_err}"
                }

        return {
            "status": "SUCCEEDED",
            "updated_count": updated_count,
            "delivery_no": delivery_no,
            "tracking_file": tracking_file.name,
            "file_path": str(tracking_file),
            "details": sync_details
        }

    @classmethod
    def _sync_dispatch_to_xlsb(
        cls,
        tracking_file: Path,
        orders_dispatched: List[Dict[str, Any]],
        delivery_no: str
    ) -> Dict[str, Any]:
        """Đồng bộ số OB vào file .xlsb thông qua Excel COM (hỗ trợ mật khẩu)."""
        try:
            import win32com.client
        except ImportError:
            return {
                "status": "SKIPPED_XLSB",
                "updated_count": 0,
                "tracking_file": tracking_file.name,
                "message": "Cần win32com hoặc Excel để ghi file .xlsb"
            }

        excel = None
        wb = None
        updated_count = 0
        sync_details: List[Dict[str, Any]] = []

        disp_map: Dict[str, List[Dict[str, Any]]] = {}
        for o in orders_dispatched:
            m = str(o.get("material") or "").strip().upper()
            if m:
                disp_map.setdefault(m, []).append(o)

        passwords = cls.get_known_passwords()

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            opened = False
            for pwd in passwords:
                try:
                    wb = excel.Workbooks.Open(str(tracking_file.resolve()), False, False, None, pwd)
                    opened = True
                    break
                except Exception:
                    continue

            if not opened or wb is None:
                return {
                    "status": "LOCKED",
                    "updated_count": 0,
                    "tracking_file": tracking_file.name,
                    "message": "Không thể mở file .xlsb bằng các mật khẩu đã lưu"
                }

            sheet_names = [s.Name for s in wb.Sheets]
            target_sheet_name = "ODER NVl" if "ODER NVl" in sheet_names else wb.Sheets(1).Name
            ws = wb.Sheets(target_sheet_name)
            max_row = ws.UsedRange.Rows.Count

            ob_slots = [(14, 16), (17, 19), (20, 22), (23, 25)]

            for r in range(2, max_row + 1):
                mat_val = ws.Cells(r, 4).Value
                if not mat_val:
                    continue
                mat = str(mat_val).strip().upper()
                if mat not in disp_map or not disp_map[mat]:
                    continue

                row_station = str(ws.Cells(r, 6).Value or "").strip().upper()
                target_order = None
                for cand in disp_map[mat]:
                    cand_st = str(cand.get("station") or "").strip().upper()
                    if not cand_st or cand_st in row_station or row_station in cand_st:
                        target_order = cand
                        break
                if not target_order:
                    target_order = disp_map[mat][0]

                disp_qty = float(target_order.get("allocated_qty") or target_order.get("qty") or 0.0)

                chosen_slot = None
                for ob_col, sl_col in ob_slots:
                    cur_ob = ws.Cells(r, ob_col).Value
                    if cur_ob is None or str(cur_ob).strip() in ["", "0", "0.0", "None"]:
                        chosen_slot = (ob_col, sl_col)
                        break
                    elif str(cur_ob).strip() == str(delivery_no).strip():
                        chosen_slot = (ob_col, sl_col)
                        break

                if not chosen_slot:
                    chosen_slot = ob_slots[-1]

                ob_col, sl_col = chosen_slot
                ws.Cells(r, ob_col).Value = str(delivery_no)
                ws.Cells(r, sl_col).Value = disp_qty

                # Tính lại Total cấp
                total_cap = 0.0
                for _, sl_c in ob_slots:
                    val = ws.Cells(r, sl_c).Value
                    try:
                        if val is not None and str(val).strip() not in ["", "None"]:
                            total_cap += float(val)
                    except Exception:
                        pass

                ws.Cells(r, 26).Value = total_cap

                try:
                    sl_order = float(ws.Cells(r, 5).Value or 0)
                except Exception:
                    sl_order = 0.0

                gap = total_cap - sl_order
                ws.Cells(r, 27).Value = gap

                if gap >= 0:
                    ws.Cells(r, 28).Value = "Finished"
                    ws.Cells(r, 29).Value = "Đã cấp"
                else:
                    ws.Cells(r, 28).Value = "Đã ra đơn chờ cấp"
                    ws.Cells(r, 29).Value = "Chờ cấp"

                updated_count += 1
                sync_details.append({
                    "material": mat,
                    "row": r,
                    "delivery_no": delivery_no,
                    "qty": disp_qty,
                    "slot": f"OB{ob_slots.index(chosen_slot) + 1}",
                    "total_cap": total_cap,
                    "gap": gap
                })

            wb.Save()
            return {
                "status": "SUCCEEDED",
                "updated_count": updated_count,
                "tracking_file": tracking_file.name,
                "sync_details": sync_details,
                "delivery_no": delivery_no
            }
        except Exception as e:
            return {
                "status": "ERROR",
                "updated_count": 0,
                "tracking_file": tracking_file.name,
                "message": f"Lỗi COM .xlsb: {e}"
            }
        finally:
            if wb:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass

    @classmethod
    def get_wht_print_guidance(cls, doc_no: str, system_type: str = "EWM") -> Dict[str, Any]:
        """Cung cấp hướng dẫn chi tiết và cú pháp T-code in phiếu WHT trong SAP."""
        clean_doc = (doc_no or "").strip()
        is_ewm = system_type.upper() == "EWM" or (clean_doc.startswith("10") and len(clean_doc) == 10)

        if is_ewm:
            steps = [
                f"1. Vào T-Code <b>/SCWM/MON</b> (Warehouse Management Monitor).",
                f"2. Mở nhánh thư mục: <code>Documents</code> ➔ <code>Warehouse Task</code>.",
                f"3. Nhập số Outbound Delivery (OB): <code>{clean_doc or '[Số OB]'}</code> rồi nhấn Thực thi (F8).",
                f"4. Chọn các dòng nhiệm vụ kho (WT) cần in ➔ Nhấn nút <b>Print Warehouse Task</b> (hoặc menu <i>More Methods</i> ➔ <i>Print</i>).",
                f"<b>Cách 2 (/SCWM/PRDO):</b> Vào T-Code <code>/SCWM/PRDO</code> ➔ Tìm OB <code>{clean_doc}</code> ➔ Vào tab Items ➔ Menu <code>Follow-On Functions</code> ➔ <code>Warehouse Task</code> ➔ Chọn <code>Create + Save</code> để tự động kích hoạt lệnh in."
            ]
            tcodes = ["/SCWM/MON", "/SCWM/PRDO"]
        else:
            steps = [
                f"1. Đối với chứng từ MIGO chuyển kho (311): Vào T-Code <b>MB90</b>.",
                f"2. Nhập Output Type: <code>WA01</code> (hoặc <code>WE01</code>/<code>WA03</code>).",
                f"3. Nhập số chứng từ MIGO: <code>{clean_doc or '[Số MIGO]'}</code> và Năm chứng từ.",
                f"4. Chọn dòng kết quả ➔ Nhấn nút <b>Execute (F8)</b> ➔ Bấm <b>Print</b>.",
                f"<b>Đối với lệnh chuyển kho WM:</b> Dùng T-Code <code>LT31</code> (In đơn lẻ TO) hoặc <code>LT21</code>/<code>LT22</code> để xem và in."
            ]
            tcodes = ["MB90", "LT31", "LT21"]

        return {
            "status": "OK",
            "doc_no": clean_doc,
            "system_type": "SAP EWM" if is_ewm else "SAP ECC / WM",
            "tcodes": tcodes,
            "steps": steps,
            "summary": f"Hướng dẫn in phiếu WHT cho chứng từ {clean_doc} (T-Code: {', '.join(tcodes)})"
        }

    @classmethod
    def parse_production_order_text(cls, text: str) -> List[Dict[str, Any]]:
        """Bóc tách nội dung email hoặc bảng order từ xưởng thành danh sách yêu cầu chuẩn hóa."""
        if not text:
            return []

        lines = [line.strip() for line in text.splitlines() if line.strip()]
        orders: List[Dict[str, Any]] = []

        mat_pattern = re.compile(r"\b([A-Z0-9]{8,15})\b")
        qty_pattern = re.compile(r"\b(?:sl|qty|số\s*lượng)?[:\s=]*(\d+(?:[.,]\d+)?)\b", re.IGNORECASE)

        for line in lines:
            # Bỏ qua dòng tiêu đề
            if any(h in line.lower() for h in ["mã nvl", "tên nvl", "stt", "vật tư", "packing"]):
                continue

            parts = [p.strip() for p in re.split(r"\s{2,}|\t|[,|;]", line) if p.strip()]
            if len(parts) < 2:
                parts = line.split()
            if len(parts) >= 2:
                # Tìm mã và số lượng
                mat_cand = None
                qty_cand = None
                station_cand = ""
                model_cand = ""

                for p in parts:
                    m = mat_pattern.search(p)
                    if m and not mat_cand and any(c.isalpha() for c in m.group(1)) and any(c.isdigit() for c in m.group(1)):
                        mat_cand = m.group(1)
                    q = qty_pattern.search(p)
                    if q and qty_cand is None:
                        try:
                            val = float(q.group(1).replace(",", "."))
                            if val > 0:
                                qty_cand = val
                        except Exception:
                            pass
                    if any(st in p.upper() for st in ["SUB", "MAIN", "DOOR", "BODY", "LINE", "RB", "MG"]):
                        station_cand = p
                    if any(md in p.upper() for md in ["VF", "WILD", "LIMO"]):
                        model_cand = p

                if mat_cand and qty_cand:
                    orders.append({
                        "material": mat_cand,
                        "qty": qty_cand,
                        "station": station_cand or "LINE_PROD",
                        "model": model_cand or "LIMO",
                        "unit": "PC"
                    })

        return orders

    @classmethod
    def trace_material_allocation(
        cls,
        material: str,
        hint_model: Optional[str] = None,
        hint_shift: Optional[str] = None,
        hint_station: Optional[str] = None,
        tracking_file: Optional[Path] = None,
        directory: Optional[str] = None,
        sap_engine: Any = None
    ) -> Dict[str, Any]:
        """Định danh chi tiết mã vật tư cấp cho Ca nào, Model nào, Trạm nào.
        
        Xử lý đặc biệt cho các mã dùng chung (ví dụ: dùng chung cho cả Limo và VF8NP).
        Đối soát 3 nguồn: File theo dõi ca + Order sản xuất + Lịch sử xuất MB51 / Tồn MB52.
        """
        clean_mat = (material or "").strip().upper()
        if not clean_mat:
            return {
                "status": "INVALID",
                "material": "",
                "summary": "Mã vật tư không hợp lệ",
                "is_shared": False,
                "models": [],
                "records": []
            }

        # Tìm file theo dõi
        if not tracking_file:
            candidates = cls.find_tracking_files(directory=directory, shift=hint_shift)
            if not candidates:
                return {
                    "status": "NOT_FOUND",
                    "material": clean_mat,
                    "summary": f"Không tìm thấy file theo dõi ca để đối soát cho mã {clean_mat}",
                    "is_shared": False,
                    "models": [],
                    "records": []
                }
            tracking_file = candidates[0]

        try:
            sheet_rows = cls._read_tracking_sheet_rows(tracking_file)
        except NvlFilePasswordRequiredError as e:
            return {
                "status": "PASSWORD_REQUIRED",
                "material": clean_mat,
                "summary": f"🔒 <b>FILE CÓ MẬT KHẨU:</b> File theo dõi ca <code>{tracking_file.name}</code> đang được cài mật khẩu mở file.<br/>👉 Vui lòng chat: <code>mật khẩu nvl: [mật_khẩu]</code> để Bot tự động mở khóa và tra cứu.",
                "is_shared": False,
                "models": [],
                "records": [],
                "tracking_file": tracking_file.name
            }

        # Chuẩn hóa tìm kiếm: cả mã đầy đủ và 8 số cốt lõi
        clean_mat_short = clean_mat.replace("BIW", "").replace("SUB", "").replace("AA", "").replace("AB", "").replace("AC", "").replace("HA", "")

        matched_rows: List[Dict[str, Any]] = []
        for r_idx, r_vals in enumerate(sheet_rows):
            if r_idx == 0:
                continue
            r = r_idx + 1

            row_mat_val = r_vals[3] if len(r_vals) > 3 else None
            if not row_mat_val:
                continue
            row_mat = str(row_mat_val).strip().upper()
            is_match = (clean_mat in row_mat) or (len(clean_mat_short) >= 6 and clean_mat_short in row_mat)
            if not is_match:
                continue

            row_model = str(r_vals[9] or "").strip().upper() if len(r_vals) > 9 else ""
            row_ca_nhap = str(r_vals[10] or "").strip().upper() if len(r_vals) > 10 else ""
            row_ca_chay = str(r_vals[11] or "").strip().upper() if len(r_vals) > 11 else ""
            row_station = str(r_vals[5] or "").strip() if len(r_vals) > 5 else ""

            try:
                sl_order = float(Decimal(str(r_vals[4] or 0))) if len(r_vals) > 4 else 0.0
            except Exception:
                sl_order = 0.0

            try:
                total_cap = float(Decimal(str(r_vals[25] or 0))) if len(r_vals) > 25 else 0.0
            except Exception:
                total_cap = 0.0

            try:
                gap = float(Decimal(str(r_vals[26] or 0))) if len(r_vals) > 26 else (total_cap - sl_order)
            except Exception:
                gap = total_cap - sl_order

            collected_docs = []
            for sp_c in [14, 17, 20, 23]:
                s_val = str(r_vals[sp_c] or "").strip() if len(r_vals) > sp_c else ""
                if s_val.endswith(".0"):
                    s_val = s_val[:-2]
                if s_val and s_val not in ["None", "0"]:
                    doc_str = f"Phiếu {s_val}" if not s_val.lower().startswith("phiếu") else s_val
                    if doc_str not in collected_docs:
                        collected_docs.append(doc_str)

            sp_display = ", ".join(collected_docs) if collected_docs else "Chưa ra đơn" 

            status_order = str(r_vals[27] or "Pending").strip() if len(r_vals) > 27 else "Pending"
            status_pick = str(r_vals[28] or "Chưa Cấp").strip() if len(r_vals) > 28 else "Chưa Cấp"
            note_val = str(r_vals[30] or "").strip() if len(r_vals) > 30 else ""
            date_order = str(r_vals[31] or "").strip() if len(r_vals) > 31 else ""
            user_order = str(r_vals[33] or "").strip() if len(r_vals) > 33 else ""

            matched_rows.append({
                "row": r,
                "material": row_mat,
                "model": row_model,
                "shift_nhap": row_ca_nhap,
                "shift_chay": row_ca_chay,
                "station": row_station,
                "order_qty": sl_order,
                "total_cap": total_cap,
                "gap": gap,
                "sp_display": sp_display,
                "status_order": status_order,
                "status_pick": status_pick,
                "note": note_val,
                "date_order": date_order,
                "user_order": user_order
            })

        if not matched_rows:
            return {
                "status": "NOT_FOUND",
                "material": clean_mat,
                "summary": f"Mã {clean_mat} không có order nào trong file theo dõi '{tracking_file.name}'",
                "is_shared": False,
                "models": [],
                "records": []
            }

        # Phân nhóm theo Model
        models_map: Dict[str, List[Dict[str, Any]]] = {}
        for r_info in matched_rows:
            m_key = r_info["model"] or "UNKNOWN"
            models_map.setdefault(m_key, []).append(r_info)

        distinct_models = list(models_map.keys())
        is_shared = len(distinct_models) > 1

        # Lọc theo gợi ý nếu có (hint_model hoặc hint_shift)
        highlighted_model = None
        if hint_model:
            h_clean = hint_model.strip().upper()
            for m_cand in distinct_models:
                if h_clean in m_cand or m_cand in h_clean:
                    highlighted_model = m_cand
                    break

        # Tra cứu SAP MB52 (Tồn kho 1005) nếu có sap_engine
        stock_1005 = 0.0
        stock_checked = False
        if sap_engine:
            try:
                stock_res = sap_engine.check_stock_balance([clean_mat], plant="3093", sloc="1005")
                if stock_res:
                    stock_1005 = float(stock_res[0].get("qty", 0.0))
                    stock_checked = True
            except Exception as e:
                print(f"[TRACE WARN] Không thể tra cứu MB52: {e}")

        # Xây dựng đoạn tóm tắt đối soát (Quick Copy)
        lines_summary = []
        if is_shared:
            lines_summary.append(f"⚠️ <b>MÃ DÙNG CHUNG ĐA MODEL:</b> Mã <code>{clean_mat}</code> đang được order cho <b>{len(distinct_models)} Model</b> ({', '.join(distinct_models)}):")
        else:
            lines_summary.append(f"ℹ️ Mã <code>{clean_mat}</code> phục vụ cho Model <b>{distinct_models[0]}</b>:")

        for mdl, r_list in models_map.items():
            is_hl = (highlighted_model == mdl)
            prefix = "🎯 [YÊU CẦU] " if is_hl else "• "
            for r_item in r_list:
                gap_txt = f"Còn thiếu: {abs(r_item['gap']):.0f} PC" if r_item['gap'] < 0 else "Đã đủ order"
                lines_summary.append(
                    f"{prefix}<b>Model {mdl}</b> (Ca chạy: <b>{r_item['shift_chay']}</b> | Trạm: <code>{r_item['station']}</code>): "
                    f"Order: {r_item['order_qty']:.0f} | Đã cấp: {r_item['total_cap']:.0f} | {gap_txt} "
                    f"({r_item['sp_display']} - {r_item['status_pick']})"
                )

        total_gap = sum(abs(r["gap"]) for r in matched_rows if r["gap"] < 0)
        if stock_checked:
            if stock_1005 >= total_gap and total_gap > 0:
                lines_summary.append(f"📦 <b>TỒN KHO 1005:</b> Hiện còn <b>{stock_1005:.0f} PC</b> ➔ Đủ tồn để ra đơn MIGO 311 cấp nốt.")
            elif stock_1005 > 0 and total_gap > 0:
                lines_summary.append(f"📦 <b>TỒN KHO 1005:</b> Hiện chỉ còn <b>{stock_1005:.0f} PC</b> (thiếu tổng {total_gap:.0f} PC) ➔ Cần ưu tiên cấp cho ca chạy trước.")
            elif total_gap > 0:
                lines_summary.append(f"❌ <b>TỒN KHO 1005:</b> Hiện <b>HẾT HÀNG (0 PC)</b> ➔ Cần liên hệ NCC giao hàng khẩn cấp.")

        summary_text = "<br/>".join(lines_summary)

        # Cập nhật chi tiết hiển thị cho từng dòng kết quả
        for r_info in matched_rows:
            r_info["stock_1005"] = f"{stock_1005:.0f} PC" if stock_checked else "---"
            r_info["stock_1005_val"] = stock_1005 if stock_checked else 0.0
            r_info["is_target_model"] = bool(highlighted_model and highlighted_model.upper() in r_info["model"].upper())
            gap_val = r_info["gap"]
            if gap_val < 0:
                r_info["gap_display"] = f"-{abs(gap_val):.0f} PC"
            elif gap_val == 0:
                r_info["gap_display"] = "0 (Đủ)"
            else:
                r_info["gap_display"] = f"+{gap_val:.0f} PC"

        if highlighted_model:
            matched_rows.sort(key=lambda x: (0 if x["is_target_model"] else 1, x["row"]))

        html_table = cls.generate_trace_table_html(
            records=matched_rows,
            hint_model=highlighted_model,
            tracking_file=tracking_file.name
        )
        markdown_table = cls.generate_trace_markdown(
            records=matched_rows,
            hint_model=highlighted_model
        )

        return {
            "status": "OK",
            "material": clean_mat,
            "is_shared": is_shared,
            "models": distinct_models,
            "highlighted_model": highlighted_model,
            "summary": summary_text,
            "total_gap": total_gap,
            "stock_1005": stock_1005 if stock_checked else None,
            "records": matched_rows,
            "tracking_file": tracking_file.name,
            "html_table": html_table,
            "markdown_table": markdown_table
        }

    @classmethod
    def trace_multiple_materials(
        cls,
        materials: List[str],
        hint_model: Optional[str] = None,
        hint_shift: Optional[str] = None,
        tracking_file: Optional[Path] = None,
        directory: Optional[str] = None,
        sap_engine: Any = None
    ) -> Dict[str, Any]:
        """Định danh và đối soát hàng loạt mã hàng."""
        results = []
        for m in materials:
            res = cls.trace_material_allocation(
                material=m,
                hint_model=hint_model,
                hint_shift=hint_shift,
                tracking_file=tracking_file,
                directory=directory,
                sap_engine=sap_engine
            )
            results.append(res)

        all_records = []
        tracking_file_name = ""
        for r in results:
            if not tracking_file_name and r.get("tracking_file"):
                tracking_file_name = r["tracking_file"]
            all_records.extend(r.get("records", []))

        # Ưu tiên sắp xếp dòng theo Model yêu cầu (Target Model), sau đó theo mã và số dòng
        if hint_model:
            h_clean = hint_model.strip().upper()
            all_records.sort(key=lambda x: (0 if h_clean in (x.get("model") or "").upper() else 1, x.get("material", ""), x.get("row", 0)))

        full_summary = "<br/><hr style='border:0;border-top:1px dashed #d0d7de;margin:8px 0;'/>".join(
            r["summary"] for r in results if r.get("summary")
        )

        html_table = cls.generate_trace_table_html(
            records=all_records,
            hint_model=hint_model,
            tracking_file=tracking_file_name
        )
        markdown_table = cls.generate_trace_markdown(
            records=all_records,
            hint_model=hint_model
        )

        status = "OK"
        if not all_records and any(r.get("status") == "PASSWORD_REQUIRED" for r in results):
            status = "PASSWORD_REQUIRED"

        return {
            "status": status,
            "count": len(materials),
            "results": results,
            "records": all_records,
            "summary": full_summary,
            "html_table": html_table,
            "markdown_table": markdown_table
        }

    @classmethod
    def generate_trace_table_html(
        cls,
        records: List[Dict[str, Any]],
        hint_model: Optional[str] = None,
        tracking_file: str = ""
    ) -> str:
        """Tạo bảng HTML trực quan, chuyên nghiệp cho Teams/Web khi tra cứu cấp NVL."""
        if not records:
            return "<div style='color:#57606a;font-style:italic;'>Không có dữ liệu cấp NVL để hiển thị.</div>"

        total_order = sum(r.get("order_qty", 0) for r in records)
        total_cap = sum(r.get("total_cap", 0) for r in records)
        total_gap = sum(abs(r.get("gap", 0)) for r in records if r.get("gap", 0) < 0)
        
        hint_upper = hint_model.strip().upper() if hint_model else ""
        model_badge = f" - MODEL <span style='color:#0969da;'>{hint_upper}</span>" if hint_upper else ""

        rows_html = []
        for r in records:
            mat = r.get("material", "")
            mdl = r.get("model", "")
            shift = r.get("shift_chay", "")
            stn = r.get("station", "")
            sp = r.get("sp_display", "Chưa ra đơn")
            sl_ord = r.get("order_qty", 0)
            sl_cap = r.get("total_cap", 0)
            gap = r.get("gap", 0)
            gap_disp = r.get("gap_display", f"{gap:.0f}")
            tt = r.get("status_pick", "Chưa Cấp")
            stk = r.get("stock_1005", "---")
            is_tgt = r.get("is_target_model", False) or (hint_upper and hint_upper in mdl.upper())

            row_bg = "#f0fdf4" if is_tgt else "#ffffff"

            if is_tgt:
                mdl_html = f"<span style='background:#dafbe1;color:#1a7f37;padding:3px 7px;border-radius:4px;font-weight:bold;'>{mdl}</span>"
            elif hint_upper and hint_upper not in mdl.upper():
                mdl_html = f"<span style='background:#fff8c5;color:#9a6700;padding:2px 6px;border-radius:4px;font-size:11px;'>{mdl} (Chung)</span>"
            else:
                mdl_html = f"<b>{mdl}</b>"

            if "Phiếu" in sp:
                sp_html = f"<span style='background:#ddf4ff;color:#0969da;padding:3px 8px;border-radius:4px;font-weight:600;'>{sp}</span>"
            else:
                sp_html = f"<span style='background:#ffebe9;color:#cf222e;padding:3px 8px;border-radius:4px;'>Chưa ra đơn</span>"

            if gap < 0:
                gap_html = f"<span style='background:#ffebe9;color:#cf222e;padding:3px 7px;border-radius:4px;font-weight:bold;'>{gap_disp}</span>"
            else:
                gap_html = f"<span style='background:#dafbe1;color:#1a7f37;padding:3px 7px;border-radius:4px;font-weight:bold;'>{gap_disp}</span>"

            if any(w in tt.lower() for w in ["đã cấp", "finished", "du"]):
                tt_html = f"<span style='color:#1a7f37;font-weight:600;'>✅ {tt}</span>"
            elif "cấp thiếu" in tt.lower():
                tt_html = f"<span style='color:#9a6700;font-weight:600;'>⚠️ {tt}</span>"
            else:
                tt_html = f"<span style='color:#cf222e;font-weight:600;'>⏳ {tt}</span>"

            rows_html.append(f"""
            <tr style='background:{row_bg};border-bottom:1px solid #e1e4e8;font-size:12px;'>
                <td style='padding:8px 10px;font-family:Consolas,monospace;font-weight:bold;color:#0969da;'>{mat}</td>
                <td style='padding:8px 10px;text-align:center;'>{mdl_html}</td>
                <td style='padding:8px 10px;text-align:center;font-weight:600;'>{shift}</td>
                <td style='padding:8px 10px;color:#57606a;'>{stn}</td>
                <td style='padding:8px 10px;text-align:center;'>{sp_html}</td>
                <td style='padding:8px 10px;text-align:right;font-weight:600;'>{sl_ord:.0f}</td>
                <td style='padding:8px 10px;text-align:right;font-weight:600;color:#1a7f37;'>{sl_cap:.0f}</td>
                <td style='padding:8px 10px;text-align:center;'>{gap_html}</td>
                <td style='padding:8px 10px;text-align:center;'>{tt_html}</td>
            </tr>
            """)

        table_rows_str = "".join(rows_html)

        html = f"""
        <div style='font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;background:#ffffff;border:1px solid #d0d7de;border-radius:8px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.06);margin-bottom:12px;'>
            <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;border-bottom:2px solid #0969da;padding-bottom:8px;'>
                <div style='font-size:15px;font-weight:bold;color:#24292f;'>
                    📋 BẢNG THEO DÕI CẤP HÀNG NVL{model_badge}
                </div>
                <div style='font-size:11px;color:#57606a;'>
                    Thư mục: <code>{tracking_file or 'THEO DOI ODER NVL'}</code>
                </div>
            </div>

            <!-- Thống kê nhanh KPI -->
            <div style='display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:12px;'>
                <div style='background:#f6f8fa;border:1px solid #d0d7de;padding:8px;border-radius:6px;text-align:center;'>
                    <div style='font-size:11px;color:#57606a;'>Số dòng khớp</div>
                    <div style='font-size:16px;font-weight:bold;color:#24292f;'>{len(records)}</div>
                </div>
                <div style='background:#f6f8fa;border:1px solid #d0d7de;padding:8px;border-radius:6px;text-align:center;'>
                    <div style='font-size:11px;color:#57606a;'>Tổng SL Order</div>
                    <div style='font-size:16px;font-weight:bold;color:#0969da;'>{total_order:.0f}</div>
                </div>
                <div style='background:#dafbe1;border:1px solid #aceebb;padding:8px;border-radius:6px;text-align:center;'>
                    <div style='font-size:11px;color:#1a7f37;'>Đã cấp</div>
                    <div style='font-size:16px;font-weight:bold;color:#1a7f37;'>{total_cap:.0f}</div>
                </div>
                <div style='background:#ffebe9;border:1px solid #ff8182;padding:8px;border-radius:6px;text-align:center;'>
                    <div style='font-size:11px;color:#cf222e;'>Tổng thiếu (Gap)</div>
                    <div style='font-size:16px;font-weight:bold;color:#cf222e;'>-{total_gap:.0f} PC</div>
                </div>
            </div>

            <!-- Bảng chi tiết -->
            <div style='overflow-x:auto;'>
                <table style='width:100%;border-collapse:collapse;text-align:left;'>
                    <thead>
                        <tr style='background:#f6f8fa;border-bottom:2px solid #d0d7de;font-size:12px;color:#24292f;'>
                            <th style='padding:8px 10px;'>Mã NVL</th>
                            <th style='padding:8px 10px;text-align:center;'>Model</th>
                            <th style='padding:8px 10px;text-align:center;'>Ca chạy</th>
                            <th style='padding:8px 10px;'>Trạm cấp</th>
                            <th style='padding:8px 10px;text-align:center;'>Số phiếu</th>
                            <th style='padding:8px 10px;text-align:right;'>SL Order</th>
                            <th style='padding:8px 10px;text-align:right;'>Đã cấp</th>
                            <th style='padding:8px 10px;text-align:center;'>Cấp thiếu (Gap)</th>
                            <th style='padding:8px 10px;text-align:center;'>TT Nhặt hàng</th>
                        </tr>
                    </thead>
                    <tbody>
                        {table_rows_str}
                    </tbody>
                </table>
            </div>
        </div>
        """
        return html

    @classmethod
    def generate_trace_markdown(
        cls,
        records: List[Dict[str, Any]],
        hint_model: Optional[str] = None
    ) -> str:
        """Tạo bảng Markdown chuẩn GFM kèm Quick Copy cho Zalo/Teams."""
        if not records:
            return "Không có dữ liệu cấp NVL."

        hint_upper = hint_model.strip().upper() if hint_model else ""
        md_lines = []
        md_title = f"### 📋 BẢNG THEO DÕI CẤP HÀNG NVL (MODEL {hint_upper})" if hint_upper else "### 📋 BẢNG THEO DÕI CẤP HÀNG NVL"
        md_lines.append(md_title)
        md_lines.append("")
        md_lines.append("| Mã NVL | Model | Ca chạy | Trạm cấp | Số phiếu | SL Order | Đã cấp | Cấp thiếu (Gap) | TT Nhặt hàng |")
        md_lines.append("| :--- | :---: | :---: | :--- | :---: | :---: | :---: | :---: | :---: |")

        quick_copy = []
        for r in records:
            mat = r.get("material", "")
            mdl = r.get("model", "")
            shift = r.get("shift_chay", "")
            stn = r.get("station", "")
            sp = r.get("sp_display", "Chưa ra đơn")
            sl_ord = r.get("order_qty", 0)
            sl_cap = r.get("total_cap", 0)
            gap = r.get("gap", 0)
            gap_disp = r.get("gap_display", f"{gap:.0f}")
            tt = r.get("status_pick", "Chưa Cấp")
            is_tgt = r.get("is_target_model", False) or (hint_upper and hint_upper in mdl.upper())

            mdl_disp = f"**{mdl}**" if is_tgt else mdl
            gap_disp_md = f"**{gap_disp}**" if gap < 0 else gap_disp
            md_lines.append(
                f"| `{mat}` | {mdl_disp} | {shift} | {stn} | {sp} | {sl_ord:.0f} | {sl_cap:.0f} | {gap_disp_md} | {tt} |"
            )

            if gap < 0:
                quick_copy.append(f"- `{mat}` ({mdl} - {shift}): {sp} | Order: {sl_ord:.0f} | Đã cấp: {sl_cap:.0f} | **Thiếu: {abs(gap):.0f} PC** ({tt})")
            else:
                quick_copy.append(f"- `{mat}` ({mdl} - {shift}): {sp} | Order: {sl_ord:.0f} | Đã cấp: {sl_cap:.0f} | Đủ ({tt})")

        md_lines.append("")
        if quick_copy:
            md_lines.append("💬 **Quick Copy cho Zalo / Teams xưởng:**")
            md_lines.extend(quick_copy)

        return "\n".join(md_lines)
