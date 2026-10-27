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
import glob
import shutil
import datetime
from pathlib import Path
from decimal import Decimal
from typing import Dict, Any, List, Optional, Tuple

import openpyxl
from app.nvl_master import NvlMasterService


class NvlTrackingService:
    """Dịch vụ điều phối và đồng bộ luồng theo dõi cấp NVL."""

    DEFAULT_TRACKING_DIR = r"G:\SAP"

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
        all_files = glob.glob(pattern)
        if not all_files:
            # Tìm đệ quy hoặc tìm thêm định dạng xlsb
            all_files = glob.glob(str(search_dir / "**/*THEO DOI ODER NVL_P_*.xlsx"), recursive=True)

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

        wb = openpyxl.load_workbook(str(tracking_file), data_only=True)
        sheet_name = "ODER NVl" if "ODER NVl" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

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

        for r in range(2, ws.max_row + 1):
            mat_val = ws.cell(r, 4).value
            if not mat_val:
                continue

            mat = str(mat_val).strip().upper()
            row_model = str(ws.cell(r, 10).value or "").strip().upper()
            row_ca_nhap = str(ws.cell(r, 11).value or "").strip().upper()
            row_ca_chay = str(ws.cell(r, 12).value or "").strip().upper()

            # Bộ lọc Model (nếu được yêu cầu)
            if model_clean and model_clean not in row_model:
                continue

            # Bộ lọc Ca: Khớp ca chạy hoặc ca nhập
            if shift_clean and (shift_clean not in row_ca_chay and shift_clean not in row_ca_nhap):
                continue

            # Bộ lọc Ngày (nếu được yêu cầu)
            if target_date:
                date_cell = str(ws.cell(r, 32).value or ws.cell(r, 30).value or "")
                if target_date not in date_cell:
                    continue

            # Lấy thông tin số lượng
            try:
                sl_order = float(Decimal(str(ws.cell(r, 5).value or 0)))
            except Exception:
                sl_order = 0.0

            try:
                total_cap = float(Decimal(str(ws.cell(r, 26).value or 0)))
            except Exception:
                total_cap = 0.0

            try:
                gap_val = float(Decimal(str(ws.cell(r, 27).value or 0)))
            except Exception:
                gap_val = total_cap - sl_order

            station = str(ws.cell(r, 6).value or "").strip()
            ob_val = str(ws.cell(r, 13).value or ws.cell(r, 14).value or "").strip()
            status_order = str(ws.cell(r, 28).value or "Pending").strip()
            status_pick = str(ws.cell(r, 29).value or "Chưa Cấp").strip()
            note_val = str(ws.cell(r, 31).value or "").strip()

            # Lấy NCC
            ncc_val = str(ws.cell(r, 1).value or "").strip()
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
            tracking_file = candidates[0]

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
