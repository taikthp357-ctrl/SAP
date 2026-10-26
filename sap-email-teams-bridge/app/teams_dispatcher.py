# -*- coding: utf-8 -*-
"""Xử lý gửi thông báo kết quả tới Teams Webhook qua Power Automate.

Tối ưu hóa dung lượng (Payload Budgeting):
- Tuân thủ nghiêm ngặt giới hạn 28 KB của Microsoft Teams Webhook.
- Tự động lấy N dòng dữ liệu gần nhất (mặc định 10-12 dòng) để chống tràn màn hình điện thoại.
- Cơ chế cắt giảm động (Dynamic Truncation) đảm bảo payload JSON luôn < 26 KB (không bao giờ dính lỗi HTTP 413).
"""

import html
import json
import urllib.request
import urllib.error
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from app.number_utils import parse_sap_float, format_qty_display

try:
    from app.config import DEFAULT_CONFIG
    SLOC_MASTER = DEFAULT_CONFIG.get("sloc_master", {})
except Exception:
    SLOC_MASTER = {}

MAX_DISPLAY_RECORDS = 50          # Số dòng tối đa hiển thị trực quan trên Teams
MAX_TEAMS_PAYLOAD_BYTES = 26000   # Trần dung lượng an toàn (dưới mốc 28 KB của Teams)


class TeamsDispatcher:
    """Đóng gói và gửi payload kết quả tới Webhook Teams."""

    @staticmethod
    def build_html_table(
        records: List[Dict[str, Any]],
        headers_map: Dict[str, str],
        total_count: Optional[int] = None
    ) -> str:
        """Tạo bảng HTML có sẵn Inline CSS hiển thị tối ưu trên Microsoft Teams."""
        if not records:
            return "<p style='color:#d9534f;font-weight:bold;'>⚠️ Không tìm thấy dữ liệu thỏa mãn điều kiện trên SAP.</p>"

        total = total_count if total_count is not None else len(records)
        displayed = len(records)

        th_style = (
            "background-color:#005a9e;color:#ffffff;padding:5px 8px;"
            "border:1px solid #d0d7de;text-align:left;font-size:12px;white-space:nowrap;"
        )
        td_style = (
            "padding:5px 8px;border:1px solid #d0d7de;text-align:left;font-size:12px;"
        )

        tbl = ["<table style='border-collapse:collapse;width:100%;font-family:Segoe UI,sans-serif;'>"]
        tbl.append("<thead><tr>")
        for label in headers_map.values():
            tbl.append(f"<th style='{th_style}'>{html.escape(str(label))}</th>")
        tbl.append("</tr></thead><tbody>")

        for idx, row in enumerate(records):
            bg_color = "#f6f8fa" if idx % 2 == 1 else "#ffffff"
            tbl.append(f"<tr style='background-color:{bg_color};'>")
            for key in headers_map.keys():
                val = row.get(key, "")
                desc = SLOC_MASTER.get(str(val).strip().upper(), "") if key == "sloc" else ""
                esc_val = html.escape(str(val))
                if desc:
                    esc_desc = html.escape(str(desc))
                    tbl.append(f"<td style='{td_style}' title='{esc_desc}'>{esc_val}</td>")
                else:
                    tbl.append(f"<td style='{td_style}'>{esc_val}</td>")
            tbl.append("</tr>")
        tbl.append("</tbody></table>")

        # Bỏ dòng thông báo giới hạn theo yêu cầu sếp

        return "".join(tbl)

    @classmethod
    def build_kpi_badges(cls, records: List[Dict[str, Any]], command: str) -> str:
        """Tạo thanh thẻ tóm tắt KPI (Executive Summary Badges) ngay trên bảng dữ liệu."""
        if not records:
            return ""

        cmd_upper = (command or "").upper()
        badges = []
        badge_style = (
            "display:inline-block;padding:3px 8px;margin-right:6px;margin-bottom:6px;"
            "border-radius:3px;font-size:11px;font-weight:600;font-family:Segoe UI,sans-serif;"
        )

        try:
            if any(k in cmd_upper for k in ["CHECK_ATT", "ZFIR266"]):
                total_cnt = len(records)
                full_cnt = sum(1 for r in records if r.get("status_code") == "FULL_ATT" or "Đã có cả" in str(r.get("overall_status", "")))
                missing_vat = sum(1 for r in records if r.get("status_code") == "MISSING_VAT" or "thiếu VAT" in str(r.get("overall_status", "")))
                missing_bb = sum(1 for r in records if r.get("status_code") == "MISSING_BBBG" or "thiếu BBBG" in str(r.get("overall_status", "")))
                no_att = sum(1 for r in records if r.get("status_code") == "NO_ATT" or "Chưa có cả" in str(r.get("overall_status", "")))

                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"✅ Đủ cả 2 file: {full_cnt}/{total_cnt} DOC</span>"
                )
                if missing_vat > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⚠️ Thiếu VAT: {missing_vat}</span>"
                    )
                if missing_bb > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⚠️ Thiếu BBBG: {missing_bb}</span>"
                    )
                if no_att > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"❌ Chưa có file nào: {no_att}</span>"
                    )

            elif any(k in cmd_upper for k in ["CHECK_FG_OB", "DIEN_OB"]):
                tot_fg = sum(parse_sap_float(r.get("fgStock", 0.0)) for r in records if r.get("fgStock") != "N/A")
                tot_ob = sum(parse_sap_float(r.get("openObQty", 0.0)) for r in records if r.get("openObQty") != "N/A")
                tot_avail = sum(parse_sap_float(r.get("availableQty", 0.0)) for r in records if r.get("availableQty") != "N/A")
                unit = next((r.get("unit") for r in records if r.get("unit")), "PC")
                has_error = any("Lỗi" in str(r.get("workAdvice", "")) or "Chưa" in str(r.get("workAdvice", "")) for r in records)

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🏢 Tồn FG 3093: {format_qty_display(tot_fg)}</span>"
                )
                if tot_ob > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"🚚 OB chờ điền: {format_qty_display(tot_ob)}</span>"
                    )
                if has_error:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⚠️ Một nguồn gặp sự cố</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ Khả dụng xuất: {format_qty_display(tot_avail)}</span>"
                    )

            elif any(k in cmd_upper for k in ["CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL", "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING"]):
                def _is_nvl_ready(r):
                    if r.get("eligible") is not None:
                        return bool(r.get("eligible"))
                    adv = str(r.get("migoAdvice", ""))
                    return "Sẵn sàng" in adv and "chờ" not in adv.lower() and "chưa" not in adv.lower()
                tot_avail = sum(parse_sap_float(r.get("qty", 0.0)) for r in records if _is_nvl_ready(r))
                tot_all = sum(parse_sap_float(r.get("qty", 0.0)) for r in records if r.get("qty") not in ["0", "N/A"])
                bins = {r.get("bin") for r in records if r.get("bin") and r.get("bin") not in ["-", "Chưa có Bin", "Chưa xác định (Lỗi /SCWM/POST)"]}
                batches = {r.get("batch") for r in records if r.get("batch") and r.get("batch") != "-"}
                ready_count = sum(1 for r in records if _is_nvl_ready(r))
                qc_count = sum(1 for r in records if "QC" in str(r.get("migoAdvice", "")) or "chờ kiểm" in str(r.get("migoAdvice", "")).lower())
                unit = next((r.get("unit") for r in records if r.get("unit")), "PC")

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🏢 Plant 3093 - Kho NVL 1005</span>"
                )
                if bins:
                    badges.append(
                        f"<span style='{badge_style}background:#f0f3f6;color:#24292f;border:1px solid #d0d7de;'>"
                        f"📍 {len(bins)} Vị trí Bin</span>"
                    )
                if batches:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"🏷️ {len(batches)} Lô (Batch)</span>"
                    )
                # BẢO ĐẢM R02: tot_avail = 0 thì hiển thị Khả dụng = 0, TUYỆT ĐỐI KHÔNG dùng `tot_avail or tot_all`!
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"📦 Khả dụng: {format_qty_display(tot_avail)} {unit}</span>"
                )
                if qc_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⚠️ {qc_count} Lô chờ QC</span>"
                    )
                if ready_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"⚡ Sẵn sàng ra đơn MIGO 311</span>"
                    )
                elif tot_all == 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"🔴 Hết hàng kho 1005</span>"
                    )
                elif qc_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⚠️ Không có tồn khả dụng (Chờ QC/Khóa)</span>"
                    )

            elif any(k in cmd_upper for k in ["VITRI", "BIN", "POST"]):
                bins = {r.get("bin") for r in records if r.get("bin")}
                batches = {r.get("batch") for r in records if r.get("batch")}
                tot_avail = 0.0
                tot_stk = 0.0
                tot_held = 0.0
                unit = ""
                for r in records:
                    tot_avail += parse_sap_float(r.get("availQty") or r.get("qty"))
                    tot_stk += parse_sap_float(r.get("totalQty") or r.get("availQty") or r.get("qty"))
                    tot_held += parse_sap_float(r.get("heldQty", 0.0))
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📍 {len(bins)} Vị trí Bin</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                    f"🏷️ {len(batches)} Lô (Batch)</span>"
                )
                if tot_stk > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#f0f3f6;color:#24292f;border:1px solid #d0d7de;'>"
                        f"📦 Tổng tồn: {format_qty_display(tot_stk)} {unit}</span>"
                    )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"⚡ Khả dụng: {format_qty_display(tot_avail)} {unit}</span>"
                )

            elif (any(k in cmd_upper for k in ["TON", "MB52"]) and "THAP" not in cmd_upper):
                batches = {r.get("batch") for r in records if r.get("batch") and not r.get("is_transit") and str(r.get("sloc", "")).upper() != "TRANSIT"}
                unrest_sum = 0.0
                transit_sum = 0.0
                unit = ""
                sloc_totals = {}
                target_plant = ""
                for r in records:
                    sloc = str(r.get("sloc", "")).strip()
                    is_tr = bool(r.get("is_transit") or sloc.upper() == "TRANSIT")
                    u_val = parse_sap_float(r.get("unrest"))
                    t_val = parse_sap_float(r.get("transit", "0"))
                    if not target_plant and r.get("plant"):
                        target_plant = str(r.get("plant")).strip()
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                    if is_tr:
                        transit_sum += (t_val if t_val > 0 else u_val)
                    else:
                        unrest_sum += u_val
                        if sloc:
                            sloc_totals[sloc] = sloc_totals.get(sloc, 0.0) + u_val

                # 1. Badge Tổng tồn khả dụng (nếu có hàng transit thì ghi rõ Tồn khả dụng để không cộng lẫn)
                if transit_sum > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ Tồn khả dụng: {format_qty_display(unrest_sum)} {unit}</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ Tồn: {format_qty_display(unrest_sum)} {unit}</span>"
                    )

                # 2. Badge Hàng đang đi đường / In-Transit (tách biệt rõ ràng)
                if transit_sum > 0:
                    transit_label = "🚚 Chờ nhập (Transit từ 3093)" if target_plant in ("4007", "4004") else "🚚 Chờ nhập (Transit)"
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"{transit_label}: {format_qty_display(transit_sum)} {unit}</span>"
                    )

                # 3. Phân rã theo SLoc (nếu có từ 2 kho trở lên hoặc có transit)
                if len(sloc_totals) > 1 or transit_sum > 0 or ("1005" in sloc_totals and "3001" in sloc_totals):
                    if "1005" in sloc_totals:
                        badges.append(
                            f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                            f"🏢 Tồn kho 1005: {format_qty_display(sloc_totals['1005'])} {unit}</span>"
                        )
                    if "3001" in sloc_totals:
                        badges.append(
                            f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                            f"🏭 Tồn line nhỏ 3001: {format_qty_display(sloc_totals['3001'])} {unit}</span>"
                        )
                    for sloc_k, sloc_qty in sorted(sloc_totals.items()):
                        if sloc_k not in ("1005", "3001"):
                            badges.append(
                                f"<span style='{badge_style}background:#f0f3f6;color:#24292f;border:1px solid #d0d7de;'>"
                                f"📦 SLoc {sloc_k}: {format_qty_display(sloc_qty)} {unit}</span>"
                            )

                # 4. Badge Số lượng Lô (Batch)
                if batches:
                    badges.append(
                        f"<span style='{badge_style}background:#f0f3f6;color:#24292f;border:1px solid #d0d7de;'>"
                        f"🏷️ {len(batches)} Lô (Batch)</span>"
                    )

            elif any(k in cmd_upper for k in ["LICHSU", "MB51"]):
                docs = {r.get("docNo") for r in records if r.get("docNo")}
                batches = {r.get("batch") for r in records if r.get("batch")}
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📝 {len(records)} Giao dịch</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"📄 {len(docs)} Chứng từ</span>"
                )
                if batches:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"🏷️ {len(batches)} Lô</span>"
                    )

            elif any(k in cmd_upper for k in ["IB_TREO", "VL06I", "INBOUND"]):
                ib_set = {r.get("ibNo") for r in records if r.get("ibNo")}
                batches = {r.get("batch") for r in records if r.get("batch")}
                total_qty = 0.0
                unit = ""
                for r in records:
                    total_qty += parse_sap_float(r.get("qty"))
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                    f"⏳ {len(ib_set)} Đơn IB treo</span>"
                )
                if total_qty > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"📦 Tổng: {format_qty_display(total_qty)} {unit}</span>"
                    )
                if batches:
                    badges.append(
                        f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                        f"🏷️ {len(batches)} Lô</span>"
                    )

            elif any(k in cmd_upper for k in ["STO", "ME2N"]):
                sto_set = {r.get("stoNo") for r in records if r.get("stoNo")}
                total_ord = 0.0
                total_rem = 0.0
                unit = ""
                for r in records:
                    total_ord += parse_sap_float(r.get("orderQty"))
                    total_rem += parse_sap_float(r.get("remainingQty"))
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📋 {len(sto_set)} Đơn STO</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"📦 Đặt: {format_qty_display(total_ord)} {unit}</span>"
                )
                if total_rem > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⏳ Còn nợ: {format_qty_display(total_rem)} {unit}</span>"
                    )

            elif any(k in cmd_upper for k in ["TON_THAP", "LOW_STOCK"]):
                zero_count = sum(1 for r in records if "hết" in str(r.get("status", "")).lower() or (parse_sap_float(r.get("kho_ma", "0")) == 0 and parse_sap_float(r.get("line_ma", "0")) == 0))
                priority_body = sum(1 for r in records if "ưu tiên xuất" in str(r.get("status", "")).lower())
                priority_line = sum(1 for r in records if "ưu tiên kéo line" in str(r.get("status", "")).lower())
                transit_count = sum(1 for r in records if "chờ xuất" in str(r.get("status", "")).lower() or parse_sap_float(r.get("transit", "0")) > 0)

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📊 Báo cáo Tồn & Cảnh báo: {len(records)} mã linh kiện</span>"
                )
                if zero_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"🔴 {zero_count} Mã MA hết tồn</span>"
                    )
                if priority_body > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⚡ {priority_body} Mã ưu tiên xuất Body</span>"
                    )
                if priority_line > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#f0f3f6;color:#24292f;border:1px solid #d0d7de;'>"
                        f"🏭 {priority_line} Mã ưu tiên kéo line</span>"
                    )
                if transit_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"🚚 {transit_count} Mã có hàng chờ xuất</span>"
                    )

            elif any(k in cmd_upper for k in ["ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO"]):
                tot_items = len(records)
                success_count = sum(
                    1 for r in records
                    if r.get("is_succ") is True
                    or str(r.get("overall_status", "")).upper() in ("DONE", "HOÀN TẤT", "THÀNH CÔNG", "OK")
                    or any(k in str(r.get("overall_status", "")).lower() for k in ["done", "thành công", "thanh cong", "hoàn tất"])
                    or str(r.get("vat_status", "")).upper().startswith("DONE")
                    or str(r.get("bbbg_status", "")).upper().startswith("DONE")
                    or str(r.get("headtext_status", "")).upper() in ("HEADTEXT_DONE", "HEADTEXT_OK")
                )
                ht_count = sum(1 for r in records if "headtext_done" in str(r.get("headtext_status", "")).lower() or "headtext_ok" in str(r.get("headtext_status", "")).lower() or "mô phỏng" in str(r.get("headtext_msg", "")).lower())

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📑 {tot_items} Chứng từ MIGO</span>"
                )
                if ht_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"📝 {ht_count} Số phiếu HeadText</span>"
                    )
                if success_count == tot_items and tot_items > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"⚡ {success_count}/{tot_items} Thành công (100%)</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⚡ {success_count}/{tot_items} Thành công</span>"
                    )

            elif any(k in cmd_upper for k in ["CHECK_BOM", "BOM", "CS03"]):
                comp_count = sum(1 for r in records if r.get("component") not in ("", "-"))
                bom_ok = any("CÓ BOM" in str(r.get("status", "")).upper() for r in records)
                if bom_ok:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"🧩 Có BOM ({comp_count} Linh kiện con)</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⚠️ Không có BOM</span>"
                    )

            elif any(k in cmd_upper for k in ["TAO_BBBG", "XEM_BBBG", "BBBG"]):
                tot = len(records)
                matched = sum(1 for r in records if r.get("status") == "MATCHED")
                mismatch = sum(1 for r in records if r.get("status") == "MISMATCH_QTY")
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📑 Đối soát BBBG FG: {tot} dòng</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"✅ Khớp 100%: {matched}</span>"
                )
                if mismatch > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⚠️ Lệch SL: {mismatch}</span>"
                    )

            elif any(k in cmd_upper for k in ["RA_DON_NVL", "KE_HOACH_NVL"]):
                tot_items = len(records)
                tot_qty = sum(parse_sap_float(r.get("allocated_qty", 0.0)) for r in records)
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🏭 Ra đơn NVL V31 Zone: {tot_items} dòng</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"📦 Phân bổ: {format_qty_display(tot_qty)}</span>"
                )

            elif any(k in cmd_upper for k in ["DIEU_PHOI_OB", "RC65"]):
                tot_items = len(records)
                deficits = sum(1 for r in records if r.get("status") == "DEFICIT")
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🚚 Điều phối Outbound RC65: {tot_items} items</span>"
                )
                if deficits > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⛔ CHẶN CỔNG: {deficits} dòng thiếu tồn</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ CỔNG ĐẠT: Sẵn sàng xuất kho</span>"
                    )

            elif any(k in cmd_upper for k in ["CHECK_QUEUE", "QUEUE"]):
                badges.append(
                    f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                    f"⚠️ {len(records)} Lỗi Queue kẹt</span>"
                )

            elif any(k in cmd_upper for k in ["CHECK_OB", "VL06O", "OB"]):
                ob_set = {r.get("obNo") for r in records if r.get("obNo")}
                total_qty = sum(parse_sap_float(r.get("qty")) for r in records)
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📦 {len(ob_set)} Đơn OB mở</span>"
                )
                if total_qty > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"🔢 Tổng SL: {format_qty_display(total_qty)}</span>"
                    )

            elif any(k in cmd_upper for k in ["LICHSU_BIN", "MON", "WHT"]):
                cf_count = sum(1 for r in records if "Hoàn thành" in str(r.get("status", "")))
                can_count = sum(1 for r in records if "Đã hủy" in str(r.get("status", "")))
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🔄 {len(records)} Lệnh chuyển Bin</span>"
                )
                if cf_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ {cf_count} Hoàn thành</span>"
                    )
                if can_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⚠️ {can_count} Đã hủy (Cancel)</span>"
                    )

            elif any(k in cmd_upper for k in ["CHECK_IB", "IB_DETAIL", "KIEMTRA_IB"]):
                ib_set = {r.get("ibNo") for r in records if r.get("ibNo")}
                match_count = sum(1 for r in records if "Khớp" in str(r.get("qtyCheck", "")))
                diff_count = sum(1 for r in records if "Lệch" in str(r.get("qtyCheck", "")) or "Thiếu" in str(r.get("qtyCheck", "")))
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🚚 IB: {', '.join(ib_set) if ib_set else len(records)} ({len(records)} Item)</span>"
                )
                if match_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ {match_count} Khớp SL</span>"
                    )
                if diff_count > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⚠️ {diff_count} Lệch/Thiếu SL</span>"
                    )

            elif any(k in cmd_upper for k in ["CHECK_PO", "PO_STATUS", "KIEMTRA_PO"]):
                po_set = {r.get("poNo") for r in records if r.get("poNo")}
                open_items = sum(1 for r in records if parse_sap_float(r.get("openQty", 0.0)) > 0)
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📋 PO: {', '.join(po_set) if po_set else len(records)} ({len(records)} Item)</span>"
                )
                if open_items > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⏳ {open_items} Item còn nợ hàng</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ Đã giao đủ</span>"
                    )

            elif any(k in cmd_upper for k in ["TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR"]):
                gr_docs = {r.get("mblnr") for r in records if r.get("mblnr")}
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📑 {len(gr_docs)} Chứng từ GR ({len(records)} dòng)</span>"
                )

            elif any(k in cmd_upper for k in ["TIM_IB", "FIND_IB"]):
                ib_set = {r.get("ibNo") for r in records if r.get("ibNo")}
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🔍 Tìm thấy {len(ib_set)} Inbound Delivery</span>"
                )

            elif any(k in cmd_upper for k in ["CHECK_FG_OB", "DIEN_OB"]):
                tot_fg = sum(parse_sap_float(r.get("fgStock", 0.0)) for r in records)
                tot_ob = sum(parse_sap_float(r.get("openObQty", 0.0)) for r in records)
                tot_avail = sum(parse_sap_float(r.get("availableQty", 0.0)) for r in records)
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"🏢 Tồn FG 3093: {format_qty_display(tot_fg)}</span>"
                )
                if tot_ob > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"🚚 OB chờ điền: {format_qty_display(tot_ob)}</span>"
                    )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"✅ Khả dụng xuất: {format_qty_display(tot_avail)}</span>"
                )

            elif any(k in cmd_upper for k in ["CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL"]):
                tot_wh = sum(parse_sap_float(r.get("whStock", 0.0)) for r in records)
                tot_line = sum(parse_sap_float(r.get("lineStock", 0.0)) for r in records)
                need_migo = sum(1 for r in records if "Cần ra đơn MIGO" in str(r.get("migoAdvice", "")))
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📦 Kho NVL 1005: {format_qty_display(tot_wh)}</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#f0f3f6;color:#24292f;border:1px solid #d0d7de;'>"
                    f"🏭 Line Body: {format_qty_display(tot_line)}</span>"
                )
                if need_migo > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                        f"⚡ {need_migo} Mã cần cấp MIGO 311</span>"
                    )
                else:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"✅ Line đủ hàng</span>"
                    )
        except Exception:
            return ""

        if not badges:
            return ""

        return (
            "<div style='margin-bottom:8px;padding-top:2px;'>"
            + "".join(badges)
            + "</div>"
        )

    @staticmethod
    def build_help_html(sender_name: str = "Đồng nghiệp") -> str:
        """Tạo thẻ HTML hướng dẫn cú pháp thân thiện gửi trực tiếp lên Teams."""
        return (
            f"<div style='font-family:Segoe UI,sans-serif;background:#f0f7ff;border:1px solid #cce4ff;"
            f"border-left:4px solid #005a9e;padding:12px 14px;border-radius:4px;color:#24292f;'>"
            f"<div style='font-weight:bold;color:#005a9e;font-size:13px;margin-bottom:6px;'>"
            f"👋 Chào {sender_name}! Bot sẵn sàng hỗ trợ tra cứu SAP EWM/ECC:</div>"
            f"<div style='font-size:12px;margin-bottom:8px;'>Bạn chỉ cần tag tên bot và nhập nội dung tự nhiên. Ví dụ:</div>"
            f"<ul style='margin:0 0 10px 18px;padding:0;font-size:12px;line-height:1.6;'>"
            f"<li>📍 <strong>Tra cứu Vị trí lưu kho & Giờ GR (/SCWM/POST):</strong><br>"
            f"<code>@Bot check vị trí 32013765 lô B2610</code> hoặc <code>@Bot 32013765</code></li>"
            f"<li>📦 <strong>Kiểm tra Tồn kho tổng hợp (MB52):</strong><br>"
            f"<code>@Bot kiểm tra tồn kho 32013765 plant 3093</code></li>"
            f"<li>⏱️ <strong>Xem Lịch sử xuất nhập (MB51):</strong><br>"
            f"<code>@Bot xem lịch sử 32013765 trong 14 ngày</code></li>"
            f"<li>🚚 <strong>Kiểm tra Inbound (IB) treo chưa GR (VL06I):</strong><br>"
            f"<code>@Bot check ib treo 32013765</code> hoặc <code>@Bot 32013765 có ib treo không</code></li>"
            f"<li>📋 <strong>Kiểm tra đơn điều chuyển STO (ME2N):</strong><br>"
            f"<code>@Bot check sto 32013765</code> hoặc <code>@Bot kiểm tra đơn sto 32013765</code></li>"
            f"<li>⚠️ <strong>Cảnh báo tồn thấp theo dòng xe (BODY 3093+3001 / XMĐ 3093+2001+3001):</strong><br>"
            f"<code>@Bot check cho tôi 20 mã limo tồn thấp</code> hoặc <code>@Bot check tồn thấp xmd</code></li>"
            f"<li>📦 <strong>Kiểm tra đơn Outbound (OB) đang mở (VL06O):</strong><br>"
            f"<code>@Bot check ob có không mã 32013765</code> hoặc <code>@Bot check ob 80012345</code></li>"
            f"<li>🔄 <strong>Xem Lịch sử chuyển Bin-to-Bin & Cancel WHT (/SCWM/MON):</strong><br>"
            f"<code>@Bot lịch sử bin 32013765</code> hoặc <code>@Bot xem lịch sử bin 10-A-01</code></li>"
            f"<li>📑 <strong>Tra cứu Chứng từ nhập kho GR (MB51):</strong><br>"
            f"<code>@Bot tìm chứng từ gr theo inbound 1800123456</code> hoặc <code>@Bot chứng từ gr 5000123456</code></li>"
            f"<li>🚚 <strong>Kiểm tra Chi tiết Inbound & Đối chiếu SL (VL06I):</strong><br>"
            f"<code>@Bot chi tiết inbound 1800123456 sl 500</code> hoặc <code>@Bot kiểm tra inbound 1800123456</code></li>"
            f"<li>🔍 <strong>Tìm kiếm đơn Inbound theo PO / Vận đơn (VL06I):</strong><br>"
            f"<code>@Bot tìm inbound theo PO 4500123456</code> hoặc <code>@Bot tìm inbound bill BL123456</code></li>"
            f"<li>📋 <strong>Kiểm tra Tình trạng đơn hàng PO/SA (ME2N):</strong><br>"
            f"<code>@Bot kiểm tra PO 4500123456</code> hoặc <code>@Bot kiểm tra PO 4500123456 item 10</code></li>"
            f"<li>🏢 <strong>Điều độ Thành Phẩm FG & Đơn OB (VL06O + MB52):</strong><br>"
            f"<code>@Bot điền ob 32013765</code> hoặc <code>@Bot check ob fg 32013765</code></li>"
            f"<li>🏭 <strong>Kiểm tra Posting Change NVL Kho 1005 (Plant 3093) & Ra đơn MIGO 311:</strong><br>"
            f"<code>@Bot check nvl 32013765</code> hoặc <code>@Bot check postingchange 32013765</code></li>"
            f"<li>📎 <strong>Kiểm tra File đính kèm BBBG & VAT theo DOC (ZFIR266 / MIGO GOS):</strong><br>"
            f"<code>@Bot check att 5007145690</code> hoặc <code>@Bot kiểm tra đính kèm 5007145690 5007145691</code></li>"
            f"</ul>"
            f"<div style='font-size:11px;color:#57606a;'>"
            f"💡 <em>Mẹo: Bạn có thể gõ mã 8 số rút gọn (như <code>32013765</code>), hệ thống sẽ tự động "
            f"đối chiếu sang mã SAP đầy đủ <code>BIW32013765AA</code> qua master data.</em></div>"
            f"</div>"
        )

    @classmethod
    def prepare_safe_payload(
        cls,
        request_id: str,
        meta: Dict[str, Any],
        status: str,
        summary: str,
        records: List[Dict[str, Any]],
        headers_map: Dict[str, str],
        custom_html: Optional[str] = None
    ) -> Tuple[Dict[str, Any], bytes]:
        """Tự động co giãn và cắt giảm đảm bảo payload UTF-8 TUYỆT ĐỐI không bao giờ vượt quá 26 KB."""
        total_records = len(records)
        display_records = list(records[:MAX_DISPLAY_RECORDS])
        command = meta.get("command", "")

        # TÍNH KPI TRÊN TOÀN BỘ DỮ LIỆU THỰC TẾ (Không bị giới hạn bởi số dòng hiển thị)
        kpi_html = cls.build_kpi_badges(records, command) if not custom_html else ""

        # Khởi tạo HTML ban đầu đảm bảo luôn tồn tại biến kể cả khi records rỗng hoặc dùng custom_html
        if custom_html:
            cur_html = custom_html
        else:
            cur_html = kpi_html + cls.build_html_table(display_records, headers_map, total_count=total_records)

        cur_summary = summary

        def _make_payload(cur_recs, cur_tbl_html, cur_sum, include_records_array=True):
            cmd_u = (command or "").upper()
            if status == "FAILED":
                greeting = "Sếp ơi, lần này em chưa tra cứu được SAP ạ. Chi tiết em để bên dưới nhé."
            elif status == "INFO":
                greeting = "Sếp ơi, em gửi hướng dẫn tra cứu bên dưới ạ 👇"
            elif cmd_u == "PING":
                greeting = "Sếp ơi, em nhận được yêu cầu kiểm tra kết nối rồi ạ 👇"
            elif total_records == 0:
                greeting = "Sếp ơi, em chưa tìm thấy dữ liệu phù hợp ạ. Sếp kiểm tra lại mã hàng, lô hoặc plant giúp em nhé."
            elif cmd_u in ["VITRI", "BIN", "POST"]:
                greeting = "Sếp ơi, em đã tra vị trí Bin trên hệ thống rồi ạ 👇"
            elif cmd_u in ["TON_KHO", "TONKHO"]:
                greeting = "Sếp ơi, em đã tra tồn kho (Kho 1005 & 3093) cho sếp rồi ạ 👇"
            elif cmd_u in ["TON", "MB52"]:
                greeting = "Sếp ơi, em đã tra tổng hợp tồn các kho cho sếp rồi ạ 👇"
            elif cmd_u in ["IB_TREO", "VL06I", "INBOUND"]:
                greeting = "Sếp ơi, em đã tra danh sách đơn Inbound (IB) treo cho sếp rồi ạ 👇"
            elif cmd_u in ["STO", "ME2N", "MB5T", "DIEUCHUYEN"]:
                greeting = "Sếp ơi, em đã tra danh sách đơn điều chuyển STO cho sếp rồi ạ 👇"
            elif cmd_u in ["LICHSU", "MB51"]:
                greeting = "Sếp ơi, em đã tra lịch sử giao dịch cho sếp rồi ạ 👇"
            elif cmd_u in ["TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR"]:
                greeting = "Sếp ơi, em đã tra chứng từ nhập kho GR (MB51) cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_IB", "IB_DETAIL", "KIEMTRA_IB"]:
                greeting = "Sếp ơi, em đã kiểm tra chi tiết đơn Inbound (VL06I) cho sếp rồi ạ 👇"
            elif cmd_u in ["TIM_IB", "FIND_IB"]:
                greeting = "Sếp ơi, em đã tìm kiếm đơn Inbound theo tham chiếu cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_PO", "PO_STATUS", "KIEMTRA_PO"]:
                greeting = "Sếp ơi, em đã kiểm tra tình trạng đơn PO (ME2N) cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_OB", "VL06O"]:
                greeting = "Sếp ơi, em đã kiểm tra đơn Outbound (OB) cho sếp rồi ạ 👇"
            elif cmd_u in ["LICHSU_BIN", "MON", "WHT"]:
                greeting = "Sếp ơi, em đã tra lịch sử chuyển Bin-to-Bin cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_FG_OB", "DIEN_OB"]:
                greeting = "Sếp ơi, em đã kiểm tra điều độ tồn FG và đơn OB mở cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL", "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING"]:
                greeting = "Sếp ơi, em đã kiểm tra Posting Change kho NVL 1005 (Plant 3093) cho sếp rồi ạ! 📋"
            elif cmd_u in ["TON_THAP", "LOW_STOCK"]:
                greeting = f"Sếp ơi, em đã kiểm tra cảnh báo tồn kho Body (Model {meta.get('model', 'LIMO')}) cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_BOM", "BOM", "CS03"]:
                greeting = "Sếp ơi, em đã kiểm tra cấu trúc định mức BOM (CS03) cho sếp rồi ạ 👇"
            elif cmd_u in ["CHECK_QUEUE", "QUEUE"]:
                greeting = "Sếp ơi, em đã quét danh sách kẹt Queue EWM (/SCWM/MON) cho sếp rồi ạ 👇"
            elif cmd_u in ["ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO"]:
                if not meta.get("folderPath"):
                    greeting = "Sếp ơi, em đã tự động điền số phiếu vào Header Text MIGO cho sếp rồi ạ 👇"
                else:
                    greeting = "Sếp ơi, em đã thực hiện xong tự động đính kèm và điền số phiếu MIGO cho sếp rồi ạ 👇"
            else:
                greeting = "Sếp ơi, em tra được rồi ạ. Kết quả của sếp đây 👇"

            pl = {
                "requestId": request_id,
                "messageId": meta.get("messageId", ""),
                "command": command,
                "status": status,
                "greeting": greeting,
                "senderEmail": meta.get("senderEmail", ""),
                "senderName": meta.get("senderName", "Đồng nghiệp"),
                "completedAt": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "summary": cur_sum,
                "totalRecords": total_records,
                "displayedRecords": len(cur_recs),
                "table_html": cur_tbl_html,
                "records": cur_recs if include_records_array else []
            }
            d = json.dumps(pl, ensure_ascii=False).encode("utf-8")
            return pl, d

        # Kiểm tra nhanh: nếu payload ban đầu đã dưới ngưỡng 26 KB thì trả về ngay
        payload, data = _make_payload(display_records, cur_html, cur_summary, include_records_array=True)
        if len(data) <= MAX_TEAMS_PAYLOAD_BYTES:
            return payload, data

        # Giai đoạn 1: Giảm dần số dòng hiển thị từ 12 xuống 1 (nếu có records)
        while len(display_records) > 1:
            display_records = display_records[:max(1, len(display_records) - 2)]
            if not custom_html:
                cur_html = kpi_html + cls.build_html_table(display_records, headers_map, total_count=total_records)

            payload, data = _make_payload(display_records, cur_html, cur_summary, include_records_array=True)
            if len(data) <= MAX_TEAMS_PAYLOAD_BYTES:
                return payload, data

        # Giai đoạn 2: Bỏ mảng JSON `records` nếu payload vẫn > 26 KB (chỉ giữ lại table_html)
        payload, data = _make_payload(display_records, cur_html, cur_summary, include_records_array=False)
        if len(data) <= MAX_TEAMS_PAYLOAD_BYTES:
            return payload, data

        # Giai đoạn 3: Cắt ngắn độ dài ký tự trong các ô nếu vẫn > 26 KB
        if display_records:
            truncated_records = [
                {k: (str(v)[:30] + "..." if len(str(v)) > 30 else v) for k, v in display_records[0].items()}
            ]
            cur_html = kpi_html + cls.build_html_table(truncated_records, headers_map, total_count=total_records)
            payload, data = _make_payload(truncated_records, cur_html, cur_summary, include_records_array=False)
            if len(data) <= MAX_TEAMS_PAYLOAD_BYTES:
                return payload, data

        # Giai đoạn 4: Thu gọn bảng sang Compact Card
        compact_card = (
            f"<div style='font-family:Segoe UI,sans-serif;padding:10px 12px;background:#fff8c5;"
            f"border-left:4px solid #d4a72c;font-size:12px;color:#24292f;'>"
            f"⚠️ <strong>Dữ liệu chi tiết ({total_records} dòng) vượt giới hạn kích thước Teams.</strong><br>"
            f"Hệ thống đã thu gọn để đảm bảo gửi thành công. Vui lòng lọc cụ thể hơn theo Lô hoặc Mã hàng."
            f"</div>"
        )
        cur_html = (kpi_html + compact_card) if not custom_html else custom_html
        payload, data = _make_payload([], cur_html, cur_summary, include_records_array=False)
        if len(data) <= MAX_TEAMS_PAYLOAD_BYTES:
            return payload, data

        # Giai đoạn 5: Cắt giảm triệt để summary và table_html (Bảo vệ tuyệt đối chống vượt 26 KB)
        if len(cur_summary) > 200:
            cur_summary = cur_summary[:200] + "..."
            payload, data = _make_payload([], cur_html, cur_summary, include_records_array=False)
            if len(data) <= MAX_TEAMS_PAYLOAD_BYTES:
                return payload, data

        # Cắt ngắn custom_html hoặc table_html nếu bản thân HTML quá dài (> 20 KB)
        while len(data) > MAX_TEAMS_PAYLOAD_BYTES and len(cur_html) > 200:
            cur_html = cur_html[:int(len(cur_html) * 0.7)] + "</div>"
            payload, data = _make_payload([], cur_html, cur_summary, include_records_array=False)

        # Cắt ngắn summary xuống tối đa nếu vẫn còn vượt
        if len(data) > MAX_TEAMS_PAYLOAD_BYTES:
            cur_summary = cur_summary[:50] + "..."
            payload, data = _make_payload([], "<p>Kết quả quá dài, vui lòng thu hẹp phạm vi tìm kiếm.</p>", cur_summary, include_records_array=False)

        return payload, data

    @classmethod
    def send_reply(
        cls,
        webhook_url: str,
        request_id: str,
        meta: Dict[str, Any],
        status: str,
        summary: str,
        records: List[Dict[str, Any]],
        headers_map: Dict[str, str],
        custom_html: Optional[str] = None
    ) -> bool:
        """Gửi kết quả kèm kiểm soát dung lượng < 28 KB tới Teams."""
        # Tương thích nếu gọi theo dạng tắt send_reply(req_id, meta, status, ...)
        if isinstance(webhook_url, str) and (webhook_url.startswith("SAP-") or webhook_url.startswith("REQ-")):
            custom_html = headers_map if isinstance(headers_map, str) else None
            headers_map = records if isinstance(records, dict) else {}
            records = summary if isinstance(summary, list) else []
            summary = status if isinstance(status, str) else ""
            status = meta if isinstance(meta, str) else "SUCCEEDED"
            meta = request_id if isinstance(request_id, dict) else {}
            request_id = webhook_url
            try:
                from app.config import load_config
                webhook_url = load_config().get("teams_webhook_url", "")
            except Exception:
                webhook_url = ""

        payload, data = cls.prepare_safe_payload(
            request_id, meta, status, summary, records, headers_map, custom_html
        )
        size_bytes = len(data)
        size_kb = size_bytes / 1024

        if not webhook_url or "YOUR_WEBHOOK_URL_HERE" in webhook_url:
            print(f"[DRY-RUN] Dữ liệu sẽ gửi lên Teams ({size_kb:.1f} KB / tối đa 28 KB):")
            print(f"  • Tổng dòng: {payload['totalRecords']} | Hiển thị: {payload['displayedRecords']}")
            return True

        req = urllib.request.Request(
            webhook_url,
            data=data,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST"
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                print(f"[TEAMS] Đã gửi kết quả Message ID '{meta.get('messageId')}' ({size_kb:.1f} KB - HTTP {resp.status})")
                return True
        except urllib.error.URLError as e:
            print(f"[ERROR] Lỗi gửi Webhook Teams: {e}")
            return False

    @classmethod
    def send_progress_update(
        cls,
        webhook_url: str,
        request_id: str,
        meta: Dict[str, Any],
        current: int,
        total: int,
        current_doc: str = "",
        success_count: int = 0
    ) -> bool:
        """Gửi thông báo tiến độ xử lý danh sách lớn (100 - 300 chứng từ) tới Teams mỗi 3-5 phút."""
        percent = int(current * 100 / total) if total > 0 else 0
        summary = f"Đang đính kèm & điền HeadText MIGO: {current}/{total} chứng từ ({percent}%)"
        progress_html = (
            f"<div style='font-family:Segoe UI,sans-serif;padding:12px;background:#f0f7ff;border:1px solid #cce4ff;"
            f"border-left:4px solid #0078d4;border-radius:4px;color:#24292f;'>"
            f"<div style='font-size:13px;font-weight:bold;color:#0078d4;margin-bottom:8px;'>"
            f"⏳ Tiến độ xử lý: {current}/{total} ({percent}%)</div>"
            f"<div style='background:#e0e0e0;border-radius:4px;height:10px;width:100%;overflow:hidden;margin-bottom:8px;'>"
            f"<div style='background:#0078d4;height:100%;width:{percent}%;'></div>"
            f"</div>"
            f"<div style='font-size:12px;color:#444;'>"
            f"Đã xử lý: <b>{current}</b> | Thành công: <b style='color:#107c10;'>{success_count}</b>"
            f" | Đang chạy DOC: <code>{current_doc}</code></div>"
            f"<div style='font-size:11px;color:#666;margin-top:6px;'>"
            f"⏱️ <em>Danh sách dài cần vài phút, sếp đợi em thêm chút nhé. Em đang tiếp tục tự động xử lý ạ!</em></div>"
            f"</div>"
        )
        payload, data = cls.prepare_safe_payload(
            request_id,
            meta,
            status="INFO",
            summary=summary,
            records=[],
            headers_map={},
            custom_html=progress_html
        )
        if not webhook_url:
            return False
        return cls.send_raw_payload(webhook_url, payload)

    @classmethod
    def deliver_payload(cls, webhook_url: str, payload: Dict[str, Any]) -> Dict[str, str]:
        """2xx means Flow accepted; transport failures have uncertain delivery."""
        if not webhook_url or "YOUR_WEBHOOK_URL_HERE" in webhook_url:
            return {"status": "SIMULATED", "detail": "Webhook not configured; nothing sent"}
        try:
            data = json.dumps(payload, ensure_ascii=False, allow_nan=False).encode("utf-8")
            req = urllib.request.Request(webhook_url, data=data,
                headers={"Content-Type": "application/json; charset=utf-8", "X-Request-ID": str(payload.get("requestId", ""))}, method="POST")
            with urllib.request.urlopen(req, timeout=15) as resp:
                code = int(resp.status)
                return {"status": "ACCEPTED" if 200 <= code < 300 else "AMBIGUOUS", "detail": f"HTTP {code}; delivery in Teams not verified"}
        except urllib.error.HTTPError as exc:
            return {"status": "REJECTED" if 400 <= exc.code < 500 else "AMBIGUOUS", "detail": f"HTTP {exc.code}"}
        except (urllib.error.URLError, TimeoutError, OSError):
            # Do not log URLs/tokens or assume a timeout prevented receipt.
            return {"status": "AMBIGUOUS", "detail": "Transport interrupted; check Flow before retry"}

    @classmethod
    def send_raw_payload(cls, webhook_url: str, payload: Dict[str, Any]) -> bool:
        return cls.deliver_payload(webhook_url, payload)["status"] == "ACCEPTED"

    @classmethod
    def send_to_teams(
        cls,
        webhook_url: str,
        request_id: str,
        command: str,
        status: str,
        sender: str,
        summary: str,
        records: List[Dict[str, Any]],
        headers_map: Dict[str, str]
    ) -> bool:
        """Hàm tương thích ngược."""
        meta = {
            "messageId": "",
            "command": command,
            "senderEmail": sender,
            "senderName": sender
        }
        return cls.send_reply(webhook_url, request_id, meta, status, summary, records, headers_map)
