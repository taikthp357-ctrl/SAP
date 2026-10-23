# -*- coding: utf-8 -*-
"""Xử lý gửi thông báo kết quả tới Teams Webhook qua Power Automate.

Tối ưu hóa dung lượng (Payload Budgeting):
- Tuân thủ nghiêm ngặt giới hạn 28 KB của Microsoft Teams Webhook.
- Tự động lấy N dòng dữ liệu gần nhất (mặc định 10-12 dòng) để chống tràn màn hình điện thoại.
- Cơ chế cắt giảm động (Dynamic Truncation) đảm bảo payload JSON luôn < 26 KB (không bao giờ dính lỗi HTTP 413).
"""

import json
import urllib.request
import urllib.error
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from app.number_utils import parse_sap_float, format_qty_display

MAX_DISPLAY_RECORDS = 12          # Số dòng tối đa hiển thị trực quan trên Teams
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

        html = ["<table style='border-collapse:collapse;width:100%;font-family:Segoe UI,sans-serif;'>"]
        html.append("<thead><tr>")
        for label in headers_map.values():
            html.append(f"<th style='{th_style}'>{label}</th>")
        html.append("</tr></thead><tbody>")

        for idx, row in enumerate(records):
            bg_color = "#f6f8fa" if idx % 2 == 1 else "#ffffff"
            html.append(f"<tr style='background-color:{bg_color};'>")
            for key in headers_map.keys():
                val = row.get(key, "")
                html.append(f"<td style='{td_style}'>{val}</td>")
            html.append("</tr>")
        html.append("</tbody></table>")

        # Nếu số dòng thực tế nhiều hơn số dòng hiển thị, thêm thanh thông báo
        if total > displayed:
            html.append(
                f"<div style='margin-top:6px;padding:6px 10px;background:#eef4fa;"
                f"border-left:4px solid #005a9e;font-size:12px;color:#24292f;font-family:Segoe UI,sans-serif;'>"
                f"📊 <strong>Đang hiển thị {displayed} / {total} dòng gần nhất</strong> "
                f"(Dữ liệu được nén tối ưu &lt; 28KB cho màn hình Teams di động)."
                f"</div>"
            )

        return "".join(html)

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
            if any(k in cmd_upper for k in ["VITRI", "BIN", "POST"]):
                bins = {r.get("bin") for r in records if r.get("bin")}
                batches = {r.get("batch") for r in records if r.get("batch")}
                total_qty = 0.0
                unit = ""
                for r in records:
                    total_qty += parse_sap_float(r.get("qty"))
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
                if total_qty > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                        f"📦 Tổng: {format_qty_display(total_qty)} {unit}</span>"
                    )

            elif any(k in cmd_upper for k in ["TON", "MB52"]):
                batches = {r.get("batch") for r in records if r.get("batch")}
                unrest_sum = 0.0
                qual_sum = 0.0
                unit = ""
                for r in records:
                    unrest_sum += parse_sap_float(r.get("unrest"))
                    qual_sum += parse_sap_float(r.get("qual"))
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"✅ Khả dụng: {format_qty_display(unrest_sum)} {unit}</span>"
                )
                if qual_sum > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⏳ Chờ kiểm (QI): {format_qty_display(qual_sum)} {unit}</span>"
                    )
                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
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
            pl = {
                "requestId": request_id,
                "messageId": meta.get("messageId", ""),
                "command": command,
                "status": status,
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
    def send_raw_payload(cls, webhook_url: str, payload: Dict[str, Any]) -> bool:
        """Gửi trực tiếp payload đã chuẩn hóa (dùng cho hàng đợi outbox retry)."""
        if not webhook_url or "YOUR_WEBHOOK_URL_HERE" in webhook_url:
            print(f"[DRY-RUN] Gửi raw payload: {payload.get('requestId')}")
            return True

        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            webhook_url,
            data=data,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST"
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                print(f"[TEAMS RAW] Đã gửi lại payload {payload.get('requestId')} (HTTP {resp.status})")
                return True
        except urllib.error.URLError as e:
            print(f"[TEAMS RAW ERROR] Lỗi gửi raw payload {payload.get('requestId')}: {e}")
            return False

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
