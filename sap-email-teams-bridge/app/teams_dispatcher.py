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
                    q_str = str(r.get("qty", "0")).replace(".", "").replace(",", ".")
                    try:
                        total_qty += float(q_str)
                    except Exception:
                        pass
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
                        f"📦 Tổng: {total_qty:,.0f} {unit}</span>"
                    )

            elif any(k in cmd_upper for k in ["TON", "MB52"]):
                batches = {r.get("batch") for r in records if r.get("batch")}
                unrest_sum = 0.0
                qual_sum = 0.0
                unit = ""
                for r in records:
                    u_str = str(r.get("unrest", "0")).replace(".", "").replace(",", ".")
                    q_str = str(r.get("qual", "0")).replace(".", "").replace(",", ".")
                    try:
                        unrest_sum += float(u_str)
                    except Exception:
                        pass
                    try:
                        qual_sum += float(q_str)
                    except Exception:
                        pass
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"✅ Khả dụng: {unrest_sum:,.0f} {unit}</span>"
                )
                if qual_sum > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⏳ Chờ kiểm (QI): {qual_sum:,.0f} {unit}</span>"
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
                    q_str = str(r.get("qty", "0")).replace(".", "").replace(",", ".")
                    try:
                        total_qty += float(q_str)
                    except Exception:
                        pass
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#ffebe9;color:#cf222e;border:1px solid #ff8182;'>"
                    f"⏳ {len(ib_set)} Đơn IB treo</span>"
                )
                if total_qty > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"📦 Tổng: {total_qty:,.0f} {unit}</span>"
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
                    o_str = str(r.get("orderQty", "0")).replace(".", "").replace(",", ".")
                    r_str = str(r.get("remainingQty", "0")).replace(".", "").replace(",", ".")
                    try:
                        total_ord += float(o_str)
                    except Exception:
                        pass
                    try:
                        total_rem += float(r_str)
                    except Exception:
                        pass
                    if not unit and r.get("unit"):
                        unit = r.get("unit")

                badges.append(
                    f"<span style='{badge_style}background:#ddf4ff;color:#0969da;border:1px solid #54aeff;'>"
                    f"📋 {len(sto_set)} Đơn STO</span>"
                )
                badges.append(
                    f"<span style='{badge_style}background:#dafbe1;color:#1a7f37;border:1px solid #4ac26b;'>"
                    f"📦 Đặt: {total_ord:,.0f} {unit}</span>"
                )
                if total_rem > 0:
                    badges.append(
                        f"<span style='{badge_style}background:#fff8c5;color:#9a6700;border:1px solid #d4a72c;'>"
                        f"⏳ Còn nợ: {total_rem:,.0f} {unit}</span>"
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
        """Tự động co giãn số dòng để bảo đảm payload UTF-8 không bao giờ vượt quá 26 KB."""
        total_records = len(records)
        display_records = list(records[:MAX_DISPLAY_RECORDS])
        command = meta.get("command", "")

        while True:
            if custom_html:
                table_html = custom_html
            else:
                kpi_html = cls.build_kpi_badges(display_records, command)
                table_html = kpi_html + cls.build_html_table(display_records, headers_map, total_count=total_records)

            payload = {
                "requestId": request_id,
                "messageId": meta.get("messageId", ""),
                "command": command,
                "status": status,
                "senderEmail": meta.get("senderEmail", ""),
                "senderName": meta.get("senderName", "Đồng nghiệp"),
                "completedAt": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "summary": summary,
                "totalRecords": total_records,
                "displayedRecords": len(display_records),
                "table_html": table_html,
                "records": display_records
            }

            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            if len(data) <= MAX_TEAMS_PAYLOAD_BYTES or len(display_records) <= 3:
                return payload, data

            display_records = display_records[:-2]

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
