# -*- coding: utf-8 -*-
"""Handler chuyên trách các lệnh hệ thống: PING, WHO_IS_TAI và HELP."""
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler


class SystemHandler(BaseHandler):
    """Xử lý nhóm lệnh hệ thống và tiện ích:
    - PING (Kiểm tra đường truyền kết nối)
    - WHO_IS_TAI (Giới thiệu kiến trúc sư trưởng)
    - HELP (Hướng dẫn sử dụng bot)
    """

    def handle(self, ctx: RequestContext) -> CommandResult:
        cmd = (ctx.command or "").upper()
        req_id_display = ctx.meta.get("requestId") or ctx.request_id

        # -------------------------------------------------------------
        # 1. LỆNH PING KẾT NỐI
        # -------------------------------------------------------------
        if cmd == "PING":
            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#dafbe1;border:1px solid #4ac26b;"
                f"border-left:4px solid #1a7f37;padding:12px 14px;border-radius:4px;color:#24292f;'>"
                f"<div style='font-weight:bold;color:#1a7f37;font-size:14px;margin-bottom:6px;'>"
                f"🏓 PONG! Kết nối hoàn tất thành công!</div>"
                f"<div style='font-size:12px;line-height:1.6;'>"
                f"• <strong>Request ID:</strong> {req_id_display}<br>"
                f"• <strong>Plant:</strong> {ctx.plant}<br>"
                f"• <strong>Trạng thái SAP:</strong> Kết nối bình thường<br>"
                f"• <strong>Ghi chú:</strong> {ctx.meta.get('note', 'Kiểm tra đường truyền hoàn tất.')}"
                f"</div></div>"
            )
            summary = f"Kiểm tra kết nối PING-PONG thành công (Request ID: {req_id_display})"

        # -------------------------------------------------------------
        # 2. LỆNH WHO_IS_TAI
        # -------------------------------------------------------------
        elif cmd == "WHO_IS_TAI":
            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#f0f8ff;border:1px solid #b6d4fe;"
                f"border-left:4px solid #0969da;padding:14px 16px;border-radius:6px;color:#24292f;'>"
                f"<div style='font-weight:bold;color:#0969da;font-size:15px;margin-bottom:8px;'>"
                f"🦾✨ Dạ kính chào sếp! Em xin phép được giới thiệu về 'Cha đẻ' của em ạ!</div>"
                f"<div style='font-size:13px;line-height:1.7;'>"
                f"• <strong>Tác giả / Kiến trúc sư trưởng:</strong> Kỹ sư <strong>Phạm Tiến Tài (FVN-KSX-NMVFHP)</strong>.<br>"
                f"• <strong>Sứ mệnh của em:</strong> Được anh Tài khai sinh, lập trình và nạp năng lượng 24/7 để đồng hành cùng các sếp và anh chị em xưởng sản xuất, "
                f"giải phóng mọi người khỏi cảnh quay cuồng giữa rừng mã SAP, F3 mỏi tay hay trích xuất Excel thủ công!<br>"
                f"• <strong>Năng lực cốt lõi:</strong> Soi vị trí Bin EWM (<code>/SCWM/POST</code>), kiểm tra tồn kho MB52, "
                f"săn đơn xuất OB (<code>VL06O</code>), kiểm tra hàng về IB (<code>VL06I</code>/<code>ZMME240</code>), theo dõi STO (<code>ME2N</code>), "
                f"tra cứu định mức BOM (<code>CS03</code>), đối soát BBBG và đính kèm MIGO GOS tự động.<br>"
                f"• <strong>Lời nhắn từ trợ lý:</strong> Sếp cần tra cứu mã hàng, vị trí hay đơn hàng nào, cứ ra lệnh cho em nhé! "
                f"Nếu em có điểm nào chưa chuẩn, sếp cứ 'méc' anh Tài để em được... nâng cấp thêm ạ! 😄📦"
                f"</div></div>"
            )
            summary = "Giới thiệu tác giả: Kỹ sư Phạm Tiến Tài (FVN-KSX-NMVFHP) - Kiến trúc sư trưởng SAP Bridge"

        # -------------------------------------------------------------
        # 3. LỆNH TRỢ GIÚP / HELP
        # -------------------------------------------------------------
        else:
            custom_html = (
                f"<div style='font-family:Segoe UI,sans-serif;background:#f6f8fa;border:1px solid #d0d7de;"
                f"border-left:4px solid #0969da;padding:12px 14px;border-radius:4px;color:#24292f;'>"
                f"<div style='font-weight:bold;color:#0969da;font-size:14px;margin-bottom:6px;'>"
                f"💡 Danh sách lệnh hỗ trợ chính của Trợ lý SAP Bridge</div>"
                f"<div style='font-size:12px;line-height:1.6;'>"
                f"• <strong>Tra cứu tồn & Vị trí:</strong> <code>VITRI [mã]</code>, <code>TON [mã]</code>, <code>LICHSU [mã]</code>, <code>TON_THAP [model]</code><br>"
                f"• <strong>Logistics & Mua hàng:</strong> <code>CHECK_IB [số IB]</code>, <code>CHECK_PO [số PO]</code>, <code>STO [mã]</code>, <code>ASN [số IB]</code><br>"
                f"• <strong>Sản xuất & Điều phối:</strong> <code>BOM [mã]</code>, <code>OB [số OB]</code>, <code>QUEUE</code><br>"
                f"• <strong>Đính kèm & MIGO:</strong> <code>CHECK_ATT [số DOC]</code>, <code>ATT_MIGO [thư mục]</code><br>"
                f"• <strong>Cấp phát NVL:</strong> <code>RA_DON_NVL [bảng hoặc mã kèm SL/Trạm]</code>"
                f"</div></div>"
            )
            summary = "Hướng dẫn cú pháp lệnh SAP Bridge"

        return CommandResult(
            status="SUCCEEDED",
            summary=summary,
            records=[],
            headers={},
            custom_html=custom_html,
            audit_target=req_id_display,
            audit_message=summary
        )
