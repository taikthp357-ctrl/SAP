# -*- coding: utf-8 -*-
"""Command Router - Điều phối các yêu cầu nghiệp vụ SAP đến đúng Handler chuyên trách."""
from typing import Dict, Optional, Type
from app.core.context import RequestContext, CommandResult
from app.handlers.base_handler import BaseHandler
from app.handlers.att_handler import AttHandler
from app.handlers.nvl_handler import NvlHandler
from app.handlers.stock_handler import StockHandler
from app.handlers.logistics_handler import LogisticsHandler
from app.handlers.production_handler import ProductionHandler
from app.handlers.bbbg_handler import BbbgHandler
from app.handlers.system_handler import SystemHandler


class CommandRouter:
    """Bộ điều phối trung tâm (Dispatcher) ánh xạ Intent sang Handler chuyên trách."""

    def __init__(self):
        # Khởi tạo các singleton handler
        self.att_handler = AttHandler()
        self.nvl_handler = NvlHandler()
        self.stock_handler = StockHandler()
        self.logistics_handler = LogisticsHandler()
        self.production_handler = ProductionHandler()
        self.bbbg_handler = BbbgHandler()
        self.system_handler = SystemHandler()

        # Bảng ánh xạ lệnh (Command / Intent) -> Handler
        self.routes: Dict[str, BaseHandler] = {}
        self._register_default_routes()

    def register(self, command: str, handler: BaseHandler) -> None:
        """Đăng ký một lệnh với handler tương ứng."""
        self.routes[command.strip().upper()] = handler

    def _register_default_routes(self) -> None:
        """Đăng ký toàn bộ các lệnh nghiệp vụ được hỗ trợ."""

        # 1. Nhóm Đính kèm & Tra cứu ZFIR266 (AttHandler)
        att_commands = [
            "CHECK_ATT", "CHECK_ATT_DOC", "CHECK_DINH_KEM", "ZFIR266",
            "ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO"
        ]
        for cmd in att_commands:
            self.register(cmd, self.att_handler)

        # 2. Nhóm Cấp phát NVL & MIGO 311 (NvlHandler)
        nvl_commands = [
            "RA_DON_NVL", "KE_HOACH_NVL",
            "CHECK_NVL_MIGO", "RA_DON_MIGO", "MIGO_311", "CHECK_NVL",
            "CHECK_POSTING_CHANGE", "POSTING_CHANGE", "POSTING"
        ]
        for cmd in nvl_commands:
            self.register(cmd, self.nvl_handler)

        # 3. Nhóm Tồn kho, Vị trí & Lịch sử (StockHandler)
        stock_commands = [
            "VITRI", "VỊ TRÍ", "BIN", "POST",
            "TON_KHO", "TONKHO",
            "TON", "TỒN", "MB52",
            "LICHSU", "LỊCH SỬ", "MB51",
            "TON_THAP", "LOW_STOCK",
            "LICHSU_BIN", "MON", "WHT",
            "TIM_GR_DOC", "GR_DOC", "CHUNGTU_GR"
        ]
        for cmd in stock_commands:
            self.register(cmd, self.stock_handler)

        # 4. Nhóm Mua hàng & Logistics (LogisticsHandler)
        logistics_commands = [
            "IB_TREO", "IB", "INBOUND", "VL06I",
            "STO", "ME2N", "MB5T", "DIEUCHUYEN",
            "CHECK_IB", "IB_DETAIL", "KIEMTRA_IB",
            "TIM_IB", "FIND_IB",
            "CHECK_PO", "PO_STATUS", "KIEMTRA_PO",
            "TRA_CUU_ASN", "ASN"
        ]
        for cmd in logistics_commands:
            self.register(cmd, self.logistics_handler)

        # 5. Nhóm Quản lý Sản xuất & Điều phối (ProductionHandler)
        prod_commands = [
            "CHECK_RES", "MB23", "MB25",
            "CHECK_OB", "VL06O",
            "CHECK_BOM", "BOM", "CS03",
            "CHECK_QUEUE", "QUEUE",
            "CHECK_FG_OB", "DIEN_OB",
            "DIEU_PHOI_OB", "RC65", "XAC_NHAN_OB"
        ]
        for cmd in prod_commands:
            self.register(cmd, self.production_handler)

        # 6. Nhóm Biên bản giao nhận Thành phẩm (BbbgHandler)
        bbbg_commands = [
            "TAO_BBBG", "XEM_BBBG", "BBBG_FG"
        ]
        for cmd in bbbg_commands:
            self.register(cmd, self.bbbg_handler)

        # 7. Nhóm Hệ thống & Tiện ích (SystemHandler)
        system_commands = [
            "PING", "WHO_IS_TAI", "HELP"
        ]
        for cmd in system_commands:
            self.register(cmd, self.system_handler)

    def get_handler(self, command: str) -> Optional[BaseHandler]:
        """Lấy handler chuyên trách cho lệnh."""
        cmd_key = (command or "").strip().upper()
        return self.routes.get(cmd_key)

    def dispatch(self, ctx: RequestContext) -> CommandResult:
        """Điều phối request tới đúng handler.
        
        Nếu không tìm thấy handler, trả về CommandResult thất bại với thông báo rõ ràng.
        """
        cmd_key = (ctx.command or "").strip().upper()
        handler = self.get_handler(cmd_key)

        if not handler:
            print(f"[ROUTER WARN] Không tìm thấy handler cho lệnh: '{ctx.command}'")
            return CommandResult(
                status="FAILED",
                summary=f"Lệnh không được hỗ trợ: '{ctx.command}'",
                records=[],
                headers={},
                custom_html=(
                    f"<div style='font-family:Segoe UI,sans-serif;background:#fff1f0;border:1px solid #ffa39e;"
                    f"border-left:4px solid #cf1322;padding:12px 14px;border-radius:4px;color:#24292f;'>"
                    f"<div style='font-weight:bold;color:#cf1322;font-size:14px;margin-bottom:6px;'>"
                    f"⚠️ Lệnh không xác định: <code>{ctx.command}</code></div>"
                    f"<div style='font-size:12px;line-height:1.6;'>"
                    f"Hệ thống chưa hỗ trợ xử lý cho lệnh này. Vui lòng gửi lệnh <code>HELP</code> để xem danh sách cú pháp."
                    f"</div></div>"
                ),
                audit_target=ctx.command,
                audit_message=f"Lệnh không tìm thấy handler: {ctx.command}"
            )

        # Gọi trực tiếp handler chuyên trách
        return handler.handle(ctx)


# Global default router instance
router = CommandRouter()
