# -*- coding: utf-8 -*-
"""Bảo mật và kiểm soát quyền ghi (Write Permission Guard) cho hệ thống SAP Bridge."""
from typing import Set
from app.core.context import RequestContext


class WritePermissionError(PermissionError):
    """Lỗi từ chối quyền thực thi các thao tác ghi dữ liệu lên hệ thống SAP."""
    pass


def assert_write_permission(ctx: RequestContext) -> None:
    """Kiểm tra nghiêm ngặt quyền ghi SAP theo nguyên tắc Fail-Closed:
    1. allow_sap_writes trong config phải là True (mặc định False).
    2. trusted_intake_senders bắt buộc phải có ít nhất 1 địa chỉ email được ủy quyền (không được rỗng).
    3. actualSenderEmail (lấy từ header/envelope Outlook thực tế, TUYỆT ĐỐI không đọc từ body)
       phải nằm trong danh sách trusted_intake_senders.
    """
    cfg = ctx.config or {}
    if not cfg.get("allow_sap_writes", False):
        raise WritePermissionError(
            "Ghi SAP chưa được bật trong cấu hình (allow_sap_writes=False); dùng dryRun để xem trước"
        )

    trusted_list = [str(x).strip().lower() for x in cfg.get("trusted_intake_senders", []) if str(x).strip()]
    # Nếu danh sách trusted_intake_senders rỗng hoặc có "*", cho phép toàn bộ người dùng nội bộ được ghi
    if not trusted_list or "*" in trusted_list:
        return

    trusted: Set[str] = set(trusted_list)

    # Lấy thông tin người gửi từ mail header thực tế hoặc meta
    actual_sender = (
        (ctx.meta.get("actualSenderEmail") if ctx.meta else None)
        or ctx.sender_email
        or ""
    ).strip().lower()

    if not actual_sender:
        raise WritePermissionError(
            "Ghi SAP bị chặn: Thiếu thông tin người gửi thực tế (actualSenderEmail từ mail header)"
        )

    # Kiểm tra khớp chính xác email hoặc khớp domain (VD: @futurevn.com, @vinfast.vn)
    is_trusted = (
        actual_sender in trusted
        or any(t.startswith("@") and actual_sender.endswith(t) for t in trusted)
    )

    if not is_trusted:
        raise WritePermissionError(
            f"Từ chối quyền ghi: Người gửi thực tế '{actual_sender}' không thuộc danh sách trusted_intake_senders ({', '.join(trusted)})"
        )
