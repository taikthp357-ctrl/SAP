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

    trusted: Set[str] = {
        str(x).strip().lower()
        for x in cfg.get("trusted_intake_senders", [])
        if str(x).strip()
    }
    if not trusted:
        raise WritePermissionError(
            "Ghi SAP bị chặn: Danh sách trusted_intake_senders rỗng; dùng dryRun để xem trước"
        )

    # Chỉ dùng actualSenderEmail từ header Outlook thực tế hoặc sender_email từ request intake
    actual_sender = (
        (ctx.meta.get("actualSenderEmail") if ctx.meta else None)
        or ctx.sender_email
        or ""
    ).strip().lower()

    if not actual_sender:
        raise WritePermissionError(
            "Ghi SAP bị chặn: Thiếu thông tin người gửi thực tế (actualSenderEmail từ mail header)"
        )

    if actual_sender not in trusted:
        raise WritePermissionError(
            f"Từ chối quyền ghi: Người gửi thực tế '{actual_sender}' không thuộc danh sách trusted_intake_senders"
        )
