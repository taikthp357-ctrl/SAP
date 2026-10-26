# -*- coding: utf-8 -*-
"""Định nghĩa hợp đồng dữ liệu chuẩn hóa (Data Contracts & DTOs) cho hệ thống Router & Handlers."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass
class RequestContext:
    """Ngữ cảnh dữ liệu đầy đủ của một yêu cầu gửi tới SAP Bridge."""
    request_id: str
    command: str
    message_id: str = ""
    sender_email: str = ""
    meta: Dict[str, Any] = field(default_factory=dict)
    sap: Any = None                                     # Đối tượng SapEngine
    resolver: Any = None                                # Đối tượng MaterialResolver
    state_mgr: Any = None                               # Đối tượng StateManager
    config: Dict[str, Any] = field(default_factory=dict)# Cấu hình config.json
    data_dir: Optional[Path] = None                     # Thư mục lưu trữ dữ liệu
    raw_materials: List[str] = field(default_factory=list)
    resolved_materials: List[str] = field(default_factory=list)
    plant: str = "3093"
    batch: str = ""
    days: int = 7
    session: Any = None                                 # SAP GuiSession (nếu có sẵn)


@dataclass
class CommandResult:
    """Kết quả trả về chuẩn hóa từ mọi nghiệp vụ Handler."""
    status: str = "SUCCEEDED"                   # 'SUCCEEDED', 'FAILED', 'PARTIAL', 'PREVIEW', 'SIMULATED'
    summary: str = ""                           # Tóm tắt hiển thị
    records: List[Dict[str, Any]] = field(default_factory=list)  # Danh sách dòng dữ liệu
    headers: Dict[str, str] = field(default_factory=dict)        # Bản đồ tiêu đề cột
    custom_html: Optional[str] = None           # HTML hiển thị tùy biến
    audit_target: str = ""                      # Đối tượng / chứng từ ghi log audit
    audit_message: str = ""                     # Nội dung thông điệp ghi log audit
    error_message: str = ""                     # Chi tiết lỗi nếu thất bại
