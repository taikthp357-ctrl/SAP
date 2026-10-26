# -*- coding: utf-8 -*-
"""Định nghĩa giao diện trừu tượng BaseHandler cho toàn bộ các nghiệp vụ SAP."""
from abc import ABC, abstractmethod
from app.core.context import RequestContext, CommandResult


class BaseHandler(ABC):
    """Giao diện cơ sở cho toàn bộ các Handler xử lý nghiệp vụ SAP."""

    @abstractmethod
    def handle(self, ctx: RequestContext) -> CommandResult:
        """Thực thi nghiệp vụ cụ thể dựa trên RequestContext và trả về CommandResult chuẩn hóa."""
        raise NotImplementedError("Handler con phải hiện thực phương thức handle()")
