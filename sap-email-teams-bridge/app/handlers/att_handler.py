# -*- coding: utf-8 -*-
"""Handler chuyên trách nghiệp vụ đính kèm file GOS, điền số phiếu HeadText và tra cứu ZFIR266."""
from datetime import datetime
from typing import Dict, Any, List
from app.core.context import RequestContext, CommandResult
from app.core.security import assert_write_permission, WritePermissionError
from app.handlers.base_handler import BaseHandler
from app.sap_attachment import MigoAttachmentRunner
from app.teams_dispatcher import TeamsDispatcher


class AttHandler(BaseHandler):
    """Xử lý nhóm lệnh: CHECK_ATT, CHECK_ATT_DOC, CHECK_DINH_KEM, ZFIR266, ATT_MIGO, ATTACH_VAT_BBBG, MIGO_ATT."""

    def handle(self, ctx: RequestContext) -> CommandResult:
        cmd = (ctx.command or "").upper()

        # -------------------------------------------------------------
        # NHÁNH 1: TRA CỨU ĐÍNH KÈM ZFIR266 / MIGO DISPLAY (READ-ONLY)
        # -------------------------------------------------------------
        if cmd in ["CHECK_ATT", "CHECK_ATT_DOC", "CHECK_DINH_KEM", "ZFIR266"]:
            att_docs = ctx.meta.get("attachmentDocs", [])
            if not att_docs and ctx.raw_materials:
                att_docs = [m for m in ctx.raw_materials if m.isdigit() and len(m) >= 8]
            if not att_docs and ctx.meta.get("attachmentItems"):
                att_docs = [item["doc"] for item in ctx.meta["attachmentItems"] if item.get("doc")]
            if not att_docs:
                raise ValueError("Vui lòng cung cấp ít nhất một số chứng từ DOC (10 chữ số) để kiểm tra đính kèm")

            company_code = ctx.meta.get("companyCode") or ctx.meta.get("bukrs") or "2301"
            target_year = (ctx.meta.get("year") or "").strip()
            if not target_year:
                target_year = str(datetime.now().year)
            ticket_no = ctx.meta.get("refNo") or ctx.meta.get("ticket_no") or ""

            records = ctx.sap.check_attachment_status(
                docs=att_docs,
                year=target_year,
                company_code=company_code,
                ticket_no=ticket_no
            )

            headers = {
                "doc": "Số DOC",
                "company_code": "Company",
                "year": "Năm",
                "bbbg_status": "BBBG",
                "vat_status": "VAT",
                "overall_status": "Kết luận",
                "files_found": "File phát hiện"
            }
            full_count = sum(1 for r in records if r.get("status_code") == "FULL_ATT" or "Đã có cả" in str(r.get("overall_status", "")))
            summary = f"Kiểm tra file đính kèm BBBG & VAT cho {len(records)} DOC: {full_count}/{len(records)} DOC đủ cả hai file"

            return CommandResult(
                status="SUCCEEDED",
                summary=summary,
                records=records,
                headers=headers,
                audit_target=", ".join(att_docs[:5]),
                audit_message=f"Kiểm tra ZFIR266 cho {len(att_docs)} DOC"
            )

        # -------------------------------------------------------------
        # NHÁNH 2: ĐÍNH KÈM FILE GOS & ĐIỀN SỐ PHIẾU HEADTEXT (MIGO A12)
        # -------------------------------------------------------------
        elif cmd in ["ATT_MIGO", "ATTACH_VAT_BBBG", "MIGO_ATT", "ATTACH_MIGO"]:
            is_preview = ctx.meta.get("dryRun", False) is True
            if not ctx.meta.get("explicitWriteIntent"):
                raise ValueError("ATT_MIGO cần lệnh đính kèm/điền HeadText rõ ràng")

            if not is_preview:
                # Chốt chặn quyền ghi an toàn & xác thực người gửi thực tế từ Mail Header
                assert_write_permission(ctx)
                expected_sys = ctx.config.get("sap_system", "")
                if expected_sys and ctx.sap.session_identity and ctx.sap.session_identity.get("system", "").upper() != expected_sys.upper():
                    raise ValueError(f"Hệ thống SAP ({ctx.sap.session_identity.get('system')}) không khớp cấu hình ({expected_sys})")
                ctx.sap.ensure_session()

            folder_path = ctx.meta.get("folderPath", "")
            items_to_att = ctx.meta.get("attachmentItems", [])
            if not items_to_att:
                att_docs = ctx.meta.get("attachmentDocs", [])
                if not att_docs and ctx.raw_materials:
                    att_docs = [m for m in ctx.raw_materials if m.isdigit() and len(m) >= 8]
                items_to_att = [{"doc": d, "ticket_no": ""} for d in att_docs]

            runner = MigoAttachmentRunner(sap_engine=ctx.sap, expected_session_identity=ctx.sap.session_identity if not is_preview else None)

            def on_att_progress(done_count, total_count, cur_doc, cur_success):
                try:
                    webhook_url = ctx.config.get("teams_webhook_url", "")
                    if webhook_url:
                        TeamsDispatcher.send_progress_update(
                            webhook_url, ctx.request_id, ctx.meta, done_count, total_count, cur_doc, cur_success
                        )
                        print(f"[PROGRESS TEAMS] {ctx.request_id} | Đã cập nhật tiến độ {done_count}/{total_count} ({int(done_count*100/total_count)}%) lên Teams!")
                except Exception as ex:
                    print(f"[PROGRESS WARN] Không thể gửi cập nhật tiến độ: {ex}")

            target_year = (ctx.meta.get("year") or "").strip()
            if not target_year:
                raise ValueError("Yêu cầu đính kèm thiếu năm chứng từ tài chính (MISSING_DOC_YEAR); vui lòng cung cấp năm chứng từ (ví dụ: 2026)")

            att_summary = runner.process_attachment_batch(
                folder_path,
                items_to_att,
                fill_headtext=True,
                dry_run=is_preview,
                year=target_year,
                progress_callback=on_att_progress
            )
            records = att_summary["results"]

            has_partial = any("PARTIAL" in str(r.get("overall_status", "")).upper() for r in records)
            all_done = all(r.get("overall_status") == "DONE" for r in records) and len(records) > 0
            any_success = att_summary.get("total_success", 0) > 0 or any(r.get("is_succ") for r in records)

            if is_preview:
                result_status = "SIMULATED"
            elif all_done and not has_partial:
                result_status = "SUCCEEDED"
            elif has_partial or any_success:
                result_status = "PARTIAL"
            else:
                result_status = "FAILED"

            if not folder_path:
                headers = {
                    "doc": "Số DOC",
                    "ticket_no": "Số phiếu",
                    "headtext_status": "Điền HeadText",
                    "overall_status": "Trạng thái"
                }
                summary = f"Điền số phiếu HeadText MIGO ({att_summary['total_success']}/{att_summary['total_items']} thành công)"
            else:
                headers = {
                    "doc": "Số DOC",
                    "ticket_no": "Số phiếu",
                    "headtext_status": "Điền HeadText",
                    "bbbg_status": "BBBG",
                    "vat_status": "VAT",
                    "overall_status": "Trạng thái"
                }
                summary = f"Đính kèm VAT, BBBG & Điền HeadText MIGO ({att_summary['total_success']}/{att_summary['total_items']} thành công)"

            return CommandResult(
                status=result_status,
                summary=summary,
                records=records,
                headers=headers,
                audit_target=f"{len(items_to_att)} chứng từ",
                audit_message=f"Đính kèm MIGO hoàn tất: {att_summary['total_success']}/{att_summary['total_items']} thành công"
            )

        raise ValueError(f"AttHandler không hỗ trợ lệnh: {cmd}")
