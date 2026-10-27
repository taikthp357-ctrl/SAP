#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Công cụ quét và gửi lại các tin nhắn đang bị kẹt trong hàng đợi Outbox (status AMBIGUOUS/PENDING/FAILED).
Sử dụng khi SAP đã tra cứu thành công nhưng Webhook gặp sự cố mạng tạm thời.
"""

import sys
import os
import json
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from app.teams_dispatcher import TeamsDispatcher

def resend_stuck_messages(db_path=None):
    if not db_path:
        db_path = BASE_DIR / "data" / "bridge_state.db"
    
    db_path = Path(db_path)
    if not db_path.exists():
        print(f"[ERROR] Không tìm thấy file DB tại: {db_path}")
        return

    print(f"[INFO] Đang mở DB: {db_path}")
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row

    with conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT id, req_id, webhook_url, payload_json, attempts, status, last_error
            FROM outbox_queue
            WHERE status IN ('AMBIGUOUS', 'PENDING', 'FAILED', 'SENDING')
            ORDER BY id ASC
        """)
        stuck_items = cur.fetchall()

        if not stuck_items:
            print("[INFO] Không có tin nhắn nào bị kẹt trong hàng đợi Outbox. Toàn bộ đã gửi thành công (ACCEPTED)!")
            return

        print(f"[INFO] Tìm thấy {len(stuck_items)} tin nhắn chưa gửi thành công qua Webhook:")
        for r in stuck_items:
            print(f"  • ID {r['id']} | Req: {r['req_id']} | Trạng thái: {r['status']} | Số lần thử: {r['attempts']} | Lỗi trước: {r['last_error']}")

        print("\n[INFO] Đang tiến hành gửi lại lên Webhook...")
        for r in stuck_items:
            req_id = r["req_id"]
            webhook_url = r["webhook_url"]
            try:
                payload = json.loads(r["payload_json"])
            except Exception as e:
                print(f"  ❌ {req_id}: Payload JSON không hợp lệ ({e})")
                continue

            print(f"  -> Đang gửi {req_id}...")
            outcome = TeamsDispatcher.deliver_payload(webhook_url, payload, max_tries=3)
            print(f"  => Kết quả: {outcome['status']} ({outcome['detail']})")

            if outcome["status"] == "ACCEPTED":
                conn.execute("""
                    UPDATE outbox_queue
                    SET status = 'ACCEPTED', accepted_at = datetime('now', 'localtime'), last_error = ?, attempts = attempts + 1
                    WHERE req_id = ?
                """, (outcome["detail"], req_id))
            else:
                conn.execute("""
                    UPDATE outbox_queue
                    SET last_error = ?, attempts = attempts + 1
                    WHERE req_id = ?
                """, (outcome["detail"], req_id))

        conn.commit()
    print("\n[HOÀN TẤT] Đã xử lý xong toàn bộ hàng đợi Outbox!")

if __name__ == "__main__":
    target_db = sys.argv[1] if len(sys.argv) > 1 else None
    resend_stuck_messages(target_db)
