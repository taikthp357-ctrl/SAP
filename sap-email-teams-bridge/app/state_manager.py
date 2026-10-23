# -*- coding: utf-8 -*-
"""Quản lý trạng thái bền vững (SQLite) cho SAP Email Teams Bridge:
- Bảng processed_messages: Chống xử lý trùng lặp tin nhắn bền vững qua các lần khởi động lại bot.
- Bảng outbox_queue: Hàng đợi lưu kết quả đã tra cứu SAP thành công nhưng gặp sự cố mạng khi gửi Webhook Teams.
  Cho phép retry gửi lại Webhook trực tiếp mà KHÔNG CẦN CHẠY LẠI SAP GUI!
"""

import json
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple


class StateManager:
    """Quản trị hàng đợi outbox và chống trùng lặp bằng SQLite độc lập."""

    def __init__(self, db_path: Optional[Path] = None):
        if db_path is None:
            db_path = Path(__file__).resolve().parent.parent / "data" / "bridge_state.db"
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), timeout=10.0)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        with self._get_connection() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS processed_messages (
                    message_id TEXT PRIMARY KEY,
                    entry_id TEXT,
                    req_id TEXT,
                    command TEXT,
                    status TEXT,
                    processed_at TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS outbox_queue (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    req_id TEXT UNIQUE,
                    webhook_url TEXT,
                    payload_json TEXT,
                    attempts INTEGER DEFAULT 0,
                    last_attempt TEXT,
                    last_error TEXT,
                    status TEXT DEFAULT 'PENDING'
                )
            """)
            conn.commit()

    def is_message_processed(self, message_id: str, entry_id: Optional[str] = None) -> bool:
        """Kiểm tra tin nhắn đã được xử lý và phản hồi thành công trước đây hay chưa."""
        if not message_id and not entry_id:
            return False
        with self._get_connection() as conn:
            cur = conn.cursor()
            if message_id:
                cur.execute("SELECT 1 FROM processed_messages WHERE message_id = ?", (message_id,))
                if cur.fetchone():
                    return True
            if entry_id:
                cur.execute("SELECT 1 FROM processed_messages WHERE entry_id = ?", (entry_id,))
                if cur.fetchone():
                    return True
        return False

    def mark_message_processed(
        self,
        message_id: str,
        entry_id: Optional[str],
        req_id: str,
        command: str,
        status: str
    ):
        """Đánh dấu tin nhắn đã hoàn tất vào cơ sở dữ liệu bền vững."""
        key = message_id or entry_id or req_id
        if not key:
            return
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO processed_messages
                (message_id, entry_id, req_id, command, status, processed_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (key, entry_id or "", req_id, command, status, now_str))
            conn.commit()

    def enqueue_outbox(self, req_id: str, webhook_url: str, payload: Dict[str, Any]):
        """Lưu kết quả tra cứu vào outbox để gửi hoặc thử lại nếu mất mạng."""
        payload_str = json.dumps(payload, ensure_ascii=False)
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO outbox_queue
                (req_id, webhook_url, payload_json, attempts, last_attempt, status)
                VALUES (?, ?, ?, 0, ?, 'PENDING')
            """, (req_id, webhook_url, payload_str, now_str))
            conn.commit()

    def get_pending_outbox(self, max_attempts: int = 3) -> List[Dict[str, Any]]:
        """Lấy danh sách các phản hồi chưa gửi thành công và chưa vượt quá số lần thử."""
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, req_id, webhook_url, payload_json, attempts
                FROM outbox_queue
                WHERE status = 'PENDING' AND attempts < ?
                ORDER BY id ASC
            """, (max_attempts,))
            rows = cur.fetchall()
            results = []
            for r in rows:
                try:
                    pl = json.loads(r["payload_json"])
                except Exception:
                    pl = {}
                results.append({
                    "id": r["id"],
                    "req_id": r["req_id"],
                    "webhook_url": r["webhook_url"],
                    "payload": pl,
                    "attempts": r["attempts"]
                })
            return results

    def record_outbox_success(self, req_id: str):
        """Xóa hoặc đánh dấu hoàn tất mục trong outbox khi gửi Webhook thành công."""
        with self._get_connection() as conn:
            conn.execute("DELETE FROM outbox_queue WHERE req_id = ?", (req_id,))
            conn.commit()

    def record_outbox_failure(self, req_id: str, error_msg: str, max_attempts: int = 3):
        """Ghi nhận lần gửi thất bại; nếu vượt quá max_attempts thì chuyển trạng thái 'FAILED'."""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT attempts FROM outbox_queue WHERE req_id = ?", (req_id,))
            row = cur.fetchone()
            if not row:
                return
            new_attempts = row["attempts"] + 1
            new_status = "FAILED" if new_attempts >= max_attempts else "PENDING"
            conn.execute("""
                UPDATE outbox_queue
                SET attempts = ?, last_attempt = ?, last_error = ?, status = ?
                WHERE req_id = ?
            """, (new_attempts, now_str, error_msg, new_status, req_id))
            conn.commit()

    def prune_old_records(self, days: int = 30):
        """Dọn dẹp lịch sử cũ để giữ kích thước file DB luôn nhỏ gọn."""
        cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        with self._get_connection() as conn:
            conn.execute("DELETE FROM processed_messages WHERE processed_at < ?", (cutoff,))
            conn.commit()
