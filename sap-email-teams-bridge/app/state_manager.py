# -*- coding: utf-8 -*-
"""Quản lý trạng thái bền vững (SQLite) cho SAP Email Teams Bridge:
- Bảng processed_messages: Chống xử lý trùng lặp tin nhắn bền vững qua các lần khởi động lại bot.
- Bảng outbox_queue: Hàng đợi lưu kết quả đã tra cứu SAP thành công nhưng gặp sự cố mạng khi gửi Webhook Teams.
  Cho phép retry gửi lại Webhook trực tiếp mà KHÔNG CẦN CHẠY LẠI SAP GUI!
"""

import json
import sqlite3
import uuid
from contextlib import contextmanager
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

    @contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(str(self.db_path), timeout=10.0)
        conn.row_factory = sqlite3.Row
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        # SQLite backup API includes committed WAL content; never copy a live DB file.
        if self.db_path.exists():
            with self._get_connection() as source:
                version = source.execute("PRAGMA user_version").fetchone()[0]
                if version < 2:
                    backup = self.db_path.with_name(self.db_path.name + ".pre-v1049-" + uuid.uuid4().hex + ".bak")
                    dest = sqlite3.connect(str(backup))
                    try:
                        source.backup(dest)
                    finally:
                        dest.close()
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
            conn.execute("""
                CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    req_id TEXT,
                    action_type TEXT,
                    user_id TEXT,
                    user_name TEXT,
                    target_doc TEXT,
                    status TEXT,
                    payload_json TEXT,
                    message TEXT,
                    created_at TEXT
                )
            """)
            conn.execute("CREATE TABLE IF NOT EXISTS request_claims (request_id TEXT PRIMARY KEY, status TEXT NOT NULL, created_at TEXT NOT NULL)")
            columns = {r[1] for r in conn.execute("PRAGMA table_info(outbox_queue)")}
            for name, definition in [("lease_token", "TEXT"), ("lease_until", "TEXT"), ("accepted_at", "TEXT")]:
                if name not in columns:
                    conn.execute(f"ALTER TABLE outbox_queue ADD COLUMN {name} {definition}")
            if conn.execute("PRAGMA user_version").fetchone()[0] < 2:
                # Old PENDING may already have reached Flow even with attempts=0.
                conn.execute("UPDATE outbox_queue SET status='AMBIGUOUS', last_error='Legacy delivery requires reconciliation' WHERE status='PENDING'")
            conn.execute("PRAGMA user_version=2")
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
                INSERT OR IGNORE INTO processed_messages
                (message_id, entry_id, req_id, command, status, processed_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (key, entry_id or "", req_id, command, status, now_str))
            conn.commit()

    @staticmethod
    def _insert_outbox(conn, req_id, webhook_url, payload):
        payload_str = json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False)
        existing = conn.execute("SELECT webhook_url,payload_json FROM outbox_queue WHERE req_id=?", (req_id,)).fetchone()
        if existing:
            if existing["webhook_url"] != webhook_url or json.loads(existing["payload_json"]) != payload:
                raise ValueError("Request ID conflict: immutable result already exists")
            return
        conn.execute("INSERT INTO outbox_queue(req_id,webhook_url,payload_json,status) VALUES(?,?,?,'PENDING')", (req_id,webhook_url,payload_str))

    def enqueue_outbox(self, req_id, webhook_url, payload):
        with self._get_connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            self._insert_outbox(conn, req_id, webhook_url, payload)

    def complete_request(self, message_id, entry_id, req_id, command, status, webhook_url, payload):
        """Persist result and dedup marker atomically before any HTTP attempt."""
        with self._get_connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            self._insert_outbox(conn, req_id, webhook_url, payload)
            conn.execute("INSERT OR IGNORE INTO processed_messages VALUES(?,?,?,?,?,?)", (message_id or entry_id or req_id,entry_id or "",req_id,command,status,datetime.now().isoformat()))
            conn.execute("UPDATE request_claims SET status='COMPLETED' WHERE request_id=?", (req_id,))

    def claim_request(self, req_id):
        with self._get_connection() as conn:
            result = conn.execute("INSERT OR IGNORE INTO request_claims VALUES(?,'RUNNING',?)", (req_id,datetime.now().isoformat()))
            return result.rowcount == 1

    def recover_interrupted(self):
        with self._get_connection() as conn:
            conn.execute("UPDATE request_claims SET status='RECOVERY_REQUIRED' WHERE status='RUNNING'")
            conn.execute("UPDATE outbox_queue SET status='AMBIGUOUS', last_error='Worker stopped during HTTP delivery' WHERE status='SENDING'")

    @staticmethod
    def request_id(namespace, message_id, entry_id):
        identity = message_id or entry_id
        if not identity:
            raise ValueError("Missing durable message identity")
        return "REF-" + uuid.uuid5(uuid.NAMESPACE_URL, namespace + "|" + identity).hex

    def begin_delivery(self, req_id):
        token = uuid.uuid4().hex
        with self._get_connection() as conn:
            changed = conn.execute("UPDATE outbox_queue SET status='SENDING',lease_token=?,lease_until=?, attempts=attempts+1,last_attempt=? WHERE req_id=? AND status='PENDING'", (token,(datetime.now()+timedelta(seconds=60)).isoformat(),datetime.now().isoformat(),req_id)).rowcount
        return token if changed else None

    def finish_delivery(self, req_id, token, outcome, detail=""):
        if outcome not in {"ACCEPTED", "AMBIGUOUS", "REJECTED", "SIMULATED"}:
            raise ValueError("Invalid delivery outcome")
        with self._get_connection() as conn:
            conn.execute("UPDATE outbox_queue SET status=?,last_error=?,accepted_at=?,lease_until=NULL WHERE req_id=? AND lease_token=? AND status='SENDING'", (outcome,detail,datetime.now().isoformat() if outcome=="ACCEPTED" else None,req_id,token))

    def requeue_reviewed(self, req_id, *, confirmed_not_delivered=False):
        if not confirmed_not_delivered:
            raise ValueError("Verify Flow delivery before retrying an ambiguous result")
        with self._get_connection() as conn:
            conn.execute("UPDATE outbox_queue SET status='PENDING', attempts=0,lease_token=NULL WHERE req_id=? AND status IN ('AMBIGUOUS','REJECTED','FAILED')", (req_id,))

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
                    conn.execute("UPDATE outbox_queue SET status='CORRUPT',last_error='Invalid JSON payload' WHERE id=?", (r["id"],))
                    continue
                if not isinstance(pl, dict):
                    conn.execute("UPDATE outbox_queue SET status='CORRUPT' WHERE id=?", (r["id"],))
                    continue
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
            conn.execute("UPDATE outbox_queue SET status='ACCEPTED',accepted_at=? WHERE req_id=?", (datetime.now().isoformat(),req_id))
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

    def log_audit(
        self,
        req_id: str,
        action_type: str,
        status: str,
        target_doc: str = "",
        user_id: str = "",
        user_name: str = "",
        payload: Optional[Dict[str, Any]] = None,
        message: str = ""
    ) -> int:
        """Ghi nhận nhật ký kiểm toán bất biến (Audit Ledger) cho các thao tác tác động hệ thống."""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        pl_str = json.dumps(payload, ensure_ascii=False) if payload else ""
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO audit_logs
                (req_id, action_type, user_id, user_name, target_doc, status, payload_json, message, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (req_id, action_type, user_id, user_name, target_doc, status, pl_str, message, now_str))
            conn.commit()
            return cur.lastrowid

    def query_audit_logs(
        self,
        req_id: Optional[str] = None,
        action_type: Optional[str] = None,
        target_doc: Optional[str] = None,
        limit: int = 50
    ) -> List[Dict[str, Any]]:
        """Truy vấn sổ cái kiểm toán theo điều kiện lọc."""
        query = "SELECT * FROM audit_logs WHERE 1=1"
        params: List[Any] = []
        if req_id:
            query += " AND req_id = ?"
            params.append(req_id)
        if action_type:
            query += " AND action_type = ?"
            params.append(action_type)
        if target_doc:
            query += " AND target_doc = ?"
            params.append(target_doc)
        query += " ORDER BY id DESC LIMIT ?"
        params.append(limit)

        results = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(query, tuple(params))
            for r in cur.fetchall():
                try:
                    pl = json.loads(r["payload_json"]) if r["payload_json"] else {}
                except Exception:
                    pl = {}
                results.append({
                    "id": r["id"],
                    "req_id": r["req_id"],
                    "action_type": r["action_type"],
                    "user_id": r["user_id"],
                    "user_name": r["user_name"],
                    "target_doc": r["target_doc"],
                    "status": r["status"],
                    "payload": pl,
                    "message": r["message"],
                    "created_at": r["created_at"]
                })
        return results
