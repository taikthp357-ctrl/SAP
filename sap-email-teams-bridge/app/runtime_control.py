"""Local process ownership and cooperative shutdown; no broad process killing."""
import atexit
import json
import os
import time
import uuid
from pathlib import Path


class WorkerLease:
    def __init__(self, data_dir):
        self.directory = Path(data_dir).resolve()
        self.directory.mkdir(parents=True, exist_ok=True)
        self.token = uuid.uuid4().hex
        self.handle = None

    def acquire(self):
        import msvcrt
        self.handle = open(self.directory / "worker.lock", "a+b")
        self.handle.seek(0)
        if not self.handle.read(1):
            self.handle.write(b"0")
            self.handle.flush()
        self.handle.seek(0)
        try:
            msvcrt.locking(self.handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            self.handle.close()
            self.handle = None
            raise RuntimeError("Bridge đã chạy với thư mục dữ liệu này; không mở bản thứ hai.") from exc
        self._write("worker.json", {"pid": os.getpid(), "token": self.token, "started": time.time()})
        try:
            (self.directory / "bridge.pid").write_text(str(os.getpid()), encoding="utf8")
        except Exception:
            pass
        atexit.register(self.close)
        return self

    def _write(self, name, value):
        temp = self.directory / (name + ".tmp")
        temp.write_text(json.dumps(value), encoding="utf8")
        os.replace(temp, self.directory / name)

    def should_stop(self):
        try:
            return json.loads((self.directory / "stop.json").read_text(encoding="utf8")).get("token") == self.token
        except (OSError, ValueError):
            return False

    def close(self):
        try:
            pid_file = self.directory / "bridge.pid"
            if pid_file.exists():
                pid_file.unlink()
        except Exception:
            pass
        if self.handle:
            import msvcrt
            self.handle.seek(0)
            msvcrt.locking(self.handle.fileno(), msvcrt.LK_UNLCK, 1)
            self.handle.close()
            self.handle = None

    @classmethod
    def request_stop(cls, data_dir):
        lease = cls(data_dir)
        try:
            state = json.loads((lease.directory / "worker.json").read_text(encoding="utf8"))
        except (OSError, ValueError):
            return False
        lease._write("stop.json", {"token": state["token"]})
        return True


def data_directory(config, base_dir):
    path = Path(config.get("data_dir") or "data")
    return (path if path.is_absolute() else Path(base_dir) / path).resolve()
