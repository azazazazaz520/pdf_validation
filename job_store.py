from __future__ import annotations

import sqlite3
import threading
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable


_JOB_COLUMNS = (
    "job_id",
    "filename",
    "workspace",
    "input_path",
    "output_path",
    "progress_path",
    "stage_log_path",
    "cancel_path",
    "page_count",
    "created_at",
    "started_at",
    "finished_at",
    "status",
    "progress",
    "route",
    "route_reason",
    "worker_pid",
    "error",
    "attempt",
    "max_retries",
)


class JobStore:
    """使用 SQLite 持久化任务状态，支持服务重启后恢复任务记录。"""

    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(
            self.path,
            timeout=30,
            isolation_level=None,
            check_same_thread=False,
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA busy_timeout = 30000")
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA synchronous = NORMAL")
        return connection

    @contextmanager
    def _connection(self):
        connection = self._connect()
        try:
            yield connection
        finally:
            connection.close()

    def _initialize(self) -> None:
        with self._lock, self._connection() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    job_id TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    workspace TEXT NOT NULL,
                    input_path TEXT NOT NULL,
                    output_path TEXT NOT NULL,
                    progress_path TEXT NOT NULL,
                    stage_log_path TEXT NOT NULL,
                    cancel_path TEXT NOT NULL,
                    page_count INTEGER NOT NULL,
                    created_at TEXT NOT NULL,
                    started_at TEXT,
                    finished_at TEXT,
                    status TEXT NOT NULL,
                    progress INTEGER NOT NULL,
                    route TEXT NOT NULL,
                    route_reason TEXT,
                    worker_pid INTEGER,
                    error TEXT,
                    attempt INTEGER NOT NULL DEFAULT 0,
                    max_retries INTEGER NOT NULL DEFAULT 0
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_jobs_status_created "
                "ON jobs(status, created_at)"
            )

    def save(self, record: dict[str, Any]) -> None:
        values = [record.get(column) for column in _JOB_COLUMNS]
        placeholders = ", ".join("?" for _ in _JOB_COLUMNS)
        columns = ", ".join(_JOB_COLUMNS)
        updates = ", ".join(
            f"{column}=excluded.{column}" for column in _JOB_COLUMNS if column != "job_id"
        )
        with self._lock, self._connection() as connection:
            connection.execute(
                f"INSERT INTO jobs ({columns}) VALUES ({placeholders}) "
                f"ON CONFLICT(job_id) DO UPDATE SET {updates}",
                values,
            )

    def load_all(self) -> list[dict[str, Any]]:
        with self._lock, self._connection() as connection:
            rows = connection.execute(
                "SELECT " + ", ".join(_JOB_COLUMNS) + " FROM jobs ORDER BY created_at"
            ).fetchall()
        return [dict(row) for row in rows]

    def delete(self, job_id: str) -> None:
        with self._lock, self._connection() as connection:
            connection.execute("DELETE FROM jobs WHERE job_id = ?", (job_id,))

    def delete_many(self, job_ids: Iterable[str]) -> None:
        ids = list(job_ids)
        if not ids:
            return
        placeholders = ", ".join("?" for _ in ids)
        with self._lock, self._connection() as connection:
            connection.execute(
                f"DELETE FROM jobs WHERE job_id IN ({placeholders})", ids
            )
