"""
progress_tracker.py
────────────────────
断点续传模块：使用 SQLite 记录每条记录的处理状态。
支持：
- 查询某条记录是否已处理
- 标记成功/失败
- 清空进度（全量重跑）
- 导出处理统计
"""

from __future__ import annotations

import hashlib
import sqlite3
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class ProgressRecord:
    record_id: str          # 唯一标识（文件名+sheet+行号的哈希）
    source_file: str
    sheet_name: str
    row_index: int
    status: str             # pending | done | failed | skipped
    output_file: str = ""   # 成功时记录输出文件路径
    error_msg: str = ""
    created_at: float = 0.0
    updated_at: float = 0.0


class ProgressTracker:
    """
    基于 SQLite 的断点续传追踪器。
    每条 EventRecord 生成唯一 ID（文件名+sheet+行号的 MD5）。
    """

    _CREATE_TABLE = """
    CREATE TABLE IF NOT EXISTS progress (
        record_id   TEXT PRIMARY KEY,
        source_file TEXT NOT NULL,
        sheet_name  TEXT NOT NULL,
        row_index   INTEGER NOT NULL,
        status      TEXT NOT NULL DEFAULT 'pending',
        output_file TEXT DEFAULT '',
        error_msg   TEXT DEFAULT '',
        created_at  REAL NOT NULL,
        updated_at  REAL NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_status ON progress(status);
    CREATE INDEX IF NOT EXISTS idx_source ON progress(source_file);
    """

    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    @staticmethod
    def make_record_id(source_file: str, sheet_name: str, row_index: int) -> str:
        """生成唯一记录 ID。"""
        key = f"{source_file}::{sheet_name}::{row_index}"
        return hashlib.md5(key.encode("utf-8")).hexdigest()

    def is_done(self, source_file: str, sheet_name: str, row_index: int) -> bool:
        """判断某条记录是否已成功处理。"""
        record_id = self.make_record_id(source_file, sheet_name, row_index)
        with self._conn() as conn:
            row = conn.execute(
                "SELECT status FROM progress WHERE record_id = ?", (record_id,)
            ).fetchone()
            return row is not None and row[0] == "done"

    def mark_done(
        self,
        source_file: str,
        sheet_name: str,
        row_index: int,
        output_file: str = "",
    ):
        """标记为已完成。"""
        self._upsert(source_file, sheet_name, row_index, "done", output_file=output_file)

    def mark_failed(
        self,
        source_file: str,
        sheet_name: str,
        row_index: int,
        error_msg: str = "",
    ):
        """标记为失败。"""
        self._upsert(source_file, sheet_name, row_index, "failed", error_msg=error_msg)

    def mark_skipped(self, source_file: str, sheet_name: str, row_index: int, reason: str = ""):
        """标记为跳过（内容无效等）。"""
        self._upsert(source_file, sheet_name, row_index, "skipped", error_msg=reason)

    def get_stats(self) -> dict[str, int]:
        """获取各状态的记录数量统计。"""
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT status, COUNT(*) FROM progress GROUP BY status"
            ).fetchall()
            return {row[0]: row[1] for row in rows}

    def get_failed_records(self) -> list[ProgressRecord]:
        """获取所有失败记录，便于重试。"""
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM progress WHERE status = 'failed' ORDER BY source_file, row_index"
            ).fetchall()
            return [self._row_to_record(r) for r in rows]

    def reset(self, source_file: Optional[str] = None):
        """
        重置进度。
        - source_file=None：清空全部进度（全量重跑）
        - source_file='xxx.xlsx'：仅清空指定文件的进度
        """
        with self._conn() as conn:
            if source_file:
                conn.execute("DELETE FROM progress WHERE source_file = ?", (source_file,))
            else:
                conn.execute("DELETE FROM progress")

    def count_total(self) -> int:
        """总记录数。"""
        with self._conn() as conn:
            return conn.execute("SELECT COUNT(*) FROM progress").fetchone()[0]

    # ── 内部方法 ──────────────────────────────────────────────────────────────

    def _upsert(
        self,
        source_file: str,
        sheet_name: str,
        row_index: int,
        status: str,
        output_file: str = "",
        error_msg: str = "",
    ):
        record_id = self.make_record_id(source_file, sheet_name, row_index)
        now = time.time()
        with self._conn() as conn:
            conn.execute(
                """
                INSERT INTO progress (record_id, source_file, sheet_name, row_index,
                    status, output_file, error_msg, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(record_id) DO UPDATE SET
                    status=excluded.status,
                    output_file=excluded.output_file,
                    error_msg=excluded.error_msg,
                    updated_at=excluded.updated_at
                """,
                (record_id, source_file, sheet_name, row_index,
                 status, output_file, error_msg, now, now),
            )

    @contextmanager
    def _conn(self):
        conn = sqlite3.connect(str(self.db_path), timeout=10)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def _init_db(self):
        with self._conn() as conn:
            conn.executescript(self._CREATE_TABLE)

    @staticmethod
    def _row_to_record(row) -> ProgressRecord:
        return ProgressRecord(
            record_id=row["record_id"],
            source_file=row["source_file"],
            sheet_name=row["sheet_name"],
            row_index=row["row_index"],
            status=row["status"],
            output_file=row["output_file"] or "",
            error_msg=row["error_msg"] or "",
            created_at=row["created_at"],
            updated_at=row["updated_at"],
        )
