"""
저장된 학습 보고서(SavedReport) 저장소

- SQLite 테이블: saved_reports
"""
from __future__ import annotations

import json
from datetime import datetime
from typing import List, Optional

from core.models import SavedReport
from database.sqlite_connection import SQLiteConnection, row_to_dict, json_col


class ReportRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection

    def create(self, report: SavedReport) -> str:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        r = SavedReport.from_dict(report.to_dict())
        if r.created_at is None:
            r.created_at = datetime.now()
        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO saved_reports (
                student_id, period_start, period_end, comment, created_at, snapshot_json
            ) VALUES (?, ?, ?, ?, ?, ?)""",
            (
                (r.student_id or "").strip(),
                (r.period_start or "").strip(),
                (r.period_end or "").strip(),
                (r.comment or "").strip(),
                r.created_at.isoformat() if r.created_at else None,
                json_col(r.snapshot, "{}"),
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def list_by_student(self, student_id: str) -> List[SavedReport]:
        sid = (student_id or "").strip()
        if not sid:
            return []
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM saved_reports WHERE student_id = ? ORDER BY created_at DESC",
                (sid,),
            ).fetchall()
            out = []
            for row in rows:
                d = row_to_dict(row)
                d["snapshot"] = _parse_json(d.get("snapshot_json"), {})
                out.append(SavedReport.from_dict(d))
            return out
        except Exception:
            return []

    def get_by_id(self, report_id: str) -> Optional[SavedReport]:
        rid = (report_id or "").strip()
        if not rid:
            return None
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM saved_reports WHERE id = ?", (int(rid),)
            ).fetchone()
            if not row:
                return None
            d = row_to_dict(row)
            d["snapshot"] = _parse_json(d.get("snapshot_json"), {})
            return SavedReport.from_dict(d)
        except (ValueError, TypeError):
            return None

    def update_comment(self, report_id: str, comment: str) -> bool:
        rid = (report_id or "").strip()
        if not rid:
            return False
        try:
            cur = self._db.get_conn().execute(
                "UPDATE saved_reports SET comment = ? WHERE id = ?",
                ((comment or "").strip(), int(rid)),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def delete(self, report_id: str) -> bool:
        rid = (report_id or "").strip()
        if not rid:
            return False
        try:
            cur = self._db.get_conn().execute(
                "DELETE FROM saved_reports WHERE id = ?", (int(rid),)
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False


def _parse_json(s, default):
    if not s:
        return default
    if isinstance(s, str):
        try:
            return json.loads(s)
        except Exception:
            return default
    return default if s is None else s
