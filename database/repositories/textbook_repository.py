"""
Textbook 저장소

- SQLite 테이블: textbooks
"""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from core.models import Textbook
from database.sqlite_connection import SQLiteConnection, row_to_dict


class TextbookRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection

    def create(self, textbook: Textbook) -> str:
        if textbook.created_at is None:
            textbook.created_at = datetime.now()
        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO textbooks (
                name, subject, major_unit, sub_unit, created_at, parsed_at,
                is_parsed, problem_count
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                textbook.name or "",
                textbook.subject or "",
                textbook.major_unit or "",
                textbook.sub_unit,
                textbook.created_at.isoformat() if textbook.created_at else None,
                textbook.parsed_at.isoformat() if textbook.parsed_at else None,
                1 if textbook.is_parsed else 0,
                textbook.problem_count or 0,
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def find_by_id(self, textbook_id: str) -> Optional[Textbook]:
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM textbooks WHERE id = ?", (int(textbook_id),)
            ).fetchone()
            if not row:
                return None
            return Textbook.from_dict(_textbook_row_to_dict(row))
        except (ValueError, TypeError):
            return None

    def find_by_metadata(
        self,
        name: str,
        subject: str,
        major_unit: str,
        sub_unit: Optional[str] = None,
    ) -> Optional[Textbook]:
        try:
            if sub_unit is not None:
                row = self._db.get_conn().execute(
                    """SELECT * FROM textbooks
                       WHERE name = ? AND subject = ? AND major_unit = ? AND sub_unit = ?""",
                    (name, subject, major_unit, sub_unit),
                ).fetchone()
            else:
                row = self._db.get_conn().execute(
                    """SELECT * FROM textbooks
                       WHERE name = ? AND subject = ? AND major_unit = ? AND (sub_unit IS NULL OR sub_unit = '')""",
                    (name, subject, major_unit),
                ).fetchone()
            if not row:
                return None
            return Textbook.from_dict(_textbook_row_to_dict(row))
        except Exception:
            return None

    def list_all(self) -> List[Textbook]:
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM textbooks ORDER BY created_at DESC"
            ).fetchall()
            return [Textbook.from_dict(_textbook_row_to_dict(r)) for r in rows]
        except Exception:
            return []

    def update(self, textbook: Textbook) -> bool:
        if not textbook.id:
            return False
        try:
            conn = self._db.get_conn()
            cur = conn.execute(
                """UPDATE textbooks SET
                    name = ?, subject = ?, major_unit = ?, sub_unit = ?,
                    created_at = ?, parsed_at = ?, is_parsed = ?, problem_count = ?
                WHERE id = ?""",
                (
                    textbook.name or "",
                    textbook.subject or "",
                    textbook.major_unit or "",
                    textbook.sub_unit,
                    textbook.created_at.isoformat() if textbook.created_at else None,
                    textbook.parsed_at.isoformat() if textbook.parsed_at else None,
                    1 if textbook.is_parsed else 0,
                    textbook.problem_count or 0,
                    int(textbook.id),
                ),
            )
            conn.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def update_parsed_status(
        self,
        textbook_id: str,
        is_parsed: bool,
        problem_count: int,
    ) -> bool:
        try:
            parsed_at = datetime.now().isoformat() if is_parsed else None
            conn = self._db.get_conn()
            cur = conn.execute(
                """UPDATE textbooks SET is_parsed = ?, problem_count = ?, parsed_at = ?
                WHERE id = ?""",
                (1 if is_parsed else 0, problem_count, parsed_at, int(textbook_id)),
            )
            conn.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def delete(self, textbook_id: str) -> bool:
        try:
            cur = self._db.get_conn().execute("DELETE FROM textbooks WHERE id = ?", (int(textbook_id),))
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False


def _textbook_row_to_dict(row) -> dict:
    d = row_to_dict(row)
    d["created_at"] = _parse_dt(d.get("created_at"))
    d["parsed_at"] = _parse_dt(d.get("parsed_at"))
    d["is_parsed"] = bool(d.get("is_parsed"))
    return d


def _parse_dt(s):
    if not s:
        return None
    if isinstance(s, str):
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
        except Exception:
            return None
    return s
