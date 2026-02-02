"""
반(Class) 저장소

- SQLite 테이블: classes
"""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from core.models import SchoolClass
from database.sqlite_connection import SQLiteConnection, row_to_dict, json_col


class ClassRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection

    def create(self, school_class: SchoolClass) -> str:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        c = SchoolClass.from_dict(school_class.to_dict())
        if c.created_at is None:
            c.created_at = datetime.now()
        c.updated_at = datetime.now()
        c.deleted_at = None

        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO classes (
                grade, name, teacher, note, student_ids_json,
                created_at, updated_at, deleted_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                c.grade or "",
                c.name or "",
                c.teacher or "",
                c.note or "",
                json_col(c.student_ids, "[]"),
                c.created_at.isoformat() if c.created_at else None,
                c.updated_at.isoformat() if c.updated_at else None,
                None,
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def update(self, school_class: SchoolClass) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        if not school_class.id:
            return False
        now = datetime.now()
        try:
            cur = self._db.get_conn().execute(
                """UPDATE classes SET
                    grade = ?, name = ?, teacher = ?, note = ?, student_ids_json = ?,
                    updated_at = ?
                WHERE id = ?""",
                (
                    (school_class.grade or "").strip(),
                    (school_class.name or "").strip(),
                    (school_class.teacher or "").strip(),
                    (school_class.note or "").strip(),
                    json_col([str(x) for x in (school_class.student_ids or [])], "[]"),
                    now.isoformat(),
                    int(school_class.id),
                ),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def soft_delete(self, class_id: str) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        try:
            now = datetime.now()
            cur = self._db.get_conn().execute(
                "UPDATE classes SET deleted_at = ?, updated_at = ? WHERE id = ?",
                (now.isoformat(), now.isoformat(), int(class_id)),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def find_by_id(self, class_id: str) -> Optional[SchoolClass]:
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM classes WHERE id = ?", (int(class_id),)
            ).fetchone()
            if not row:
                return None
            d = row_to_dict(row)
            d["student_ids"] = _parse_json(d.get("student_ids_json"), [])
            return SchoolClass.from_dict(d)
        except (ValueError, TypeError):
            return None

    def list_all(self, *, include_deleted: bool = False) -> List[SchoolClass]:
        try:
            if include_deleted:
                rows = self._db.get_conn().execute(
                    "SELECT * FROM classes ORDER BY grade, name"
                ).fetchall()
            else:
                rows = self._db.get_conn().execute(
                    "SELECT * FROM classes WHERE deleted_at IS NULL ORDER BY grade, name"
                ).fetchall()
            out = []
            for row in rows:
                d = row_to_dict(row)
                d["student_ids"] = _parse_json(d.get("student_ids_json"), [])
                d.setdefault("deleted_at", None)
                out.append(SchoolClass.from_dict(d))
            return out
        except Exception:
            return []


def _parse_json(s, default):
    if not s:
        return default
    if isinstance(s, str):
        try:
            import json
            return json.loads(s)
        except Exception:
            return default
    return default if s is None else s
