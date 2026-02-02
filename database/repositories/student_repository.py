"""
Student 저장소

- SQLite 테이블: students
"""
from __future__ import annotations

from datetime import datetime
from typing import Iterable, List, Optional

from core.models import Student
from database.sqlite_connection import SQLiteConnection, row_to_dict


class StudentRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection

    def create(self, student: Student) -> str:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        now = datetime.now()
        s = Student.from_dict(student.to_dict())
        if s.created_at is None:
            s.created_at = now
        s.updated_at = now
        s.deleted_at = None

        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO students (
                grade, status, name, school_name, parent_phone, student_phone,
                created_at, updated_at, deleted_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                s.grade or "",
                s.status or "재원",
                s.name or "",
                s.school_name or "",
                s.parent_phone or "",
                s.student_phone or "",
                s.created_at.isoformat() if s.created_at else None,
                s.updated_at.isoformat() if s.updated_at else None,
                None,
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def update(self, student: Student) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        if not student.id:
            return False
        now = datetime.now()
        try:
            cur = self._db.get_conn().execute(
                """UPDATE students SET
                    grade = ?, status = ?, name = ?, school_name = ?,
                    parent_phone = ?, student_phone = ?, updated_at = ?
                WHERE id = ?""",
                (
                    (student.grade or "").strip(),
                    (student.status or "").strip(),
                    (student.name or "").strip(),
                    (student.school_name or "").strip(),
                    (student.parent_phone or "").strip(),
                    (student.student_phone or "").strip(),
                    now.isoformat(),
                    int(student.id),
                ),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def soft_delete(self, student_id: str) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        try:
            now = datetime.now()
            cur = self._db.get_conn().execute(
                "UPDATE students SET deleted_at = ?, updated_at = ? WHERE id = ?",
                (now.isoformat(), now.isoformat(), int(student_id)),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def find_by_id(self, student_id: str) -> Optional[Student]:
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM students WHERE id = ?", (int(student_id),)
            ).fetchone()
            if not row:
                return None
            return Student.from_dict(row_to_dict(row))
        except (ValueError, TypeError):
            return None

    def list_all(self, *, include_deleted: bool = False) -> List[Student]:
        try:
            if include_deleted:
                rows = self._db.get_conn().execute(
                    "SELECT * FROM students ORDER BY created_at DESC"
                ).fetchall()
            else:
                rows = self._db.get_conn().execute(
                    "SELECT * FROM students WHERE deleted_at IS NULL ORDER BY created_at DESC"
                ).fetchall()
            return [Student.from_dict(row_to_dict(r)) for r in rows]
        except Exception:
            return []

    def bulk_upsert(self, students: Iterable[Student]) -> dict:
        inserted = 0
        updated = 0
        skipped = 0
        now = datetime.now()
        conn = self._db.get_conn()

        for s in students:
            st = Student.from_dict(s.to_dict())
            if st.created_at is None:
                st.created_at = now
            st.updated_at = now
            st.deleted_at = None

            phone = (st.student_phone or "").strip()
            if phone:
                row = conn.execute(
                    "SELECT id FROM students WHERE student_phone = ? AND deleted_at IS NULL",
                    (phone,),
                ).fetchone()
            else:
                name = (st.name or "").strip()
                if not name:
                    skipped += 1
                    continue
                row = conn.execute(
                    """SELECT id FROM students WHERE name = ? AND school_name = ? AND parent_phone = ? AND deleted_at IS NULL""",
                    (name, (st.school_name or "").strip(), (st.parent_phone or "").strip()),
                ).fetchone()

            if row:
                conn.execute(
                    """UPDATE students SET
                        grade = ?, status = ?, name = ?, school_name = ?,
                        parent_phone = ?, student_phone = ?, updated_at = ?, deleted_at = ?
                    WHERE id = ?""",
                    (
                        st.grade or "",
                        st.status or "재원",
                        st.name or "",
                        st.school_name or "",
                        st.parent_phone or "",
                        st.student_phone or "",
                        st.updated_at.isoformat(),
                        None,
                        row["id"],
                    ),
                )
                updated += 1
            else:
                conn.execute(
                    """INSERT INTO students (
                        grade, status, name, school_name, parent_phone, student_phone,
                        created_at, updated_at, deleted_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        st.grade or "",
                        st.status or "재원",
                        st.name or "",
                        st.school_name or "",
                        st.parent_phone or "",
                        st.student_phone or "",
                        st.created_at.isoformat() if st.created_at else None,
                        st.updated_at.isoformat(),
                        None,
                    ),
                )
                inserted += 1
        conn.commit()
        return {"inserted": inserted, "updated": updated, "skipped": skipped}
