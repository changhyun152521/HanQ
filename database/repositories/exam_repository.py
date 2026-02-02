"""
Exam 저장소

- SQLite 테이블: exams
"""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from core.models import Exam
from database.sqlite_connection import SQLiteConnection, row_to_dict


class ExamRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection

    def create(self, exam: Exam) -> str:
        if exam.created_at is None:
            exam.created_at = datetime.now()
        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO exams (
                grade, semester, exam_type, school_name, year,
                created_at, parsed_at, is_parsed, problem_count
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                exam.grade or "",
                exam.semester or "",
                exam.exam_type or "",
                exam.school_name or "",
                exam.year or "",
                exam.created_at.isoformat() if exam.created_at else None,
                exam.parsed_at.isoformat() if exam.parsed_at else None,
                1 if exam.is_parsed else 0,
                exam.problem_count or 0,
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def find_by_id(self, exam_id: str) -> Optional[Exam]:
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM exams WHERE id = ?", (int(exam_id),)
            ).fetchone()
            if not row:
                return None
            return Exam.from_dict(_exam_row_to_dict(row))
        except (ValueError, TypeError):
            return None

    def find_by_metadata(
        self,
        grade: str,
        semester: str,
        exam_type: str,
        school_name: str,
        year: str,
    ) -> Optional[Exam]:
        try:
            row = self._db.get_conn().execute(
                """SELECT * FROM exams
                   WHERE grade = ? AND semester = ? AND exam_type = ? AND school_name = ? AND year = ?""",
                (grade, semester, exam_type, school_name, year),
            ).fetchone()
            if not row:
                return None
            return Exam.from_dict(_exam_row_to_dict(row))
        except Exception:
            return None

    def list_all(self) -> List[Exam]:
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM exams ORDER BY created_at DESC"
            ).fetchall()
            return [Exam.from_dict(_exam_row_to_dict(r)) for r in rows]
        except Exception:
            return []

    def update(self, exam: Exam) -> bool:
        if not exam.id:
            return False
        try:
            conn = self._db.get_conn()
            cur = conn.execute(
                """UPDATE exams SET
                    grade = ?, semester = ?, exam_type = ?, school_name = ?, year = ?,
                    created_at = ?, parsed_at = ?, is_parsed = ?, problem_count = ?
                WHERE id = ?""",
                (
                    exam.grade or "",
                    exam.semester or "",
                    exam.exam_type or "",
                    exam.school_name or "",
                    exam.year or "",
                    exam.created_at.isoformat() if exam.created_at else None,
                    exam.parsed_at.isoformat() if exam.parsed_at else None,
                    1 if exam.is_parsed else 0,
                    exam.problem_count or 0,
                    int(exam.id),
                ),
            )
            conn.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def update_parsed_status(
        self,
        exam_id: str,
        is_parsed: bool,
        problem_count: int,
    ) -> bool:
        try:
            parsed_at = datetime.now().isoformat() if is_parsed else None
            conn = self._db.get_conn()
            cur = conn.execute(
                """UPDATE exams SET is_parsed = ?, problem_count = ?, parsed_at = ?
                WHERE id = ?""",
                (1 if is_parsed else 0, problem_count, parsed_at, int(exam_id)),
            )
            conn.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def delete(self, exam_id: str) -> bool:
        try:
            cur = self._db.get_conn().execute("DELETE FROM exams WHERE id = ?", (int(exam_id),))
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False


def _exam_row_to_dict(row) -> dict:
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
